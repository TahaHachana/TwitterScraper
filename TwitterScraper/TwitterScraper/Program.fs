open System
open System.Collections.Generic
open System.Net
open System.Text.RegularExpressions
open Microsoft.FSharp.Reflection
open Microsoft.Office.Interop.Excel
open Spidy.Types
open Spidy.Crawler
open SpreadSharp

type Data =
    {
        FullName   : string
        ScreenName : string
        Location   : string
        Bio        : string
    }

type MatchData = URI | Content

let makeDataRecord values = FSharpValue.MakeRecord(typeof<Data>, values) :?> Data

/// Comiles a pattern into a regex object.
let compileRegex pattern = Regex(pattern, RegexOptions.Compiled)

/// Decodes HTML encoded characters.
let decodeHtml html = WebUtility.HtmlDecode html

let concatLines (str : string) =
    str.Split([|'\n'|], StringSplitOptions.RemoveEmptyEntries)
    |> String.concat " "

let htmlTagPattern = "<[^>]*>"
let htmlTagRegex = compileRegex htmlTagPattern

/// Removes comments, inline JS/CSS and HTML tags.
let cleanHtml html =
    htmlTagRegex.Replace(html, "")
    |> decodeHtml
    |> concatLines

/// Generic data scraping function.
let scrapeData data (regex : Regex) = regex.Matches data |> Seq.cast<Match> |> List.ofSeq

/// Returns the value of the specified regex group and removes HTML tags from it.
let groupValue item (idx : int) (matchList : Match list) =
    let matchOption = try List.nth matchList item |> Some with _ -> None
    match matchOption with
        | None   -> ""
        | Some x -> x.Groups.[idx].Value.Trim() |> cleanHtml

/// Returns the fields' names of a record type.
let recordFieldsNames recordType =
    FSharpType.GetRecordFields recordType
    |> Array.map (fun x -> x.Name)

/// Sets the headers of a worksheet.
let setWorksheetHeaders worksheet headers =
    let length = Array.length headers
    let range = Range.range worksheet "A1" <| Some (string (char (length + 64)) + "1")
    range.Value2 <- headers

/// Sets the value of a worksheet cell.
let setCellValue worksheet (col : string) row value =
    let range = Range.range worksheet (col + row) None
    range.Value2 <- value

let screenNameHashset = HashSet<string>()

let wasntScraped screenName = screenNameHashset.Add screenName

let startExcelAgent (recordType : Type) worksheet = MailboxProcessor<obj>.Start(fun inbox ->
    let length = FSharpType.GetRecordFields recordType |> Array.length
    let chars = [|1 .. length|] |> Array.map (fun x -> string (char <| 64 + x))
    let setCellValue' = setCellValue worksheet
    let rec loop count =
        async {
            let! record = inbox.Receive()
            let fields = FSharpValue.GetRecordFields record
            let screenName = fields.[1].ToString()
            match wasntScraped screenName with
                | false -> return! loop count
                | true  ->
                    let row = string count
                    Array.zip chars fields
                    |> Array.iter (fun (cell, field) ->
                        let field' = field.ToString()
                        setCellValue' cell row field')
                    return! loop (count + 1)
        }
    loop 2)

let printAgent = MailboxProcessor.Start(fun inbox ->
    let rec loop count =
        async {
            let! msg = inbox.Receive()
            printfn "%d: %s" count msg
            return! loop (count + 1) }
    loop 1)

let patterns =
    [
        "<h2>Follow[^>]+</h2>" // page
        "(?s)<h1\ class=\"fullname\">(.+?)</h1>" // full name
        "(?s)<span\ class=\"screen-name\">(.+?)</span>" // screen name
        "(?s)<span\ class=\"location\">(.+?)</span>" // location
        "(?s)<p\ class=\"bio\ \">(.+?)</p>" // bio
    ]

let regexObjs = List.map compileRegex patterns

// Excel interop
let excel = Excel.start true
let workbook = Workbook.addWorkbook excel
let worksheet = Worksheet.worksheetAtIndex workbook 1

let excelAgent = startExcelAgent typeof<Data> worksheet

let headers = recordFieldsNames typeof<Data>
setWorksheetHeaders worksheet headers

let range' x y  = Range.range worksheet x y

["A1", 30; "B1", 30; "C1", 30; "D1", 50]
|> List.iter (fun (x, y) ->
    range' x None
    |> (fun rng -> rng.EntireColumn.ColumnWidth <- y)
    |> ignore)

range' "A1" None |> (fun x ->
    let entireRow = x.EntireRow
    entireRow.Font.Bold <- true
    entireRow.HorizontalAlignment <- XlHAlign.xlHAlignCenter)

let scrapeData' html =
    regexObjs
    |> List.tail
    |> List.map (fun x ->
        scrapeData html x
        |> groupValue 0 1)
    |> List.toArray
    |> Array.map box

let isDataUrl (regex : Regex) str = regex.IsMatch str

let f regex matchData (httpData : HttpData) =
    async {
        let status = httpData.StatusCode
        match status with
            | HttpStatusCode.OK ->
                let contentOption = httpData.Content
                match contentOption with
                    | None         -> ()
                    | Some content ->
                        let requestUri = httpData.RequestUri
                        printAgent.Post requestUri
                        let matchData' =
                            match matchData with
                                | URI     -> requestUri
                                | Content -> content
                        let boolVal = isDataUrl regex matchData'
                        match boolVal with
                            | false -> ()
                            | true  ->
                                let data = scrapeData' content 
                                let record = makeDataRecord data
                                excelAgent.Post record
            | _ -> ()
    }

let seed = Uri "http://twitter.com/TahaHachana"
let httpDataFunc = f regexObjs.Head Content
let completionFunc = async { printAgent.Post "Done." }

let config =
    {
        Seeds          = [seed]
        Depth          = None
        Limit          = None
        AllowedHosts   = Some [seed.Host]
        RogueMode      = RogueMode.ON
        HttpDataFunc   = httpDataFunc
        CompletionFunc = completionFunc
    }

let main = crawl config
let canceler = main |> Async.RunSynchronously

while true do ()