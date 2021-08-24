module YearGraphs.ExcelParser

open OfficeOpenXml
open System.Collections.Generic
open System.IO
open YearGraphs.Log
open YearGraphs.Utils

let private getSummaryColumnNumbers (worksheet: ExcelWorksheet)
                                    (fileDimensions: ExcelCellAddress * ExcelCellAddress)
                                    : Result<int list, string> =
    let fileStart, fileEnd = fileDimensions

    let getRow (rowNumber: int): string seq =
        seq {
            for i in [ fileStart.Column .. fileEnd.Column ] do
                yield worksheet.Cells.[rowNumber, i].Text
        }

    let headerRow = getRow fileStart.Row |> Seq.toList
    logDebug $"Extracted header row: {headerRow |> stringifySeq}"

    let validHeaders = Queue([ 0 .. 10 .. 10 * fileEnd.Column ])

    let classifiedValues =
        headerRow
        |> List.map (function
                     | "" -> Ok -1
                     | str when (str |> convertToInt).IsNone -> Error "Invalid character"
                     | str when validHeaders.Contains (str |> convertToInt).Value -> Ok (str |> convertToInt).Value
                     | _ -> Error "Invalid header number")

    let invalidValues =
        classifiedValues
        |> List.filter (function | Error _ -> true | _ -> false)

    if not (List.isEmpty invalidValues) then
        Error $"Encountered invalid values: {stringifySeq invalidValues}"
    else
        classifiedValues
        |> List.indexed
        |> List.map (fun (i, x) -> i + 1, x)
        |> List.choose (function | i, Ok num when num <> -1 -> Some (i, num) | _ -> None)
        |> List.choose (fun (i, x) ->
            if validHeaders.Peek() = x then
                validHeaders.Dequeue() |> ignore
                Some i
            else
                None
        )
        |> Ok

let parseExcel (excelFile: FileInfo) =
    ExcelPackage.LicenseContext <- LicenseContext.NonCommercial

    use package = new ExcelPackage(excelFile)
    let worksheet = package.Workbook.Worksheets.[0]
    let fileStart, fileEnd = worksheet.Dimension.Start, worksheet.Dimension.End
    let summaryColumnNumbers = getSummaryColumnNumbers worksheet (fileStart, fileEnd)

    match summaryColumnNumbers with
    | Error message ->
        logError message
        1
    | Ok summaryColumnNumbers ->
        let getHeaderValueByIndex index =
            worksheet.Cells.[fileStart.Row, index].Text

        let headers = summaryColumnNumbers |> List.map getHeaderValueByIndex
        logInformation $"Extracted headers: {stringifySeq headers}"
        0
