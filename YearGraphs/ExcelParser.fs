module YearGraphs.ExcelParser

open OfficeOpenXml
open System.IO
open YearGraphs.Log
open YearGraphs.Utils

let private getSummaryColumnNumbers (worksheet: ExcelWorksheet)
                                    (fileDimensions: ExcelCellAddress * ExcelCellAddress)
                                    : int list =
    let fileStart, fileEnd = fileDimensions

    let getRow (rowNumber: int): string seq =
        seq {
            for i in [ fileStart.Column .. fileEnd.Column ] do
                yield worksheet.Cells.[rowNumber, i].Text
        }

    let headerRow = getRow fileStart.Row |> Seq.toList
    logDebug $"Extracted header row: {headerRow |> stringifySeq}"

    let validHeaders = [ 0 .. 10 .. 10 * fileEnd.Column ] |> set

    let classifiedValues =
        headerRow
        |> List.map (function
                     | "" -> Error "Empty value"
                     | str when (str |> convertToInt).IsNone -> Error "Invalid character"
                     | str when validHeaders.Contains (str |> convertToInt).Value -> Ok (str |> convertToInt).Value
                     | _ -> Error "Invalid header number")

    let invalidValues =
        classifiedValues
        |> List.filter (function
                        | Error "Invalid character"
                        | Error "Invalid header number" -> true
                        | _ -> false)

    if not (List.isEmpty invalidValues) then
        logWarning $"Encountered invalid values: {stringifySeq invalidValues}"

    classifiedValues
    |> List.indexed
    |> List.filter (function | _, Ok _ -> true | _, _-> false)
    |> List.map (fun (i, _) -> i + 1)

let parseExcel (excelFile: FileInfo) =
    ExcelPackage.LicenseContext <- LicenseContext.NonCommercial

    use package = new ExcelPackage(excelFile)
    let worksheet = package.Workbook.Worksheets.[0]
    let fileStart, fileEnd = worksheet.Dimension.Start, worksheet.Dimension.End
    let summaryColumnNumbers = getSummaryColumnNumbers worksheet (fileStart, fileEnd)

    let getHeaderValueByIndex index =
        worksheet.Cells.[fileStart.Row, index].Text

    let headers = summaryColumnNumbers |> List.map getHeaderValueByIndex
    logInformation $"Extracted headers: {stringifySeq headers}"
    0
