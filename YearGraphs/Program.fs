open Argu
open System
open System.IO
open YearGraphs.Arguments
open YearGraphs.Common
open YearGraphs.ExcelParser
open YearGraphs.Log

[<EntryPoint>]
let main (argv: string array) =
    let parser = ArgumentParser.Create<Arguments>(errorHandler = ProcessExiter())
    let parseResults = parser.ParseCommandLine(inputs = argv)

    if parseResults.Contains Version then
        printfn $"YearGraphs version {getApplicationVersion()}"
        0
    else
        let workingDirectory =
            parseResults.GetResult(Result_Directory, defaultValue = AppDomain.CurrentDomain.BaseDirectory)

        Serilog.Log.Logger <- createLogger workingDirectory

        let version = getApplicationVersion ()
        logInformation $"Application version {version}"

        logDebug $"""Received following parameters: {argv |> String.concat " "}"""

        let excelPath =
            parseResults.GetResult Excel_Path
            |> FileInfo
            |> function
            | fileInfo when not fileInfo.Exists
                -> Error $"File {fileInfo.FullName} not exists."
            | fileInfo when not (fileInfo.Extension = ".xls" || fileInfo.Extension = ".xlsx")
                -> Error "Given file is not an Excel file. Excel files have 'xls' or 'xlsx' extension."
            | fileInfo -> Ok fileInfo

        try
            match excelPath with
            | Error err ->
                logError err
                1
            | Ok file ->
                logDebug ("Executing application with following parameters: " +
                          $"Excel path: {file.FullName}; " +
                          $"Result directory: {workingDirectory}.")

                let summaries = parseExcel file

                match summaries with
                | Ok summaries ->
                    let writeToFile (summary: Summary): unit =
                        let summaryText =
                            (summary.First, summary.Second)
                            ||> List.zip
                            |> List.map (fun (x, y) -> $"{x}\t{y}")
                            |> String.concat Environment.NewLine

                        let path = Path.Combine(workingDirectory, $"{summary.First.Head}.txt")
                        logInformation $"Writing file {path}"

                        File.WriteAllText(path = path, contents = summaryText)

                    summaries
                    |> List.iter writeToFile

                    0
                | Error message ->
                    logError message
                    1
        with ex ->
            logError ex.Message
            1