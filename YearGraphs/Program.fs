open System
open Argu
open System.IO
open YearGraphs.Arguments
open YearGraphs.Common
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

        match excelPath with
        | Error err ->
            logError err
            1
        | Ok file ->
            let nl = Environment.NewLine
            logDebug ($"Executing application with following parameters:{nl}" +
                      $"Excel path: {file.FullName}{nl}" +
                      $"Result directory: {workingDirectory}")
            0