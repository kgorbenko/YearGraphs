module YearGraphs.Log

open Serilog
open Serilog.Events
open Serilog.Formatting.Json
open System.IO

let makeLogFilePath (workingDirectory: string)
                    (fileName: string)
                    : string =
    Path.Combine([| workingDirectory; "logs"; fileName |])

let createLogger (workingDirectory: string) =
    LoggerConfiguration()
        .MinimumLevel.Debug()
        .WriteTo.Console(outputTemplate = "{Timestamp} {Level:u3} {Message}{NewLine}{Exception}",
                         restrictedToMinimumLevel = LogEventLevel.Information)
        .WriteTo.File(JsonFormatter(),
                      makeLogFilePath workingDirectory "log.json",
                      rollingInterval = RollingInterval.Minute,
                      restrictedToMinimumLevel = LogEventLevel.Debug)
        .CreateLogger();

let logDebug message = Log.Debug message

let logInformation message = Log.Information message

let logWarning message = Log.Warning message

let logError message = Log.Error message