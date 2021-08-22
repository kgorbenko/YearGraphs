module YearGraphs.Common

open System.Diagnostics
open System.Reflection

let getApplicationVersion() = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location)
                                             .FileVersion