module YearGraphs.Arguments

open Argu

type Arguments =
    | [<First; CliPrefix(CliPrefix.None)>] Version
    | [<AltCommandLine("-e")>] Excel_Path of path: string
    | [<AltCommandLine("-r");>] Result_Directory of directory: string

    interface IArgParserTemplate with
        member this.Usage =
            match this with
            | Version -> "print version information"
            | Excel_Path _ -> "specify path to a Year Graphs file"
            | Result_Directory _ -> "specify path to a folder to place results"