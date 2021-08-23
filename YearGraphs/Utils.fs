module YearGraphs.Utils

open System

let convertToInt (str: string): int option =
    match Int32.TryParse str with
    | true, num -> Some num
    | _ -> None

let stringifySeq (seq: 'a seq): string =
    sprintf "%A" seq