module ExcelInterop

open ExcelDna.Integration
open System

type Optional = obj
type ExcelReturn = obj

[<Literal>]
let ExcelErrorValue = ExcelError.ExcelErrorValue
[<Literal>]
let ExcelErrorNA = ExcelError.ExcelErrorNA

let defaultWith<'T> parse defaultValue (value:Optional) =
    match value with
    | :? ExcelMissing -> defaultValue
    | :? 'T as value -> value
    | _ -> 
        match parse (value.ToString()) with 
        | (true, value) -> value
        | (false, _) -> defaultValue

let defaultInt defaultValue (value:Optional) = defaultWith Int32.TryParse defaultValue value
let defaultBool defaultValue (value:Optional) = defaultWith Boolean.TryParse defaultValue value
let defaultDouble defaultValue (value:Optional) = defaultWith Double.TryParse defaultValue value
let defaultString defaultValue (value:Optional) = defaultWith (fun s -> (true, s)) defaultValue value