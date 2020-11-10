module ExcelExtraFunctions.Strings

open ExcelInterop
open ExcelDna.Integration
open System

module String =
    let comparisonType = StringComparison.CurrentCultureIgnoreCase 

    let (|Positive|Negative|Zero|) =
        function
            | i when i > 0 -> Positive
            | i when i < 0 -> Negative
            | _ -> Zero
    
    let indexOfAtN occurrence value source =
        match source, value, occurrence with
        | ("", _, _) -> None
        | (_, "", _) -> None
        | (_, _, Zero) -> None
        | (_, _, Positive) ->
            let rec indexOfAtN occurrence index =
                match occurrence, index with
                | 0, _ -> index
                | _, None -> None
                | _, Some index ->
                    source.IndexOf(value, index + 1, comparisonType)
                    |> fun index -> if index > -1
                                    then Some (index + String.length value)
                                    else None
                    |> indexOfAtN (occurrence - 1)
                    
            indexOfAtN (abs occurrence) (Some -1)
        | (_, _, Negative) ->
            let rec lastIndexOfAtN occurrence index =
                match occurrence, index with
                | _, None -> None
                | 0, _ -> index
                | _, Some index ->
                    source.LastIndexOf(value, index - 1, comparisonType)
                    |> fun index -> if index > -1
                                    then Some index
                                    else None
                    |> lastIndexOfAtN (occurrence - 1)
                    
            lastIndexOfAtN (abs occurrence) (Some (String.length source))
            |> Option.map (fun index -> index + 1)

[<ExcelFunction(Category = "EXF Strings", Name = "STR.BETWEEN",
    Description = "Returns the string between two other strings")>]
let Between source leftDelimiter rightDelimiter (leftCount:Optional) (rightCount:Optional) =
    match source |> String.length with 
        | 0 -> Error ExcelErrorValue
        | _ -> (Ok source)
    |> Result.bind 
        (fun source -> 
            match String.indexOfAtN (leftCount |> defaultInt 1) leftDelimiter source  with 
                | None -> Error ExcelErrorValue
                | Some index -> Ok (source.Substring(index))
        )
    |> Result.bind
        (fun source -> 
            match String.indexOfAtN (rightCount |> defaultInt 1) rightDelimiter source with
                | None -> Error ExcelErrorValue
                | Some index -> Ok (source.Remove(index - 1))
        )
    |> function
        | Error excelError -> excelError :> ExcelReturn
        | Ok result -> result :> ExcelReturn