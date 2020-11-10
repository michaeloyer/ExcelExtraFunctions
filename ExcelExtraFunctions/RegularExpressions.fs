module ExcelExtraFunctions.RegularExpressions

open ExcelDna.Integration;
open ExcelInterop
open System.Text.RegularExpressions;

let (|ValidPattern|_|) = function
    | "" -> None
    | _ -> Some ()

[<ExcelFunction(Category = "EXF Regular Expression", Name = "RE.ISMATCH",
    Description = "Returns TRUE if a single pattern match is found in the input, otherwise FALSE")>]
let IsMatch input pattern = 
    match pattern with 
    | ValidPattern -> Regex.IsMatch(input, pattern) :> ExcelReturn
    | _ -> ExcelErrorValue :> ExcelReturn

[<ExcelFunction(Category = "EXF Regular Expression", Name = "RE.MATCH",
    Description = "Returns the first matched pattern in the input string.")>]
let Match input pattern = 
    match pattern with
    | ValidPattern -> 
        match Regex.Match(input, pattern) with
        | regex when regex.Success -> regex.Value :> ExcelReturn
        | _ -> ExcelErrorNA :> ExcelReturn
    | _ -> ExcelErrorValue :> ExcelReturn

[<ExcelFunction(Category = "EXF Regular Expression", Name = "RE.MATCHES",
    Description = "Returns array of matched patterns in the input string.")>]
let Matches input pattern = 
    match pattern with
    | ValidPattern -> 
        match Regex.Matches(input, pattern) with
        | matches when matches.Count = 0 -> ExcelErrorNA :> ExcelReturn
        | matches -> matches |> Seq.cast<Match> |> Seq.map (fun m -> m.Value) |> Seq.toArray :> ExcelReturn
    | _ -> ExcelErrorValue :> ExcelReturn

[<ExcelFunction(Category = "EXF Regular Expression", Name = "RE.GROUPS",
    Description = "Returns 2D array of the groups of each match. Can return full match")>]
let Groups input pattern 
            ([<ExcelArgument("Includes the full match along with the group of matches at the front of the array")>] 
             includeFullMatch) =
    match pattern with
    | ValidPattern -> 
        match Regex.Matches(input, pattern) with
        | matches when matches.Count = 0 || matches.[0].Groups.Count = 0 ->
            ExcelErrorNA :> ExcelReturn
        | matches -> 
            let includeFullMatchInt = if includeFullMatch then 0 else 1

            Array2D.init 
                matches.Count 
                (matches.[0].Groups.Count - includeFullMatchInt)
                (fun x y -> matches.[x].Groups.[y + includeFullMatchInt].Value) :> ExcelReturn
    | _ -> ExcelErrorValue :> ExcelReturn

[<ExcelFunction(Category = "EXF Regular Expression", Name = "RE.REPLACE",
    Description = "Replaces all pattern matches in the input with the replacement string.")>]
let Replace input pattern 
            ([<ExcelArgument("If a capture group is used it can be reference with $1, or if is explicitly referenced you can use $Name")>]
            replacement : string) =
    match pattern with
    | ValidPattern -> Regex.Replace(input, pattern, replacement) :> ExcelReturn
    | _ -> ExcelErrorValue :> ExcelReturn

[<ExcelFunction(Category = "EXF Regular Expression", Name = "RE.SPLIT",
    Description = "Splits the input string at each matched pattern and returns the array.")>]
let Split input pattern =
    match pattern with
    | ValidPattern -> Regex.Split(input, pattern) :> ExcelReturn
    | _ -> ExcelErrorValue :> ExcelReturn

[<ExcelFunction(Category = "EXF Regular Expression", Name = "RE.COUNT",
    Description = "Counts the number of pattern matches in the input.")>]
let Count input pattern =
    match pattern with
    | ValidPattern -> Regex.Matches(input, pattern).Count :> ExcelReturn
    | _ -> ExcelErrorValue :> ExcelReturn

[<ExcelFunction(Category = "EXF Regular Expression", Name = "RE.ESCAPE",
    Description = "Puts a '\' (backslash) character in front of all regex modifier characters")>]
let Escape pattern = Regex.Escape(pattern)
