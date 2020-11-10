module ExcelExtraFunctions.Tests.RegularExpressions

open System.Net
open Xunit
open ExcelExtraFunctions.RegularExpressions
open FsUnit.Xunit
open ExcelInterop

[<Fact>]
let ``IsMatch returns 'true' with match`` () =
    IsMatch "abc123" "[a-c]{3}[1-3]{3}" |> should equal true
    
[<Fact>]
let ``IsMatch returns 'false' with no match`` () =
    IsMatch "abc123" "[x-z]{3}[4-9]{3}" |> should equal false
 
[<Fact>]   
let ``IsMatch returns Value Error on empty`` () =
    IsMatch "abc123" "" |> should equal ExcelErrorValue

[<Theory>]
[<InlineData("abc123", "[a-c][1-3]", "c1")>]
[<InlineData("a1b2c3", "[b-c][1-3]", "b2")>]
let ``Match returns match value`` source pattern expected =
    Match source pattern |> should equal expected
    
let ``Match returns NaError on no match`` () =
    Match "abc123" "[x-z]{3}[4-9]{3}" |> should equal ExcelErrorNA

let ``Match returns ErrorValue on Blank Pattern`` () =
    Match "abc123" "" |> should equal ExcelErrorValue

[<Fact>]
let ``Matches returns array on success`` () =
    Matches "a1b2c3" "[a-z][0-9]" |> should equal [|"a1";"b2";"c3"|]

[<Fact>]
let ``Matches returns NaError on No Matches`` () =
    Matches "abc123" "[x-z]{3}[4-9]{3}" |> should equal ExcelErrorNA
    
[<Fact>]
let ``Matches returns ValueError on empty pattern`` () =
    Matches "abc123" "" |> should equal ExcelErrorValue

[<Fact>]
let ``Groups returns array on Success`` () =
    Groups "a1b2c3" "[a-z]([0-9])" false
    |> should equal (array2D [ [ "1" ]; [ "2" ]; [ "3" ] ])

[<Fact>]
let ``Groups on success return array with full match`` () =
    Groups "a1b2c3" "[a-z]([0-9])" true
    |> should equal (array2D [
        [ "a1"; "1" ];
        [ "b2"; "2" ];
        [ "c3"; "3" ];
    ])

[<Fact>]
let ``Groups on match failure return ExcelNA Errors`` () =
    Groups "a1b2c3" "[x-z]{3}([0-9]{3})" false |> should equal ExcelErrorNA
    
[<Fact>]
let ``Groups with empty pattern return ExcelErrorValue`` () =
    Groups "a1b2c3" "" false |> should equal ExcelErrorValue
    
[<Theory>]
[<InlineData("", "")>]
[<InlineData("abc", "abc")>]
[<InlineData("(abc)", @"\(abc\)")>]
let ``Escape Tests`` pattern result =
    Escape pattern |> should equal result

    
[<Theory>]
[<InlineData("a1b2c3", "[a-z]", "", "123")>]
[<InlineData("a1b2c3", @"\d", "x", "axbxcx")>]
[<InlineData("abc123", "[x-z]{3}[4-9]{3}", "", "abc123")>]
let ``Replace will replace matched patterns`` source pattern replacement expected =
    Replace source pattern replacement |> should equal expected
    
[<Fact>]
let ``Replace on empty pattern returns ExcelErrorValue`` () =
    let pattern = ""
    Replace "" pattern "" |> should equal ExcelErrorValue
    
[<Theory>]
[<InlineData("a1b2c3", "[a-c][1-3]", 3)>]
[<InlineData("abc123", "[a-c][1-3]", 1)>]
[<InlineData("a1b2c3", "[b-c][1-3]", 2)>]
[<InlineData("abc123", "[x-z]{3}[4-9]{3}", 0)>]
let ``Count Tests`` source pattern result =
    Count source pattern |> should equal result

[<Fact>]
let ``Count Returns ExcelValueError on empty pattern`` () =
    let pattern = ""
    Count "" pattern |> should equal ExcelErrorValue

[<Fact>]
let ``Split on Success returns array`` () =
    Split "a1b2c3" @"\d" |> should equal [| "a"; "b"; "c"; "" |]

[<Fact>]
let ``Split on no match returns array containing source string`` () =
    Split "a1b2c3" "z" |> should equal [| "a1b2c3" |]

[<Fact>]
let ``Split empty pattern returns ExcelErrorValue`` () =
    Split "abc123" "" |> should equal ExcelErrorValue