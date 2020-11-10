module ExcelExtraFunctions.Tests.Strings

open System.Net
open ExcelDna.Integration
open Xunit
open ExcelExtraFunctions.Strings
open FsUnit.Xunit
open ExcelInterop

[<Theory>]
[<InlineData("1abc2abc3abc4abc5abc", "1", "2", 1, 1, "abc")>]
[<InlineData("1abc2abc3abc4abc5abc", "a", "a", 1, 2, "bc2abc3")>]
[<InlineData("1abc2abc3abc4abc5abc", "a", "a", 2, 2, "bc3abc4")>]
[<InlineData("1abc2abc3abc4abc5abc", "a", "a", -2, 1, "bc5")>]
[<InlineData("1abc2abc3abc4abc5abc", "a", "a", 1, -1, "bc2abc3abc4abc5")>]
[<InlineData("1abc2abc3abc4abc5abc", "a", "c", 1, -3, "bc2abc3ab")>]
[<InlineData("1abc2abc3abc4abc5abc", "ab", "ab", 2, -1, "c3abc4abc5")>]
let ``Between Tests``
    source left right leftCount rightCount result =
        Between source left right leftCount rightCount
        |> should equal result

[<Theory>]
[<InlineData("abc", "d", "d", 500, 500, ExcelErrorValue)>]
[<InlineData("abc", "d", "d", 1, 1, ExcelErrorValue)>]
[<InlineData("abc", "a", "d", 1, 1, ExcelErrorValue)>]
[<InlineData("abc", "d", "c", 1, 1, ExcelErrorValue)>]
let ``Between Failures return ExcelErrorValue``
    source left right leftCount rightCount (result:ExcelError) =
        Between source left right leftCount rightCount
        |> should equal result
