(numberToCheck as number) as logical =>
let
    StringOf1s = Text.Repeat("1", numberToCheck),
    Script = "<script>" &
        "var regex = /^.?$|^(..+?)\1+$/;" &
        "var res = regex.test('" & StringOf1s & "');" &
        "document.write(res)" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0}
in
    Expression.Evaluate("not " & ExtractedText)
