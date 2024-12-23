(numberToCheck as number) as logical =>
let
    Script = "<script>" &
        "var regex = /^.?$|^(..+?)\1+$/;" &
        "var res = regex.test('" & Text.Repeat("1", numberToCheck) & "');" &
        "document.write(res)" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0}
in
    Expression.Evaluate("not " & ExtractedText)
