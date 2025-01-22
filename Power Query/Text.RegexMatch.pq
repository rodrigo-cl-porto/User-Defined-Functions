(str as text, regex as text) as nullable logical =>
let 
    Script = "<script>" &
        "var str = """ & str & """;" &
        "var regex = " & regex & ";" &
        "var result = regex.test(text);" &
        "document.write(result)" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0}
in
    try Expression.Evaluate(ExtractedText) otherwise null
