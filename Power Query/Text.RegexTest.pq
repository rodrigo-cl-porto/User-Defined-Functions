(str as text, regex as text, optional caseSensitive as nullable logical) as any =>
let 
    i = if caseSensitive = true then "" else "i",
    Script = "<script>" &
        "var text = """ & str & """;" &
        "var regex = /" & regex & "/" & i & ";" &
        "var result = regex.test(text);" &
        "document.write(result);" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0},
    Result = try Expression.Evaluate(ExtractedText) otherwise null
in
    Result
