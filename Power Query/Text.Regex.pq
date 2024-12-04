(regex as text, str as text) as text =>
let 
    Script = "<script>" &
        "var regex = " & regex & ";" &
        "var str = """ & str & """;" &
        "var res = str.match(regex);" &
        "document.write(res)" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0}
in 
    ExtractedText
