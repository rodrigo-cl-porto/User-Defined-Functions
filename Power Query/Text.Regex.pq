(str as text, regex as text) as text =>
let 
    Script = "<script>" &
        "var str = """ & str & """;" &
        "var regex = /" & regex & "/;" &
        "var res = str.match(regex);" &
        "document.write(res)" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0}
in 
    try ExtractedText otherwise null
