(str as text, regex as text, replacer as text) as nullable text =>
let 
    Script = "<script>" &
        "var str = """ & str & """;" &
        "var regex = " & regex & ";" &
        "var replacer = """ & replacer & """;" &
        "var res = str.replace(regex, replacer);" &
        "document.write(res)" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0}
in 
    try ExtractedText otherwise null
