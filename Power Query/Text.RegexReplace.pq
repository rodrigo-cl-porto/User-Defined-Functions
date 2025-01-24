(str as text, regex as text, replacer as text) as nullable text =>
let 
    Script = "<script type='text/javascript'>" &
        "var str = """ & str & """;" &
        "var regex = " & regex & ";" &
        "var replacer = """ & replacer & """;" &
        "var result = str.replace(regex, replacer);" &
        "document.write(result)" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0}
in 
    try ExtractedText otherwise null
