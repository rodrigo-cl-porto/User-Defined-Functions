(str as text, regex as text) as nullable text =>
let 
    Script = "<script type='text/javascript'>" &
        "var str = '" & str & "';" &
        "var regex = " & regex & ";" &
        "var res = str.match(regex);" &
        "document.write(res)" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0}
in try
    if ExtractedText = "null" then null else ExtractedText
otherwise error Error.Record(
    "Regex.Error",
    "Regex failed to match the expression in string",
    "Regex: " & regex
)
