(numberToCheck as number) as logical =>
let
    Script = "<script>" &
        "var str = """ & Text.Repeat("1", numberToCheck) & """;" &
        "var regex = /^.?$|^(..+?)\1+$/g;" &
        "var res = str.match(regex);" &
        "document.write(res)" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0}
in
    ExtractedText = "null"
