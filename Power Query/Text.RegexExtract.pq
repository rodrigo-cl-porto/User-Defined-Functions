(textToExtract as text, regex as text) as any =>
// The Power Query's JavaScript parser doesn't support some reserved words like 'let'.
// It also doesn't support some Regex patterns like lookbehind '(?<=pattern)' and negative lookbehind '(?<!pattern)'.
// It also doens't seem to support some Regex flags like 's', 'u', 'v', 'd' and 'y'.
// Thus, the only Regex flags available are 'g', 'm' and 'i'.
let 
    NonPrintableChars = {
        {"\", "\\"},
        {"#(lf)", "\n"},
        {"#(cr)", "\r"},
        {"#(tab)", "\t"},
        {"#(#)", "#"},
        {"'", "\'"}
    },
    ReplacedChars = List.Accumulate(
        NonPrintableChars,
        textToExtract,
        (state as text, chars as list) as text => Text.Replace(state, chars{0}, chars{1})
    ),
    Separator = ";;;;",
    Script = "<script type='text/javascript'>" &
        "var str = '" & ReplacedChars & "';" &
        "var regex = " & regex & ";" &
        "var matches = str.match(regex);" &
        "if (matches.length > 1){" &
            "var result = matches[0];" &
            "for(var i = 1; i < matches.length; i++){" &
                "result += '" & Separator & "' + matches[i];" &
            "}" &
        "} else {" &
            "var result = matches[0];" &
        "}" &
	"document.write(result);" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0},
    SplittedText = Text.Split(ExtractedText, Separator)
in 
    try 
        if SplittedText{0} = "null" then
            null
        else if List.Count(SplittedText) = 1 then
            SplittedText{0}
        else 
            SplittedText
    otherwise
        error Error.Record(
            "Regex.Error",
            "Regex failed to match the expression in string",
            "Regex: " & regex
        )
