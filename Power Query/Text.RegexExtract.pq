(textToExtract as text, regexPattern as text, optional global as nullable logical, optional caseSensitive as nullable logical, optional multiline as nullable logical) as any =>
// The Power Query's JavaScript parser doesn't support some reserved words like 'let'.
// It also doesn't support some Regex patterns like lookbehind '(?<=pattern)' and negative lookbehind '(?<!pattern)'.
// It also doens't seem to support some Regex flags like 's', 'u', 'v', 'd' and 'y'.
// Thus, the only Regex flags available are 'g', 'i' and 'm'.
let
    ReplacedUnprintableChars = List.Accumulate(
        {
            {"\", "\\"},
            {"#(lf)", "\n"},
            {"#(cr)", "\r"},
            {"#(tab)", "\t"},
            {"#(#)", "#"},
            {"'", "\'"}
        },
        textToExtract,
        (state as text, chars as list) as text => Text.Replace(state, chars{0}, chars{1})
    ),
    RegexFlags = List.Accumulate(
        {
            {global, "g"},
            {caseSensitive, "i"},
            {multiline, "m"}
        },
        "",
        (flagSet as text, flag as list) as text => if flag{0} = true then flagSet & flag{1} else flagSet
    ),
    Separator = ";;;;",
    Script = "<script type='text/javascript'>" &
        "var str = '" & ReplacedUnprintableChars & "';" &
        "var regex = /" & regexPattern & "/" & RegexFlags & ";" &
        "var matches = str.match(regex);" &
        "if (matches.length > 1){" &
            "var result = matches[0];" &
            "for(var i = 1; i < matches.length; i++){" &
                "result += '" & Separator & "' + matches[i];" &
            "}" &
        "} else {" &
            "var result = matches[0];" &
        "}" &
        "var replaced = result.replace(/\n|\r/g, '&lt;br&gt;');" &
        "replaced = replaced.replace(/\t/g, '&emsp;');" &
        "replaced = replaced.replace('<', '&lt;');" &
        "replaced = replaced.replace('>', '&gt;');" &
	    "document.write(replaced);" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0},
    ReplacedBreakRows = Text.Replace(ExtractedText, "<br>", "#(lf)"),
    SplitText = Text.Split(ReplacedBreakRows, Separator),
    ExtractedMatches = 
    try
        if SplitText{0} = "null" then
            null
        else if global = true then 
            SplitText
        else
            SplitText{0}
    otherwise error Error.Record(
        "Regex.Error",
        "Regex failed to match the expression in string",
        "Regex: " & regexPattern
    )
in 
    ExtractedMatches
