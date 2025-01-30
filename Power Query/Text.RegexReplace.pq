(textToReplace as text, regexPattern as text, replacer as text, optional global as nullable logical, optional caseInsensitive as nullable logical, optional multiline as nullable logical) as nullable text =>
// The Power Query's JavaScript parser doesn't support some reserved words like 'let'.
// It also doesn't support some Regex patterns like lookbehind '(?<=pattern)' and negative lookbehind '(?<!pattern)'.
// It also doens't seem to support some Regex flags like 's', 'u', 'v', 'd' and 'y'.
// Thus, the only Regex flags available are 'g', 'm' and 'i'.
let 
    ReplacedUnprintableChars = List.Accumulate(
        {
            {"\", "\\"},
            {"#(cr)", "\r"},
            {"#(lf)", "\n"},
            {"#(tab)", "\t"},
            {"#(#)", "#"},
            {"'", "\'"}
        },
        textToReplace,
        (state as text, chars as list) as text => Text.Replace(state, chars{0}, chars{1})
    ),
    RegexFlags = List.Accumulate(
        {{global, "g"}, {caseInsensitive, "i"}, {multiline, "m"}},
        "",
        (flags as text, flag as list) as text => if flag{0} = true then flags & flag{1} else flags
    ),
    Script = "<script type='text/javascript'>" &
        "var str = '" & ReplacedUnprintableChars & "';" &
        "var regex = /" & regexPattern & "/" & RegexFlags & ";" &
        "var replacer = '" & replacer & "';" &
        "var result = str.replace(regex, replacer);" &
        "var replaced = result.replace(/\r\n|\n\r|\r|\n/g, '&lt;br&gt;');" &
        "replaced = replaced.replace(/\t/g, '&emsp;');" &
        "replaced = replaced.replace('<', '&lt;');" &
        "replaced = replaced.replace('>', '&gt;');" &
        "document.write(replaced);" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0},
    ReplacedBreakRows = Text.Replace(ExtractedText, "<br>", "#(lf)"),
    ReplacedText =
    try 
        if ReplacedBreakRows = "null" then 
            null
        else
            ReplacedBreakRows
    otherwise
            error Error.Record(
            "Regex.Error",
            "Regex failed to match the expression in string",
            "Regex: " & regexPattern
        )
in
    ReplacedText
