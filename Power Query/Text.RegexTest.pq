(textToTest as text, regexPattern as text, optional caseInsensitive as nullable logical, optional multiline as nullable logical) as nullable logical =>
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
        textToTest,
        (state as text, chars as list) as text => Text.Replace(state, chars{0}, chars{1})
    ),
    RegexFlags = List.Accumulate(
        {{caseInsensitive, "i"}, {multiline, "m"}},
        "",
        (flagSet as text, flag as list) as text =>
        if flag{0} = true then flagSet & flag{1} else flagSet
    ),
    Script = "<script type='text/javascript'>" &
        "var text = '" & textToTest & "';" &
        "var regex = /" & regexPattern & "/" & RegexFlags & ";" &
        "var result = regex.test(text);" &
        "document.write(result);" &
    "</script>",
    ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0},
    Result = try Expression.Evaluate(ExtractedText) otherwise null
in
    Result
