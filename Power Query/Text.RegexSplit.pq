let 
    func = (textToSplit as text, regexPattern as text, optional caseInsensitive as logical, optional multiline as logical) as any => let
        ReplacedUnprintableChars = List.Accumulate(
            {
                {"\", "\\"},
                {"#(lf)", "\n"},
                {"#(cr)", "\r"},
                {"#(tab)", "\t"},
                {"#(#)", "#"},
                {"'", "\'"}
            },
            textToSplit,
            (state as text, chars as list) as text => Text.Replace(state, chars{0}, chars{1})
        ),
        RegexFlags = List.Accumulate(
            {{caseInsensitive, "i"}, {multiline, "m"}},
            "g",
            (flagSet as text, flag as list) as text => if flag{0} = true then flagSet & flag{1} else flagSet
        ),
        Separator = ";;;;",
        Script = "<script type='text/javascript'>" &
            "var str = '" & ReplacedUnprintableChars & "';" &
            "var regex = /(^|" & regexPattern & ")([\s\S](?!" & regexPattern & "))*[\s\S](?=" & regexPattern & "|$)/" & RegexFlags & ";" &
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
        ExtractedMatches = try
            if SplitText{0} = "null" then null else SplitText
            otherwise error Error.Record(
                "Regex.Error",
                "Regex failed to match the expression in string",
                "Regex: " & regexPattern
            )
        in ExtractedMatches,

    documentation = type function (
        textToSplit as (type text meta [
            Documentation.FieldCaption = "textToSplit",
            Documentation.FieldDescription = "The input text to apply the regular expression on"
        ]),
        regexPattern as (type text meta [
            Documentation.FieldCaption = "regexPattern",
            Documentation.FieldDescription = "The regular expression pattern to match"
        ]),
        optional caseInsensitive as (type logical meta [
            Documentation.FieldCaption = "caseInsensitive",
            Documentation.FieldDescription = "If true, performs case-insensitive matching",
            Documentation.AllowedValues = {true, false},
            Documentation.DefaultValue = false
        ]),
        optional multiline as (type logical meta [
            Documentation.FieldCaption = "multiline",
            Documentation.FieldDescription = "If true, enables multiline mode for the regular expression",
            Documentation.AllowedValues = {true, false},
            Documentation.DefaultValue = false
        ])) as any meta [
            Documentation.Name = "Text.RegexSplit",
            Documentation.Description = "Splits text into a list using a regular expression pattern.",
            Documentation.LongDescription = "This function allows you to split a given text into a list of substrings based on a specified regular expression pattern. It supports optional flags for case insensitivity and multiline mode.",
            Documentation.Category = "Text",
            Documentation.Version = "1.0.0",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Examples = {
                [
                    Description = "Basic usage of Text.RegexSplit",
                    Code = "Text.RegexSplit(""apple, banana; cherry|date"", ""[,\s;|]+"")",
                    Result = "{""apple"", ""banana"", ""cherry"", ""date""}"
                ],
                [
                    Description = "Using case insensitive flag",
                    Code = "Text.RegexSplit(""Apple, banana; Cherry|date"", ""[a-z]+"", true)",
                    Result = "{"""", "", "", ""; "", ""|"", """"}"
                ]
            }
        ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in 
    documentedFunction