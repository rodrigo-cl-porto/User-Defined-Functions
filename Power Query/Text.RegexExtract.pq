let
    func = (textToExtract as nullable text, regexPattern as text, optional global as logical, optional caseInsensitive as logical, optional multiline as logical) as any =>
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
            (state as text, char as list) as text => Text.Replace(state, char{0}, char{1})
        ),
        RegexFlags = List.Accumulate(
            {
                {global,          "g"},
                {caseInsensitive, "i"},
                {multiline,       "m"}
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
        ExtractedMatches = try
            if SplitText{0} = "null" then
                null
            else if global=true then 
                SplitText
            else
                SplitText{0}
        otherwise 
            if global=true then {} else ""
    in 
        ExtractedMatches,

    documentedType = type function (
        textToExtract as (type nullable text meta [
            Documentation.FieldCaption = "textToExtract",
            Documentation.FieldDescription = "The input text to apply the regular expression on"
        ]),
        regexPattern as (type text meta [
            Documentation.FieldCaption = "regexPattern",
            Documentation.FieldDescription = "The regular expression pattern to match"
        ]),
        optional global as (type logical meta [
            Documentation.FieldCaption = "global",
            Documentation.FieldDescription = "If true, returns all matches; otherwise, returns only the first match"
        ]),
        optional caseInsensitive as (type logical meta [
            Documentation.FieldCaption = "caseInsensitive",
            Documentation.FieldDescription = "If true, performs case-insensitive matching"
        ]),
        optional multiline as (type logical meta [
            Documentation.FieldCaption = "multiline",
            Documentation.FieldDescription = "If true, enables multiline mode"
        ])
    ) as any meta [
        Documentation.Name = "Text.RegexExtract",
        Documentation.Description = "Extracts text from a string using a regular expression pattern.",
        Documentation.LongDescription = 
            "This function applies a regular expression to a given text and returns the matched result. "
            & "It supports optional flags for global matching, case insensitivity, and multiline mode. "
            & "Due to Power Query's JavaScript parser limitations, some advanced regex features like the 𝚕𝚘𝚘𝚔𝚋𝚎𝚑𝚒𝚗𝚍 ""(?<=)"" and certain flags (𝚜, 𝚞, 𝚟, 𝚍, 𝚢) are not supported. "
            & "Only the flags 𝚐, 𝚒, 𝚖 are available.",
        Documentation.Category = "Text",
        Documentation.Version = "1.0.0",
        Documentation.Author = "Rodrigo Celso de Lima Porto",
        Documentation.Examples = {[
            Description = "Extracts the first number from a string.",
            Code = "Text.RegexExtract(""Order #12345"", ""\d+"")",
            Result = "12345"
        ], [
            Description = "Extracts all words using global flag.",
            Code = "Text.RegexExtract(""Hello world!"", ""\w+"", true)",
            Result = "{""Hello"", ""world""}"
        ], [
            Description = "Returns empty text when no match is found.",
            Code = "Text.RegexExtract(""abc"", ""\d+"")",
            Result = """"""
        ], [
            Description = "Returns empty list when no global match is found.",
            Code = "Text.RegexExtract(""abc"", ""\d+"", true)",
            Result = "{}"
        ]}
    ],
    documentedFunction = Value.ReplaceType(func, documentedType)
in 
    documentedFunction
