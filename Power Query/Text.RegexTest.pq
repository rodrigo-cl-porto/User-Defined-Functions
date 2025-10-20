let 
    func = (textToTest as text, regexPattern as text, optional caseInsensitive as logical, optional multiline as logical) as logical => let
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
        Result = try Expression.Evaluate(ExtractedText) otherwise false
    in
        Result,

    documentation = type function (
        textToTest as (type text meta [
            Documentation.FieldCaption = "textToTest",
            Documentation.FieldDescription = "The text to be tested against the regex pattern"
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
            Documentation.FieldDescription = "If true, enables multiline mode",
            Documentation.AllowedValues = {true, false},
            Documentation.DefaultValue = false
        ])) as logical meta [
            Documentation.Name = "Text.RegexTest",
            Documentation.Description = "Tests if the input text matches the given regular expression pattern.",
            Documentation.LongDescription = "Uses JavaScript's RegExp engine to test if the provided text matches the specified regex pattern. Supports optional case-insensitive and multiline flags.",
            Documentation.Category = "Text",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Version = "1.0.0",
            Documentation.Examples = {[
                Description = "Basic usage - checks if 'Hello World' contains 'World'.",
                Code = "Text.RegexTest(""Hello World"", ""World"")",
                Result = "true"
            ], [
                Description = "Multiline matching - checks if 'Line1\nLine2' starts with 'Line2' on a new line.",
                Code = "Text.RegexTest(""Line1#(lf)Line2"", ""^Line2"", false, true)",
                Result = "true"
            ], [
                Description = "Case-insensitive matching - checks if 'POWER QUERY' contains 'power' ignoring case.",
                Code = "Text.RegexTest(""POWER QUERY"", ""power"", true)",
                Result = "true"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in 
    documentedFunction
