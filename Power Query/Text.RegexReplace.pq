let
    func = (textToModify as text, regexPattern as text, replacer as text, optional global as logical, optional caseInsensitive as logical, optional multiline as logical) as nullable text => let 
        ReplacedUnprintableChars = List.Accumulate(
            {
                {"\", "\\"},
                {"#(cr)", "\r"},
                {"#(lf)", "\n"},
                {"#(tab)", "\t"},
                {"#(#)", "#"},
                {"'", "\'"}
            },
            textToModify,
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
        ReplacedText,

    documentation = type function (
        textToModify as (type text meta [
            Documentation.FieldCaption = "textToModify",
            Documentation.FieldDescription = "The input text to apply the regular expression on"
        ]),
        regexPattern as (type text meta [
            Documentation.FieldCaption = "regexPattern",
            Documentation.FieldDescription = "The regular expression pattern to match"
        ]),
        optional global as (type logical meta [
            Documentation.FieldCaption = "global",
            Documentation.FieldDescription = "If true, replaces all matches; otherwise, replaces only the first match",
            Documentation.AllowedValues = {true, false},
            Documentation.DefaultValue = false
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
        ])) as any meta [
            Documentation.Name = "Text.RegexReplace",
            Documentation.Description = "Replaces text in a string using a regular expression pattern.",
            Documentation.LongDescription = "This function allows you to replace occurrences of a pattern in a given text using regular expressions. You can specify whether to replace all occurrences or just the first one, and you can also enable case-insensitive and multiline matching.",
            Documentation.Category = "Text",
            Documentation.Version = "1.0.0",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Examples = {[
                Description = "Replace all digits with '#' in the given text.",
                Code = "Text.RegexReplace(""Order 1234"", ""\d"", ""#"", true)",
                Result = """" & "Order ####" & """"
            ], [
                    Description = "Replace the first occurrence of 'cat' with 'dog', case-insensitively.",
                    Code = "Text.RegexReplace(""The Cat sat on the mat."", ""cat"", ""dog"", false, true)",
                    Result = """" & "The dog sat on the mat." & """"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in 
    documentedFunction
