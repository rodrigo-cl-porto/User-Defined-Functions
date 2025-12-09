let 
    func = (textToRemove as nullable text, optional replacer as text) as text =>
        if textToRemove = null or textToRemove = "" then
            ""
        else let
            Punctuations = Text.ToList("!""#$%&'()*+,-./:;<=>?@[]\^_`´{|}~ºª°§"),
            Chars = Text.ToList(textToRemove),
            ListWithoutPunctuation = List.Transform(
                Chars,
                each if List.Contains(Punctuations, _) then
                    if replacer = null then "" else replacer
                else 
                    _
            ),
            TextWithoutPunctuation = Text.Combine(ListWithoutPunctuation),
            SplittedText = Text.Split(TextWithoutPunctuation, " "),
            RemovedEmpties = List.RemoveItems(SplittedText, {"", null}),
            RemovedDoubleSpaces = Text.Combine(RemovedEmpties, " ")
            in RemovedDoubleSpaces,
    
    documentation = type function (
        textToRemove as (type nullable text meta [
            Documentation.FieldCaption = "textToRemove",
            Documentation.FieldDescription = "The text string from which punctuation characters will be removed."
        ]),
        optional replacer as (type text meta [
            Documentation.FieldCaption = "replacer",
            Documentation.FieldDescription = "An optional string to replace each punctuation character. If not provided, punctuation characters will be removed without replacement."
        ])) as text meta [
            Documentation.Name = "Text.RemovePunctuations",
            Documentation.Description = "Removes all punctuation characters from a text string, with an option to replace them with a specified character or string.",
            Documentation.LongDescription = "This function takes a text string and removes all punctuation characters from it. If a replacer string is provided, it replaces each punctuation character with the specified string. The function also ensures that multiple consecutive spaces resulting from the removal are reduced to a single space.",
            Documentation.Category = "Text",
            Documentation.Examples = {[
                Description = "Removing punctuations from a simple text.",
                Code = "Text.RemovePunctuations(""Hello, World!"")",
                Result = """Hello World"""
            ], [
                Description = "Removing punctuations and replacing them with a hyphen.",
                Code = "Text.RemovePunctuations(""Hello, World!"", ""-"")",
                Result = """Hello- World-"""
            ]}
        ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
