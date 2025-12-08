let
    func = (textToRemove as nullable text) as text =>
        if textToRemove = null or textToRemove = "" then
            ""
        else let
            SplittedText = Text.Split(textToRemove, " "),
            RemovedEmpties = List.RemoveItems(SplittedText, {"", null}),
            CombinedText = Text.Combine(RemovedEmpties, " ")
            in CombinedText,
    
    documentation = type function (
        textToRemove as (type nullable text meta [
            Documentation.FieldCaption = "textToRemove",
            Documentation.FieldDescription = "The input text from which to remove double spaces"
        ])) as text meta [
            Documentation.Name = "Text.RemoveDoubleSpaces",
            Documentation.Description = "Removes all double spaces from the input text, replacing them with single spaces.",
            Documentation.LongDescription = "Scans the provided text for occurrences of double spaces and removes them by replacing them with single spaces. This process continues until no double spaces remain in the text.",
            Documentation.Category = "Text",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Version = "1.0.0",
            Documentation.Examples = {[
                Description = "Removes double spaces from a string.",
                Code = "Text.RemoveDoubleSpaces(""This  is   an example    text."")",
                Result = """This is an example text."""
            ], [
                Description = "Returns an empty string when input is null or empty.",
                Code = "Text.RemoveDoubleSpaces(null)",
                Result = """"
            ]}
        ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
