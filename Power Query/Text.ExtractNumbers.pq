let
    func = (textToExtract as nullable text) as list =>
        if textToExtract = null or textToExtract = "" then
            {}
        else let
            Digits = {"0".."9", " "},
            ExtractedDigits = List.Select(Text.ToList(textToExtract), each List.Contains(Digits, _)),
            ExtractedNumbers = List.RemoveItems(Text.Split(Text.Combine(ExtractedDigits), " "), {"", null}),
            ConvertedToListNumbers = List.Transform(ExtractedNumbers, Number.FromText)
        in
            ConvertedToListNumbers,
    
    documentation = type function (
        textToExtract as (type nullable text meta [
            Documentation.FieldCaption = "textToExtract",
            Documentation.FieldDescription = "The input text from which to extract numbers"
        ])) as list meta [
            Documentation.Name = "Text.ExtractNumbers",
            Documentation.Description = "Extracts all numeric values from the input text and returns them as a list of numbers.",
            Documentation.LongDescription = "Scans the provided text for sequences of digits, extracts them, and converts them into a list of numeric values. Non-numeric characters are ignored.",
            Documentation.Category = "Text",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Version = "1.0.0",
            Documentation.Examples = {[
                Description = "Extracts numbers from a string containing mixed characters.",
                Code = "Text.ExtractNumbers(""Order #12345: 67 items at $89 each."")",
                Result = "{12345, 67, 89}"
            ], [
                Description = "Returns an empty list when no numbers are present.",
                Code = "Text.ExtractNumbers(""No numbers here!"")",
                Result = "{}"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
