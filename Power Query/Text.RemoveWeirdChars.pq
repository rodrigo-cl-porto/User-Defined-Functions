let 
    func = (textToClean as text, optional replacer as text) as text => let
        WeirdAsciiChars = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157, 160},
        repl = if replacer = null then " " else replacer,
        CleanedText = List.Accumulate(
            WeirdAsciiChars,
            textToClean,
            (currentText as text, AsciiChar as number) as text =>
            Text.Replace(currentText, Character.FromNumber(AsciiChar), repl)
        )
    in 
        CleanedText,
    
    documentation = type function (
        textToClean as (type text meta [
            Documentation.FieldCaption = "textToClean",
            Documentation.FieldDescription = "The input text from which weird/non-printable ASCII characters will be removed."
        ]),
        optional replacer as (type text meta [
            Documentation.FieldCaption = "replacer",
            Documentation.FieldDescription = "The string that will replace each weird/non-printable ASCII character found in the input text."
        ])) as text meta [
            Documentation.Name = "Text.RemoveWeirdChars",
            Documentation.Description = "Removes weird/non-printable ASCII characters from the input text and replaces them with the specified replacer string.",
            Documentation.LongDescription = "This function scans the input text for non-printable ASCII characters (ASCII codes 0-31, 127, 129, 141, 143, 144, 157, and 160) and replaces each occurrence with the provided replacer string. This is useful for cleaning up text data that may contain unwanted control characters.",
            Documentation.Category = "Text",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Version = "1.0.0",
            Documentation.Examples = {[
                Description = "Basic usage with space as replacer",
                Code = "Text.RemoveWeirdChars(""Hello"" & Character.FromNumber(10) & ""World!"")",
                Result = "Hello World!"
            ], [
                Description = "Replacing by custom replacer",
                Code = "Text.RemoveWeirdChars(""Data"" & Character.FromNumber(127) & ""Science"", ""_"")",
                Result = "Data_Science"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction