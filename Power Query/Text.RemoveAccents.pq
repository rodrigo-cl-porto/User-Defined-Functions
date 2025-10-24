let 
    func = (textToConvert as nullable text) as text =>
        if textToConvert = null or textToConvert = "" then
            ""
        else let
            textBinary = Text.ToBinary(textToConvert, 1251),
            textASCII  = Text.FromBinary(textBinary, TextEncoding.Ascii)
            in textASCII,
    
    documentation = type function (
        textToConvert as (type nullable text meta [
            Documentation.FieldCaption = "textToConvert",
            Documentation.FieldDescription = "The input text containing accented characters to be converted"
        ])) as text meta [
            Documentation.Name = "Text.RemoveAccents",
            Documentation.Description = "Removes accented characters from the input text, converting them to their unaccented ASCII equivalents.",
            Documentation.LongDescription = "Processes the provided text to replace accented characters with their closest unaccented ASCII counterparts, effectively stripping accents while preserving the base characters.",
            Documentation.Category = "Text",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Version = "1.0.0",
            Documentation.Examples = {[
                Description = "Basic usage - removes accents from a string.",
                Code = "Text.RemoveAccents(""Café à la mode"")",
                Result = "Cafe a la mode"
            ]}
        ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
