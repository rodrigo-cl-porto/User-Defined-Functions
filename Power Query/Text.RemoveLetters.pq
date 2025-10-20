let 
    func = (textToRemove as nullable text) as text =>
        if textToRemove = null or textToRemove = "" then
            ""
        else let
            textBinary = Text.ToBinary(textToRemove, 1251),
            textASCII  = Text.FromBinary(textBinary, TextEncoding.Ascii),
            RemovedLetters = List.RemoveItems(Text.ToList(textASCII), {"a".."z", "A".."Z"}),
            CombinedText = Text.Combine(RemovedLetters)
        in 
            CombinedText,

    documentation = type function (
        textToRemove as (type nullable text meta [
            Documentation.FieldCaption = "textToRemove",
            Documentation.FieldDescription = "The input text from which to remove letters"
        ])) as text meta [
            Documentation.Name = "Text.RemoveLetters",
            Documentation.Description = "Removes all letters from the input text, leaving only non-letter characters.",
            Documentation.LongDescription = "Scans the provided text for alphabetic characters (both uppercase and lowercase) and removes them, returning a string that contains only non-letter characters such as numbers, punctuation, and special symbols.",
            Documentation.Category = "Text",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Version = "1.0.0",
            Documentation.Examples = {[
                Description = "Removes letters from a mixed string.",
                Code = "Text.RemoveLetters(""Hello World! 123"")",
                Result = """ ! 123"""
            ], [
                Description = "Returns an empty string when input is null or empty.",
                Code = "Text.RemoveLetters(null)",
                Result = """"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
