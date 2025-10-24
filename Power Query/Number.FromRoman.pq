let 
    func = (romanToConvert as text) as number => let
        Chars = Text.ToList(Text.Upper(romanToConvert))
        in if not List.MatchesAll(Chars, each List.Contains({"I", "V", "X", "L", "C", "D", "M"}, _)) then
            error Error.Record(
                "Expresion.Error",
                "Text is not a Roman numeral.",
                "The text has characters other than ""I"", ""V"", ""X"", ""L"", ""C"", ""D"" or ""M""."
            )
        else let
            ListNumbers = List.Transform(
                Chars,
                each if _ = "I" then 1
                else if _ = "V" then 5
                else if _ = "X" then 10
                else if _ = "L" then 50
                else if _ = "C" then 100
                else if _ = "D" then 500
                else if _ = "M" then 1000
                else 0
            ),
            ReversedListNumbers = List.Reverse(ListNumbers),
            N = List.Count(ReversedListNumbers),
            ConvertedToNumber = List.Accumulate(
                {0..N-1},
                0,
                (state as number, i as number) as number =>
                    if i = 0 then
                        state + ReversedListNumbers{0}
                    else if ReversedListNumbers{i} >= ReversedListNumbers{i-1} then
                        state + ReversedListNumbers{i}
                    else
                        state - ReversedListNumbers{i}
            )
            in ConvertedToNumber,

    documentation = type function (
        romanToConvert as (type text meta [
            Documentation.FieldCaption = "romanToConvert",
            Documentation.FieldDescription = "The Roman numeral in text format to be converted to a number."
        ])) as number meta [
            Documentation.Name = "Number.FromRoman",
            Documentation.Description = "Converts a Roman numeral (text) to a number.",
            Documentation.LongDescription = "This function takes a Roman numeral represented as text and converts it to its corresponding numerical value. It supports the standard Roman numeral characters: I, V, X, L, C, D, and M. If the input text contains invalid characters, an error is raised.",
            Documentation.Category = "Number",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Version = "1.0.0",
            Documentation.Examples = {[
                Description = "Convert a valid Roman numeral to a number.",
                Code = "Number.FromRoman(""XIV"")",
                Result = "14"
            ], [
                Description = "Handle invalid Roman numeral input.",
                Code = "Number.FromRoman(""ABCD"")",
                Result = "Error: Text is not a Roman numeral."
            ]},
            Documentation.SeeAlso = {"Number.ToRoman"}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
