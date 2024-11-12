(textToClean as text, replacer as text) as text =>
let
    WeirdAsciiChars = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157, 160},
    CleanedText = List.Accumulate(
        WeirdAsciiChars,
        textToClean,
        (currentText as text, AsciiChar as number) as text =>
        Text.Replace(currentText, Character.FromNumber(AsciiChar), replacer)
    )
in 
    CleanedText
