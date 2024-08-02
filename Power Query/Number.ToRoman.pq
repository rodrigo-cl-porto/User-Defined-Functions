(numberToConvert as number) as any =>
// Adapted from Livio
// Title: Converting a Number to Roman with Power Query
// URL: https://xcelanz.com/converting-a-number-to-roman-with-power-query/
// Date: September 21, 2018
// Access Date: August 1st, 2024

if numberToConvert < 1 or numberToConvert > 3999 then
    error Error.Record(
        "Expression.Error",
        "It's not possible do convert the number to Roman numeral.",
        "Number outside the range between 1 and 3999."
    )
else if Number.Mod(numberToConvert, 1) <> 0 then
    error Error.Record(
        "Expression.Error",
        "It's not possible do convert the number to Roman numeral.",
        "Number is not a integer."
    )
else let
    Symbols = {"I", "V", "X", "L", "C", "D", "M"},
    RomanSymbol = List.Accumulate( 
        {0, 2, 4, 6},
        [Roman="", TempVal=numberToConvert],
        (state as record, current as number) as record =>
        let
            digit = Number.Mod(state[TempVal], 10),
            str = 
            if digit >= 1 and digit <= 3 then
                Text.Repeat(Symbols{current}, digit)
            else if digit = 4 then
                Symbols{current} & Symbols{current + 1}
            else if digit >= 5 and digit <= 8 then
                Symbols{current + 1} & Text.Repeat( Symbols{current}, Number.Mod(digit, 5) )
            else if digit = 9 then
                Symbols{current} & Symbols{current + 2}
            else ""
        in
            [Roman = str & state[Roman], TempVal = Number.IntegerDivide(state[TempVal], 10)]
    )
in
    RomanSymbol[Roman]
