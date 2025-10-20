let
    func = (numberToConvert as number) as any =>
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
        RomanSymbol[Roman],
    
    documentation = type function (
        numberToConvert as (type number meta [
            Documentation.FieldCaption = "numberToConvert",
            Documentation.FieldDescription = "The integer number to be converted to a Roman numeral."
        ])) as text meta [
            Documentation.Name = "Number.ToRoman",
            Documentation.Description = "Converts an integer number to a Roman numeral (between 1 and 3999).",
            Documentation.LongDescription = "This function converts an integer number in the range between 1 and 3999 to its corresponding Roman numeral representation. If the input number is outside this range or not an integer, an error is raised.",
            Documentation.Category = "Number",
            Documentation.Author = "Adapted from Livio - xcelanz.com",
            Documentation.Source = "https://xcelanz.com/converting-a-number-to-roman-with-power-query/",
            Documentation.Version = "1.0.1",
            Documentation.Examples = {[
                Documentation.Description = "Convert the number 12 to a Roman numeral.",
                Documentation.Code = "Number.ToRoman(12)",
                Documentation.Result = "XII"
            ], [
                Documentation.Description = "Attempt to convert an invalid number (e.g., 0) to a Roman numeral, resulting in an error.",
                Documentation.Code = "Number.ToRoman(0)",
                Documentation.Result = "Error"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in 
    documentedFunction
