let
    func = (numberToCheck as number) as logical => if Number.Mod(numberToCheck, 1) = 0 then true else false,

    documentation = type function (
        numberToCheck as (type number meta [
            Documentation.FieldCaption = "numberToCheck",
            Documentation.FieldDescription = "The number to be checked if it is an integer."
        ])) as (type logical meta [
            Documentation.Name = "Number.IsInteger",
            Documentation.Description = "Checks if a given number is an integer.",
            Documentation.LongDescription = "This function evaluates whether the provided number is an integer by checking if there is no fractional part. It returns true if the number is an integer and false otherwise.",
            Documentation.Category = "Number",
            Documentation.Version = "1.0.0",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Examples = {
                [
                    Description = "Check if a whole number is an integer.",
                    Code = "Number.IsInteger(10)",
                    Result = "true"
                ],
                [
                    Description = "Check if a decimal number is an integer.",
                    Code = "Number.IsInteger(10.5)",
                    Result = "false"
                ]
            }
        ]),

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
