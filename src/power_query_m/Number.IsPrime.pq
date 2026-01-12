let
    func = (numberToCheck as number) as logical => let
        StringOf1s = Text.Repeat("1", numberToCheck),
        Script = "<script>" &
            "var regex = /^.?$|^(..+?)\1+$/;" &
            "var res = regex.test('" & StringOf1s & "');" &
            "document.write(res)" &
            "</script>",
        ExtractedText = Web.Page(Script)[Data]{0}[Children]{0}[Children]{1}[Text]{0},
        EvaluatedExpression = Expression.Evaluate("not " & ExtractedText)
        in EvaluatedExpression,

    documentation = type function (
        numberToCheck as (type number meta [
            Documentation.FieldCaption = "numberToCheck",
            Documentation.FieldDescription = "The number to be checked if it is a prime number."
        ])) as logical meta [
            Documentation.Name = "Number.IsPrime",
            Documentation.Description = "Checks if a given number is a prime number.",
            Documentation.LongDescription = "This function determines whether the provided number is a prime number." & "#(lf)" &
            "A prime number is defined as a natural number greater than 1 that cannot be formed by multiplying two smaller natural numbers." & "#(lf)" &
            "The function returns true if the number is prime and false otherwise.",
            Documentation.Category = "Number",
            Documentation.Version = "1.0.1",
            Documentation.Author = "Abigail",
            Documentation.Source = "http://test.neilk.net/blog/2000/06/01/abigails-regex-to-test-for-prime-numbers/",
            Documentation.Examples = {[
                Description = "Check if a prime number is identified correctly.",
                Code = "Number.IsPrime(7)",
                Result = "true"
            ], [
                Description = "Check if a non-prime number is identified correctly.",
                Code = "Number.IsPrime(100)",
                Result = "false"
            ]}
        ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
