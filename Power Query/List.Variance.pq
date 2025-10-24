let
    func = (values as list, optional isPopulational as logical) as number =>
        if isPopulational = null or isPopulational = false then
            // Uses sampling variance formula, which takes (n-1) on the denominator.
            Number.Power(List.StandardDeviation(values), 2)
        else
            // Uses population variance formula, which takes (n) on the denominator.
            List.Covariance(values, values),

    documentation = type function (
        values as (type {number} meta [
            Documentation.FieldCaption = "values",
            Documentation.FieldDescription = "The numeric list whose variance is to be calculated."
        ]),
        optional isPopulational as (type logical meta [
            Documentation.FieldCaption = "isPopulational",
            Documentation.FieldDescription = "If true, calculates population variance; otherwise, sample variance",
            Documentation.AllowedValues = {true, false},
            Documentation.DefaultValue = true
        ])) as number meta [
            Documentation.Name = "List.Variance",
            Documentation.Description = "Calculates the variance of a numerical list, either population or sample-based.",
            Documentation.LongDescription = 
                "This function returns the variance of a list of numbers. "
                & "By default, it calculates the sample variance, using n-1 in the denominator. "
                & "If the optional parameter populational is set to true, it calculates the population variance, using n in the denominator.",
            Documentation.Category = "List",
            Documentation.Version = "1.0.0",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Examples = {[
                Description = "Calculates sample variance (default behavior).",
                Code = "List.Variance({1, 2, 3, 4, 5})",
                Result = "2.5"
            ], [
                Description = "Calculates population variance.",
                Code = "List.Variance({1, 2, 3, 4, 5}, true)",
                Result = "2"
            ], [
                Description = "Returns 0 for a list with identical values.",
                Code = "List.Variance({7, 7, 7})",
                Result = "0"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in 
    documentedFunction
