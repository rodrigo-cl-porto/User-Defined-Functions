let 
    func = (values as list) as number => let 
        result = Number.Sqrt(List.Covariance(values, values))
        in Decimal.From(result),

    documentedType = type function (
        values as (type {number} meta [
            Documentation.FieldCaption = "values",
            Documentation.FieldDescription = "Numerical list to calculate the population standard deviation"
        ])) as number meta [
            Documentation.Name = "List.PopulationStdDev",
            Documentation.Description = "Calculates the population standard deviation of a numerical list.",
            Documentation.LongDescription = 
                "This function returns the population standard deviation of a list of numbers."
                & "The population standard deviation is a measure of the amount of variation or dispersion of a set of values."
                & "It is calculated using the formula: σ = √(Σ (xi - μ)² / N), where μ is the mean of the list and N is the number of values in the list.",
            Documentation.Category = "List",
            Documentation.Version = "1.0.0",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Examples = {[
                Description = "Calculates the population standard deviation of a numerical list.",
                Code = "List.PopulationStdDev({2, 4, 4, 4, 5, 5, 7, 9})",
                Result = "2"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentedType)
in  
    documentedFunction