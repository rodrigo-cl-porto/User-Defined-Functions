let 
    func = (X as list, Y as list) as number =>
        let 
            N = List.Count(X)
        in if N > 0 and List.Count(Y) = N then
            let
                SumX = List.Sum(X),
                SumY = List.Sum(Y),
                SumXY = List.Accumulate({0..N-1}, 0, (sum, i) => sum + X{i} * Y{i}),
                SumX2 = List.Accumulate({0..N-1}, 0, (sum, i) => sum + Number.Power(X{i}, 2))
            in  
                (SumY * SumX2 - SumX * SumXY) / (N * SumX2 - Number.Power(SumX, 2))
        else
            null,
    documentedType = type function(
        X as (type {number} meta [
            Documentation.FieldCaption = "X",
            Documentation.FieldDescription = "Numerical list for independent values"
        ]),
        Y as (type {number} meta [
            Documentation.FieldCaption = "Y",
            Documentation.FieldDescription = "Numerical list for dependent values" 
        ])
    ) as nullable number meta [
        Documentation.Name = "List.Intercept",
        Documentation.Description = "Calculates the intercept of the linear regression line between two numerical lists 𝚇 and 𝚈.",
        Documentation.LongDescription = 
            "This function returns the intercept of the simple linear regression line between two lists of numbers. "
            & "It is useful for statistical analysis and predictive modeling. "
            & "The intercept represents the expected value of 𝚈 when 𝚇 is zero, based on the least squares method.",
        Documentation.Category = "Statistics",
        Documentation.Version = "1.0.0",
        Documentation.Author = "Rodrigo Celso de Lima Porto",
        Documentation.Examples = {[
            Description = "Calculates the intercept between two numerical lists of equal length.",
            Code = "List.Intercept({1, 2, 3}, {4, 5, 6})",
            Result = "3"
        ], [
            Description = "Returns 𝚗𝚞𝚕𝚕 when the lists have different lengths.",
            Code = "List.Intercept({1, 2}, {3})",
            Result = "null"
        ]}
    ],
    documentedFunction = Value.ReplaceType(func, documentedType)
in 
    documentedFunction
