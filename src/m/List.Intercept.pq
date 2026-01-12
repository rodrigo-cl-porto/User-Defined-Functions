let
    func = (X as list, Y as list) as number => let
        RemovedNonNumericValues = List.Accumulate(
            List.Zip({X, Y}),
            {},
            (state as list, current as list) as list => let
                x = current{0},
                y = current{1}
                in if Value.Is(x, type number) and Value.Is(y, type number) then
                    state & {current}
                else
                    state
        ),
        x = List.Transform(RemovedNonNumericValues, each _{0}),
        y = List.Transform(RemovedNonNumericValues, each _{1}),
        N = List.Count(x)
        in if N > 0 and List.Count(y) = N then
            if N > 1 then let
                SumX = List.Sum(x),
                SumY = List.Sum(y),
                SumXY = List.Accumulate({0..N-1}, 0, (sum, i) => sum + x{i} * y{i}),
                SumX2 = List.Accumulate({0..N-1}, 0, (sum, i) => sum + Number.Power(x{i}, 2)),
                Intercept = (SumY * SumX2 - SumX * SumXY) / (N * SumX2 - Number.Power(SumX, 2))
                in Intercept
            else
                0
        else
            null,
    
    documentation = type function(
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
        Documentation.Description = "Calculates the intercept of the linear regression line between two numerical lists X and Y.",
        Documentation.LongDescription = 
            "This function returns the intercept of the simple linear regression line between two lists of numbers. "
            & "It is useful for statistical analysis and predictive modeling. "
            & "The intercept represents the expected value of Y when X is zero, based on the least squares method.",
        Documentation.Category = "Statistics",
        Documentation.Version = "1.0.0",
        Documentation.Author = "Rodrigo Celso de Lima Porto",
        Documentation.Examples = {[
            Description = "Calculates the intercept between two numerical lists of equal length.",
            Code = "List.Intercept({1, 2, 3}, {4, 5, 6})",
            Result = "3"
        ], [
            Description = "Returns null when the lists have different lengths.",
            Code = "List.Intercept({1, 2}, {3})",
            Result = "null"
        ], [
            Description = "Ignores pairs with non-numeric values and calculates the intercept with the remaining ones.",
            Code = "List.Intercept({1, ""a"", 3}, {4, 5, 6})",
            Result = "3"
        ]}
    ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in 
    documentedFunction