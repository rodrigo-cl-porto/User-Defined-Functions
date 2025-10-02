let
    func = (X as list, Y as list) as nullable number =>
        let
            N = List.Count(X)
        in if N > 0 and List.Count(Y) = N then
            let
                SumX = List.Sum(X),
                SumY = List.Sum(Y),
                SumXY = List.Accumulate({0..N-1}, 0, (sum, i) => sum + X{i} * Y{i}),
                SumX2 = List.Accumulate({0..N-1}, 0, (sum, i) => sum + Number.Power(X{i}, 2)),
                Slope = (N * SumXY - SumX * SumY) / (N * SumX2 - Number.Power(SumX, 2))
            in
                Slope
        else
            null,
    documentedType = type function (
        X as (type {number} meta [
          Documentation.FieldCaption = "X",
          Documentation.FieldDescription = "Numerical list for independent values"
        ]),
        Y as (type {number} meta [
          Documentation.FieldCaption = "Y",
          Documentation.FieldDescription = "Numerical list for dependent values"
        ])
      ) as nullable number meta [
        Documentation.Name = "List.Slope",
        Documentation.Description = "Calculates the slope of the linear regression between two numerical lists 𝚇 and 𝚈.",
        Documentation.LongDescription = 
            "This function returns the slope (coefficient) of the simple linear regression line between two lists of numbers."
            & "It is useful for statistical analysis and predictive modeling. "
            & "The formula used is based on the least squares method, minimizing the squared error between observed and predicted values.",
        Documentation.Category = "Statistics",
        Documentation.Tags = {"regression", "statistics", "linear model", "slope"},
        Documentation.Version = "1.0.0",
        Documentation.Author = "Rodrigo Celso de Lima Porto",
        Documentation.Since = "2025-10-02",
        Documentation.Examples = {[
            Description = "Calculates the slope between two numerical lists of equal length.",
            Code = "List.Slope({1, 2, 3}, {4, 5, 6})",
            Result = "1"
        ], [ 
            Description = "Returns 𝚗𝚞𝚕𝚕 when the lists have different lengths.",
            Code = "List.Slope({1, 2}, {3})",
            Result = "null"
        ]}
    ],
    documentedFunction = Value.ReplaceType(func, documentedType)
in
    documentedFunction
