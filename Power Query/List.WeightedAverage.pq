let 
    func = (values as list, weights as list) as nullable number => let  
        N = List.Count(values)
        in if N < 1 then 
            null
        else let
            ZippedNumbers = List.Zip({values, weights}),
            Quotient = List.Accumulate(
                ZippedNumbers,
                0,
                (sum as number, current as list) as number => let  
                    value = current{0},
                    weight = current{1}
                    in sum + weight * value
            ),
            Divisor = List.Sum(weights)
            in Quotient / Divisor,
    
    documentation = type function (
        values as (type {number} meta [
            Documentation.FieldCaption = "values",
            Documentation.FieldDescription = "The list of numbers for which the weighted average is to be calculated."
        ]),
        weights as (type {number} meta [
            Documentation.FieldCaption = "weights",
            Documentation.FieldDescription = "The list of weights corresponding to each number in the values list."
        ])) as (type nullable number meta [
        Documentation.Name = "List.WeightedAverage",
        Documentation.Description = "Calculates the weighted average of a list of numbers, where weights are their respective positions in the list.",
        Documentation.LongDescription =
            "This function computes the weighted average of a list of numbers by assigning weights based on their positions in the list. "
                & "The first element has a weight of 1, the second a weight of 2, and so on, up to the N-th element which has a weight of N. "
                & "The weighted average is calculated by multiplying each number by its weight, summing these products, and then dividing by the sum of the weights.",
            Documentation.Category = "List",
            Documentation.Version = "1.0.0",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Examples = {[
                Description = "Calculates the weighted average of a list of numbers.",
                Code = "List.WeightedAverage({10, 20, 30}, {1, 2, 3})",
                Result = "23.3333333333333"
            ], [
                Description = "Returns null for an empty list.",
                Code = "List.WeightedAverage({}, {})",
                Result = "null"
            ]}
        ]),

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
