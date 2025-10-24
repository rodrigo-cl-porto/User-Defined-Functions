let
    func = (values as list, optional order as number) as list =>
    let 
        AddedIndex = List.Accumulate(
            values,
            {},
            (state as list, current as number) as list =>
            if state = {} then 
                List.InsertRange(state, 0, {{0, current}})
            else let
                N = List.Count(state),
                values = state{N-1},
                previousIndex = values{0}
                in List.InsertRange(state, N, {{previousIndex + 1, current}})
        ),
        SortedValues = List.Sort(
            AddedIndex,
            (x as list, y as list) as number => let
                x_value = x{1},
                y_value = y{1}
                in  if order = Order.Ascending then
                    Value.Compare(x_value, y_value)
                else
                    (-1) * Value.Compare(x_value, y_value)
        ),
        AddedRank = List.Accumulate(
            SortedValues,
            {},
            (state as list, current as list) as list => let
                N = List.Count(state),
                currentIndex = current{0},
                currentValue = current{1},
                previousValues = if state = {} then null else state{N-1},
                previousValue = if previousValues = null then null else previousValues{1},
                previousRank = if previousValues = null then null else previousValues{2}
                in if state = {} then
                    List.InsertRange(state, 0, {{currentIndex, currentValue, 1}})
                else if currentValue = previousValue then 
                    List.InsertRange(state, N, {{currentIndex, currentValue, previousRank}})
                else
                    List.InsertRange(state, N, {{currentIndex, currentValue, previousRank+1}})
        ),
        SortedValueByIndex = List.Sort(
            AddedRank,
            (x as list, y as list) as number => let
                x_index = x{0},
                y_index = y{0}
                in Value.Compare(x_index, y_index)
        ),
        ListOfRanks = List.Transform(SortedValueByIndex, each _{2})
    in
        ListOfRanks,
    documentation = type function (
        values as (type {number} meta [
            Documentation.FieldCaption = "Values",
            Documentation.FieldDescription = "A list of numeric values to be ranked.",
            Documentation.SampleValues = {"{1,2,3}"}
        ]),
        optional order as (Order.Type meta [
            Documentation.FieldCaption = "Order",
            Documentation.FieldDescription = "An optional parameter specifying the order of ranking. Use Order.Ascending for ascending order (default) or Order.Descending for descending order.",
            Documentation.DefaultValue = Order.Descending,
            Documentation.AllowedValues = {Order.Ascending, Order.Descending}
        ])
    ) as list meta [
        Documentation.Name = "List.Rank",
        Documentation.Description = "Returns a list of ranks corresponding to the input list of values. Ties are assigned the same rank.",
        Documentation.LongDescription = "This function takes a list of numeric values and returns a list of ranks. The ranks are assigned based on the order specified (ascending or descending). In case of ties, the same rank is assigned to tied values.",
        Documentation.Category = "List",
        Documentation.Version = "1.0.0",
        Documentation.Author = "Rodrigo Celso de Lima Porto",
        Documentation.Examples = {[
            Description = "Rank a list of numeric values in descending order.",
            Code = "List.Rank({10, 20, 20, 30}) // default = Order.Descending",
            Result = "{3, 2, 2, 1}"
        ], [
            Description = "Rank a list of numeric values in ascending order.",
            Code = "List.Rank({10, 20, 20, 30}, Order.Ascending)",
            Result = "{1, 2, 2, 3}"
        ], [
            Description = "Rank a list with multiple ties. Tied values receive the same rank.",
            Code = "List.Rank({10, 10, 30, 30, 2})",
            Result = "{2, 2, 1, 1, 3}"
        ]}
    ],
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
