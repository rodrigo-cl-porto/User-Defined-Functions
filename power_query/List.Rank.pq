(values as list, optional order as nullable number) as list =>
let 
    AddedIndex = List.Accumulate(
        values,
        {},
        (state as list, current as number) as list =>
        if state = {} then 
            List.InsertRange(state, 0, {{0, current}})
        else 
            let 
                N = List.Count(state),
                values = state{N-1},
                previousIndex = values{0}
            in
                List.InsertRange(state, N, {{previousIndex + 1, current}})
    ),
    SortedValues = List.Sort(
        AddedIndex,
        (x as list, y as list) as number =>
        let
            x_value = x{1},
            y_value = y{1}
        in  
            if order = Order.Ascending then 
                Value.Compare(x_value, y_value)
            else
                (-1) * Value.Compare(x_value, y_value)
    ),
    AddedRank = List.Accumulate(
        SortedValues,
        {},
        (state as list, current as list) as list =>
        let
            N = List.Count(state),
            currentIndex = current{0},
            currentValue = current{1},
            previousValues = if state = {} then null else state{N-1},
            previousValue = if previousValues = null then null else previousValues{1},
            previousRank = if previousValues = null then null else previousValues{2}
        in
            if state = {} then 
                List.InsertRange(state, 0, {{currentIndex, currentValue, 1}})
            else if currentValue = previousValue then 
                List.InsertRange(state, N, {{currentIndex, currentValue, previousRank}})
            else
                List.InsertRange(state, N, {{currentIndex, currentValue, previousRank+1}})
    ),
    SortedValueByIndex = List.Sort(
        AddedRank,
        (x as list, y as list) as number =>
        let
            x_index = x{0},
            y_index = y{0}
        in  
            Value.Compare(x_index, y_index)
    ),
    ListOfRanks = List.Transform(SortedValueByIndex, each _{2})
in
    ListOfRanks
