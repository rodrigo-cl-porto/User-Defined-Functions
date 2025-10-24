let
    fxRank = (values as list, optional order as number) as list => let
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
            in
                List.InsertRange(state, N, {{previousIndex + 1, current}})
        ),
        SortedValues = List.Sort(
            AddedIndex,
            (x as list, y as list) as number => let
                x_value = x{1},
                y_value = y{1}
            in if order = Order.Ascending then 
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
            in  
                Value.Compare(x_index, y_index)
        ),
        ListOfRanks = List.Transform(SortedValueByIndex, each _{2})
        in ListOfRanks,
    
    func = (list1 as list, list2 as list, optional typeCorrelation as text) as number => let
        handleList = (values as list, typeCorr as text) as list => let
            CleanedList = List.Transform(values, each if _ = null or _ = "" or not Value.Is(_, type number) then 0 else _),
            HandledList = if Text.Upper(typeCorr) = "SPEARMAN" then
                fxRank(CleanedList)
            else if Text.Upper(typeCorr) = "PEARSON" or typeCorr = null then
                CleanedList
            else Error.Record(
                "Not a valid correlation type!",
                "The correlation type informed is not valid for this function.",
                "Only pearson or spearman are valid. Please, check if there's any typo."
            )
            in HandledList,
        HandledList1 = handleList(list1, typeCorrelation),
        HandledList2 = handleList(list2, typeCorrelation),
        x_mean = List.Average(HandledList1),
        y_mean = List.Average(HandledList2),
        ZippedValues = List.Zip({HandledList1, HandledList2}),
        Numerator = List.Accumulate(
            ZippedValues,
            0,
            (state as number, current as list) => let 
                x = current{0},
                y = current{1}
                in state + (x - x_mean)*(y - y_mean)
        ),
        x_req = Number.Sqrt(List.Accumulate(
            ZippedValues,
            0,
            (state as number, current as list) => let 
                x = current{0}
                in state + Number.Power(x - x_mean, 2)
        )),
        y_req = Number.Sqrt(List.Accumulate(
            ZippedValues,
            0,
            (state as number, current as list) => let
                y = current{1}
                in state + Number.Power(y - y_mean, 2)
        )),
        Correlation = Decimal.From(Numerator/(x_req * y_req))
        in Correlation,
    
    documentation = type function (
        list1 as (type {number} meta [
            Documentation.FieldCaption = "First List",
            Documentation.FieldDescription = "The first list of numeric values for correlation calculation."
        ]),
        list2 as (type {number} meta [
            Documentation.FieldCaption = "Second List",
            Documentation.FieldDescription = "The second list of numeric values for correlation calculation."
        ]),
        optional typeCorrelation as (type text meta [
            Documentation.FieldCaption = "Correlation Type",
            Documentation.FieldDescription = "The type of correlation to compute: 'Pearson' for Pearson correlation or 'Spearman' for Spearman rank correlation. Default is 'Pearson'.",
            Documentation.AllowedValues = {"Pearson", "Spearman"},
            Documentation.DefaultValue = "Pearson"
        ])
    ) as number meta [
        Documentation.Name = "List.Correlation",
        Documentation.Description = "Calculates the correlation coefficient between two lists of numbers.",
        Documentation.LongDescription = "This function computes the correlation coefficient between two lists of numeric values. It supports both Pearson and Spearman correlation types. The Pearson correlation measures the linear relationship between two datasets, while the Spearman correlation assesses the monotonic relationship based on the ranks of the data.",
        Documentation.Category = "List",
        Documentation.Version = "1.0.0",
        Documentation.Author = "Rodrigo Celso de Lima Porto",
        Documentation.Tags = {"correlation", "pearson", "spearman", "statistics", "data analysis"},
        Documentation.Examples = {[
            Description = "Calculate the Pearson correlation between two lists of numbers.",
            Code = "List.Correlation({0, 1, 3, 4}, {4, 5, 10, 30})",
            Result = "0.858575902776297"
        ], [
            Description = "Calculate the Spearman correlation between two lists of numbers.",
            Code = "List.Correlation({0, 1, 3, 4}, {4, 5, 10, 30}, ""Spearman"")",
            Result = "1"
        ], [
            Description = "Calculate the correlation even when lists contain nulls and non-numeric values. In this case, such values are treated as zeros.",
            Code = "List.Correlation({0, null, 3, ""a"", 4}, {4, 5, null, 10, 30})",
            Result = "0.556720639738652"
        ]}
    ],

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction