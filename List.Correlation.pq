(list1 as list, list2 as list, optional typeCorrelation as nullable text) as number =>
let
    HandledList1 = List.Transform(
        list1,
        each let
            CleanedList = 
                if _ = null or _ = "" or not Value.Is(_, type number) then
                    0
                else 
                    _
            ,
            HandledList = 
                if Text.Upper(typeCorrelation) = "SPEARMAN" then
                    fxRank(CleanedList)
                else if Text.Upper(typeCorrelation) = "PEARSON" or typeCorrelation = null then
                    CleanedList
                else 
                    Error.Record(
                        "Not a valid correlation type!",
                        "The correlation type informed is not valid for this function.",
                        "Only pearson or spearman are valid. Please, check if there's any typo."
                    )
        in 
            HandledList
    ),
    HandledList2 = List.Transform(
        list2,
        each let
            CleanedList = 
                if _ = null or _ = "" or not Value.Is(_, type number) then
                    0
                else
                    _
            ,
            HandledList = 
                if Text.Upper(typeCorrelation) = "SPEARMAN" then
                    fxRank(CleanedList)
                else if Text.Upper(typeCorrelation) = "PEARSON" or typeCorrelation = null then
                    CleanedList
                else 
                    Error.Record(
                        "Not valid correlation type!",
                        "The correlation type informed is not valid for this function.",
                        "Only pearson or spearman are valid. Please, check if there's any typo."
                    )
        in 
            HandledList
    ),
    x_mean = List.Average(HandledList1),
    y_mean = List.Average(HandledList2),
    ZippedValues = List.Zip({HandledList1, HandledList2}),
    Numerator = List.Accumulate(
        ZippedValues,
        0,
        (state as number, current as list) =>
        let 
            x = current{0},
            y = current{1}
        in 
            state + (x - x_mean)*(y - y_mean)
    ),
    x_req = Number.Power(
        List.Accumulate(
            ZippedValues,
            0,
            (state as number, current as list) =>
            let 
                x = current{0}
            in 
                state + Number.Power(x - x_mean, 2)
        ),
        0.5
    ),
    y_req = Number.Power(
        List.Accumulate(
            ZippedValues,
            0,
            (state as number, current as list) =>
            let
                y = current{1}
            in 
                state + Number.Power(y - y_mean, 2)
        ),
        0.5
    ),
    Correlation = Numerator/(x_req * y_req)
in
    Correlation
