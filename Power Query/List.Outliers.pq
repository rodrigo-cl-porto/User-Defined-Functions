(values as list) as list =>
let 
    NonNullValues = List.RemoveItems(values, {null, "", " "}),
    NumValues = List.Select(NonNullValues, each Value.Is(Number.From(_), type number)),
    Q1 = List.Percentile(NumValues, 0.25), // First quartile
    Q3 = List.Percentile(NumValues, 0.75), // Third quartile
    IQR = Q3 - Q1, // Interquartile Range
    L_sup = Q3 + 1.5*IQR,
    L_inf = Q1 - 1.5*IQR,
    Outliers = List.Select(NumValues, each _ > L_sup or _ < L_inf)
in  
    Outliers
