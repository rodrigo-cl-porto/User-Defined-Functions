let 
    func = (values as list) as list =>
        let 
            NonNullValues = List.RemoveItems(values, {null, "", " "}),
            NumValues = List.Select(NonNullValues, each Value.Is(Number.From(_), type number)),
            Q1 = List.Percentile(NumValues, 0.25), // First quartile
            Q3 = List.Percentile(NumValues, 0.75), // Third quartile
            IQR = Q3 - Q1, // Interquartile Range
            upperFence = Q3 + 1.5*IQR,
            lowerFence = Q1 - 1.5*IQR,
            Outliers = List.Select(NumValues, each _ > upperFence or _ < lowerFence)
        in  
            Outliers,
    documentedType = type function(
        values as (type {number} meta [
            Documentation.FieldCaption = "values",
            Documentation.FieldDescription = "A numerical sample"
        ])) as {number} meta [
            Documentation.Name = "List.Outliers",
            Documentation.Description = "Identifies outliers in a numerical list using the interquartile range (IQR) method.",
            Documentation.LongDescription = 
                "This function returns the outlier values from a list of numbers based on the interquartile range (IQR) method. "
                & "It first removes nulls, empty strings, and whitespace entries, then selects only valid numeric values. "
                & "Outliers are defined as values below 𝚀₁ - 𝟷.𝟻×𝙸𝚀𝚁 or above 𝚀₃ + 𝟷.𝟻×𝙸𝚀𝚁, where 𝚀₁ and 𝚀₃ are the first and third quartiles respectively.",
            Documentation.Category = "Statistics",
            Documentation.Version = "1.0.0",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Examples = {[
                Description = "Returns outliers from a list with extreme values.",
                Code = "List.Outliers({1, 2, 3, 4, 100})",
                Result = "{100}"
            ], [
                Description = "Returns an empty list when there are no outliers.",
                Code = "List.Outliers({10, 12, 13, 14, 15})",
                Result = "{}"
            ], [
                Description = "Ignores nulls and empty strings.",
                Code = "List.Outliers({1, null, """", 2, 3, 100})",
                Result = "{100}"
            ]}
        ],
    documentedFunction = Value.ReplaceType(func, documentedType)
in 
    documentedFunction
