let 
    func = (values as list, optional multiplier as number) as list => let 
        NonNullValues = List.RemoveItems(values, {null, "", " "}),
        NumValues = List.Select(NonNullValues, each Value.Is(Number.From(_), type number)),
        Q1 = List.Percentile(NumValues, 0.25), // First quartile
        Q3 = List.Percentile(NumValues, 0.75), // Third quartile
        IQR = Q3 - Q1, // Interquartile Range
        mult = if multiplier is null then multiplier else 1.5,
        upperFence = Q3 + mult*IQR,
        lowerFence = Q1 - mult*IQR,
        Outliers = List.Select(NumValues, each _ > upperFence or _ < lowerFence)
        in Outliers,

    documentation = type function(
        values as (type {number} meta [
            Documentation.FieldCaption = "Values",
            Documentation.FieldDescription = "A numerical list sample",
            Formatting.Format = "List"
        ]),
        optional multiplier as (type number meta [
            Documentation.FieldCaption = "Multiplier",
            Documentation.FieldDescription = "Multiplier to define outlier thresholds (default is 1.5)",
            Documentation.DefaultValue = 1.5
        ])) as {number} meta [
            Documentation.Name = "List.Outliers",
            Documentation.Description = "Identifies outliers in a numerical list using the interquartile range (IQR) method.",
            Documentation.LongDescription = "This function returns the outlier values from a list of numbers based on the interquartile range (IQR) method. It first removes nulls, empty strings, and whitespace entries, then selects only valid numeric values. Outliers are defined as values below Q₁ - 1.5×IQR or above Q₃ + 1.5×IQR, where Q₁ and Q₃ are the first and third quartiles respectively.",
            Documentation.Category = "List",
            Documentation.Version = "1.0.1",
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
            ], [
                Description = "Uses a custom multiplier to define outliers more strictly.",
                Code = "List.Outliers({1, 2, 3, 4, 50, 100}, 2)",
                Result = "{100}"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in 
    documentedFunction