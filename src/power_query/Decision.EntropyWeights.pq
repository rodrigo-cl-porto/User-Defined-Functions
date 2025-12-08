let
    func = (tbl as table, optional columnNames as list) as record => let
        ColumnNames = if columnNames = null or columnNames = {} then
            Table.ColumnsOfType(tbl, {type number, type nullable number})
        else
            columnNames,
        SelectedColumns = Table.SelectColumns(tbl, ColumnNames),
        MaxValuesRecord = let
            MaxValues = List.Transform(ColumnNames, each List.Max(Table.Column(SelectedColumns, _))),
            ConvertedToRecord = Record.FromList(MaxValues, ColumnNames)
            in ConvertedToRecord,
        MinValuesRecord = let 
            MinValues = List.Transform(ColumnNames, each List.Min(Table.Column(SelectedColumns, _))),
            ConvertedToRecord = Record.FromList(MinValues, ColumnNames)
            in ConvertedToRecord,
        NormalizationTransformations = List.Transform(
            ColumnNames,
            (columnName as text) as list =>
            {
                columnName,
                each let
                    Max = Record.Field(MaxValuesRecord, columnName),
                    Min = Record.Field(MinValuesRecord, columnName),
                    Normalized = if Max = Min then 1 else (_ - Min)/(Max - Min)
                    in Normalized,
                type number
            }
        ),
        NormalizedValues = Table.TransformColumns(SelectedColumns, NormalizationTransformations),
        IndividualEntropyTransformations = List.Transform(
            ColumnNames,
            (columnName as text) as list =>
            {
                columnName,
                each let
                    Vj = List.Sum(Table.Column(NormalizedValues, columnName)),
                    p = _ / Vj,
                    Entropy = if p > 0 then p * Number.Ln(p) else 0
                    in Entropy,
                type number
            }
        ),
        IndividualEntropyCalculated = Table.TransformColumns(NormalizedValues, IndividualEntropyTransformations),
        WeightCriteriaRecord = let
            N = Table.RowCount(tbl), //Calculating criteria's entropy
            sMax = Number.Ln(N), // Maximum Entropy
            EntropyCriteria = List.Transform(
                ColumnNames,
                (columnName as text) as number => let
                    IndividualEntropiesSummmed = List.Sum(
                        Table.Column(IndividualEntropyCalculated, columnName)
                    ),
                    sj = -IndividualEntropiesSummmed / sMax // Criterion Entropy
                    in sj
            ),

            // Calculating criteria's weights
            S = List.Sum(EntropyCriteria),
            m = List.Count(ColumnNames),
            WeightCriteria = if N > 1 then 
                List.Transform(EntropyCriteria, each (1-_)/(m-S))
            else
                List.Repeat({1/m}, m),
            ConvertedToRecord = Record.FromList(WeightCriteria, ColumnNames)
            in ConvertedToRecord
        in WeightCriteriaRecord,
    
    documentation = type function (
        tbl as (type table meta [
            Documentation.FieldCaption = "table",
            Documentation.FieldDescription = "The input table containing numeric values for each criterion."
        ]),
        optional columnNames as (type {text} meta [
            Documentation.FieldCaption = "column names",
            Documentation.FieldDescription = "A list of column names to include in the calculation. If not provided, all numeric columns in the table will be used."
        ])) as record meta [
            Documentation.Name = "Decision.EntropyWeights",
            Documentation.Description = "This function calculates the weights of decision criteria using the entropy weighting method.",
            Documentation.LongDescription = "This function calculates the weights of decision criteria using the entropy weighting method. It analyzes the distribution of values across each criterion to determine their relative importance based on information entropy.",
            Documentation.Category = "Decision",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Version = "1.0.0",
            Documentation.Examples = {[
                Description = "Entropy weights for all numeric columns. Weights are assigned based on the variability of each criterion.",
                Code = "let" & "#(lf)" &
                "   Source = #table(type table" & "#(lf)" &
                "       [Cost=number, Quality=number, Speed=number], {" & "#(lf)" &
                "       {300, 80, 60}," & "#(lf)" &
                "       {250, 70, 75}," & "#(lf)" &
                "       {400, 90, 50}" & "#(lf)" &
                "   })," & "#(lf)" &
                "   Result = Decision.EntropyWeights(Source)" & "#(lf)" &
                "in" & "#(lf)" &
                "   Result",
                Result = "[Cost = 0.36, Quality = 0.31, Speed = 0.33]"
            ], [
                Description = "Calculates entropy weights only for specified columns.",
                Code = "let" & "#(lf)" &
                "   Source = #table(type table" & "#(lf)" &
                "       [Cost=number, Quality=number, Speed=number], {" & "#(lf)" &
                "       {300, 80, 60}," & "#(lf)" &
                "       {250, 70, 75}," & "#(lf)" &
                "       {400, 90, 50}" & "#(lf)" &
                "   })," & "#(lf)" &
                "   Result = Decision.EntropyWeights(Source, {""Cost"", ""Speed""})" & "#(lf)" &
                "in" & "#(lf)" &
                "   Result",
                Result = "[Cost = 0.52, Speed = 0.48]"
            ]}
        ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction