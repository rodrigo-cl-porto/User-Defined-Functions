(tbl as table, optional listColumns as nullable list) as any =>
let
    ColumnNames = if listColumns = null or listColumns = {} then
        Table.ColumnsOfType(tbl, {type number, type nullable number})
    else 
        listColumns,
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
            in 
                Normalized,
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
                p = _ / Vj
            in 
                if p > 0 then p * Number.Ln(p) else 0,
            type number
        }
    ),
    IndividualEntropyCalculated = Table.TransformColumns(NormalizedValues, IndividualEntropyTransformations),
    WeightCriteriaRecord = 
    let 
        //Calculating criteria's entropy
        N = Table.RowCount(tbl),
        sMax = Number.Ln(N), // Maximum Entropy
        EntropyCriteria = List.Transform(
            ColumnNames,
            (columnName as text) as number =>let
                IndividualEntropiesSummmed = List.Sum(Table.Column(IndividualEntropyCalculated, columnName)),
                sj = -(1/sMax) * IndividualEntropiesSummmed // Criterion Entropy
                in sj
        ),

        // Calculating criteria's weights
        S = List.Sum(EntropyCriteria),
        m = List.Count(ColumnNames),
        WeightCriteria = if N > 1 then List.Transform(EntropyCriteria, each (1-_)/(m-S)) else List.Repeat({1/m}, m),
        ConvertedToRecord = Record.FromList(WeightCriteria, ColumnNames)
    in 
        ConvertedToRecord
in
    WeightCriteriaRecord
