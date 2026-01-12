let 
    func = (tbl as table, alternativesColumn as text, weights as record) as any => let
        listAlternatives = let
            values = Table.Column(tbl, alternativesColumn),
            TextValues = List.Transform(values, Text.From)
            in TextValues,
        CriteriaNames = Record.FieldNames(weights),
        #"Normalized Values" = let
            Denominators = List.Transform(
                CriteriaNames,
                each let
                    ColumnValues = Table.Column(tbl, _),
                    Norm2 = List.Accumulate(
                        ColumnValues,
                        0,
                        (sum as number, current as number) as number => 
                            sum + Number.Power(current, 2)
                    ),
                    Norm = Number.Power(Norm2, 0.5)
                    in Norm
            ),
            DenominatorRecord = Record.FromList(Denominators, CriteriaNames),
            transformOperations = List.Transform(
                CriteriaNames,
                (Criterion as text) as list => 
                {
                    Criterion,
                    each let
                        ColumnValues = Table.Column(tbl, Criterion),
                        Denominator = Record.Field(DenominatorRecord, Criterion),
                        NormalizedValue = if Denominator = 0 then _ else _ / Denominator
                        in NormalizedValue,
                    type number
                }
            ),
            TransformedTable = Table.TransformColumns(tbl, transformOperations)
            in TransformedTable,
        #"Weighted Values" = let
            transformOperations = List.Transform(
                CriteriaNames,
                (Criterion as text) as list => 
                {
                    Criterion,
                    each _ * Record.Field(weights, Criterion),
                    type number
                }
            ),
            TransformedTable = Table.TransformColumns(#"Normalized Values", transformOperations)
            in TransformedTable,
        PIS = let // Positive Ideal Solution
            MaxValues = List.Transform(CriteriaNames, each List.Max(Table.Column(#"Weighted Values", _))),
            ConvertedToRecord = Record.FromList(MaxValues, CriteriaNames)
            in ConvertedToRecord,
        NIS = let // Negative Ideal Solution
            MinValues = List.Transform(CriteriaNames, each List.Min(Table.Column(#"Weighted Values", _))),
            ConvertedToRecord = Record.FromList(MinValues, CriteriaNames)
            in ConvertedToRecord,
        #"Added D+" = Table.AddColumn(
            #"Weighted Values",
            "D+",
            each List.Accumulate(
                CriteriaNames,
                0,
                (sum as number, criterion as text) as number => let 
                    value = Record.Field(_, criterion),
                    pisValue = Record.Field(PIS, criterion)
                    in sum + Number.Power(value-pisValue, 2)
            ),
            type number
        ),
        #"Added D-" = Table.AddColumn(
            #"Added D+",
            "D-",
            each List.Accumulate(
                CriteriaNames,
                0,
                (sum as number, criterion as text) as number => let 
                    value = Record.Field(_, criterion),
                    nisValue = Record.Field(NIS, criterion)
                    in sum + Number.Power(value-nisValue, 2)
            ),
            type number
        ),
        #"Merged to CC" = Table.CombineColumns( // Closeness coeficient
            #"Added D-",
            {"D+", "D-"},
            (values as list) as number => let 
                #"D+" = values{0},
                #"D-" = values{1},
                CC = if #"D+" = 0 and #"D-" = 0 then 0.5 else #"D-"/(#"D+" + #"D-")
                in CC,
            "CC"
        ),
        #"Added CC" = let 
            ccRecord = Record.FromList(#"Merged to CC"[CC], Table.Column(#"Merged to CC", alternativesColumn)),
            AddedCC = Table.AddColumn(
                tbl, 
                "CC", 
                each let
                    alternative = Record.Field(_, alternativesColumn),
                    CC = Record.Field(ccRecord, alternative)
                    in CC, 
                type number
            )
            in AddedCC,
        #"Sorted Alternatives" = Table.Sort(#"Added CC", {{"CC", Order.Descending}}),
        #"Added Ranking" = let 
            ccList = List.Buffer(#"Sorted Alternatives"[CC]),
            N = List.Count(ccList),
            RankingList = List.Accumulate(
                {0..N-1},
                {},
                (state as list, i as number) as list =>
                if i = 0 then
                    {1}
                else if ccList{i} = ccList{i-1} then 
                    List.Combine({state, {List.Last(state)}})
                else
                    List.Combine({state, {List.Last(state)+1}})
            ),
            AddedIndex = Table.AddIndexColumn(#"Sorted Alternatives", "RANKING", 0, 1),
            AddedRanking = Table.TransformColumns(AddedIndex, {"RANKING", each RankingList{_}, Int64.Type})
            in AddedRanking
        in #"Added Ranking",

    documentation = type function (
        tbl as (type table meta[
            Documentation.FieldCaption = "table",
            Documentation.FieldDescription = "The input table containing alternatives and their evaluation across multiple criteria."
        ]),
        alternativesColumn as (type text meta[
            Documentation.FieldCaption = "alternativesColumn",
            Documentation.FieldDescription = "The name of the column that identifies each alternative."
        ]),
        weights as (type record meta[
            Documentation.FieldCaption = "alternativesColumn",
            Documentation.FieldDescription = "A record where each field name corresponds to a criterion column in the table, and each field value is the weight (importance) of that criterion."
        ])) as table meta [
            Documentation.Name = "Decision.TOPSIS",
			Documentation.Description = "This function applies the TOPSIS (Technique for Order Preference by Similarity to Ideal Solution) method to rank alternatives based on multiple criteria.",
            Documentation.LongDescription = "This function applies the TOPSIS (Technique for Order Preference by Similarity to Ideal Solution) method to rank alternatives based on multiple criteria. It normalizes the data, applies weights, calculates the distance to the ideal and anti-ideal solutions, and computes a closeness coefficient (CC) for each alternative. The result is a ranked table of alternatives.",
			Documentation.Category = "Decision",
			Documentation.Author = "Rodrigo Celso de Lima Porto",
			Documentation.Version = "1.0.0",
			Documentation.Examples = {[
                Description = "Ranking alternatives with three criteria.",
                Code = "let" & "#(lf)" &
                "   Source = #table(" & "#(lf)" &
                "       {""Alternative"", ""Cost"", ""Quality"", ""Speed""}, {" & "#(lf)" &
                "       {""A"", 300, 80, 60}," & "#(lf)" &
                "       {""B"", 250, 70, 75}," & "#(lf)" &
                "       {""C"", 400, 90, 50}"  & "#(lf)" &
                "   })," & "#(lf)" &
                "   Weights = [Cost = 0.4, Quality = 0.3, Speed = 0.3]," & "#(lf)" &
                "   Result = Decision.TOPSIS(Source, ""Alternative"", Weights)" & "#(lf)" &
                "in" & "#(lf)" &
                "   Result",
                Result = "#table(" & "#(lf)" &
                "   {""Alternative"", ""Cost"", ""Quality"", ""Speed"", ""CC"", ""RANKING""}," & "#(lf)" &
                "   {" & "#(lf)" &
                "       {""C"", 400, 90, 50, 0.72, 1}"  & "#(lf)" &
                "       {""B"", 250, 70, 75, 0.26, 2}," & "#(lf)" &
                "       {""A"", 300, 80, 60, 0.26, 3}," & "#(lf)" &
                "   }" & "#(lf)" &
                ")"
            ]}
        ],
        
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction