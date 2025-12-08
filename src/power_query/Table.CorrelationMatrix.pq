let
    fxCorrelation = (list1 as list, list2 as list) as number => let
        handleList = (values as list) as list => List.Transform(
            values, each if _ = null or _ = "" or not Value.Is(_, type number) then 0 else _
        ),
        HandledList1 = handleList(list1),
        HandledList2 = handleList(list2),
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

    func = (tbl as table, optional columnNames as list) => let
        tblBuffer = Table.Buffer(tbl),
        columns = if columnNames = null or columnNames = {} then 
                Table.ColumnNames(tblBuffer)
            else
                columnNames,
        N = List.Count(columns),
        RepeatedColumnNames = List.Transform(columns, each List.Repeat({_}, N)),
        AllColumnCombinations = List.Transform(RepeatedColumnNames, each List.Zip({_, columns})),
        MatrixCorrelation = List.Transform(
            AllColumnCombinations,
            each List.Transform( _,
                each let
                    column1 = Table.Column(tblBuffer, _{0}),
                    column2 = Table.Column(tblBuffer, _{1})
                    in fxCorrelation(column1, column2)
            )
        ),
        TableMatrixCorrelation = Table.Buffer(#table(columns, MatrixCorrelation)),
        AddedIndex = Table.AddIndexColumn(TableMatrixCorrelation, "Index", 0, 1),
        AddedVariableColumn = Table.AddColumn(AddedIndex, "VARIABLE", each columns{[Index]}, type text),
        RemovedIndex = Table.RemoveColumns(AddedVariableColumn, {"Index"}),
        ColumnTypesList = List.Transform(columns, each {_, type number}),
        TypeChanged = Table.TransformColumnTypes(RemovedIndex, ColumnTypesList),
        ReorderedColumns = Table.ReorderColumns(TypeChanged, List.Combine({{"VARIABLE"}, columns}))
        in ReorderedColumns,
    
    documentation = type function (
        tbl as (type table meta [
            Documentation.FieldCaption = "table",
            Documentation.FieldDescription = "The input table containing numeric columns to be analyzed."
        ]),
        optional columnNames as (type {text} meta [
            Documentation.FieldCaption = "columnNames",
            Documentation.FieldDescription = "A list of column names to include in the correlation matrix. If not provided, all numeric columns in the table will be used."
        ])) as table meta [
            Documentation.Name = "Table.CorrelationMatrix",
            Documentation.Description = "Calculates the correlation matrix for a given table.",
            Documentation.LongDescription = "Calculates the correlation matrix for a given table. It computes the Pearson correlation coefficient between each pair of numeric columns, returning a table where each cell represents the correlation between two variables.",
            Documentation.Category = "Table",
            Documentation.Version = "1.0.0",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Examples = {[
                Description = "Correlation matrix for all numeric columns.",
                Code = "let" & "#(lf)" &
                "   Source = #table({""A"", ""B"", ""C""}, {" & "#(lf)" &
                "       {1, 2, 3}," & "#(lf)" &
                "       {2, 4, 6}," & "#(lf)" &
                "       {3, 6, 9}" & "#(lf)" &
                "   })," & "#(lf)" &
                "   Result = Table.CorrelationMatrix(Source)" & "#(lf)" &
                "in" & "#(lf)" &
                "   Result",
                Result = "#table({""VARIABLE"", ""A"", ""B"", ""C""}, {" & "#(lf)" &
                "   {""A"", 1.0, 1.0, 1.0}," & "#(lf)" &
                "   {""B"", 1.0, 1.0, 1.0}," & "#(lf)" &
                "   {""C"", 1.0, 1.0, 1.0}" & "#(lf)" &
                "})"
            ], [
                Description = "Correlation matrix for selected columns.",
                Code = "let" & "#(lf)" &
                "   Source = #table({""X"", ""Y"", ""Z""}, {" & "#(lf)" &
                "       {1, 10, 100}," & "#(lf)" &
                "       {2, 20, 80}," & "#(lf)" &
                "       {3, 30, 60}," & "#(lf)" &
                "       {4, 40, 40}" & "#(lf)" &
                "   })," & "#(lf)" &
                "   Result = Table.CorrelationMatrix(Source, {""X"", ""Z""})" & "#(lf)" &
                "in" & "#(lf)" &
                "   Result",
                Result = "#table({""VARIABLE"", ""X"", ""Z""}, {" & "#(lf)" &
                "   {""X"",  1.0, -1.0}," & "#(lf)" &
                "   {""Z"", -1.0,  1.0}" & "#(lf)" &
                "})"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction