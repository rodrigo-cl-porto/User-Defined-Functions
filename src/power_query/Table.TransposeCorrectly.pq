let
    fxAddListAsColumn = (tbl as table, columnName as text, columnValues as list, optional position as number, optional columnType as type) as table => let 
        #"Added Index" = Table.AddIndexColumn(tbl, "Index", 0, 1),
        #"Added List as Column" = Table.AddColumn(
            #"Added Index",
            columnName,
            each try columnValues{[Index]} otherwise null,
            if columnType = null then type any else columnType
        ),
        #"Removed Index" = Table.RemoveColumns(#"Added List as Column", "Index")
        in if position = null then
            #"Removed Index"
        else let
            #"List of Reordered Columns" = List.InsertRange(Table.ColumnNames(tbl), position - 1, {columnName}),
            #"Reordered Columns" = Table.ReorderColumns(#"Removed Index", #"List of Reordered Columns")
            in  #"Reordered Columns",

    func = (tbl as table, optional columnNames as list, optional firstColumnName as text) as table => let 
        ColumnNames = if columnNames = null then Table.ColumnNames(tbl) else columnNames,
        #"Transposed Table" = if columnNames = null or columnNames = {} then
            Table.Transpose(tbl)
        else let
            SelectedColumns = Table.SelectColumns(tbl, ColumnNames),
            TransposedTable = Table.Transpose(SelectedColumns)
        in 
            TransposedTable,
        #"Promoted Headers" = Table.PromoteHeaders(#"Transposed Table"),
        ColumnName = if firstColumnName = null then ColumnNames{0} else firstColumnName,
        ColumnValues = List.Range(ColumnNames, 1),
        #"Added First Column" = fxAddListAsColumn(#"Promoted Headers", ColumnName, ColumnValues, 1)
        in #"Added First Column",

    documentation = type function (
        tbl as (type table meta [
            Documentation.FieldCaption = "table",
            Documentation.FieldDescription = "The input table whose column names need to be fixed."
        ]),
        optional columnNames as (type {text} meta [
            Documentation.FieldCaption = "columnNames",
            Documentation.FieldDescription = "A list of column names to transpose. If not provided, all columns in the table will be transposed."
        ]),
        optional firstColumnName as (type text meta [
            Documentation.FieldCaption = "firstColumnName",
            Documentation.FieldDescription = "The name to assign to the first column of the transposed table. If not provided, the first name from the columns list will be used."
        ])) as table meta [
            Documentation.Name = "Table.TransposeCorrectly",
			Documentation.Description = "Transposes a table by converting selected columns (or all columns if none are specified) into rows, promotes headers, and adds a new column containing the original column names. This is useful for restructuring data while preserving column identity.",
			Documentation.Category = "Table",
			Documentation.Author = "Rodrigo Celso de Lima Porto",
			Documentation.Version = "1.0.0",
			Documentation.Examples = {[
                Description = "Transposing all columns",
                Code = "let" & "#(lf)" &
                "   Source = #table({""A"", ""B"", ""C""}, {{1, 2, 3}, {4, 5, 6}})," & "#(lf)" &
                "   Result = Table.TransposeCorrectly(Source)" & "#(lf)" &
                "in" & "#(lf)" &
                "   Result",
                Result = "Table.FromRecords({[A=""B"", 1=2, 4=5], [A=""C"", 1=3, 4=6]})"
            ], [
                Description = "Transposing only the selected columns",
                Code = "let" & "#(lf)" &
                "   Source = #table({""A"", ""B"", ""C""}, {{1, 2, 3}, {4, 5, 6}})," & "#(lf)" &
                "   Result = Table.TransposeCorrectly(Source, {""A"", ""B""})" & "#(lf)" &
                "in" & "#(lf)" &
                "   Result",
                Result = "Table.FromRecords({[A=""B"", 1=2, 4=5]})"
            ], [
                Description = "Transposing the table and changing the first column name",
                Code = "let" & "#(lf)" &
                "   Source = #table({""A"", ""B"", ""C""}, {{1, 2, 3}, {4, 5, 6}})," & "#(lf)" &
                "   Result = Table.TransposeCorrectly(Source, null, ""D"")" & "#(lf)" &
                "in" & "#(lf)" &
                "   Result",
                Result = "Table.FromRecords({[D=""B"", 1=2, 4=5], [D=""C"", 1=3, 4=6]})"
            ]}
		],

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
