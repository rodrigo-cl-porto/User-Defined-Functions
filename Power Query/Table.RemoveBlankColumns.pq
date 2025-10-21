let
    func = (tbl as table) as table => let
        transformOperations = List.Transform(
            Table.ColumnNames(tbl),
            (columnName) => {columnName, each if List.Contains({null, ""}, _) then null else _}
        ),
        TransformedColumns = Table.TransformColumns(tbl, transformOperations),
        TableDescription = Table.Profile(TransformedColumns),
        FilteredRows = Table.SelectRows(TableDescription, each Record.HasFields(_, {"Count", "NullCount"}) and ([Count] > [NullCount])),
        NonBlankColumnNames = FilteredRows[Column],
        RemovedBlankColumns = Table.SelectColumns(TransformedColumns, NonBlankColumnNames)
        in FilteredRows,

    documentation = type function (
        tbl as (type table meta[
            Documentation.FieldCaption = "table",
            Documentation.FieldDescription = "The table from which blank columns will be removed."
        ])) as table meta [
            Documentation.Name = "Table.RemoveBlankColumns",
			Documentation.Description = "Removes columns from a table that contain only blank values.",
            Documentation.LongDescription = "This function scans each column in the input table and replaces any blank values (null or empty strings) with null. It then profiles the table to identify columns where all values are null and removes those columns from the final output.",
			Documentation.Category = "Table",
			Documentation.Author = "ExcelOffTheGrid",
            Documentation.Source = "https://www.youtube.com/watch?v=Zkg9ICg9i40",
			Documentation.Version = "1.0.1",
			Documentation.Examples = {[
                Description = "Remove columns with only blank values from a sample table.",
                Code = "let" & "#(lf)" &
                "   Source = #table({""A"", ""B""}, {{null, ""value1""}, {"""", ""value2""}})," & "#(lf)" &
                "   Result = Table.RemoveBlankColumns(Source)" & "#(lf)" &
                "in" & "#(lf)" &
                "   Result",
                Result = "#table({""B""}, {{""value1""}, {""value2""}})"
            ]}
        ],
        
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction