let
    func = (tbl as table, columnName as text, columnValues as list, optional position as number, optional columnType as type) as table => let 
        AddedIndex = Table.AddIndexColumn(tbl, "Index", 0, 1),
        AddedColumn = Table.AddColumn(
            AddedIndex,
            columnName,
            each try columnValues{[Index]} otherwise null,
            if columnType = null then type any else columnType
        ),
        RemovedIndex = Table.RemoveColumns(AddedColumn, "Index")
        in if position = null then
            RemovedIndex
        else let
            ListOfReorderedColumns = List.InsertRange(Table.ColumnNames(tbl), position - 1, {columnName}),
            ReorderedColumns = Table.ReorderColumns(RemovedIndex, ListOfReorderedColumns)
            in ReorderedColumns,
    
    documentation = type function (
        tbl as (type table meta [
            Documentation.FieldCaption = "Table",
            Documentation.FieldDescription = "A table to which a new column will be added."
        ]),
        columnName as (type text meta [
            Documentation.FieldCaption = "Column Name",
            Documentation.FieldDescription = "The name of the new column to be added."
        ]),
        columnValues as (type list meta [
            Documentation.FieldCaption = "Column Values",
            Documentation.FieldDescription = "The list of values to populate the new column."
        ]),
        optional position as (type number meta [
            Documentation.FieldCaption = "Position",
            Documentation.FieldDescription = "Optional. The position (1-based index) where the new column should be inserted. If not provided, the column is added at the end."
        ]),
        optional columnType as (type type meta [
            Documentation.FieldCaption = "Column Type",
            Documentation.FieldDescription = "Optional. The data type of the new column. If not provided, defaults to type any.",
            Documentation.AllowedValues = {
                type any,
                type text,
                type number,
                Int64.Type,
                Currency.Type,
                type date,
                type datetime,
                type datetimezone,
                type time,
                type logical
            },
            Documentation.DefaultValue = type any
        ])) as table meta [
            Documentation.Name = "Table.AddColumnFromList",
            Documentation.Description = "Adds a list as a new column to a table at a specified position.",
            Documentation.LongDescription = "This function takes a table and adds a new column with values from a provided list. The new column can be inserted at a specified position within the table. If no position is provided, the new column is added at the end of the table. An optional column type can also be specified for the new column.",
            Documentation.Category = "Table",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Version = "1.0.0",
            Documentation.Examples = {[
                Description = "Add a list as a new column at the end of the table.",
                Code = "let" & "#(lf)" &
                    "    Source = Table.FromRecords({[A=1, B=2], [A=3, B=4]})," & "#(lf)" &
                    "    NewColumnValues = {10, 20}," & "#(lf)" &
                    "    Result = Table.AddColumnFromList(Source, ""C"", NewColumnValues)" & "#(lf)" &
                    "in" & "#(lf)" &
                    "    Result",
                Result = "Table.FromRecords({[A=1, B=2, C=10], [A=3, B=4, C=20]})"
            ], [
                Description = "Add a list as a new column at position 2 with a specified type.",
                Code = "let" & "#(lf)" &
                    "    Source = Table.FromRecords({[A=1, B=2], [A=3, B=4]})," & "#(lf)" &
                    "    NewColumnValues = {10, 20}," & "#(lf)" &
                    "    Result = Table.AddColumnFromList(Source, ""C"", NewColumnValues, 2, Int64.Type)" & "#(lf)" &
                    "in" & "#(lf)" &
                    "    Result",
                Result = "Table.FromRecords({[A=1, C=10, B=2], [A=3, C=20, B=4]})"
            ], [
                Description = "If list has fewer items than rows, nulls are added for remaining rows.",
                Code = "let" & "#(lf)" &
                    "    Source = Table.FromRecords({[A=1, B=2], [A=3, B=4], [A=5, B=6]})," & "#(lf)" &
                    "    NewColumnValues = {10, 20}," & "#(lf)" &
                    "    Result = Table.AddColumnFromList(Source, ""C"", NewColumnValues)" & "#(lf)" &
                    "in" & "#(lf)" &
                    "    Result",
                Result = "Table.FromRecords({[A=1, B=2, C=10], [A=3, B=4, C=20], [A=5, B=6, C=null]"
            ], [
                Description = "If list has more items than rows, extra items are ignored.",
                Code = "let" & "#(lf)" &
                    "    Source = Table.FromRecords({[A=1, B=2], [A=3, B=4]})," & "#(lf)" &
                    "    NewColumnValues = {10, 20, 30, 40}," & "#(lf)" &
                    "    Result = Table.AddColumnFromList(Source, ""C"", NewColumnValues)" & "#(lf)" &
                    "in" & "#(lf)" &
                    "    Result",
                Result = "Table.FromRecords({[A=1, B=2, C=10], [A=3, B=4, C=20]})"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
