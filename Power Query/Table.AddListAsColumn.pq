(tb as table, columnName as text, columnValues as list, optional position as nullable number, optional columnType as nullable type) as table =>
let 
    #"Added Index" = Table.AddIndexColumn(tb, "Index", 0, 1),
    #"Added List as Column" = Table.AddColumn(
        #"Added Index",
        columnName,
        each columnValues{[Index]}, 
        if columnType = null then type any else columnType 
    ),
    #"Removed Index" = Table.RemoveColumns(#"Added List as Column", "Index")
in
    if position = null then
        #"Removed Index"
    else 
        let
            #"List of Reordered Columns" = List.InsertRange(Table.ColumnNames(tb), position - 1, {columnName}),
            #"Reordered Columns" = Table.ReorderColumns(#"Removed Index", #"List of Reordered Columns")
        in  
            #"Reordered Columns"
