(tb as table, optional columns as nullable list, optional firstColumnName as nullable text) as table =>
let 
    ColumnNames = if columns = null then Table.ColumnNames(tb) else columns,
    #"Transposed Table" =
        if columns = null then
            Table.Transpose(tb) 
        else let 
            SelectedColumns = Table.SelectColumns(tb, ColumnNames),
            TransposedTable = Table.Transpose(SelectedColumns)
        in
            TransposedTable
    ,
    #"Promoted Headers" = Table.PromoteHeaders(#"Transposed Table"),
    ColumnName = if firstColumnName = null then ColumnNames{0} else firstColumnName,
    ColumnValues = List.Range(ColumnNames, 1),
    #"Added First Column" = fxAddListAsColumn(#"Promoted Headers", ColumnName, ColumnValues, 1)
in 
    #"Added First Column"
