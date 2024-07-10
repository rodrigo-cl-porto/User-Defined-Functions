(tbl as table) as table =>
let
    TransposedTable = fxTransposeTable(tbl),
    RemovedNullRows = Table.SelectRows(
        TransposedTable,
        each let
            ListRow = Record.ToList(_),
            FirstValueRemoved = List.RemoveFirstN(ListRow, 1),
            AllValuesIsNull = List.MatchesAll(FirstValueRemoved, each not List.Contains({null, "", "NaN", "N/A", "None"}, _))
        in
            AllValuesIsNull
    ),
    RetransposedTable = fxTransposeTable(RemovedNullRows)
in
    RetransposedTable
