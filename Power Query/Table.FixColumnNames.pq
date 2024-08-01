(tb as table, optional TextFormat as nullable text) as table =>
let
    tbTransformedColumnNames = Table.TransformColumnNames(
        tb,
        each let
            CleanedName = Text.Clean(_),
            WeirdSpaceReplaced = Text.Replace(CleanedName, Character.FromNumber(160), " "),
            TrimmedName = Text.Trim(WeirdSpaceReplaced),
            RemovedDuplicatedSpaces = 
            let
                SplittedText = Text.Split(TrimmedName, " "),
                RemovedEmpties = List.RemoveItems(SplittedText, {""}),
                CombinedText = Text.Combine(RemovedEmpties, " ")
            in
                CombinedText
        in
            if Text.Upper(TextFormat) = "PROPER" then
                Text.Proper(RemovedDuplicatedSpaces)
            else if Text.Upper(TextFormat) = "LOWER" then
                Text.Lower(RemovedDuplicatedSpaces)
            else if Text.Upper(TextFormat) = "UPPER" then
                Text.Upper(RemovedDuplicatedSpaces)
            else
                RemovedDuplicatedSpaces
    ),
    ColumnNames = Table.ColumnNames(tbTransformedColumnNames),
    RelevantColumnsNames = List.Select(ColumnNames, each not Text.StartsWith(_, "COLUMN")),
    tbRelevantColumns = Table.SelectColumns(
        tbTransformedColumnNames,
        RelevantColumnsNames
    )
in
    tbRelevantColumns
