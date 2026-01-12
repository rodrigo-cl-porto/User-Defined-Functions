let
    func = (tbl as table, optional textFormat as text) as table => let
        tbTransformedColumnNames = Table.TransformColumnNames(
            tbl,
            each let
                CleanedName = Text.Clean(_),
                WeirdSpaceReplaced = Text.Replace(CleanedName, Character.FromNumber(160), " "),
                TrimmedName = Text.Trim(WeirdSpaceReplaced),
                RemovedDuplicatedSpaces = let
                    SplittedText = Text.Split(TrimmedName, " "),
                    RemovedEmpties = List.RemoveItems(SplittedText, {""}),
                    CombinedText = Text.Combine(RemovedEmpties, " ")
                    in CombinedText
                in if Text.Upper(textFormat) = "PROPER" then
                    Text.Proper(RemovedDuplicatedSpaces)
                else if Text.Upper(textFormat) = "LOWER" then
                    Text.Lower(RemovedDuplicatedSpaces)
                else if Text.Upper(textFormat) = "UPPER" then
                    Text.Upper(RemovedDuplicatedSpaces)
                else
                    RemovedDuplicatedSpaces
        ),
        ColumnNames = Table.ColumnNames(tbTransformedColumnNames),
        RelevantColumnsNames = List.Select(ColumnNames, each not Text.StartsWith(_, "COLUMN")),
        tbRelevantColumns = Table.SelectColumns(tbTransformedColumnNames, RelevantColumnsNames)
        in tbRelevantColumns,

    documentation = type function (
        tbl as (type table meta [
            Documentation.FieldCaption = "table",
            Documentation.FieldDescription = "The input table whose column names need to be fixed."
        ]),
        optional textFormat as (type text meta [
            Documentation.FieldCaption = "textFormat",
            Documentation.FieldDescription = "The desired text format for the column names. Accepts 'Proper', 'Lower', or 'Upper'. If not specified, no formatting is applied.",
            Documentation.AllowedValues = {"Proper", "Lower", "Upper"},
            Documentation.FieldDefaultValue = null
        ])) as table meta [
            Documentation.Name = "Table.NormalizeColumnNames",
            Documentation.Description = "Cleans and standardizes column names in a table by removing unwanted characters, trimming spaces, and applying specified text formatting (Proper, Lower, Upper). It also removes columns with default names like 'Column1', 'Column2', etc.",
            Documentation.LongDescription = "This function processes the column names of the provided table to ensure they are clean and standardized. It removes non-printable characters, trims leading and trailing spaces, replaces non-breaking spaces with regular spaces, eliminates duplicated spaces, and applies the specified text formatting (Proper, Lower, Upper). Additionally, it removes any columns that have default names such as 'Column1', 'Column2', etc., ensuring that only relevant columns remain in the output table.",
            Documentation.Category = "Table",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Version = "1.0.0",
            Documentation.Examples = {[
                Description = "Fix column names to Proper case and remove default columns.",
                Code = "let" & #(lf) &
                    "   Source = Excel.CurrentWorkbook(){[Name=""Table1""]}[Content]," & #(lf) &
                    "   FixedColumnNames = Table.NormalizeColumnNames(Source, ""Proper"")" & #(lf) &
                    "in" & #(lf) &
                    "   FixedColumnNames",
                Result = "A table with cleaned column names in Proper case and without default columns."
            ], [
                Description = "Fix column names to Upper case without removing default columns.",
                Code = "let" & #(lf) &
                    "   Source = Excel.CurrentWorkbook(){[Name=""Table1""]}[Content]," & #(lf) &
                    "   FixedColumnNames = Table.NormalizeColumnNames(Source, ""Upper"")" & #(lf) &
                    "in" & #(lf) &
                    "   FixedColumnNames",
                Result = "A table with cleaned column names in Upper case, retaining all original columns."
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
