let
    func = (tbl as table, optional columnNames as list, optional textCasing as text) as table => let 
        TextColumns = if List.Contains({null, {}}, columnNames) then
            Table.ColumnsOfType(tbl, {type text, type nullable text})
        else
            columnNames,
        ListColumnsTransformations = List.Transform(
            TextColumns,
            (columnName as text) as list => {
                columnName,
                each if _ <> null then let 
                    LineFeedReplaced = Text.Replace(_, "#(lf)", " "),
                    CleanedText = Text.Clean(LineFeedReplaced),
                    WeirdSpaceReplaced = Text.Replace(CleanedText, Character.FromNumber(160), " "),
                    TrimmedText = Text.Trim(WeirdSpaceReplaced),
                    RemovedDuplicatedSpaces = let
                        SplittedText = Text.Split(TrimmedText, " "),
                        RemovedEmpties = List.RemoveItems(SplittedText, {"", null}),
                        CombinedText = Text.Combine(RemovedEmpties, " ")
                        in CombinedText,
                    CasedText = if Text.Upper(textCasing) = "PROPER" then
                        Text.Proper(RemovedDuplicatedSpaces)
                        else if Text.Upper(textCasing) = "LOWER" then
                            Text.Lower(RemovedDuplicatedSpaces)
                        else if Text.Upper(textCasing) = "UPPER" then
                            Text.Upper(RemovedDuplicatedSpaces)
                        else
                            RemovedDuplicatedSpaces
                    in CasedText
                else "",
                type text
        }),
        TransformedTable = Table.TransformColumns(tbl, ListColumnsTransformations)
        in TransformedTable,
    
    documentation = type function (
        tbl as (type table meta [
            Documentation.FieldCaption = "table",
            Documentation.FieldDescription = "The input table containing text columns to be cleaned and formatted."
        ]),
        optional columnNames as (type {text} meta [
            Documentation.FieldCaption = "columnNames",
            Documentation.FieldDescription = "A list of column names to be processed. If not provided or empty, all columns of type text or nullable text will be processed."
        ]),
        optional textCasing as (type text meta [
            Documentation.FieldCaption = "textCasing",
            Documentation.FieldDescription = "A string indicating the desired text casing format. Accepted values are ""Proper"", ""Lower"" and ""Upper"". If not specified, casing is not changed.",
            Documentation.AllowedValues = {"Proper", "Lower", "Upper"},
            Documentation.DefaultValue = null
        ])) as table meta [
        Documentation.Name = "Table.NormalizeTextColumns",
		Documentation.Description = "This function cleans and formats text columns in a table.",
        Documentation.LongDescription = "This function cleans and formats text columns in a table. It removes line breaks, non-standard spaces, duplicated spaces, and applies optional casing (Proper, Lower, or Upper). You can specify which columns to process or let the function automatically detect all text columns.",
		Documentation.Category = "Table",
		Documentation.Author = "Rodrigo Celso de Lima Porto",
		Documentation.Version = "1.0.0",
		Documentation.Examples = {[
            Description = "Clean all text columns",
            Code = "let" & "#(lf)" &
            "   Source = #table({""Name"", ""Comment""}, {" & "#(lf)" &
            "       {""  JOHN DOE  "", ""Hello#(#)(lf)World""}," & "#(lf)" &
            "       {""  jane smith"", ""Nice to meet you""}" & "#(lf)" &
            "   })," & "#(lf)" &
            "   Result = Table.NormalizeTextColumns(Source)" & "#(lf)" &
            "in" & "#(lf)" &
            "   Result",
            Result = "#table({""Name"", ""Comment""}, {" & "#(lf)" &
            "   {""JOHN DOE"", ""Hello World""}," & "#(lf)" &
            "   {""jane smith"", ""Nice to meet you""}" & "#(lf)" &
            "})"
        ], [
            Description = "Clean and apply Proper case to selected columns",
            Code = "let" & "#(lf)" &
            "   Source = #table({""Name"", ""Note""}, {" & "#(lf)" &
            "       {""  MARIA   clara"", ""great#(#)(lf)job""}," & "#(lf)" &
            "       {""joão   SILVA"", ""excellent work""}" & "#(lf)" &
            "   })," & "#(lf)" &
            "   Result = Table.NormalizeTextColumns(Source, {""Name"", ""Note""}, ""Proper"")" & "#(lf)" &
            "in" & "#(lf)" &
            "   Result",
            Result = "#table({""Name"", ""Note""}, {" & "#(lf)" &
            "   {""Maria Clara"", ""Great Job""}," & "#(lf)" &
            "   {""João Silva"", ""Excellent Work""}" & "#(lf)" &
            "})"
        ]}
    ],

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction