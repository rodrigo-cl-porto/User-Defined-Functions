(tableToTransform as table, startDateColumn as text, endDateColumn as text) as table => let
    AddedNewStart = Table.AddColumn(
        tableToTransform, 
        "New Start",
        each let
            StartDate = Record.Field(_, startDateColumn),
            EndDate = Record.Field(_, endDateColumn), 
            EndOfStartDate = Date.AddDays(Date.StartOfDay(StartDate), 1) 
            in if EndDate > EndOfStartDate then let 
                StartOfEndDate = Date.StartOfDay(EndDate),
                N = Duration.Days(StartOfEndDate - EndOfStartDate),
                DatetimeList = {StartDate} & List.DateTimes(EndOfStartDate, N, #duration(1, 0, 0, 0)) 
                in DatetimeList  
            else
                {StartDate},
        type {datetime}
    ),
    AddedNewEnd = Table.AddColumn(
        AddedNewStart, 
        "New End",
        each let
            StartDate = Record.Field(_, startDateColumn),
            EndDate = Record.Field(_, endDateColumn),
            EndOfStartDate = Date.AddDays(Date.StartOfDay(StartDate), 1)
            in if [End] > EndOfStartDate then let
                StartOfEndDate = Date.StartOfDay(EndDate),
                N = Duration.Days(StartOfEndDate - EndOfStartDate),
                DatetimeList = List.DateTimes(EndOfStartDate, N, #duration(1, 0, 0, 0)) & {EndDate}
                in DatetimeList  
            else
                {EndDate},
        type {datetime}
    ),
    RemovedOldDates = Table.RemoveColumns(AddedNewEnd,{startDateColumn, endDateColumn}),
    MergedNewDates = Table.CombineColumns(
        RemovedOldDates,
        {"New Start", "New End"},
        each let
            ZippedList = List.Zip(_),
            ConvertedToRecord = List.Transform(
                ZippedList,
                each Record.FromList(_, {startDateColumn, endDateColumn})
            )
            in ConvertedToRecord,
        "New Dates"
    ),
    ExpandedListColumn = Table.ExpandListColumn(MergedNewDates, "New Dates"),
    ExpandedRecordColumn = Table.ExpandRecordColumn(
        ExpandedListColumn,
        "New Dates",
        {startDateColumn, endDateColumn}
    ),
    ReorderedColumns = Table.ReorderColumns(ExpandedRecordColumn, Table.ColumnNames(tableToTransform)),
    ChangedType = Table.TransformColumnTypes(ReorderedColumns, {{startDateColumn, type datetime}, {endDateColumn, type datetime}})
in
    ChangedType
