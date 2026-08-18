let 
    func = (unixTimestamp as number) as datetime => let
        RefDatetime = #datetime(1970, 1, 1, 0, 0, 0),
        SECONDS_IN_A_MINUTE = 60,
        SECONDS_IN_A_HOUR = 60 * 60,
        SECONDS_IN_A_DAY = 60 * 60 * 24,
        NumberOfDays = Number.IntegerDivide(unixTimestamp, SECONDS_IN_A_DAY),
        NumberOfRemainingHours = Number.IntegerDivide(Number.Mod(unixTimestamp, SECONDS_IN_A_DAY), SECONDS_IN_A_HOUR),
        NumberOfRemainingMinutes = Number.IntegerDivide(Number.Mod(unixTimestamp, SECONDS_IN_A_HOUR), SECONDS_IN_A_MINUTE),
        NumberOfRemainingSeconds = Number.Mod(unixTimestamp, SECONDS_IN_A_MINUTE),
        ConvertedDatetime = RefDatetime + #duration(NumberOfDays, NumberOfRemainingHours, NumberOfRemainingMinutes, NumberOfRemainingSeconds)
        in ConvertedDatetime,
    
    documentation = type function (
        unixTimestamp as (type number meta [
            Documentation.FieldCaption = "Unix timestamp",
            Documentation.FieldDescription = "Total number of seconds since midnight on January 1st, 1970."
        ])) as datetime meta [
            Documentation.Name = "DateTime.FromUnixTime",
            Documentation.Description = "Converts a Unix timestamp (the number of seconds since January 1, 1970) to datetime.",
            Documentation.LongDescription = "This function takes a integer value that represents the total number of seconds that have elapsed since the Unix epoch (January 1, 1970, at 00:00:00 UTC) and converts it to datetime.",
            Documentation.Category = "DateTime",
            Documentation.Version = "1.0.2",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Examples = {[
                Description = "Convert a specific Unix timestamp to datetime.",
                Code = "DateTime.FromUnixTime(1672531200)",
                Result = "#datetime(2023, 1, 1, 0, 0, 0)"
            ], [
                Description = "Convert a specific Unix timestamp to datetime.",
                Code = "DateTime.FromUnixTime(1786649709)",
                Result = "#datetime(2026, 8, 13, 22, 7, 13)"
            ]}
        ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
