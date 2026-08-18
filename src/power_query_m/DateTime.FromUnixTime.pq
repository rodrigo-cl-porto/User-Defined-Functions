let 
    func = (unixTimestamp as number, optional format as text) as datetime => let
        RefDatetime = #datetime(1970, 1, 1, 0, 0, 0)
        in if format = "milliseconds" then let
            MILLISECONDS_IN_A_MINUTE = 1000 * 60,
            MILLISECONDS_IN_A_HOUR = 1000 * 60 * 60,
            MILLISECONDS_IN_A_DAY = 1000 * 60 * 60 * 24,
            NumberOfDays = Number.IntegerDivide(unixTimestamp, MILLISECONDS_IN_A_DAY),
            NumberOfRemainingHours = Number.IntegerDivide(Number.Mod(unixTimestamp, MILLISECONDS_IN_A_DAY), MILLISECONDS_IN_A_HOUR),
            NumberOfRemainingMinutes = Number.IntegerDivide(Number.Mod(unixTimestamp, MILLISECONDS_IN_A_HOUR), MILLISECONDS_IN_A_MINUTE),
            NumberOfRemainingSeconds = Number.Mod(unixTimestamp, MILLISECONDS_IN_A_MINUTE),
            ConvertedDatetime = RefDatetime + #duration(NumberOfDays, NumberOfRemainingHours, NumberOfRemainingMinutes, NumberOfRemainingSeconds)
            in ConvertedDatetime
        else if format = "seconds" or format = "" or format = null then let
            SECONDS_IN_A_MINUTE = 60,
            SECONDS_IN_A_HOUR = 60 * 60,
            SECONDS_IN_A_DAY = 60 * 60 * 24,
            NumberOfDays = Number.IntegerDivide(unixTimestamp, SECONDS_IN_A_DAY),
            NumberOfRemainingHours = Number.IntegerDivide(Number.Mod(unixTimestamp, SECONDS_IN_A_DAY), SECONDS_IN_A_HOUR),
            NumberOfRemainingMinutes = Number.IntegerDivide(Number.Mod(unixTimestamp, SECONDS_IN_A_HOUR), SECONDS_IN_A_MINUTE),
            NumberOfRemainingSeconds = Number.Mod(unixTimestamp, SECONDS_IN_A_MINUTE),
            ConvertedDatetime = RefDatetime + #duration(NumberOfDays, NumberOfRemainingHours, NumberOfRemainingMinutes, NumberOfRemainingSeconds)
            in ConvertedDatetime
        else error [
            Reason = "ValueError",
            Message = "Invalid ",
            Detail = [#"Input Format" = format] 
        ],
    
    documentation = type function (
        unixTimestamp as (type number meta [
            Documentation.FieldCaption = "DateTime as Unix Timestamp",
            Documentation.FieldDescription = "DateTime as Unix time (total number of seconds or milliseconds since midnight January 1st, 1970)."
        ]),
        optional format as (type text meta [
            Documentation.FieldCaption = "Format of the informed Unix Timestamp",
            Documentation.FieldDescription = "Format as seconds or milliseconds.",
            Documentation.AllowedValues = { "seconds", "milliseconds" }
        ])) as datetime meta [
            Documentation.Name = "DateTime.FromUnixTime",
            Documentation.Description = "Converts a DateTime value to Unix time (the number of seconds since January 1, 1970).",
            Documentation.LongDescription = "This function takes a integer value that represents the total number of seconds that have elapsed since the Unix epoch (January 1, 1970, at 00:00:00 UTC) and converts it to DateTime.",
            Documentation.Category = "DateTime",
            Documentation.Version = "1.0.0",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Examples = {[
                Description = "Convert a specific DateTime to Unix time.",
                Code = "DateTime.FromUnixTime(1672531200)",
                Result = "#datetime(2023, 1, 1, 0, 0, 0)"
            ], [
                Description = "Convert to Unix timestamp by informing that it's in milliseconds instead of seconds.",
                Code = "DateTime.FromUnixTime(1786649709133, ""milliseconds"")",
                Result = "#datetime(2026, 8, 13, 22, 7, 13)"
            ]}
        ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
