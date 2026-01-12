let 
    func = (datetimeToConvert as datetime) as number => let
        RefDatetime = #datetime(1970, 1, 1, 0, 0, 0),
        TotalSeconds = Duration.TotalSeconds(datetimeToConvert-RefDatetime)
        in TotalSeconds,
    
    documentation = type function (
        datetimeToConvert as (type datetime meta [
            Documentation.FieldCaption = "DateTime to Convert",
            Documentation.FieldDescription = "The DateTime value to be converted to Unix time."
        ])) as number meta [
            Documentation.Name = "DateTime.ToUnixTime",
            Documentation.Description = "Converts a DateTime value to Unix time (the number of seconds since January 1, 1970).",
            Documentation.LongDescription = "This function takes a DateTime value as input and converts it to Unix time, which is represented as the total number of seconds that have elapsed since the Unix epoch (January 1, 1970, at 00:00:00 UTC). The result is returned as a numeric value.",
            Documentation.Category = "DateTime",
            Documentation.Version = "1.0.0",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Examples = {[
                Description = "Convert a specific DateTime to Unix time.",
                Code = "DateTime.ToUnixTime(#datetime(2023, 1, 1, 0, 0, 0))",
                Result = "1672531200"
            ]}
        ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
