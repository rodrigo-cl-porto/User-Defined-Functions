(datetimeToConvert as datetime) as number =>
let
    RefDatetime = #datetime(1970, 1, 1, 0, 0, 0),
    TotalSeconds = Duration.TotalSeconds(datetimeToConvert-RefDatetime)
in
    TotalSeconds
