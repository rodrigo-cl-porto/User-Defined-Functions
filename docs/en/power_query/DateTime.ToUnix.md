# [`DateTime.ToUnixTime`](/src/power_query/DateTime.ToUnixTime.pq)

Converts a Power Query datetime value to Unix time (seconds since 1970-01-01 00:00:00).

## Syntax

```fs
DateTime.ToUnixTime(
    datetimeToConvert as datetime
) as number
```

## Parameters

- `datetimeToConvert`: A datetime value to convert.

## Return Value

Converts `datetime` to Unixtime, which consists of a number representing the total seconds between `datetimeToConvert` and the Unix epoch (1970-01-01 00:00:00). Values are negative for datetimes before the epoch.

# Remarks

- No timezone conversion is performed — treat the input as UTC if you need UTC-based Unix time.

## Example

```fs
DateTime.ToUnixTime(#datetime(2023, 1, 1, 0, 0, 0))
```

**Result**

```fs
1672531200
```