# [`PreviousMonthNumber`](/src/vba/PreviousMonthNumber.vba)

Returns the numeric month (1–12) that precedes the month of a given date.

## Syntax

```vb
PreviousMonthNumber( _
    dt As Date _
) As Integer
```

## Parameters

- `dt`: Date value used to determine the previous month

## Return Value

Returns an Integer from 1 to 12 representing the previous month. For dates in January, returns 12 (December).

## Example

```vb
Dim prev As Integer

prev = PreviousMonthNumber(DateSerial(2025, 3, 15))
Debug.Print prev ' returns 2 (February)

prev = PreviousMonthNumber(DateSerial(2025, 1, 10))
Debug.Print prev ' returns 12 (December)
```
