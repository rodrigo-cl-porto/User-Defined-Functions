# [`GetMonthNumberFromName`](/src/vba/GetMonthNumberFromName.vba)

Converts a month name to its corresponding numeric value (1-12).

## Syntax

```vb
GetMonthNumberFromName( _
    MonthName As String _
) As Integer
```

## Parameters

- `MonthName`: The name of the month (full or abbreviated, in any language supported by Excel)

## Return Value

Returns an integer from 1 to 12 representing the month number.

## Remarks

- Works with month names in any language supported by Excel
Accepts both full month names and abbreviated forms
- Case-insensitive
- Returns error if month name is invalid

## Example

```vb
Dim monthNum As Integer

monthNum = GetMonthNumberFromName("January")   
Debug.Print monthNum ' Prints 1

monthNum = GetMonthNumberFromName("Jan")
Debug.Print monthNum ' Prints 1

monthNum = GetMonthNumberFromName("Janeiro")
Debug.Print monthNum ' Prints 1 (Portuguese)

monthNum = GetMonthNumberFromName("Janvier") 
Debug.Print monthNum ' Returns 1 (French)
```
