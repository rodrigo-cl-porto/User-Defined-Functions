# [`StringStartsWith`](/src/vba/StringStartsWith.vba)

Checks whether a string starts with a specified substring, with optional case sensitivity.

## Syntax

```vb
StringStartsWith( _
    str1 As String, _
    str2 As String, _
    Optional caseSensitive As Boolean = False _
) As Boolean
```

## Parameters

- `str1`: The main string to check.
- `str2`: The prefix substring to look for.
- `caseSensitive`: (_optional_) If `True`, comparison is case-sensitive; if `False` (default), comparison is case-insensitive.

## Return Value

Returns `True` if `str1` starts with `str2`; otherwise returns `False`. Also returns `False` if `str2` is longer than `str1`.

## Example

```vb
Dim result As Boolean

result = StringStartsWith("Report.xlsx", "Report")
Debug.Print result ' True

result = StringStartsWith("Report.xlsx", "report")
Debug.Print result ' True (case-insensitive)

result = StringStartsWith("Report.xlsx", "report", True)
Debug.Print result ' False (case-sensitive)

result = StringStartsWith("Test", "LongPrefix")
Debug.Print result ' False
```
