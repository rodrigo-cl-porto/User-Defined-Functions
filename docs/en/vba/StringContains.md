# [`StringContains`](/src/vba/StringContains.vba)

Checks if a string contains another string as a substring, with optional case sensitivity.

## Syntax

```vb
StringContains( _
    str1 As String, _
    str2 As String, _
    Optional caseSensitive As Boolean = False _
) As Boolean
```

## Parameters

- `str1`: The main string to search in
- `str2`: The substring to search for
- `caseSensitive`: (_optional_) Boolean flag to enable case-sensitive search. Default is `False`

## Return Value

Returns `True` if `str2` is found within `str1`, `False` otherwise.

## **Use Cases**

- Text validation
- String searching
- Pattern matching without regular expressions
- Case-insensitive text comparisons

## Example

```vb
Dim result As Boolean

result = StringContains("Hello World", "world")
Debug.Print result ' Returns True

result = StringContains("Hello World", "WORLD")
Debug.Print result ' Returns True

result = StringContains("Hello World", "world", True)
Debug.Print result ' Returns False

result = StringContains("Test", "xyz")
Debug.Print result ' Returns False
```
