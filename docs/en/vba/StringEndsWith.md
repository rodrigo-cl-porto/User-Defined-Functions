# [`StringEndsWith`](/src/vba/StringEndsWith.vba)

Checks if a string ends with another string, with optional case sensitivity.

## Syntax

```vb
StringEndsWith( _
    str1 As String, _
    str2 As String, _
    Optional caseSensitive As Boolean = False _
) As Boolean
```

## Parameters

- `str1`: The main string to check
- `str2`: The ending string to look for
- `caseSensitive`: (_optional_) Boolean flag to enable case-sensitive comparison. Default is `False`

## Return Value

Returns `True` if `str1` ends with `str2`, `False` otherwise. Also returns `False` if `str2` is longer than `str1`.

## **Use Cases**

- File extension validation
- Text suffix checking
- String pattern matching
- Domain name validation

## Example

```vb
Dim result As Boolean

result = StringEndsWith("Hello World", "world")
Debug.Print result ' Returns True

result = StringEndsWith("Hello World", "WORLD")
Debug.Print result ' Returns True

result = StringEndsWith("Hello World", "World", True)
Debug.Print result ' Returns True

result = StringEndsWith("Test", "xyz")
Debug.Print result ' Returns False
```
