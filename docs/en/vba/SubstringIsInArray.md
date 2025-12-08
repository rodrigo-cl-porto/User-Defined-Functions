# [`SubstringIsInArray`](/src/vba/SubstringIsInArray.vba)

Searches a one-dimensional array for any string element that contains a specified substring and returns `True` on the first match.

## Syntax

```vb
StringStartsWith( _
    str1 As String, _
    str2 As String, _
    Optional caseSensitive As Boolean = False _
) As Boolean
```

## Parameters

- `subStr`: The substring to search for.
- `srcArray`: One-dimensional array containing elements to search.
- `caseSensitive`: (_optional_) If `True`, performs a case-sensitive search; default is `False`.

## Return Value

Returns `True` if any string element in `srcArray` contains `subStr`; otherwise returns `False`.

## Remarks

- Only inspects elements typed as `String`; non-string elements are ignored.

## Dependencies

- Depends on the helper function [`StringContains`](#stringcontains) for substring checks.

## Example

```vb
Dim arr As Variant
arr = Array("Hello World", "Sample", "Test")

Debug.Print SubstringIsInArray("world", arr)       ' True (case-insensitive)
Debug.Print SubstringIsInArray("WORLD", arr, True) ' False (case-sensitive)
```
