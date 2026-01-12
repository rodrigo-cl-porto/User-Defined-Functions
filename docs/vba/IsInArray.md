# [`IsInArray`](/src/vba/IsInArray.vba)

Checks whether a value exists in a one-dimensional array.

## Syntax

```vb
IsInArray( _
    ValueToBeFound As Variant, _
    SourceArray As Variant _
) As Boolean
```

## Parameters

- `ValueToBeFound`: The value to search for (any Variant).
- `SourceArray`: The one-dimensional array to search (Variant).

## Return Value

Returns `True` if the value is found in the array, otherwise returns `False`.

## Remarks

- Expects a one-dimensional array; passing an uninitialized or multi-dimensional array may cause errors.

## Example

```vb
Dim arr As Variant
arr = Array("apple", "banana", "cherry")

If IsInArray("banana", arr) Then
    Debug.Print "Found"
Else
    Debug.Print "Not found"
End If
```
