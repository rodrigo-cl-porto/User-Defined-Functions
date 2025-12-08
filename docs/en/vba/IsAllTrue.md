# [`IsAllTrue`](/src/vba/IsAllTrue.vba)

Checks if all elements in a boolean array are `True`.

## Syntax

```vb
IsAllTrue( _
    blnArray As Variant _
) As Boolean
```

## Parameters

- `blnArray`: Array containing boolean values to be checked

## Return Value

Returns `True` if all elements in the array are boolean `True`, otherwise returns `False`.

## **Use Cases**

- Validating that multiple conditions are all met
- Checking status of multiple boolean flags
- Quality control checks where all criteria must be true

## Remarks

- Returns `False` if any element is not a boolean type
- Returns `False` if any element is `False`
- Early exit when first non-true value is found
- Can handle arrays of any dimension
- Array must be passed as `Variant` type

## Example

```vb
Dim testArray As Variant

testArray = Array(True, True, True)
Debug.Print IsAllTrue(testArray) ' Returns True

testArray = Array(True, False, True)
Debug.Print IsAllTrue(testArray) ' Returns False

testArray = Array(True, "True", True)
Debug.Print IsAllTrue(testArray) ' Returns False (non-boolean element)
```
