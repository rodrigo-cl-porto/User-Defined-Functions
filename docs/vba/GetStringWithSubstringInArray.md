# [`GetStringWithSubstringInArray`](/src/vba/GetStringWithSubstringInArray.vba)

Searches through an array of strings and returns the first string that contains a specified substring.

## Syntax

```vb
GetStringWithSubstringInArray( _
    SubString As String, _ 
    SourceArray As Variant, _
    Optional CaseSensitive As Boolean = False _
) As String
```

## Parameters

- `SubString`: The text to search for within each array element
- `SourceArray`: Array containing strings to search through
- `CaseSensitive`: (_optional_) Boolean flag to enable case-sensitive search. Default is False

## Return Value

Returns the first string from the array containing the substring. Returns an empty string if no match is found.

## Remarks

- Only processes elements that are strings (type `vbString`)
- Ignores non-string elements in the array
- Case-insensitive by default
- Returns first match found and exits
- Works with arrays of any dimension

## Dependencies

- Requires [`StringContains`]() function

## Example

```vb
Dim testArray As Variant
Dim result    As String

testArray = Array("Hello World", "Test String", "Another Text")

result = GetStringWithSubstringInArray("World", testArray)
Debug.Print result ' Returns "Hello World"

result = GetStringWithSubstringInArray("text", testArray)
Debug.Print result ' Returns "Another Text"

result = GetStringWithSubstringInArray("none", testArray)
Debug.Print result  ' Returns ""
```
