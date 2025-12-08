# [`GetStringBetween`](/src/vba/GetStringBetween.vba)

Extracts a substring between two specified delimiter strings.

## Syntax

```vb
GetStringBetween( _
    str As String, _
    startStr As String, _
    endStr As String _
) As String
```

## Parameters

- `str`: The input string to search in
- `startStr`: The starting delimiter string
- `endStr`: The ending delimiter string

## Return Value

Returns the text found between the start and end strings. Returns an empty string if no match is found.

## Remarks

- Uses VBScript RegExp for pattern matching
- Creates RegExp object using late binding to avoid explicit reference requirement
- Case-insensitive search
- Non-greedy matching (returns shortest match)
- Returns only the first match if multiple exist
- Removes the delimiter strings from the result

## Example

```vb
Dim result As String

result = GetStringBetween("Hello [World] Test", "[", "]")
Debug.Print result ' Returns "World"

result = GetStringBetween("<tag>Content</tag>", "<tag>", "</tag>")
Debug.Print result ' Returns "Content"

result = GetStringBetween("No delimiters here", "[", "]")
Debug.Print result  ' Returns ""
```
