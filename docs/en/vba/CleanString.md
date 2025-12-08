# [`CleanString`](/src/vba/CleanString.vba)

Cleans a string by removing or replacing special characters and control characters with spaces.

## Syntax

```vb
CleanString( _
    myString As String, _
    Optional ReplaceBySpace As Boolean = True, _
    Optional ConvertNonBreakingSpace As Boolean = True _
) As String
```

## Parameters

- `myString`: The input string to be cleaned
- `ReplaceBySpace`: (_optional_) Boolean flag that determines if special characters should be replaced by spaces
    - `True` (Default): Replaces special characters with spaces
    - `False`: Removes special characters without replacement
- `ConvertNonBreakingSpace`: (_optional_) Boolean flag to handle non-breaking spaces
    - `True` (Default): Converts non-breaking spaces (ASCII 160) to regular spaces
    - `False`: Leaves non-breaking spaces unchanged

## Return Value

Returns the cleaned string with special characters either removed or replaced by spaces.

## Remarks
- Removes ASCII control characters (0-31)
- Handles special characters like ASCII 127, 129, 141, 143, 144, and 157
- Converts non-breaking spaces to regular spaces (when enabled)
- Trims leading and trailing spaces from the final result
- Preserves all other printable characters

## Example

```vb
Dim cleanedStr As String

' Replace special characters with spaces
cleanedStr = CleanString("Hello" & Chr(0) & "World")
Debug.Print cleanedStr ' Result: "Hello World"

' Remove special characters
cleanedStr = CleanString("Hello" & Chr(0) & "World", False)
Debug.Print cleanedStr ' Result: "HelloWorld"

' Keep non-breaking spaces
cleanedStr = CleanString("Hello" & Chr(160) & "World", True, False)
Debug.Print cleanedStr ' Result: Original string unchanged
```
