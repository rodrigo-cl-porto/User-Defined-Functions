# [`GetLetters`](/src/vba/GetLetters.vba)

Extracts only ASCII letters (a–z) from a string and returns them in lowercase.

## Syntax

```vb
GetLetters( _
    Text As String _
) As String
```

## Parameters

- `Text`: The input string to process.

## Return Value

Returns a string containing only the letters a–z (converted to lowercase). Returns an empty string if no ASCII letters are found.

## Remarks

- Filters characters using ASCII range 97–122 (letters a–z).
- Converts characters to lowercase before testing and Result.
- Does not preserve original letter case.
- Does not include accented letters, non-Latin characters, or other alphabetic Unicode letters.
- Useful for normalizing or sanitizing input to ASCII letters only.

## Example

```vb
Dim result As String

result = GetLetters("Hello, World! 123")   
Debug.Print result ' "helloworld"

result = GetLetters("Ábç Def")
Debug.Print result ' "def" (accented letters removed)
```
