# [`FileExists`](/src/vba/FileExists.vba)

Checks if a file exists at the specified file path.

## Syntax

```vb
FileExists( _
    FilePath As String _
) As Boolean
```

## Parameters

- `FilePath`: The complete path to the file being checked

## Return Value

Returns `True` if the file exists, `False` otherwise.

## Remarks

- Uses VBA's `Dir` function to test file existence
- Works with any file type
- Path must be accessible from the current environment
- Case-insensitive file path checking

## Example

```vb
Dim exists As Boolean
exists = FileExists("C:\Documents\myfile.xlsx")

If exists Then
    Debug.Print "File exists"
Else
    Debug.Print "File not found"
End If
```

## **Credits**

- Original source: [www.TheSpreadsheetGuru.com/The-Code-Vault](www.TheSpreadsheetGuru.com/The-Code-Vault)
- Resource: [http://www.rondebruin.nl/win/s9/win003.htm](http://www.rondebruin.nl/win/s9/win003.htm)
