# [`FileNameIsValid`](/src/vba/FileNameIsValid.vba)

Checks if a given string can be used as a valid file name by checking for illegal characters.

## Syntax

```vb
FileNameIsValid( _
    FileName As String _
) As Boolean
```

## Parameters

- `FileName`: The string to be validated as a file name

## Return Value

Returns `True` if the file name is valid, `False` if it contains illegal characters or is empty.

## Remarks

- Checks for the following illegal characters: `\ / : * ? < > | [ ] "`
- Returns `False` for empty strings
- Case-sensitive validation
- Does not check file name length restrictions
- Does not validate against reserved Windows file names

## Example

```vb
Dim isValid As Boolean

isValid = FileNameIsValid("my_file.txt")
Debug.Print isValid ' True

isValid = FileNameIsValid("file*.txt") 
Debug.Print isValid ' False

isValid = FileNameIsValid("folder/file.txt")
Debug.Print isValid  ' False
```

## **Credits**

- Author: Jon Peltier
- Source: [www.TheSpreadsheetGuru.com/the-code-vault](www.TheSpreadsheetGuru.com/the-code-vault)
