# [`GetAllFileNames`](/src/vba/GetAllFileNames.vba)

Retrieves an array of all file names from a specified folder and its subfolders, with optional file extension filtering.

## Syntax

```vb
GetAllFileNames( _
    FolderPath As String, _
    Optional fileExt As String _
) As String()
```

## Parameters

- `FolderPath`: The path to the folder to search in
- `fileExt`: (_optional_) File extension to filter results. If omitted, returns all files

## Return Value

Returns a zero-based string array containing all matching file names.

## Remarks

- Recursively searches through all subfolders
- Case-insensitive file extension matching
- Uses `FileSystemObject` for file system operations
- Returns only file names, not full paths
- Extension filter doesn't require the dot prefix
- Empty array if no files are found
- Requires reference to Microsoft Scripting Runtime (or late binding)

## Dependencies

- `Scripting.FileSystemObject` reference

## Example

```vb
Dim files() As String
Dim i As Long

' Get all Excel files
files = GetAllFiles("C:\Documents", "xlsx")

' Get all files regardless of extension
files = GetAllFiles("C:\Documents")

' Print all found files
For i = 0 To UBound(files)
    Debug.Print files(i)
Next i
```
