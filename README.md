# User Defined Functions

This repo contains custom functions I've developed throughout my experience as a programmer.

## Table of Contents

- [VBA](#vba)
    - [`AreArraysEquals`](#arearraysequals)
    - [`AutoFillFormulas`](#autofillformulas)
    - [`CleanString`](#cleanstring)
    - [`DisableRefreshAll`](#disablerefreshall)
    - [`EnableRefreshAll`](#enablerefreshall)
    - [`FileExists`](#fileexists)
    - [`FileNameIsValid`](#filenameisvalid)
    - [`GetAllFileNames`](#getallfilenames)
    - [`GetLettersOnly`]()
    - [`IsAllTrue`](#isalltrue)
    - [`RangeHasAnyFormula`](#rangehasanyformula)
- [Power Query (M Code)](#power-query-m-code)

## VBA

### `AreArraysEquals`

Compares two arrays to check if they are equal, meaning they have the same size and identical elements in the same order.

#### **Syntax**

```vb
AreArraysEqual(Array1 As Variant, Array2 As Variant) As Boolean
```

#### **Parameters**
- `Array1`: First array to compare
- `Array2`: Second array to compare

#### **Return Value**

Returns `True` if both arrays are equal, `False` otherwise.

#### **Remarks**

- Arrays must have the same upper and lower bounds
- Arrays must have identical elements in the same positions
- The function performs an element-by-element comparison
- Returns `False` if arrays have different sizes
- Can compare arrays of any type since parameters are declared as Variant

#### **Example**

```vb
Dim arr1 As Variant
Dim arr2 As Variant
arr1 = Array(1, 2, 3)
arr2 = Array(1, 2, 3)

If AreArraysEqual(arr1, arr2) Then
    Debug.Print "Arrays are equal"
Else
    Debug.Print "Arrays are different"
End If
```

### `AutoFillFormulas`

Automatically fills formulas across a range using a reference cell's formula. The reference cell can be either the first or last cell containing a formula in the range.

#### **Syntax**

```vb
AutoFillFormulas(rng As Range, Optional UseLastCellAsRef As Boolean = False)
```

#### **Parameters**

- `rng`: The range where formulas will be filled
- `UseLastCellAsRef`: (Optional) Boolean flag to determine which cell to use as reference
    - `False` (Default): Uses the first cell with formula as reference
    - `True`: Uses the last cell with formula as reference

#### **Remarks**

- Does nothing if the range is empty (Nothing) or contains only one cell
- Only works if the range contains at least one formula
- Uses R1C1 formula notation to ensure proper relative references when filling
- Only fills formulas in cells that are part of the specified range
- Requires the helper function [`RangeHasAnyFormula`](#rangehasanyformula) to check for formulas in the range

#### **Dependecies**

- Requires `RangeHasAnyFormula` function to work properly

#### **Example**

```vb
Dim rng As Range
Set rng = Range("A1:A10")
AutoFillFormulas rng 'Uses first formula cell as reference

'Or using the last cell as reference:
AutoFillFormulas rng, True
```

### `CleanString`

Cleans a string by removing or replacing special characters and control characters with spaces.

#### **Syntax**

```vb
CleanString(ByVal myString As String, Optional ReplaceBySpace As Boolean = True, Optional ConvertNonBreakingSpace As Boolean = True) As String
```

#### **Parameters**

- `myString`: The input string to be cleaned
- `ReplaceBySpace`: (Optional) Boolean flag that determines if special characters should be replaced by spaces
    - `True` (Default): Replaces special characters with spaces
    - `False`: Removes special characters without replacement
- `ConvertNonBreakingSpace`: (Optional) Boolean flag to handle non-breaking spaces
    - `True` (Default): Converts non-breaking spaces (ASCII 160) to regular spaces
    - `False`: Leaves non-breaking spaces unchanged

#### **Return Value**

Returns the cleaned string with special characters either removed or replaced by spaces.

#### **Remarks**
- Removes ASCII control characters (0-31)
- Handles special characters like ASCII 127, 129, 141, 143, 144, and 157
- Converts non-breaking spaces to regular spaces (when enabled)
- Trims leading and trailing spaces from the final result
- Preserves all other printable characters

#### **Example**

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

### `DisableRefreshAll`

Disables the "Refresh All" functionality for OLEDB connections in a specified workbook.

#### **Syntax**

```vb
DisableRefreshAll(ByRef wb As Workbook)
```

#### **Parameters**

- `wb`: Reference to the workbook where OLEDB connections will be modified

#### **Use Cases**

- Improve performance by preventing unnecessary data refreshes
- Control which connections should be updated during a "Refresh All" operation
- Selectively manage data refresh behavior in workbooks with multiple connections

#### **Remarks**

- Only affects OLEDB type connections
- Does not modify PowerPivot or other connection types
- Changes are applied to each connection individually
- The connections will still be refreshable individually, just not through "Refresh All" option
- Changes are made directly to the workbook passed as parameter

#### **Example**

```vb
Dim wb As Workbook
Set wb = ThisWorkbook
DisableRefreshAll wb
```

### `EnableRefreshAll`

Enables the "Refresh All" functionality for OLEDB connections in a specified workbook.

#### **Syntax**

```vb
EnableRefreshAll(ByRef wb As Workbook)
```

#### **Parameters**

- `wb`: Reference to the workbook where OLEDB connections will be modified

#### **Use Cases**

- Restore default refresh behavior for OLEDB connections
- Enable batch updates of multiple connections
- Ensure all OLEDB connections are included in "Refresh All" operations
- Manage data refresh settings after temporary disablement

#### **Remarks**

- Only affects OLEDB type connections
- Does not modify PowerPivot or other connection types
- Changes are applied to each connection individually
- Allows connections to be updated when using "Refresh All" command
- Changes are made directly to the workbook passed as parameter

#### **Example**

```vb
Dim wb As Workbook
Set wb = ThisWorkbook
EnableRefreshAll wb
```

### `FileExists`

Checks if a file exists at the specified file path.

#### **Syntax**

```vb
FileExists(FilePath As String) As Boolean
```

#### **Parameters**

- `FilePath`: The complete path to the file being checked

#### **Return Value**

Returns `True` if the file exists, `False` otherwise.

#### **Remarks**

- Uses VBA's `Dir` function to test file existence
- Works with any file type
- Path must be accessible from the current environment
- Case-insensitive file path checking

#### **Example**

```vb
Dim exists As Boolean
exists = FileExists("C:\Documents\myfile.xlsx")

If exists Then
    Debug.Print "File exists"
Else
    Debug.Print "File not found"
End If
```

#### **Credits**

- Original source: [www.TheSpreadsheetGuru.com/The-Code-Vault](www.TheSpreadsheetGuru.com/The-Code-Vault)
- Resource: [http://www.rondebruin.nl/win/s9/win003.htm](http://www.rondebruin.nl/win/s9/win003.htm)

### `FileNameIsValid`

Validates if a given string can be used as a valid file name by checking for illegal characters.

#### **Syntax**

```vb
FileNameIsValid(FileName As String) As Boolean
```

#### **Parameters**

- `FileName`: The string to be validated as a file name

#### **Return Value**

Returns `True` if the file name is valid, `False` if it contains illegal characters or is empty.

#### **Remarks**

- Checks for the following illegal characters: `\ / : * ? < > | [ ] "`
- Returns `False` for empty strings
- Case-sensitive validation
- Does not check file name length restrictions
- Does not validate against reserved Windows file names

#### **Example**

```vb
Dim isValid As Boolean

isValid = FileNameIsValid("my_file.txt")
Debug.Print isValid ' True

isValid = FileNameIsValid("file*.txt") 
Debug.Print isValid ' False

isValid = FileNameIsValid("folder/file.txt")
Debug.Print isValid  ' False
```

#### **Credits**

- Author: Jon Peltier
- Source: [www.TheSpreadsheetGuru.com/the-code-vault](www.TheSpreadsheetGuru.com/the-code-vault)

### `GetAllFileNames`

Retrieves an array of all file names from a specified folder and its subfolders, with optional file extension filtering.

#### **Syntax**

```vb
GetAllFileNames(FolderPath As String, Optional fileExt As String) As String()
```

#### **Parameters**

- `FolderPath`: The path to the folder to search in
- `fileExt`: (Optional) File extension to filter results. If omitted, returns all files

#### **Return Value**

Returns a zero-based string array containing all matching file names.

#### **Remarks**

- Recursively searches through all subfolders
- Case-insensitive file extension matching
- Uses `FileSystemObject` for file system operations
- Returns only file names, not full paths
- Extension filter doesn't require the dot prefix
- Empty array if no files are found
- Requires reference to Microsoft Scripting Runtime (or late binding)

#### **Dependencies**

- `Scripting.FileSystemObject` reference

#### **Example**

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

### `GetLettersOnly`

Extracts only ASCII letters (a–z) from a string and returns them in lowercase.

#### **Syntax**

```vb
GetLettersOnly(Text As String) As String
```

#### **Parameters**

- `Text`: The input string to process.

#### **Return Value**

Returns a string containing only the letters a–z (converted to lowercase). Returns an empty string if no ASCII letters are found.

#### **Remarks**

- Filters characters using ASCII range 97–122 (letters a–z).
- Converts characters to lowercase before testing and output.
- Does not preserve original letter case.
- Does not include accented letters, non-Latin characters, or other alphabetic Unicode letters.
- Useful for normalizing or sanitizing input to ASCII letters only.

#### **Example**

```vb
Dim result As String

result = GetLettersOnly("Hello, World! 123")   
Debug.Print result ' "helloworld"

result = GetLettersOnly("Ábç Def")
Debug.Print result ' "def" (accented letters removed)
```

### `IsAllTrue`

Checks if all elements in a boolean array are `True`.

#### **Syntax**

```vb
IsAllTrue(blnArray As Variant) As Boolean
```

#### **Parameters**

- `blnArray`: Array containing boolean values to be checked

#### **Return Value**

Returns `True` if all elements in the array are boolean `True`, otherwise returns `False`.

#### **Use Cases**

- Validating that multiple conditions are all met
- Checking status of multiple boolean flags
- Quality control checks where all criteria must be true

#### **Remarks**

- Returns `False` if any element is not a boolean type
- Returns `False` if any element is `False`
- Early exit when first non-true value is found
- Can handle arrays of any dimension
- Array must be passed as `Variant` type

#### **Example**

```vb
Dim testArray As Variant

testArray = Array(True, True, True)
Debug.Print IsAllTrue(testArray) ' Returns True

testArray = Array(True, False, True)
Debug.Print IsAllTrue(testArray) ' Returns False

testArray = Array(True, "True", True)
Debug.Print IsAllTrue(testArray) ' Returns False (non-boolean element)
```

### `RangeHasAnyFormula`

Checks if a given range contains any cells with formulas.

#### **Syntax**

```vb
RangeHasAnyFormula(ByVal rng As Range) As Boolean
```

#### **Parameters**

- `rng`: The range to be checked for formulas

#### **Return Value**

Returns `True` if the range contains at least one formula, `False` otherwise.

#### **Remarks**

- Returns `False` if the range is Nothing
- Uses error handling to detect the presence of formulas
- Shows an error message if any unexpected error occurs during execution
- Uses Excel's `SpecialCells` method with `xlCellTypeFormulas` to perform the check

#### **Example**

```vb
Dim rng As Range
Set rng = Range("A1:D10")

If RangeHasAnyFormula(rng) Then
    Debug.Print "Range contains at least one formula"
Else
    Debug.Print "Range contains no formulas"
End If
```

#### **Error Handling**

- Displays a message box with error details if an unexpected error occurs
- Properly handles the "No cells were found" error which indicates no formulas are present

## Power Query (M Code)

|Function|Description|
|:-------|:----------|
|`Binary.Unzip`||
|`DateTime.ToUnixTime`||
