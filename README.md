# User Defined Functions

This repo contains custom functions I've developed throughout my experience as a programmer.

## Table of Contents

- [Power Query (M Code)](#power-query-m-code)
    - [`Binary.Unzip`](#binaryunzip)
    - [`DateTime.ToUnixTime`](#datetimetounixtime)
    - [`List.Correlation`](#listcorrelation)
- [VBA](#vba)
    - [`AreArraysEquals`](#arearraysequals)
    - [`AutoFillFormulas`](#autofillformulas)
    - [`CleanString`](#cleanstring)
    - [`DisableRefreshAll`](#disablerefreshall)
    - [`EnableRefreshAll`](#enablerefreshall)
    - [`FileExists`](#fileexists)
    - [`FileNameIsValid`](#filenameisvalid)
    - [`GetAllFileNames`](#getallfilenames)
    - [`GetLettersOnly`](#getlettersonly)
    - [`GetMonthNumberFromName`](#getmonthnumberfromname)
    - [`GetStringBetween`](#getstringbetween)
    - [`GetStringWithSubstringInArray`](#getstringwithsubstringinarray)
    - [`GetTableColumnNames`](#gettablecolumnnames)
    - [`IsAllTrue`](#isalltrue)
    - [`IsInArray`](#isinarray)
    - [`ListObjectExists`](#listobjectexists)
    - [`PreviousMonthNumber`](#previousmonthnumber)
    - [`RangeHasAnyFormula`](#rangehasanyformula)
    - [`RangeHasConstantValues`](#rangehasconstantvalues)
    - [`RangeIsHidden`](#rangeishidden)
    - [`RangeToHtml`](#rangetohtml)
    - [`SendEmail`](#sendemail)
    - [`SetQueryFormula`](#setqueryformula)
    - [`StringContains`](#stringcontains)
    - [`StringEndsWith`](#stringendswith)
    - [`StringStartsWith`](#stringstartswith)
    - [`SubstringIsInArray`](#substringisinarray)
    - [`Summation`](#summation)
    - [`TableHasQuery`](#tablehasquery)
    - [`WorksheetHasListObject`](#worksheethaslistobject)

## Power Query (M Code)

### `Binary.Unzip`

Extracts files from a ZIP archive and returns a table of entries with file names and decompressed content.

#### **Syntax**
```m
Binary.Unzip(ZIPFile as binary) as table
```

#### **Parameters**

- `ZIPFile` — A binary containing a ZIP archive (for example, the result of `File.Contents`).

#### **Return value**

A table with the following columns:
- `FileName` (text) — The entry name inside the ZIP.
- `Content` (binary or null) — The decompressed file content; `null` if decompression failed or entry unsupported.

#### **Example**

```m
let
    Source = Binary.Unzip(File.Contents("C:\Temp\archive.zip"))
in
    Source
```

This yields a table you can expand or transform. To read the content of the first file as text:

```m
let
    Files = Binary.Unzip(File.Contents("C:\Temp\archive.zip")),
    FirstBinary = Files{0}[Content],
    FirstText = if FirstBinary <> null then Text.FromBinary(FirstBinary) else null
in
    FirstText
```

#### **Credits**

- Author: Ignacio Barrau
- Source: [ExtractZIP.pq](https://github.com/ibarrau/PowerBi-code/blob/master/PowerQuery/ExtractZIP.pq)

### `DateTime.ToUnixTime`

Converts a Power Query datetime value to Unix time (seconds since 1970-01-01 00:00:00).

#### Syntax
```m
DateTime.ToUnixTime(datetimeToConvert as datetime) as number
```

#### Parameters

- `datetimeToConvert`: A datetime value to convert.

#### Return Value

Returns a number representing the total seconds (can be fractional) between `datetimeToConvert` and the Unix epoch (1970-01-01 00:00:00). Values are negative for datetimes before the epoch.

#### Remarks

- No timezone conversion is performed — treat the input as UTC if you need UTC-based Unix time.

#### Example

```m
let
    UnixSeconds = DateTime.ToUnixTime(#datetime(2023, 1, 1, 0, 0, 0))
in
    UnixSeconds \\ Returns 1672531200
```

### `List.Correlation`

## VBA

### [`AreArraysEquals`](VBA\AreArraysEqual.vba)

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

### [`AutoFillFormulas`](VBA\AutoFillFormulas.vba)

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

### `GetMonthNumberFromName`

Converts a month name to its corresponding numeric value (1-12).

#### **Syntax**

```vb
GetMonthNumberFromName(MonthName As String) As Integer
```

#### **Parameters**

- `MonthName`: The name of the month (full or abbreviated, in any language supported by Excel)

#### **Return Value**

Returns an integer from 1 to 12 representing the month number.

#### **Remarks**

- Works with month names in any language supported by Excel
Accepts both full month names and abbreviated forms
- Case-insensitive
- Returns error if month name is invalid

#### **Example**

```vb
Dim monthNum As Integer

monthNum = GetMonthNumberFromName("January")   
Debug.Print monthNum ' Prints 1

monthNum = GetMonthNumberFromName("Jan")
Debug.Print monthNum ' Prints 1

monthNum = GetMonthNumberFromName("Janeiro")
Debug.Print monthNum ' Prints 1 (Portuguese)

monthNum = GetMonthNumberFromName("Janvier") 
Debug.Print monthNum ' Returns 1 (French)
```

### `GetStringBetween`

Extracts a substring between two specified delimiter strings.

#### **Syntax**

```vb
GetStringBetween(str As String, startStr As String, endStr As String) As String
```

#### **Parameters**

- `str`: The input string to search in
- `startStr`: The starting delimiter string
- `endStr`: The ending delimiter string

#### **Return Value**

Returns the text found between the start and end strings. Returns an empty string if no match is found.

#### **Remarks**

- Uses VBScript RegExp for pattern matching
- Creates RegExp object using late binding to avoid explicit reference requirement
- Case-insensitive search
- Non-greedy matching (returns shortest match)
- Returns only the first match if multiple exist
- Removes the delimiter strings from the result

#### **Example**

```vb
Dim result As String

result = GetStringBetween("Hello [World] Test", "[", "]")
Debug.Print result ' Returns "World"

result = GetStringBetween("<tag>Content</tag>", "<tag>", "</tag>")
Debug.Print result ' Returns "Content"

result = GetStringBetween("No delimiters here", "[", "]")
Debug.Print result  ' Returns ""
```

### `GetStringWithSubstringInArray`

Searches through an array of strings and returns the first string that contains a specified substring.

#### **Syntax**

```vb
GetStringWithSubstringInArray(SubString As String, SourceArray As Variant, Optional CaseSensitive As Boolean = False) As String
```

#### **Parameters**

- `SubString`: The text to search for within each array element
- `SourceArray`: Array containing strings to search through
- `CaseSensitive`: (Optional) Boolean flag to enable case-sensitive search. Default is False

#### **Return Value**

Returns the first string from the array containing the substring. Returns an empty string if no match is found.

#### **Remarks**

- Only processes elements that are strings (type `vbString`)
- Ignores non-string elements in the array
- Case-insensitive by default
- Returns first match found and exits
- Works with arrays of any dimension

#### **Dependencies**

- Requires [`StringContains`]() function

#### **Example**

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

### `GetTableColumnNames`

Returns the header names of an Excel ListObject (table) as a zero-based string array.

#### **Syntax**

```vb
GetTableColumnNames(lo As ListObject) As String()
```

#### **Parameters**

- `lo`: The ListObject (Excel table) to read column headers from

#### **Return value**

Returns a zero-based array of strings containing the table column header values in left-to-right order.

#### **Remarks**

- Includes hidden columns and preserves the table column order.

#### **Example**

```vb
Dim colNames() As String
Dim i          As Long

Set tbl = ThisWorkbook.Worksheets("Sheet1").ListObjects("Table1")
colNames = GetTableColumnNames(tbl)

For i = 0 To UBound(colNames)
    Debug.Print colNames(i)
Next i
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

### `IsInArray`

Checks whether a value exists in a one-dimensional array.

#### **Syntax**

```vb
IsInArray(ValueToBeFound As Variant, SourceArray As Variant) As Boolean
```

#### **Parameters**

- `ValueToBeFound`: The value to search for (any Variant).
- `SourceArray`: The one-dimensional array to search (Variant).

#### **Return Value**

Returns `True` if the value is found in the array, otherwise returns `False`.

#### **Remarks**

- Expects a one-dimensional array; passing an uninitialized or multi-dimensional array may cause errors.

#### **Example**

```vb
Dim arr As Variant
arr = Array("apple", "banana", "cherry")

If IsInArray("banana", arr) Then
    Debug.Print "Found"
Else
    Debug.Print "Not found"
End If
```

### `ListObjectExists`

Checks whether a ListObject (Excel table) with a given name exists in a workbook.

#### **Syntax**

```vb
ListObjectExists(ByRef wb As Workbook, ByVal loName As String) As Boolean
```

#### **Parameters**

- `wb`: Workbook to search.
- `loName`: Name of the table (`ListObject`) to find.

#### **Return Value**

Returns `True` if a ListObject with the specified name is found in any worksheet of the workbook; otherwise returns `False`.

#### **Remarks**

- Performs a direct name comparison (behavior may be affected by the project's Option Compare setting).

#### **Example**

```vb
Dim exists As Boolean
exists = ListObjectExists(ThisWorkbook, "Table1")

If exists Then
    Debug.Print "Table exists"
Else
    Debug.Print "Table not found"
End If
```

### `PreviousMonthNumber`

Returns the numeric month (1–12) that precedes the month of a given date.

#### **Syntax**

```vb
PreviousMonthNumber(dt As Date) As Integer
```

#### **Parameters**

- `dt`: Date value used to determine the previous month

#### **Return Value**

Returns an Integer from 1 to 12 representing the previous month. For dates in January, returns 12 (December).

#### **Example**

```vb
Dim prev As Integer

prev = PreviousMonthNumber(DateSerial(2025, 3, 15))
Debug.Print prev ' returns 2 (February)

prev = PreviousMonthNumber(DateSerial(2025, 1, 10))
Debug.Print prev ' returns 12 (December)
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

### `RangeHasConstantValues`

Checks whether a given range contains any constant (non-formula) cells.

#### **Syntax**

```vb
RangeHasConstantValues(rng As Range) As Boolean
```

#### **Parameters**

- `rng`: Range to check for constant values.

#### **Return Value**

Returns `True` if the range contains at least one constant cell; otherwise returns False. If `rng` is `Nothing` the function returns `False`.

#### **Example**

```vb
Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Sheet1").Range("A1:A10")

If RangeHasConstantValues(rng) Then
    Debug.Print "Range contains constants"
Else
    Debug.Print "Range contains no constants or is invalid"
End If
```

### `RangeIsHidden`

Determines whether a given range is entirely hidden (no visible cells).

#### **Syntax**

```vb
RangeIsHidden(rng As Range) As Boolean
```

#### **Parameters**

- `rng`: The Range to check for visibility.

#### **Return Value**

Returns `True` if the range contains no visible cells (i.e., is hidden). Returns `False` if at least one cell in the range is visible or if `rng` is `Nothing`.

#### **Example**

```vb
Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Sheet1").Range("A1:A10")

If RangeIsHidden(rng) Then
    Debug.Print "Range is hidden (no visible cells)."
Else
    Debug.Print "Range has visible cells."
End If
```

### `RangeToHtml`

Converts an Excel Range into an HTML string by copying the range to a temporary workbook, publishing that sheet as an HTML file, and returning the file contents.

#### **Syntax**

```vba
RangeToHtml(rng As Range) As String
```

#### **Parameters**

- `rng`: The Range to convert to HTML.

#### **Return Value**

Returns a string containing the HTML representation of the provided range. Returns an empty string if an error occurs.

#### **Remarks**

- Creates a temporary workbook, pastes the range (values and formats) and removes drawing objects before publishing.
- Uses the system temporary folder (Environ$("temp")) to create an intermediate .htm file.
- Reads the generated HTML file into memory and deletes the temporary file and workbook.
- Replaces `align=center` with `align=left` in the resulting HTML.
- Images/drawing objects are deleted in the temporary workbook to avoid embedding them in the HTML.

#### **Example**

```vba
Dim html As String
html = RangeToHtml(ThisWorkbook.Worksheets("Sheet1").Range("A1:D10"))
' html now contains the HTML representation of the range
```

### `SendEmail`

Sends an HTML email using CDO (Collaboration Data Objects) with NTLM authentication, typically used in corporate environments with Exchange Server.

#### **Syntax**

```vb
SendEmail( _
    Sender As String, _
    Recipient As String, _
    Subject As String, _
    Message As String, _
    Optional CarbonCopy As String, _
    Optional BlindCarbonCopy As String)
```

#### **Parameters**

- `Sender`: Email address of the sender
- `Recipient`: Email address(es) of the recipient(s)
- `Subject`: Subject line of the email
- `Message`: HTML-formatted body of the email
- `CarbonCopy`: (Optional) Email address(es) for CC recipients
- `BlindCarbonCopy`: (Optional) Email address(es) for BCC recipients

#### **Remarks**

- Uses CDO with NTLM authentication (Windows Authentication)
- Configured for SMTP with STARTTLS (port 587)
- Supports HTML formatting in the message body
- Multiple recipients can be specified using semicolon (;) as separator
- No explicit error handling is implemented

#### **Configuration Constants**

- `CDO_DEFAULT_SETTINGS`: -1 (Use system default settings)
- `CDO_NTLM_AUTHENTICATION`: 2 (Windows Authentication)
- `CDO_SEND_USING_PORT`: 2 (Direct SMTP)
- `CDO_SERVER_PORT`: 587 (STARTTLS port)
- `CDO_SMTP_SERVER`: "mailhost.yourdomain.net" (SMTP server address)

#### **Dependencies**

- Requires CDO to be available on the system
- Requires proper SMTP server configuration
- Requires appropriate network/firewall access

#### **Example**

```vb
Call SendEmail( _
    "sender@company.com", _
    "recipient@company.com", _
    "Test Subject", _
    "<h1>Hello</h1><p>This is a test email.</p>", _
    "cc@company.com", _
    "bcc@company.com")
```

### `SetQueryFormula`

Modifies a Power Query formula in the current workbook based on a given value, handling different data types appropriately.

#### **Syntax**

```vb
SetQueryFormula(queryName As String, value As Variant)
```

#### **Parameters**

- `queryName`: Name of the Power Query to modify
- `value`: Value to set in the query formula (supports `String`, `Date`, and `Byte Array`)

#### **Dependencies**

- Requires Excel version that supports Power Query

#### **Example**

```vb
' Set a string value
SetQueryFormula "MyQuery", "Hello ""World"""  ' Results in: "Hello ""World"""

' Set a date value
SetQueryFormula "MyQuery", DateSerial(2023, 10, 17)  ' Results in: #date(2023,10,17)

' Set a byte array
Dim byteArr() As Byte
byteArr = Array(1, 2, 3)
SetQueryFormula "MyQuery", byteArr  ' Results in: {1,2,3}
```

### `StringContains`

Checks if a string contains another string as a substring, with optional case sensitivity.

#### **Syntax**

```vb
StringContains(str1 As String, str2 As String, Optional caseSensitive As Boolean = False) As Boolean
```

#### **Parameters**

- `str1`: The main string to search in
- `str2`: The substring to search for
- `caseSensitive`: (Optional) Boolean flag to enable case-sensitive search. Default is `False`

#### **Return Value**

Returns `True` if `str2` is found within `str1`, `False` otherwise.

#### **Use Cases**

- Text validation
- String searching
- Pattern matching without regular expressions
- Case-insensitive text comparisons

#### **Example**

```vb
Dim result As Boolean

result = StringContains("Hello World", "world")
Debug.Print result ' Returns True

result = StringContains("Hello World", "WORLD")
Debug.Print result ' Returns True

result = StringContains("Hello World", "world", True)
Debug.Print result ' Returns False

result = StringContains("Test", "xyz")
Debug.Print result ' Returns False
```

### `StringEndsWith`

Checks if a string ends with another string, with optional case sensitivity.

#### **Syntax**

```vb
StringEndsWith(str1 As String, str2 As String, Optional caseSensitive As Boolean = False) As Boolean
```

#### **Parameters**

- `str1`: The main string to check
- `str2`: The ending string to look for
- `caseSensitive`: (Optional) Boolean flag to enable case-sensitive comparison. Default is `False`

#### **Return Value**

Returns `True` if `str1` ends with `str2`, `False` otherwise. Also returns `False` if `str2` is longer than `str1`.

#### **Use Cases**

- File extension validation
- Text suffix checking
- String pattern matching
- Domain name validation

#### **Example**

```vb
Dim result As Boolean

result = StringEndsWith("Hello World", "world")
Debug.Print result ' Returns True

result = StringEndsWith("Hello World", "WORLD")
Debug.Print result ' Returns True

result = StringEndsWith("Hello World", "World", True)
Debug.Print result ' Returns True

result = StringEndsWith("Test", "xyz")
Debug.Print result ' Returns False
```

### `StringStartsWith`

Checks whether a string starts with a specified substring, with optional case sensitivity.

#### **Syntax**

```vb
StringStartsWith(str1 As String, str2 As String, Optional caseSensitive As Boolean = False) As Boolean
```

#### **Parameters**

- `str1`: The main string to check.
- `str2`: The prefix substring to look for.
- `caseSensitive`: (Optional) If `True`, comparison is case-sensitive; if `False` (default), comparison is case-insensitive.

#### **Return Value**

Returns `True` if `str1` starts with `str2`; otherwise returns `False`. Also returns `False` if `str2` is longer than `str1`.

#### **Example**

```vb
Dim result As Boolean

result = StringStartsWith("Report.xlsx", "Report")
Debug.Print result ' True

result = StringStartsWith("Report.xlsx", "report")
Debug.Print result ' True (case-insensitive)

result = StringStartsWith("Report.xlsx", "report", True)
Debug.Print result ' False (case-sensitive)

result = StringStartsWith("Test", "LongPrefix")
Debug.Print result ' False
```

### `SubstringIsInArray`

Searches a one-dimensional array for any string element that contains a specified substring and returns `True` on the first match.

#### **Syntax**

```vb
StringStartsWith(str1 As String, str2 As String, Optional caseSensitive As Boolean = False) As Boolean
```

#### **Parameters**

- `subStr`: The substring to search for.
- `srcArray`: One-dimensional array containing elements to search.
- `caseSensitive`: (Optional) If `True`, performs a case-sensitive search; default is `False`.

#### **Return Value**

Returns `True` if any string element in `srcArray` contains `subStr`; otherwise returns `False`.

#### **Remarks**

- Only inspects elements typed as `String`; non-string elements are ignored.

#### **Dependencies**

- Depends on the helper function [`StringContains`](#stringcontains) for substring checks.

#### **Example**

```vb
Dim arr As Variant
arr = Array("Hello World", "Sample", "Test")

Debug.Print SubstringIsInArray("world", arr)        ' True (case-insensitive)
Debug.Print SubstringIsInArray("WORLD", arr, True) ' False (case-sensitive)
```

### `Summation`

Computes the numeric summation of a mathematical expression over an integer index range.

#### **Syntax**

```vb
Summation(Expression As String, First As Long, Last As Long) As Double
```

#### **Parameters**

- `Expression`: A string representing the math expression in terms of a variable (e.g. `"2*n-1"` or `"1/x^2"`). The function extracts the variable name as the last alphabetical character found in the expression.
- `First`: Starting integer index.
- `Last`: Ending integer index.

#### **Return Value**

Returns the summation's result from expression evaluated for the index running from `First` to `Last`.

#### **Remarks**

- The variable used in Expression is determined by extracting letters from the expression and taking the last letter. Ensure your expression contains the intended variable and that it is the last letter in the expression if multiple letters appear

#### **Dependecies**

- Depends on the helper function [`GetLettersOnly`](#getlettersonly) in order to identify the variable in expression

#### **Example**

```vb
Debug.Print Summation("2*n-1", 1, 10) ' prints 100
Debug.Print Summation("1/x^2", 1, 1000000) ' ≈ 1.64 (approaches π²/6)
Debug.Print Summation("n^2", 1, 5) ' prints 55
```

### `TableHasQuery`

Checks whether a ListObject (Excel table) has an associated QueryTable.

#### **Syntax**

```vb
TableHasQuery(tbl As ListObject) As Boolean
```

#### **Parameters**

- `tbl`: The ListObject (table) to check.

#### **Return Value**

Returns `True` if the table has an associated `QueryTable`; otherwise returns `False`. If `tbl` is `Nothing`, the function returns `False`.

#### **Example**

```vb
Dim tbl As ListObject
Dim hasQuery As Boolean

Set tbl = ThisWorkbook.Worksheets("Sheet1").ListObjects("Table1")
hasQuery = TableHasQuery(tbl)

If hasQuery Then
    Debug.Print "Table has a QueryTable"
Else
    Debug.Print "Table does not have a QueryTable"
End If
```

### `WorksheetHasListObject`

Checks whether a worksheet contains at least one ListObject (table).

#### **Syntax**

```vb
WorksheetHasListObject(ws As Worksheet) As Boolean
```

#### **Parameters**

- `ws`: Worksheet to check for ListObjects.

#### **Return Value**

Returns `True` if the worksheet contains one or more `ListObjects`; otherwise returns `False`.

#### **Example**

```vb
Dim hasTable As Boolean
hasTable = WorksheetHasListObject(ThisWorkbook.Worksheets("Sheet1"))

If hasTable Then
    Debug.Print "Sheet1 contains at least one table."
Else
    Debug.Print "Sheet1 contains no tables."
End If
```