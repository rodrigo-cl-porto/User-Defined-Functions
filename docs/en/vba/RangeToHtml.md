# [`RangeToHtml`](/src/vba/RangeToHtml.vba)

Converts an range into an HTML string by copying the range to a temporary workbook, publishing that sheet as an HTML file, and returning the file contents.

## Syntax

```vb
RangeToHtml( _
    rng As Range _
) As String
```

## Parameters

- `rng`: The Range to convert to HTML.

## Return Value

Returns a string containing the HTML representation of the provided range. Returns an empty string if an error occurs.

## Remarks

- Creates a temporary workbook, pastes the range (values and formats) and removes drawing objects before publishing.
- Uses the system temporary folder (Environ$("temp")) to create an intermediate .htm file.
- Reads the generated HTML file into memory and deletes the temporary file and workbook.
- Replaces `align=center` with `align=left` in the resulting HTML.
- Images/drawing objects are deleted in the temporary workbook to avoid embedding them in the HTML.

## Example

```vb
Dim html As String
html = RangeToHtml(ThisWorkbook.Worksheets("Sheet1").Range("A1:D10")) ' HTML representation of the range
```

## Credits

- Website: [Ron de Bruin Excel Automation](https://jkp-ads.com/rdb/win/s1/outlook/bmail2.htm)
    - Author: Ron de Bruin
    - Date: 2006, Oct. 28th
