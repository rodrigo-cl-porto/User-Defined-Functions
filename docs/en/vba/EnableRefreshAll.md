# [`EnableRefreshAll`](/src/vba/EnableRefreshAll.vba)

Enables the "Refresh All" functionality for OLEDB connections in a specified workbook.

## Syntax

```vb
EnableRefreshAll( _
    ByRef wb As Workbook _
)
```

## Parameters

- `wb`: Reference to the workbook where OLEDB connections will be modified

## **Use Cases**

- Restore default refresh behavior for OLEDB connections
- Enable batch updates of multiple connections
- Ensure all OLEDB connections are included in "Refresh All" operations
- Manage data refresh settings after temporary disablement

## Remarks

- Only affects OLEDB type connections
- Does not modify PowerPivot or other connection types
- Changes are applied to each connection individually
- Allows connections to be updated when using "Refresh All" command
- Changes are made directly to the workbook passed as parameter

## Example

```vb
Dim wb As Workbook
Set wb = ThisWorkbook
EnableRefreshAll wb
```
