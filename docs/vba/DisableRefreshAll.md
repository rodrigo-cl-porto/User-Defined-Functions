# [`DisableRefreshAll`](/src/vba/DisableRefreshAll.vba)

Disables the "Refresh All" functionality for OLEDB connections in a specified workbook.

## Syntax

```vb
DisableRefreshAll( _
    ByRef wb As Workbook _
)
```

## Parameters

- `wb`: Reference to the workbook where OLEDB connections will be modified

## **Use Cases**

- Improve performance by preventing unnecessary data refreshes
- Control which connections should be updated during a "Refresh All" operation
- Selectively manage data refresh behavior in workbooks with multiple connections

## Remarks

- Only affects OLEDB type connections
- Does not modify PowerPivot or other connection types
- Changes are applied to each connection individually
- The connections will still be refreshable individually, just not through "Refresh All" option
- Changes are made directly to the workbook passed as parameter

## Example

```vb
Dim wb As Workbook
Set wb = ThisWorkbook
DisableRefreshAll wb
```
