# [`Table.CorrelationMatrix`](/src/power_query/Table.CorrelationMatrix.pq)

Calculates the correlation matrix for a given table. It computes the Pearson correlation coefficient between each pair of numeric columns, returning a table where each cell represents the correlation between two variables.

## Syntax

```fs
Table.CorrelationMatrix(
    tbl as table,
    optional columnNames as type {text}
) as table
```

## Parameters

- `tbl`: The input table containing numeric columns to be analyzed.
- `columnNames` (_optional_): A list of column names to include in the correlation matrix. If not provided, all numeric columns in the table will be used.

## Return Value

The result is a symmetric matrix with correlation values ranging from -1 to 1. The output table includes a "VARIABLE" column indicating the row variable, followed by columns representing correlations with other variables.

## Remarks

- The function uses the Pearson correlation formula to measure linear relationships between columns.
- Nulls, empty strings, and non-numeric values are treated as 0 during computation.

## Examples

**Example 1**: Correlation matrix for all numeric columns.

```fs
let
    Source = #table(
        {"A", "B", "C"}, {
        {1, 2, 3},
        {2, 4, 6},
        {3, 6, 9}
    }),
    Result = Table.CorrelationMatrix(Source)
in
    Result
```

**Result**

|VARIABLE|A  |B  |C  |
|:------:|:-:|:-:|:-:|
|A       |1.0|1.0|1.0|
|B       |1.0|1.0|1.0|
|C       |1.0|1.0|1.0|

**Example 2**: Correlation matrix for selected columns.

```fs
let
    Source = #table(
        {"X", "Y", "Z"}, {
        {1, 10, 100},
        {2, 20, 80},
        {3, 30, 60},
        {4, 40, 40}
    }),
    Result = Table.CorrelationMatrix(Source, {"X", "Z"})
in
    Result
```

**Result**

|VARIABLE|X   |Z   |
|:------:|:--:|:--:|
|X       | 1.0|-1.0|
|Z       |-1.0| 1.0|
