# [`List.Correlation`](/src/m/List.Correlation.pq)

Calculates the correlation coefficient between two lists of numeric values. Supports Pearson (linear) and Spearman (rank-based) correlation.

## Syntax

```fs
List.Correlation(
    list1 as list,
    list2 as list,
    optional typeCorrelation as text
) as number
```

## Parameters

- `list1`: list of numeric values (nulls and non-numeric values are treated as 0).
- `list2`: list of numeric values (nulls and non-numeric values are treated as 0).
- `typeCorrelation` (_optional_): "Pearson" (default) or "Spearman". Case-insensitive.

## Return Value

A number representing the correlation coefficient:

- Pearson: standard Pearson correlation (linear relationship).
- Spearman: Spearman rank correlation (uses dense ranking; tied values receive the same rank).

## Remarks

- Input lists must be the same length; otherwise, an error is raised.
- Null, empty string, or non-numeric entries are converted to 0 before calculation.
- Result is returned as a decimal number (can be negative, positive, or `NaN` if degenerate).

## Examples

**Example 1**: Calculates the default Pearson Correlation.

```fs
List.Correlation({0, 1, 3, 4}, {4, 5, 10, 30})
```

**Result**

```fs
0.858575902776297  (Pearson, default)
```

**Example 2**: Calculates the Spearman (monotonic) Correlation if specified.

```fs
List.Correlation({0, 1, 3, 4}, {4, 5, 10, 30}, "Spearman")
```

**Result**

```fs
1
```

**Example 3**: Non-numeric values are treated as 0.

```fs
List.Correlation({0, null, 3, "a", 4}, {4, 5, null, 10, 30})
```

**Result**

```fs
0.556720639738652
```