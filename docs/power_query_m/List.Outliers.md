# [`List.Outliers`](/src/m/List.Outliers.pq)

Identifies outliers in a list of numerical values using the Interquartile Range (IQR) method.

## Syntax

```fs
List.Outliers(
    values as list,
    optional multiplier as number
) as list
```

## Parameters

- `values`: A list of numerical values to analyze for outliers.
- `multiplier` (_optional_): A number to adjust the IQR threshold for defining outliers. Default is 1.5.

## Return Value

Returns a list of outlier values identified in the input list based on the IQR method. If no outliers are found, the function returns an empty list.

## Remarks

- The function first removes nulls, empty strings, and whitespace entries, then selects only valid numeric values.
- Outliers are defined as values below $Q_{1} - 1.5 \cdot IQR$ or above $Q_{3} + 1.5 \cdot IQR$, where $Q_1$ and $Q_3$ are the first and third quartiles respectively.

## Examples

**Example 1**: Returns outliers from a list with extreme values.

```fs
List.Outliers({1, 2, 3, 4, 5, 6, 50, 100})
```

**Result**

```fs
{50, 100}
```

**Example 2**: With a higher multiplier, identifying outliers becomes stricter.

```fs
List.Outliers({1, 2, 3, 4, 5, 6, 50, 100}, 3)
```

**Result**

```fs
{100}
```

**Example 3**: Returns an empty list if there's no outlier.

```fs
List.Outliers({10, 12, 14, 15, 16, 18, 20})
```

**Result**

```fs
{}
```

**Example 4**: Ignores nulls and empty strings.

```fs
List.Outliers({1, null, "", 2, 3, 4, 5, null, 6, 50, 100})
```

**Result**

```fs
{50, 100}
```