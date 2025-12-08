# [`List.Slope`](/src/power_query/List.Slope.pq)

Calculates the slope of the linear regression between two numerical lists X and Y.

## Syntax

```fs
List.Slope(
    X as list,
    Y as list
) as nullable number
```

## Parameters

- `X`: A list of numerical values representing the independent variable.
- `Y`: A list of numerical values representing the dependent variable.

## Return Value

Returns a number representing the slope of the linear regression line calculated using the least squares method. If the lists have different lengths, the function returns `null`.

## Remarks

- Both input lists must have the same length; otherwise, the function returns `null`.
- The function uses the least squares method to calculate the slope.
- Non-numeric values in the lists will cause an error during calculation.

## Examples

**Example 1**: Calculates the slope of linear regression.

```fs
List.Slope({1, 2, 3}, {4, 5, 6})
```

**Result**

```fs
1
```

**Example 2**: Returns `null` if input lists have different lenghts.

```fs
List.Slope({1, 2}, {3})
```

**Result**

```fs
null
```