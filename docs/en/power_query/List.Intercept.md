# [`List.Intercept`](/src/power_query/List.Intercept.pq)

Calculates the intercept of the linear regression line between two numerical lists X and Y.

## Syntax

```fs
List.Intercept(
    X as list,
    Y as list
) as number
```

## Parameters

- `X`: A list of numerical values representing the independent variable.
- `Y`: A list of numerical values representing the dependent variable.

## Return Value

Returns a number representing the intercept of the linear regression line calculated using the least squares method. If the lists have different lengths, the function returns `null`.

## Remarks

- Both input lists must have the same length; otherwise, the function returns `null`.
- The function uses the least squares method to calculate the intercept.
- Non-numeric values in the lists will cause an error during calculation.

## Example

**Example 1**: Calculating the intercept value of linear regression.

```fs
List.Intercept({1, 2, 3}, {4, 5, 6})
```

**Result**

```fs
-1
```

**Example 2**: Returns `null` if the input lists have different lengths.

```fs
List.Intercept({1, 2}, {3})
```

**Result**

```fs
null
```