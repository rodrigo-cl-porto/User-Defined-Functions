# [`List.Variance`](/src/m/List.Variance.pq)

Calculates the population variance of a list of numerical values.

## Syntax

```fs
List.Variance(
    values as {number}
) as nullable number
```

## Parameters

- `values`: A list of numerical values to calculate the population variance.

## Return Value

Returns a number representing the population variance of the input list. If the list is empty or contains no numeric values, the function returns `null`.

## Remarks

- The function calculates the population variance using the formula: $\sigma^{2} = \frac{1}{N} \sum_{i=1}^{N} (x_i - \mu)^2$
    - $N$ is the number of values
    - $x_i$ are the individual values
    - $\mu$ is the mean of the values
- Non-numeric values, nulls, and empty strings are ignored in the calculation.

## Examples

**Example 1**: Calculating sample variance value.

```fs
List.Variance({1, 2, 3, 4, 5})
```

**Result**

```fs
2.5
```

**Example 2**: Calculating population variance value.

```fs
List.Variance({1, 2, 3, 4, 5}, true)
```

**Result**

```fs
2
```

**Example 3**: Returns `null` if receives a empty list.

```fs
List.Variance({})
```

**Result**

```fs
null
```