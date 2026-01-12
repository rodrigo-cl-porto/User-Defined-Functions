# [`List.PopulationStdDev`](/src/m/List.PopulationStdDev.pq)

Calculates the population standard deviation of a list of numerical values.

## Syntax

```fs
List.PopulationStdDev(
    values as list
) as nullable number
```

## Parameters

- `values`: A list of numerical values to calculate the population standard deviation.

## Return Value

Returns a number representing the population standard deviation of the input list. If the list is empty or contains no numeric values, the function returns `null`.

## Remarks

- The function calculates the population standard deviation using the formula:
    - $\sigma = \sqrt{\frac{1}{N} \sum_{i=1}^{N} (x_i - \mu)^2}$
    - where:
        - $N$ is the number of values
        - $x_i$ are the individual values
        - $\mu$ is the mean of the values
- Non-numeric values, nulls, and empty strings are ignored in the calculation.

## Examples

**Example 1**: Calculates population standard deviation for a numeric list.

```fs
List.PopulationStdDev({2, 4, 4, 4, 5, 5, 7, 9})
```

**Result**

```fs
2.8284271247461903
```

**Example 2**: Returns `null` for empty lists.

```fs
List.PopulationStdDev({})
```

**Result**

```fs
null
```