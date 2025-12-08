# [`List.WeightedAverage`](/src/power_query/List.WeightedAverage.pq)

Calculates the weighted average of a list of values given a corresponding list of weights.

## Syntax

```fs
List.WeightedAverage(
    values as list,
    weights as list
) as nullable number
```

## Parameters

- `values`: A list of numerical values to calculate the weighted average.
- `weights`: A list of numerical weights corresponding to each value.

## Return Value

Returns a number representing the weighted average of the input values. If the lists have different lengths or if the sum of weights is zero, the function returns `null`.

## Remarks

- The function calculates the weighted average using the formula: $\text{Weighted Average} = \frac{\sum (x_i \times w_i)}{\sum w_i}$
    - $x_i$ are the individual values
    - $w_i$ are the corresponding weights

## Example

Calculing the weighted average for a list with specified weights.

```fs
List.WeightedAverage({1, 2, 3}, {4, 5, 6})
```

**Result**

```fs
2.3333333333333335
```