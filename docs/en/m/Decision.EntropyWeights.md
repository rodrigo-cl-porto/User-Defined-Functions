# [`Decision.EntropyWeights`](/src/m/Decision.EntropyWeights.pq)

This function calculates the weights of decision criteria using the **entropy weighting method**. It analyzes the distribution of values across each criterion to determine their relative importance based on information entropy.

## Syntax

```fs
Decision.EntropyWeights(
    tbl as table,
    optional columnNames as {text}
) as record
```

## Parameters

- `tbl`: The input table containing numeric values for each criterion.
- `columnNames` (_optional_): A list of column names to include in the calculation. If not provided, all numeric columns in the table will be used.

## Remarks

- The entropy weighting method is a data-driven approach to determine the importance of each criterion based on its variability.
- The function performs the following steps:
    - Normalization of each criterion using min-max scaling.
    - Calculation of entropy for each criterion based on the normalized values.
    - Computation of weights:
        - $w_{j}=\frac{1−s_{j}}{m−\sum_{j=1}^{m}{s_{j}}}​$
    - where:
        - $s_j$​ is the entropy of criterion $j$,
        - $m$ is the number of criteria.
- If all values in a column are equal, its entropy is the maximum and its weight is zero.
- If the table has only one row, all criteria receive equal weight.

## Return Value

Returns a record where each field name corresponds to a criterion and each field value is the calculated weight.

## Examples

**Example 1**: Entropy weights for all numeric columns. Weights are assigned based on the variability of each criterion.

```fs
let
    Source = #table(type table
        [Cost=number, Quality=number, Speed=number], {
        {300, 80, 60},
        {250, 70, 75},
        {400, 90, 50}
    }),
    Result = Decision.EntropyWeights(Source)
in
    Result
```

**Result**

```fs
[Cost = 0.36, Quality = 0.31, Speed = 0.33]
```

**Example 2**: Calculates entropy weights only for specified columns.

```fs
let
    Source = #table(type table
        [Cost=number, Quality=number, Speed=number], {
        {300, 80, 60},
        {250, 70, 75},
        {400, 90, 50}
    }),
    Result = Decision.EntropyWeights(Source, {"Cost", "Speed"})
in
    Result
```

**Result**

```fs
[Cost = 0.52, Speed = 0.48]
```