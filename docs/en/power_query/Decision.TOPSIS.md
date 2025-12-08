# [`Decision.TOPSIS`](/src/power_query/Decision.TOPSIS.pq)

This function applies the **TOPSIS** (_Technique for Order Preference by Similarity to Ideal Solution_) method to rank alternatives based on multiple criteria. It normalizes the data, applies weights, calculates the distance to the ideal and anti-ideal solutions, and computes a **Closeness Coefficient (CC)** for each alternative. The result is a ranked table of alternatives.

## Syntax

```fs
Decision.TOPSIS(
    tbl as table,
    alternativesColumn as text,
    weights as record
) as table
```

## Parameters

- `table`: The input table containing alternatives and their evaluation across multiple criteria.
- `alternativesColumn`: The name of the column that identifies each alternative.
- `weights`: A record where each field name corresponds to a criterion column in the table, and each field value is the weight (importance) of that criterion.

## Return Value

A ranked table of alternatives ordered by the closeness coeficient (CC).

## Remarks

- The function performs the following steps:
    - Normalization of each criterion using Euclidean norm.
    - Weighting of normalized values using the provided weights.
    - Calculation of Positive Ideal Solution (PIS) and Negative Ideal Solution (NIS).
    - Distance to PIS ($D^{+}$) and Distance to NIS ($D^{-}$) for each alternative.
    - Closeness Coefficient (CC):
        $CC = \frac{D^{-}}{D^{+} + D^{-}}​$
    - Ranking: Alternatives are sorted by CC in descending order. Ties receive the same rank.
- If both $D^{+}$ and $D^{-}$ are zero, CC is set to 0.5.
- This method is widely used in multi-criteria decision analysis (MCDA), especially when criteria have different units or scales.

## Example

Ranking alternatives with three criteria.

```fs
let
    Source = #table(
        {"Alternative", "Cost", "Quality", "Speed"}, {
        {"A", 300, 80, 60},
        {"B", 250, 70, 75},
        {"C", 400, 90, 50}
    }),
    Weights = [Cost = 0.4, Quality = 0.3, Speed = 0.3],
    Result = Decision.TOPSIS(Source, "Alternative", Weights)
in
    Result
```

**Result**

|Alternative|Cost|Quality|Speed|CC  |RANKING|
|:---------:|:--:|:-----:|:---:|:--:|:-----:|
|C          |400 |90     |50   |0.74|1      |
|B          |250 |70     |75   |0.26|2      |
|A          |300 |80     |60   |0.26|3      |

## References

- Hwang, C.L. and Yoon, K. (1981) Multiple Attribute Decision Making: Methods and Applications. Springer-Verlag, New York. http://dx.doi.org/10.1007/978-3-642-48318-9
- Madanchian M, Taherdoost H. A comprehensive guide to the TOPSIS method for multi-criteria decision making. Sustainable Social Development 2023; 1(1): 2220. doi: 10.54517/ssd.v1i1.2220