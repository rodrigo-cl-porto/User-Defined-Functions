# [`List.Rank`](/src/m/List.Rank.pq)

Returns a list of ranks for a given list of values. Tied values receive the same rank (dense ranking). The Result list preserves the input order.

## Syntax

```fs
List.Rank(
    values as list,
    optional order as Order.Type
) as list
```

## Parameters
- `values`: A list of values to rank. Values must be comparable (numbers, texts, dates, etc.).
- `order` (_optional_): Use `Order.Ascending` or `Order.Descending`. If omitted, the function treats the ordering as descending (i.e., highest value gets rank 1).

## Return Value
A list of integers with the same length as `values`, where each element is the rank (1-based) of the corresponding input value.

## Remarks
- Ranks are "dense": equal values receive the same rank and the next distinct value's rank increases by 1.
  - Example (descending default): values {3,1,2,3} → ranks {1,3,2,1}
- The function returns ranks in the original input order.
- Comparison uses Power Query's Value.Compare, so mixed-type comparisons follow Power Query rules.

## Examples

**Example 1**: Returns the descending rank list.

```fs
List.Rank({10, 10, 5, 7})
```

**Result**

```fs
{1, 1, 3, 2}
```

**Example 2**: Returns the ascending rank list if specified.

```fs
List.Rank({31, 11, 27, 31}, Order.Ascending)
```

**Result**

```fs
{3, 1, 2, 3}
```