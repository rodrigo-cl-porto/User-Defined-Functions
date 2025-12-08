# [`Number.FromRoman`](/src/power_query/Number.FromRoman.pq)

Converts a Roman numeral (text) to a number.

## Syntax

```fs
Number.FromRoman(
    romanText as text
) as number
```

## Parameters

- `romanText`: A text string representing a Roman numeral.

## Return Value

Returns a number corresponding to the Roman numeral. If the input contains invalid characters, an error is raised.

## Remarks

- Supports standard Roman numeral characters: I, V, X, L, C, D, M (case-insensitive).

## Example

```fs
Number.FromRoman("XII")
```

**Result**

```fs
12
```