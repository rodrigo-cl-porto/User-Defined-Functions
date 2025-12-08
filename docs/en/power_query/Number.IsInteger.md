# [`Number.IsInteger`](/src/power_query/Number.IsInteger.pq)

Checks if a given number is an integer.

## Syntax

```fs
Number.IsInteger(
    value as number
) as logical
```

## Parameters

- `value`: A number to check.

## Return Value

Returns `true` if the number is an integer, `false` otherwise.

## Examples

**Example 1**: Returns `true` if the number is an integer.

```fs
Number.IsInteger(10)
```

**Result**

```fs
true
```

**Example 2**: Returns `false` if the number is not an integer.

```fs
Number.IsInteger(10.5)
```

**Result**

```fs
false
```