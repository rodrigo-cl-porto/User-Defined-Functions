# [`List.Primes`](/src/m/List.Primes.pq)

Returns a list of prime numbers less than or equal to a given number `n`. It uses the Sieve of Eratosthenes for small values and a variation of Dijkstra’s algorithm for larger values to efficiently generate prime numbers.

## Syntax

```fs
List.Primes(
    n as Int64.Type
) as {number}
```

## Parameters

- `n`: A positive integer; if `n` < 2, the function returns an empty list.

## Return Value

The function returns a list of all prime numbers lower or equal to `n`.

## Remarks

- For `n` < 1000, the function uses the Sieve of Eratosthenes, which is efficient for small ranges.
- For `n` ≥ 1000, the function applies a Dijkstra-inspired algorithm that tracks multiples of known primes to identify new primes.

## Examples

**Example 1**: Primes up to 10.

```fs
List.Primes(10)
```

**Result**

```fs
{2, 3, 5, 7}
```

**Example 2**: Primes up to 30.

```fs
List.Primes(30)
```

**Result**

```fs
{2, 3, 5, 7, 11, 13, 17, 19, 23, 29}
```