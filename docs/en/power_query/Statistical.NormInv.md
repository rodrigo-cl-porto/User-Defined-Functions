# [`Statistical.NormInv`](/src/power_query/Statistical.NormInv.pq)

Returns the inverse of the cumulative distribution function (CDF) of the normal distribution.

## Syntax

```fs
Statistical.NormInv(
    probability as number,
    optional mean as number,
    optional sd as number
) as number
```

## Parameters

- `probability`: A probability value between 0 and 1. Values outside this range are clamped to 0 or 1.
- `mean` (_optional_): The mean ($\mu$) of the distribution. Defaults to 0 if not provided.
- `standard deviation` (_optional_): The standard deviation ($\sigma$) of the distribution. Defaults to 1 if not provided.

## Return Value

A number representing the value $x$ such that the normal distribution's cumulative probability $P(X \le x)$ equals the given `probability`. If neither mean nor standard deviation are specified, returns the value $z$ such that the **standard normal** distribution's cumulative probability  $P(Z \le z)$ equals the given `probability`.

## Remarks

- The function uses a [rational approximation algorithm](#credits-3) to compute the inverse of the standard normal distribution.
- The input probability is clamped between 0 and 1. Values outside this range are adjusted to the nearest valid bound.
- For `probability = 0`, the result is negative infinity (`Number.NegativeInfinity`).
- For `probability = 1`, the result is positive infinity (`Number.PositiveInfinity`).

## Examples

**Example 1**: Returns $x$ such that $P(X \le x) = p$ for a normal distribution with given mean and standard deviation.

```fs
Statistical.NormInv(0.9, 100, 15)
```

**Result**

```fs
119.22327346210234
```

**Example 2**: If neither mean nor standard deviation are informed, returns the value $z$ such that $P(Z \le z) = p$ under the **standard** normal distribution.

```fs
Statistical.NormInv(0.9)
```

**Result**

```fs
1.2815515641401563
```

## Credits

- [An algorithm for computing the inverse normal cumulative distribution function](https://web.archive.org/web/20151030215612/http://home.online.no/~pjacklam/notes/invnorm/)
    - Author: Peter John Acklam
    - Original Site: http://home.online.no/~pjacklam/notes/invnorm
    - Published at: May 4th, 2003