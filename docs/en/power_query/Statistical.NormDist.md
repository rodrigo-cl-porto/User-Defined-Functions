# [`Statistical.NormDist`](/src/power_query/Statistical.NormDist.pq)

Calculates the value of the **normal distribution** (also known as Gaussian distribution) for a given input `x`. It supports both the **probability density function (PDF)** and the **cumulative distribution function (CDF)**, depending on the cumulative parameter.

## Syntax

```fs
Statistical.NormDist(
    x as number,
    optional mean as number,
    optional std as number,
    optional accumulative as logical
) as number
```

## Parameters

- `x`: The value for which the normal distribution will be evaluated.
- `mean` (_optional_): The mean ($\mu$) of the distribution. Defaults to 0 if not provided.
- `standard deviation` (_optional_): The standard deviation ($\sigma$) of the distribution. Defaults to 1 if not provided.
- `cumulative` (_optional_): Logical value indicating whether to return the cumulative distribution (true) or the probability density (false). Defaults to true.

## Remarks

- When `cumulative = false`, the function returns the probability density at point x using the formula:
    - $\varphi(z)=\frac{1}{\sqrt{2\pi}} \exp(-\frac{z^2}{2})​$
    - where $z = \frac{x-\mu}{\sigma}$ is the number of standard deviations from mean.
- When `cumulative = true`, the function returns the cumulative probability up to point $x$ using the formula:
    - $\phi(z)=\frac{1}{2} + \frac{1}{\sqrt{2 \pi}} \int_{0}^{z}{e^{-t^{2}/2}dt}$.
- The integral part is calculated by [Gaussian Quadrature](#credits-2), which uses a 24-point Legendre-Gauss approximation for high accuracy.
    - $\frac{1}{\sqrt{2\pi}} \int_{0}^{z}{e^{-t^{2}/2}dt} = \frac{z}{4} \sqrt{\frac{2}{\pi}} \sum_{i=1}^{24}{w_{i} \exp(-\frac{z^{2}(t_{i}+1)^2}{8})}$
    - where $w_{i}$ and $t_{i}$ are parameters provided by a Gaussian Quadrature table for 24-point approximation
- This function is useful for statistical modeling, hypothesis testing, and data normalization.

## Return Value

Returns the normal cumulative probability up to a given $x$ by default. If `cumulative = false`, returns the normal probability density at point $x$. If neither x or y are given, returns the **standard** normal CDF up to a given $x$ (which will be treated as the Z-score), or returns the **standard** normal PDF at $x$ if `cumulative` is `false`.

## Examples

**Example 1**: Calculating the cumulative probability for a value of $x$ in a normal distribution with provided mean and standard deviation.

```fs
Statistical.NormDist(100, 80, 10)
```

**Result**

```fs
0.97724986805182079
```

**Example 2**: Calculating the normal PDF for given mean and standard deviation.

```fs
Statistical.NormDist(100, 80, 10, false)
```

**Result**

```fs
0.0539909665131881
```

**Example 3**: In order to calculate the standard normal CDF, just don't input any mean nor standard deviation.

```fs
Statistical.NormDist(1.96)
```

**Result**

```fs
0.97500210485177974
```

**Example 4**: Calculating the standard normal PDF.

```fs
Statistical.NormDist(1.96, null, null, false)
```

**Result**

```fs
0.058440944333451476
```

## Credits

- [Gaussian Quadrature Weights and Abscissae](https://pomax.github.io/bezierinfo/legendre-gauss.html)
    - Author: Mike "Pomax" Kamermans
    - Published at: June 5th, 2011