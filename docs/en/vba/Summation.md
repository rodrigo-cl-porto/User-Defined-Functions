# [`Summation`](/src/vba/Summation.vba)

Computes the numeric summation of a mathematical expression over an integer index range.

## Syntax

```vb
Summation( _
    Expression As String, _
    First As Long, _
    Last As Long _
) As Double
```

## Parameters

- `Expression`: A string representing the math expression in terms of a variable (e.g. `"2*n-1"` or `"1/x^2"`). The function extracts the variable name as the last alphabetical character found in the expression.
- `First`: Starting integer index.
- `Last`: Ending integer index.

## Return Value

Returns the summation's result from expression evaluated for the index running from `First` to `Last`.

## Remarks

- The variable used in Expression is determined by extracting letters from the expression and taking the last letter. Ensure your expression contains the intended variable and that it is the last letter in the expression if multiple letters appear
- Depends on the helper function [`GetLettersOnly`](#getlettersonly) in order to identify the variable in expression

## Examples

**Example 1**: Returns the sum of the first 10 odd numbers.

```vb
Debug.Print Summation("2*n-1", 1, 10)
```

**Result**

```vb
100
```

**Example 2**: Returns the approximated sum of the Basel problem ($\sum_{n=1}^{\infty}{\frac{1}{n^2}}=\frac{\pi}{6}$)

```vb
Debug.Print Summation("1/x^2", 1, 1000000)
```

**Result**

```vb
1.64 ' approaches π²/6
```

**Example 3**: Returns the sum of the first 5 square numbers.

```vb
Debug.Print Summation("i^2", 1, 5)
```

**Result**

```vb
55
```
