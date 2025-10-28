# User Defined Functions

This repo contains custom functions I've developed throughout my experience as a programmer.

## Table of Contents

- [Power Query (M Code)](#power-query-m-code)
    - [`Binary.Unzip`](#binaryunzip)
    - [`DateTime.ToUnixTime`](#datetimetounixtime)
    - [`Decision.EntropyWeights`](#decisionentropyweights)
    - [`Decision.TOPSIS`](#decisiontopsis)
    - [`List.Correlation`](#listcorrelation)
    - [`List.Rank`](#listrank)
    - [`List.Intercept`](#listintercept)
    - [`List.Outliers`](#listoutliers)
    - [`List.Slope`](#listslope)
    - [`List.PopulationStdDev`](#listpopulationstddev)
    - [`List.Primes`](#listprimes)
    - [`List.Variance`](#listvariance)
    - [`List.WeightedAverage`](#listweightedaverage)
    - [`Number.FromRoman`](#numberfromroman)
    - [`Number.IsInteger`](#numberisinteger)
    - [`Number.IsPrime`](#numberisprime)
    - [`Number.ToRoman`](#numbertoroman)
    - [`Statistical.NormDist`](#statisticalnormdist)
    - [`Statistical.NormInv`](#statisticalnorminv)
    - [`Table.AddListAsColumn`](#tableaddlistascolumn)
    - [`Table.CorrelationMatrix`](#tablecorrelationmatrix)
    - [`Table.FixColumnNames`](#tablefixcolumnnames)
    - [`Table.PreprocessTextColumns`](#tablepreprocesstextcolumns)
    - [`Table.RemoveBlankColumns`](#tableremoveblankcolumns)
    - [`Table.TransposeCorrectly`](#tabletransposecorrectly)
    - [`Text.CountChar`](#textcountchar)
    - [`Text.ExtractNumbers`](#textextractnumbers)
    - [`Text.HtmlToPlainText.pq`](#texthtmltoplaintextpq)
    - [`Text.RegexExtract`](#textregexextract)
    - [`Text.RegexReplace`](#textregexreplace)
    - [`Text.RegexSplit`](#textregexsplit)
    - [`Text.RegexTest`](#textregextest)
    - [`Text.RemoveAccents`](#textremoveaccents)
    - [`Text.RemoveDoubleSpaces`](#textremovedoublespaces)
    - [`Text.RemoveLetters`](#textremoveletters)
    - [`Text.RemoveNumerals`](#textremovenumerals)
    - [`Text.RemovePunctuations`](#textremovepunctuations)
    - [`Text.RemoveStopwords`](#textremovestopwords)
    - [`Text.RemoveWeirdChars`](#textremoveweirdchars)
- [VBA](#vba)
    - [`AreArraysEquals`](#arearraysequals)
    - [`AutoFillFormulas`](#autofillformulas)
    - [`CleanString`](#cleanstring)
    - [`DisableRefreshAll`](#disablerefreshall)
    - [`EnableRefreshAll`](#enablerefreshall)
    - [`FileExists`](#fileexists)
    - [`FileNameIsValid`](#filenameisvalid)
    - [`GetAllFileNames`](#getallfilenames)
    - [`GetLettersOnly`](#getlettersonly)
    - [`GetMonthNumberFromName`](#getmonthnumberfromname)
    - [`GetStringBetween`](#getstringbetween)
    - [`GetStringWithSubstringInArray`](#getstringwithsubstringinarray)
    - [`GetTableColumnNames`](#gettablecolumnnames)
    - [`IsAllTrue`](#isalltrue)
    - [`IsInArray`](#isinarray)
    - [`ListObjectExists`](#listobjectexists)
    - [`PreviousMonthNumber`](#previousmonthnumber)
    - [`RangeHasAnyFormula`](#rangehasanyformula)
    - [`RangeHasConstantValues`](#rangehasconstantvalues)
    - [`RangeIsHidden`](#rangeishidden)
    - [`RangeToHtml`](#rangetohtml)
    - [`SendEmail`](#sendemail)
    - [`SetQueryFormula`](#setqueryformula)
    - [`StringContains`](#stringcontains)
    - [`StringEndsWith`](#stringendswith)
    - [`StringStartsWith`](#stringstartswith)
    - [`SubstringIsInArray`](#substringisinarray)
    - [`Summation`](#summation)
    - [`TableHasQuery`](#tablehasquery)
    - [`WorksheetHasListObject`](#worksheethaslistobject)

## Power Query (M Code)

## [`Binary.Unzip`](/Power%20Query/Binary.Unzip.pq)

Extracts files from a ZIP archive and returns a table of entries with file names and decompressed content.

### Syntax
```fs
Binary.Unzip(
    ZIPFile as binary
) as table
```

### Parameters

- `ZIPFile` — A binary containing a ZIP archive (for example, the result of `File.Contents`).

### Return Value

A table with the following columns:
- `FileName` (text) — The entry name inside the ZIP.
- `Content` (binary or null) — The decompressed file content; `null` if decompression failed or entry unsupported.

### Example

```fs
let
    Source = Binary.Unzip(File.Contents("C:\Temp\archive.zip"))
in
    Source
```

This yields a table you can expand or transform. To read the content of the first file as text:

```fs
let
    Files = Binary.Unzip(File.Contents("C:\Temp\archive.zip")),
    FirstBinary = Files{0}[Content],
    FirstText = if FirstBinary <> null then Text.FromBinary(FirstBinary) else null
in
    FirstText
```

### **Credits**

- Author: Ignacio Barrau
- Source: [ExtractZIP.pq](https://github.com/ibarrau/PowerBi-code/blob/master/PowerQuery/ExtractZIP.pq)

<br>

## [`DateTime.ToUnixTime`](/Power%20Query/DateTime.ToUnixTime.pq)

Converts a Power Query datetime value to Unix time (seconds since 1970-01-01 00:00:00).

### Syntax
```fs
DateTime.ToUnixTime(
    datetimeToConvert as datetime
) as number
```

### Parameters

- `datetimeToConvert`: A datetime value to convert.

### Return Value

Converts `datetime` to Unixtime, which consists of a number representing the total seconds between `datetimeToConvert` and the Unix epoch (1970-01-01 00:00:00). Values are negative for datetimes before the epoch.

### Remarks

- No timezone conversion is performed — treat the input as UTC if you need UTC-based Unix time.

### Example

```fs
DateTime.ToUnixTime(#datetime(2023, 1, 1, 0, 0, 0)) // -> returns 1672531200
```

<br>

## [`Decision.EntropyWeights`](/Power%20Query/Decision.EntropyWeights.pq)

This function calculates the weights of decision criteria using the **entropy weighting method**. It analyzes the distribution of values across each criterion to determine their relative importance based on information entropy.

### Syntax

```fs
Decision.EntropyWeights(
    tbl as table,
    optional columnNames as {text}
) as record
```

### Parameters

- `tbl`: The input table containing numeric values for each criterion.
- `columnNames` (_optional_): A list of column names to include in the calculation. If not provided, all numeric columns in the table will be used.

### Remarks

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

### Return Value

Returns a record where each field name corresponds to a criterion and each field value is the calculated weight.

### Examples

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

<br>

## [`Decision.TOPSIS`](/Power%20Query/Decision.TOPSIS.pq)

This function applies the **TOPSIS** (_Technique for Order Preference by Similarity to Ideal Solution_) method to rank alternatives based on multiple criteria. It normalizes the data, applies weights, calculates the distance to the ideal and anti-ideal solutions, and computes a **closeness coefficient (CC)** for each alternative. The result is a ranked table of alternatives.

### Syntax

```fs
Decision.TOPSIS(
    tbl as table,
    alternativesColumn as text,
    weights as record
) as table
```

### Parameters

- `table`: The input table containing alternatives and their evaluation across multiple criteria.
- `alternativesColumn`: The name of the column that identifies each alternative.
- `weights`: A record where each field name corresponds to a criterion column in the table, and each field value is the weight (importance) of that criterion.

### Return Value

A ranked table of alternatives ordered by the closeness coeficient (CC).

### Remarks

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

### Example

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

<br>

## [`List.Correlation`](/Power%20Query/List.Correlation.pq)

Calculates the correlation coefficient between two lists of numeric values. Supports Pearson (linear) and Spearman (rank-based) correlation.

### Syntax

```fs
List.Correlation(
    list1 as list,
    list2 as list,
    optional typeCorrelation as text
) as number
```

### Parameters

- `list1`: list of numeric values (nulls and non-numeric values are treated as 0).
- `list2`: list of numeric values (nulls and non-numeric values are treated as 0).
- `typeCorrelation` (_optional_): "Pearson" (default) or "Spearman". Case-insensitive.

### Return Value

A number representing the correlation coefficient:

- Pearson: standard Pearson correlation (linear relationship).
- Spearman: Spearman rank correlation (uses dense ranking; tied values receive the same rank).

### Remarks

- Input lists must be the same length; otherwise, an error is raised.
- Null, empty string, or non-numeric entries are converted to 0 before calculation.
- Result is returned as a decimal number (can be negative, positive, or `NaN` if degenerate).

### Examples

```fs
List.Correlation({0, 1, 3, 4}, {4, 5, 10, 30})
// -> 0.858575902776297  (Pearson, default)

List.Correlation({0, 1, 3, 4}, {4, 5, 10, 30}, "Spearman")
// -> 1  (Spearman: monotonic/rank-perfect relationship)

List.Correlation({0, null, 3, "a", 4}, {4, 5, null, 10, 30})
// -> 0.556720639738652  (non-numeric values are treated as 0)
```

<br>

## [`List.Rank`](/Power%20Query/List.Rank.pq)

Returns a list of ranks for a given list of values. Tied values receive the same rank (dense ranking). The Result list preserves the input order.

### Syntax
```fs
List.Rank(
    values as list,
    optional order as Order.Type
) as list
```

### Parameters
- `values`: A list of values to rank. Values must be comparable (numbers, texts, dates, etc.).
- `order` (_optional_): Use `Order.Ascending` or `Order.Descending`. If omitted, the function treats the ordering as descending (i.e., highest value gets rank 1).

### Return Value
A list of integers with the same length as `values`, where each element is the rank (1-based) of the corresponding input value.

### Remarks
- Ranks are "dense": equal values receive the same rank and the next distinct value's rank increases by 1.
  - Example (descending default): values {3,1,2,3} → ranks {1,3,2,1}
- The function returns ranks in the original input order.
- Comparison uses Power Query's Value.Compare, so mixed-type comparisons follow Power Query rules.

### Examples
```fs
List.Rank({10, 10, 5, 7}) // {1, 1, 3, 2} (default: descending)
List.Rank({31, 11, 27, 31}, Order.Ascending) // {3, 1, 2, 3}
List.Rank({10, 10, 30, 30, 2}) // {2, 2, 1, 1, 3}
```

<br>

## [`List.Intercept`](/Power%20Query/List.Intercept.pq)

Calculates the intercept of the linear regression line between two numerical lists X and Y.

### Syntax

```fs
List.Intercept(
    X as list,
    Y as list
) as number
```

### Parameters

- `X`: A list of numerical values representing the independent variable.
- `Y`: A list of numerical values representing the dependent variable.

### Return Value

Returns a number representing the intercept of the linear regression line calculated using the least squares method. If the lists have different lengths, the function returns `null`.

### Remarks

- Both input lists must have the same length; otherwise, the function returns `null`.
- The function uses the least squares method to calculate the intercept.
- Non-numeric values in the lists will cause an error during calculation.

### Example

**Example 1**: Calculating the intercept value of linear regression.

```fs
List.Intercept({1, 2, 3}, {4, 5, 6})
```

**Result**

```fs
-1
```

**Example 2**: Returns `null` if the input lists have different lengths.

```fs
List.Intercept({1, 2}, {3})
```

**Result**

```fs
null
```

<br>

## [`List.Outliers`](/Power%20Query/List.Outliers.pq)

Identifies outliers in a list of numerical values using the Interquartile Range (IQR) method.

### Syntax

```fs
List.Outliers(
    values as list,
    optional multiplier as number
) as list
```

### Parameters

- `values`: A list of numerical values to analyze for outliers.
- `multiplier` (_optional_): A number to adjust the IQR threshold for defining outliers. Default is 1.5.

### Return Value

Returns a list of outlier values identified in the input list based on the IQR method. If no outliers are found, the function returns an empty list.

### Remarks

- The function first removes nulls, empty strings, and whitespace entries, then selects only valid numeric values.
- Outliers are defined as values below $Q_{1} - 1.5 \cdot IQR$ or above $Q_{3} + 1.5 \cdot IQR$, where $Q_1$ and $Q_3$ are the first and third quartiles respectively.

### Examples

**Example 1**: Returns outliers from a list with extreme values.

```fs
List.Outliers({1, 2, 3, 4, 5, 6, 50, 100})
```

**Result**

```fs
{50, 100}
```

**Example 2**: With a higher multiplier, identifying outliers becomes stricter.

```fs
List.Outliers({1, 2, 3, 4, 5, 6, 50, 100}, 3)
```

**Result**

```fs
{100}
```

**Example 3**: Returns an empty list if there's no outlier.

```fs
List.Outliers({10, 12, 14, 15, 16, 18, 20})
```

**Result**

```fs
{}
```

**Example 4**: Ignores nulls and empty strings.

```fs
List.Outliers({1, null, "", 2, 3, 4, 5, null, 6, 50, 100})
```

**Result**

```fs
{50, 100}
```

<br>

## [`List.Slope`](/Power%20Query/List.Slope.pq)

Calculates the slope of the linear regression between two numerical lists X and Y.

### Syntax

```fs
List.Slope(
    X as list,
    Y as list
) as nullable number
```

### Parameters

- `X`: A list of numerical values representing the independent variable.
- `Y`: A list of numerical values representing the dependent variable.

### Return Value

Returns a number representing the slope of the linear regression line calculated using the least squares method. If the lists have different lengths, the function returns `null`.

### Remarks

- Both input lists must have the same length; otherwise, the function returns `null`.
- The function uses the least squares method to calculate the slope.
- Non-numeric values in the lists will cause an error during calculation.

### Examples

**Example 1**: Calculates the slope of linear regression.

```fs
List.Slope({1, 2, 3}, {4, 5, 6})
```

**Result**

```fs
1
```

**Example 2**: Returns `null` if input lists have different lenghts.

```fs
List.Slope({1, 2}, {3})
```

**Result**

```fs
null
```

<br>

## [`List.PopulationStdDev`](/Power%20Query/List.PopulationStdDev.pq)

Calculates the population standard deviation of a list of numerical values.

### Syntax

```fs
List.PopulationStdDev(
    values as list
) as nullable number
```

### Parameters

- `values`: A list of numerical values to calculate the population standard deviation.

### Return Value

Returns a number representing the population standard deviation of the input list. If the list is empty or contains no numeric values, the function returns `null`.

### Remarks

- The function calculates the population standard deviation using the formula:
    - $\sigma = \sqrt{\frac{1}{N} \sum_{i=1}^{N} (x_i - \mu)^2}$
    - where:
        - $N$ is the number of values
        - $x_i$ are the individual values
        - $\mu$ is the mean of the values
- Non-numeric values, nulls, and empty strings are ignored in the calculation.

### Examples

**Example 1**: Calculates population standard deviation for a numeric list.

```fs
List.PopulationStdDev({2, 4, 4, 4, 5, 5, 7, 9})
```

**Result**

```fs
2.8284271247461903
```

**Example 2**: Returns `null` for empty lists.

```fs
List.PopulationStdDev({})
```

**Result**

```fs
null
```

<br>

## [`List.Primes`](/Power%20Query/List.Primes.pq)

Returns a list of prime numbers less than or equal to a given number `n`. It uses the Sieve of Eratosthenes for small values and a variation of Dijkstra’s algorithm for larger values to efficiently generate prime numbers.

### Syntax

```fs
List.Primes(
    n as Int64.Type
) as {number}
```

### Parameters

- `n`: A positive integer; if `n` < 2, the function returns an empty list.

### Return Value

The function returns a list of all prime numbers lower or equal to `n`.

### Remarks

- For `n` < 1000, the function uses the Sieve of Eratosthenes, which is efficient for small ranges.
- For `n` ≥ 1000, the function applies a Dijkstra-inspired algorithm that tracks multiples of known primes to identify new primes.

### Examples

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

<br>

## [`List.Variance`](/Power%20Query/List.Variance.pq)

Calculates the population variance of a list of numerical values.

### Syntax

```fs
List.Variance(
    values as {number}
) as nullable number
```

### Parameters

- `values`: A list of numerical values to calculate the population variance.

### Return Value

Returns a number representing the population variance of the input list. If the list is empty or contains no numeric values, the function returns `null`.

### Remarks

- The function calculates the population variance using the formula: $\sigma^2 = \frac{1}{N} \sum_{i=1}^{N} (x_i - \mu)^2$
    - $N$ is the number of values
    - $x_i$ are the individual values
    - $\mu$ is the mean of the values
- Non-numeric values, nulls, and empty strings are ignored in the calculation.

### Examples

**Example 1**: Calculating sample variance value.

```fs
List.Variance({1, 2, 3, 4, 5})
```

**Result**

```fs
2.5
```

**Example 2**: Calculating population variance value.

```fs
List.Variance({1, 2, 3, 4, 5}, true)
```

**Result**

```fs
2
```

**Example 3**: Returns `null` if receives a empty list.

```fs
List.Variance({})
```

**Result**

```fs
null
```

<br>

## [`List.WeightedAverage`](/Power%20Query/List.WeightedAverage.pq)

Calculates the weighted average of a list of values given a corresponding list of weights.

### Syntax

```fs
List.WeightedAverage(
    values as list,
    weights as list
) as nullable number
```

### Parameters

- `values`: A list of numerical values to calculate the weighted average.
- `weights`: A list of numerical weights corresponding to each value.

### Return Value

Returns a number representing the weighted average of the input values. If the lists have different lengths or if the sum of weights is zero, the function returns `null`.

### Remarks

- The function calculates the weighted average using the formula: $WeightedAverage = \frac{\sum (x_i \times w_i)}{\sum w_i}$
    - $x_i$ are the individual values
    - $w_i$ are the corresponding weights

### Example

Calculing the weighted average for a list with specified weights.

```fs
List.WeightedAverage({1, 2, 3}, {4, 5, 6})
```

**Result**

```fs
2.3333333333333335
```

<br>

## [`Number.FromRoman`](/Power%20Query/Number.FromRoman.pq)

Converts a Roman numeral (text) to a number.

### Syntax

```fs
Number.FromRoman(
    romanText as text
) as number
```

### Parameters

- `romanText`: A text string representing a Roman numeral.

### Return Value

Returns a number corresponding to the Roman numeral. If the input contains invalid characters, an error is raised.

### Remarks

- Supports standard Roman numeral characters: I, V, X, L, C, D, M (case-insensitive).

### Example

```fs
Number.FromRoman("XII")
```

**Result**

```fs
12
```

<br>

## [`Number.IsInteger`](/Power%20Query/Number.IsInteger.pq)

Checks if a given number is an integer.

### Syntax

```fs
Number.IsInteger(
    value as number
) as logical
```

### Parameters

- `value`: A number to check.

### Return Value

Returns `true` if the number is an integer, `false` otherwise.

### Examples

```fs
Number.IsInteger(10) // -> true
Number.IsInteger(10.5) // -> false
```

<br>

## [`Number.IsPrime`](/Power%20Query/Number.IsPrime.pq)

Checks if a given number is prime.

### Syntax

```fs
Number.IsPrime(
    value as Int64.Type
) as logical
```

### Parameters

- `value`: A number to check.

### Return Value

Returns `true` if the number is prime, `false` otherwise.

### Examples

**Example 1**: Check if 7 is prime

```fs
Number.IsPrime(7)
```

**Result**

```fs
true
```

**Example 2**: Check if 100 is prime

```fs
Number.IsPrime(100)
```

**Result**

```fs
false
```

### Credits

- Author: Abigail
- Source: [Abigail's regex to test for prime numbers](http://test.neilk.net/blog/2000/06/01/abigails-regex-to-test-for-prime-numbers/)
- YouTube Video: [How on Earth does ^.?$|^(..+?)\1+$ produce primes?](https://www.youtube.com/watch?v=5vbk0TwkokM)

<br>

## [`Number.ToRoman`](/Power%20Query/Number.ToRoman.pq)

Converts an integer number to a Roman numeral (between 1 and 3999).

### Syntax

```fs
Number.ToRoman(
    numberToConvert as number
) as text
```

### Parameters

- `numberToConvert`: The integer number to be converted to a Roman numeral.

### Return Value

Returns a text string representing the Roman numeral equivalent of the input integer. If the input number is outside the range of 1 to 3999, an error is raised.

### Examples

```fs
Number.ToRoman(12) // -> "XII"
Number.ToRoman(0) // -> Error
```

<br>

## [`Statistical.NormDist`](/Power%20Query/Statistical.NormDist.pq)

Calculates the value of the **normal distribution** (also known as Gaussian distribution) for a given input `x`. It supports both the **probability density function (PDF)** and the **cumulative distribution function (CDF)**, depending on the cumulative parameter.

### Syntax

```fs
Statistical.NormDist(
    x as number,
    optional mean as number,
    optional std as number,
    optional accumulative as logical
) as number
```

### Parameters

- `x`: The value for which the normal distribution will be evaluated.
- `mean` (_optional_): The mean ($\mu$) of the distribution. Defaults to 0 if not provided.
- `standard deviation` (_optional_): The standard deviation ($\sigma$) of the distribution. Defaults to 1 if not provided.
- `cumulative` (_optional_): Logical value indicating whether to return the cumulative distribution (true) or the probability density (false). Defaults to true.

### Remarks

- When `cumulative = false`, the function returns the probability density at point x using the formula:
    - $\varphi(z)=\frac{1}{\sqrt{2 \pi}} \exp(-\frac{z^2}{2})​$
    - where $z = \frac{x - \mu}{\sigma}$
- When `cumulative = true`, the function returns the cumulative probability up to point $x$ using the formula:
    - $\phi(z) = \frac{1}{2} + \frac{1}{\sqrt{\pi}} \int_{0}^{z / \sqrt{2}}{e^{-t^{2}}dt}$.
    - where $z = \frac{x - \mu}{\sigma}$
- The integral part is calculated by [Gaussian Quadrature](#credits-2), which uses a 24-point Legendre-Gauss approximation for high accuracy.
    - $ \frac{1}{\sqrt{\pi}} \int_{0}^{z / \sqrt{2}}{e^{-t^{2}}dt} = \sqrt{\frac{2}{\pi}} \cdot \frac{z}{4} \cdot \sum_{i=1}^{24}{w_{i} \cdot \exp(-\frac{z^{2}(t_{i}+1)^2}{8})}$
    - where $w_{i}$ and $t_{i}$ are parameters provided by a Gaussian Quadrature table for 24-point approximation
- This function is useful for statistical modeling, hypothesis testing, and data normalization.

### Return Value

Returns the normal cumulative probability up to a given $x$ by default. If `cumulative = false`, returns the normal probability density at point $x$. If neither x or y are given, returns the **standard** normal CDF up to a given $x$ (which will be treated as the Z-score), or returns the **standard** normal PDF at $x$ if `cumulative` is `false`.

### Examples

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

### Credits

- [Gaussian Quadrature Weights and Abscissae](https://pomax.github.io/bezierinfo/legendre-gauss.html)
    - Author: Mike "Pomax" Kamermans
    - Published at: June 5th, 2011

<br>

## [`Statistical.NormInv`](/Power%20Query/Statistical.NormInv.pq)

Returns the inverse of the cumulative distribution function (CDF) of the normal distribution.

### Syntax

```fs
Statistical.NormInv(
    probability as number,
    optional mean as number,
    optional sd as number
) as number
```

### Parameters

- `probability`: A probability value between 0 and 1. Values outside this range are clamped to 0 or 1.
- `mean` (_optional_): The mean ($\mu$) of the distribution. Defaults to 0 if not provided.
- `standard deviation` (_optional_): The standard deviation ($\sigma$) of the distribution. Defaults to 1 if not provided.

### Return Value

A number representing the value $x$ such that the normal distribution's cumulative probability $P(X \le x)$ equals the given `probability`. If neither mean nor standard deviation are specified, returns the value $z$ such that the **standard normal** distribution's cumulative probability  $P(Z \le z)$ equals the given `probability`.

### Remarks

- The function uses a [rational approximation algorithm](#credits-3) to compute the inverse of the standard normal distribution.
- The input probability is clamped between 0 and 1. Values outside this range are adjusted to the nearest valid bound.
- For `probability = 0`, the result is negative infinity (`Number.NegativeInfinity`).
- For `probability = 1`, the result is positive infinity (`Number.PositiveInfinity`).

### Examples

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

### Credits

- [An algorithm for computing the inverse normal cumulative distribution function](https://web.archive.org/web/20151030215612/http://home.online.no/~pjacklam/notes/invnorm/)
    - Author: Peter John Acklam
    - Original Site: http://home.online.no/~pjacklam/notes/invnorm
    - Published at: May 4th, 2003

<br>

## [`Table.AddListAsColumn`](/Power%20Query/Table.AddListAsColumn.pq)

Adds a new column to a table using values from a provided list. The new column can be inserted at a specified position and can have a defined data type.

### Syntax

```fs
Table.AddListAsColumn(
    tbl as table, 
    columnName as text, 
    columnValues as list, 
    optional position as number, 
    optional columnType as type
) as table
```

### Parameters

- `tbl`: The input table to which the new column will be added.
- `columnName`: The name of the new column to be added.
- `columnValues`: A list of values to populate the new column.
- `position` (_optional_): The position (0-based index) where the new column should be inserted. If not specified, the column is added at the end.
- `columnType` (_optional_): The data type of the new column. If not specified, the column will have type `any`.

### Return Value

Returns a new table with the added column populated with values from the provided list. If the list has fewer items than the number of rows in the table, nulls are added for the remaining rows. If the list has more items than the number of rows, extra items are ignored.

### Remarks

- If the `position` parameter is provided, the new column will be inserted at the specified index. If the index is out of bounds, an error will occur.
- If the `columnType` parameter is provided, the new column will be created with the specified data type. If not provided, the column will have type `any`.

### Examples

**Example 1**: Add a list as a new column at the end of the table.",

```fs
let
    Source = Table.FromRecords({[A=1, B=2], [A=3, B=4]}),
    NewColumnValues = {10, 20},
    Result = Table.AddListAsColumn(Source, "C", NewColumnValues)
in
    Result
```

**Result**

|A  |B  |C     |
|:-:|:-:|:----:|
|1  |2  |10    |
|3  |4  |20    |

**Example 2**: Add a list as a new column at a specific position with a defined data type.

```fs
let
    Source = Table.FromRecords({[A=1, B=2], [A=3, B=4]}),
    NewColumnValues = {10, 20},
    Result = Table.AddListAsColumn(Source, "C", NewColumnValues, 2, Int64.Type)
in
    Result
```

**Result**

|A  |C     |B  |
|:-:|:----:|:-:|
|1  |10    |2  |
|3  |20    |4  |

**Example 3**: If list has fewer items than rows, nulls are added for remaining rows.

```fs
let
    Source = Table.FromRecords({[A=1, B=2], [A=3, B=4], [A=5, B=6]}),
    NewColumnValues = {10, 20},
    Result = Table.AddListAsColumn(Source, "C", NewColumnValues)
in
    Result
```

**Result**

|A  |B  |C     |
|:-:|:-:|:----:|
|1  |2  |10    |
|3  |4  |20    |
|5  |6  |_null_|

**Example 4**: If list has more items than rows, extra items are ignored.

```fs
let
    Source = Table.FromRecords({[A=1, B=2], [A=3, B=4]}),
    NewColumnValues = {10, 20, 30, 40},
    Result = Table.AddListAsColumn(Source, "C", NewColumnValues)
in
    Result
```

**Result**

|A  |B  |C  |
|:-:|:-:|:-:|
|1  |2  |10 |
|3  |4  |20 |

<br>

## [`Table.CorrelationMatrix`](/Power%20Query/Table.CorrelationMatrix.pq)

Calculates the correlation matrix for a given table. It computes the Pearson correlation coefficient between each pair of numeric columns, returning a table where each cell represents the correlation between two variables.

### Syntax

```fs
Table.CorrelationMatrix(
    tbl as table,
    optional columnNames as type {text}
) as table
```

### Parameters

- `tbl`: The input table containing numeric columns to be analyzed.
- `columnNames` (_optional_): A list of column names to include in the correlation matrix. If not provided, all numeric columns in the table will be used.

### Return Value

The result is a symmetric matrix with correlation values ranging from -1 to 1. The output table includes a "VARIABLE" column indicating the row variable, followed by columns representing correlations with other variables.

### Remarks

- The function uses the Pearson correlation formula to measure linear relationships between columns.
- Nulls, empty strings, and non-numeric values are treated as 0 during computation.

### Examples

**Example 1**: Correlation matrix for all numeric columns.

```fs
let
    Source = #table({"A", "B", "C"}, {
        {1, 2, 3},
        {2, 4, 6},
        {3, 6, 9}
    }),
    Result = Table.CorrelationMatrix(Source)
in
    Result
```

**Result**

|VARIABLE|A  |B  |C  |
|:------:|:-:|:-:|:-:|
|A       |1.0|1.0|1.0|
|B       |1.0|1.0|1.0|
|C       |1.0|1.0|1.0|

**Example 2**: Correlation matrix for selected columns.

```fs
let
    Source = #table({"X", "Y", "Z"}, {
        {1, 10, 100},
        {2, 20, 80},
        {3, 30, 60},
        {4, 40, 40}
    }),
    Result = Table.CorrelationMatrix(Source, {"X", "Z"})
in
    Result
```

**Result**

|VARIABLE|X   |Z   |
|:------:|:--:|:--:|
|X       | 1.0|-1.0|
|Z       |-1.0| 1.0|

<br>

## [`Table.FixColumnNames`](/Power%20Query/Table.FixColumnNames.pq)

Cleans and standardizes column names in a table by removing unwanted characters, trimming spaces, and applying specified text formatting (Proper, Lower, Upper). It also removes columns with default names like 'Column1', 'Column2', etc.

### Syntax

```fs
Table.FixColumnNames(
    tbl as table,
    optional textFormat as text
) as table
```

### Parameters

- `tbl`: The input table whose column names need to be fixed.
- `textFormat` (_optional_): The desired text format for the column names. Accepts 'Proper', 'Lower', or 'Upper'. If not specified, no formatting is applied.

### Return Value

A table with cleaned and standardized column names.

### Remarks

- The function processes the column names of the provided table to ensure they are clean and standardized. It removes non-printable characters, trims leading and trailing spaces, replaces non-breaking spaces with regular spaces, eliminates duplicated spaces, and applies the specified text formatting (Proper, Lower, Upper). Additionally, it removes any columns that have default names such as 'Column1', 'Column2', etc., ensuring that only relevant columns remain in the Result table.
- If the `textFormat` parameter is not provided, the function will only clean the column names without applying any specific text formatting.

### Examples

```fs
Table.FixColumnNames(SourceTable, "Proper") // Cleans and formats column names to Proper case.
Table.FixColumnNames(SourceTable, "Lower") // Cleans and formats column names to Lower case.
Table.FixColumnNames(SourceTable, "Upper") // Cleans and formats column names to Upper case.
Table.FixColumnNames(SourceTable) // Cleans column names without applying any specific text formatting.
```

<br>

## [`Table.PreprocessTextColumns`](/Power%20Query/Table.PreprocessTextColumns.pq)

This function cleans and formats text columns in a table. It removes line breaks, non-standard spaces, duplicated spaces, and applies optional casing (Proper, Lower, or Upper). You can specify which columns to process or let the function automatically detect all text columns.

### Syntax

```fs
Table.PreprocessTextColumns(
    tbl as table,
    optional columnNames as list,
    optional textCasing as text
) as table
```

### Parameters

- `tbl`: The input table containing text columns to be cleaned and formatted.
- `columnNames`: (_optional_) A list of column names to be processed. If not provided or empty, all columns of type text or nullable text will be processed.
- `textCasing`: (_optional_) A string indicating the desired text casing format. Accepted values are:
    - "Proper": Capitalizes the first letter of each word.
    - "Lower": Converts all texts to lowercase.
    - "Upper": Converts all texts to uppercase.
    - If not specified, casing is not changed.

### Remarks

- The function replaces line feed characters (`#(lf)`) with spaces.
- It removes non-breaking spaces (`Character.FromNumber(160)`), trims leading/trailing spaces, and collapses multiple spaces into one.
- This function is useful for preparing text data for analysis, comparison, or display.

### Examples

**Example 1**: Clean all text columns

```fs
let
    Source = #table({"Name", "Comment"}, {
        {"  JOHN DOE  ", "Hello#(lf)World"},
        {"  jane smith", "Nice to meet you"}
    }),
    Result = Table.PreprocessTextColumns(Source)
in
    Result
```

**Result**

|Name          |Comment         |
|:-------------|:---------------|
|JOHN DOE      |Hello World     |
|jane smith    |Nice to meet you|

**Example 2**: Clean and apply Proper case to selected columns

```fs
let
    Source = #table({"Name", "Note"}, {
        {"  MARIA   clara", "great#(lf)job"},
        {"joão   SILVA", "excellent work"}
    }),
    Result = Table.PreprocessTextColumns(Source, {"Name", "Note"}, "Proper")
in
    Result
```

**Result**

|Name       |Note          |
|:----------|:-------------|
|Maria Clara|Great Job     |
|João Silva |Excellent Work|

<br>

## [`Table.RemoveBlankColumns`](/Power%20Query/Table.RemoveBlankColumns.pq)

Removes columns from a table that contain only blank values.

### Syntax

```fs
Table.RemoveBlankColumns(
    tbl as table
) as table
```

### Parameters

- `tbl`: The table from which blank columns will be removed.

### Example

Transposing the table and changing the first column name

```fs
let
    Source = #table({"A", "B"}, {{null, "value1"}, {"", "value2"}}),
    Result = Table.RemoveBlankColumns(Source)
in
    Result
```

**Result**

|B     |
|:----:|
|value1|
|value2|

### Credits

- Author: [Excel Off The Grid](https://exceloffthegrid.com/)
- Source: [Power Query Trick: Instantly Remove All Null Columns! 💥](https://www.youtube.com/watch?v=Zkg9ICg9i40)

<br>

## [`Table.TransposeCorrectly`](/Power%20Query/Table.TransposeCorrectly.pq)

Transposes a table by converting selected columns (or all columns if none are specified) into rows, promotes headers, and adds a new column containing the original column names. This is useful for restructuring data while preserving column identity.

### Syntax

```fs
Table.TransposeCorrectly(
    tbl as table,
    optional columns as list,
    optional firstColumnName as text
) as table
```

### Parameters

- `tbl`: The input table whose columns will be transposed.
- `columnNames`: (_optional_) A list of column names to transpose. If not provided, all columns in the table will be transposed.
- `firstColumnName`: (_optional_) The name to assign to the first column of the transposed table. If not provided, the first name from the columns list will be used.

### Remarks

- The function promotes the first row of the transposed table as headers.
- A new column is added containing the original column names, inserted at the beginning of the table.
- This function is useful for reshaping data, especially when preparing it for pivoting or normalization.

### Examples

**Example 1**: Transposing all columns

```fs
let
    Source = #table({"A", "B", "C"}, {{1, 2, 3}, {4, 5, 6}}),
    Result = Table.TransposeCorrectly(Source)
in
    Result
```

**Result**

|A  |1  |4  |
|:-:|:-:|:-:|
|B  |2  |5  |
|C  |3  |6  |

**Example 2**: Transposing only the selected columns

```fs
let
    Source = #table({"A", "B", "C"}, {{1, 2, 3}, {4, 5, 6}}),
    Result = Table.TransposeCorrectly(Source, {"A", "B"})
in
    Result
```

**Result**

|A  |1  |4  |
|:-:|:-:|:-:|
|B  |2  |5  |

**Example 3**: Transposing the table and changing the first column name

```fs
let
    Source = #table({"A", "B", "C"}, {{1, 2, 3}, {4, 5, 6}}),
    Result = Table.TransposeCorrectly(Source, null, "D")
in
    Result
```

**Result**

|D  |1  |4  |
|:-:|:-:|:-:|
|B  |2  |5  |
|C  |3  |6  |


<br>

## [`Text.CountChar`](/Power%20Query/Text.CountChar.pq)

Counts the occurrences of a specific character in a given text string.

### Syntax

```fs
Text.CountChar(
    textToCount as nullable text,
    charToCount as text
) as number
```

### Parameters

- `textToCount`: The text string in which to count occurrences of the character.
- `charToCount`: The character to count within the text string.

### Return Value

Returns a number representing the count of occurrences of the specified character in the input text. If the input text is null, returns 0.

### Remarks

- The function is case-sensitive; 'a' and 'A' are considered different characters.
- If `charToCount` is an empty string, the function returns 0.

### Examples

**Example 1**: Counts the number of vowels "o" in "hello world".

```fs
Text.CountChar("hello world", "o")
```

**Result**

```fs
2
```

**Example 2**: Returns 0 if there's no occurrence.

```fs
Text.CountChar("quite", "a")
```

**Result**

```fs
0
```

<br>

## [`Text.ExtractNumbers`](/Power%20Query/Text.ExtractNumbers.pq)

Extracts all numeric values from a given text string and returns them as a list of numbers.

### Syntax

```fs
Text.ExtractNumbers(
    inputText as text
) as {number}
```

### Parameters

- `inputText`: The text string from which to extract numeric values.

### Return Value

Returns a list of numbers extracted from the input text. If no numbers are found, returns an empty list.

### Examples

**Example 1**: Extracts numbers from a string containing mixed characters.

```fs
Text.ExtractNumbers("Order #12345: 67 items at $89 each.")
```

**Result**

```fs
{12345, 67, 89}
```

**Example 2**: Returns an empty list when no numbers are present.

```fs
Text.ExtractNumbers("No numbers here!")
```

**Result**

```fs
{}
```

<br>

## [`Text.HtmlToPlainText.pq`](/Power%20Query/Text.HtmlToPlainText.pq)

Converts HTML content to plain text by stripping HTML tags.

### Syntax

```fs
Text.HtmlToPlainText(
    htmlText as text
) as text
```

### Parameters

- `htmlText`: The HTML text to be converted to plain text.

### Return Value

Returns the plain text content extracted from the HTML input.

### Example

```fs
Text.HtmlToPlainText("<p>Hello <b>World</b>!</p>")
```

**Result**

```fs
"Hello World!"
```

<br>

## [`Text.RegexExtract`](/Power%20Query/Text.RegexExtract.pq)

Extracts a substring from a text using a regular expression pattern.

### Syntax

```fs
Text.RegexExtract(
    textToExtract as text,
    regexPattern as text,
    optional global as logical,
    optional caseInsensitive as logical,
    optional multiline as logical
) as any
```

### Parameters

- `textToExtract`: The input text from which to extract the substring.
- `regexPattern`: The regular expression pattern to use for extraction.
- `global` (_optional_): A logical value indicating whether to extract all matches (`true`) or just the first match (`false`). Default is `false`.
- `caseInsensitive` (_optional_): A logical value indicating whether the regex matching should be case insensitive. Default is `false`.
- `multiline` (_optional_): A logical value indicating whether to treat the input text as multiline. Default is `false`.

### Return Value

Returns the extracted substring(s) based on the regex pattern. If `global` is `true`, returns a list of all matches; otherwise, returns the first match or `null` if no match is found.

### Remarks

- Uses .NET regular expressions for pattern matching.
- If `global` is `true`, returns a list of all matches; otherwise, returns the first match or `null` if no match is found.
- Due to Power Query's JavaScript parser limitations, some advanced regex features like lookbehind '(?<=pattern)' and negative lookbehind '(?<!pattern)' and certain flags (`s`, `u`, `v`, `d`, `y`) are not supported.
- Only the flags `g`, `i`, `m` are available.

### Examples

**Example 1**: Extract patterns which start with "W" and end with "d".

```fs
Text.RegexExtract("Hello World", "W.*d")
```

**Result**

```fs
"World"
```

**Example 2**: Extracts all numbers from a text by activating the `global` flag.

```fs
Text.RegexExtract("abc 123 def 456", "\d+", true)
```

**Result**

```fs
{"123", "456"}
```

**Example 3**: By activating the `multiline` flag, the character "^" and "$" comes to mean, respectively, "start of line" and "end of line" instead of "start of text" and "end of text".

```fs
Text.RegexExtract("Hello#(lf)World", "^W.*?d", false, false, true)
```

**Result**

```fs
"World"
```

<br>

## [`Text.RegexReplace`](/Power%20Query/Text.RegexReplace.pq)

Replaces substrings in a text that match a regular expression pattern with a specified replacement string.

### Syntax

```fs
Text.RegexReplace(
    textToModify as text,
    regexPattern as text,
    replacer as text,
    optional global as logical,
    optional caseInsensitive as logical,
    optional multiline as logical
) as nullable text
```

### Parameters

- `textToModify`: The input text in which to perform the replacements.
- `regexPattern`: The regular expression pattern to match substrings for replacement.
- `replacer`: The string to replace matched substrings with.
- `global` (_optional_): A logical value indicating whether to replace all occurrences (`true`) or just the first occurrence (`false`). Default is `false`.
- `caseInsensitive` (_optional_): A logical value indicating whether the regex matching should be case insensitive. Default is `false`.
- `multiline` (_optional_): A logical value indicating whether to treat the input text as multiline. Default is `false`.

### Return Value

Returns the modified text with the specified replacements. If no matches are found, returns the original text.

### Remarks

- Uses .NET regular expressions for pattern matching and replacement.
- If `global` is `true`, replaces all matches; otherwise, replaces only the first match.
- Due to Power Query's JavaScript parser limitations, some advanced regex features like lookbehind '(?<=pattern)' and negative lookbehind '(?<!pattern)' and certain flags (`s`, `u`, `v`, `d`, `y`) are not supported.
- Only the flags `g`, `i`, `m` are available.

### Examples

**Example 1**: Replacing text to another one.

```fs
Text.RegexReplace("Hello World", "World", "Universe")
```

**Result**

```fs
"Hello Universe"
```

**Example 2**: Replacing all numbers in text to word "number".

```fs
Text.RegexReplace("abc 123 def 456", "\d+", "number", true)
```

**Result**

```fs
"abc number def number"
```

**Example 3**: Replacing all words at start of a line which start with a "W" and end with a "d" by "Everyone".

```fs
Text.RegexReplace("Hello#(lf)World", "^W\w*?d", "Everyone", false, false, true)
```

**Result**

```fs
"Hello#(lf)Everyone"
```

<br>

## [`Text.RegexSplit`](/Power%20Query/Text.RegexSplit.pq)

Splits a text into a list of substrings based on a regular expression pattern.

### Syntax

```fs
Text.RegexSplit(
    textToSplit as text,
    regexPattern as text,
    optional caseInsensitive as logical,
    optional multiline as logical
) as list
```

### Parameters

- `textToSplit`: The input text to be split.
- `regexPattern`: The regular expression pattern to use as the delimiter for splitting.
- `caseInsensitive` (_optional_): A logical value indicating whether the regex matching should be case insensitive. Default is `false`.
- `multiline` (_optional_): A logical value indicating whether to treat the input text as multiline. Default is `false`.

### Return Value

Returns a list of substrings obtained by splitting the input text at each match of the regex pattern.

### Remarks

- Uses .NET regular expressions for pattern matching.
- Due to Power Query's JavaScript parser limitations, some advanced regex features like lookbehind '(?<=pattern)' and negative lookbehind '(?<!pattern)' and certain flags (`s`, `u`, `v`, `d`, `y`) are not supported.
- Only the flags `i`, `m` are available.

### Examples

```fs
Text.RegexSplit("apple,banana,cherry", ",") // -> {"apple", "banana", "cherry"}
Text.RegexSplit("one1two2three3", "\d") // -> {"one", "two", "three", ""}
Text.RegexSplit("Hello\nWorld", "\n", false, true) // -> {"Hello", "World"}
```

<br>

## [`Text.RegexTest`](/Power%20Query/Text.RegexTest.pq)

Tests whether a text matches a regular expression pattern.

### Syntax

```fs
Text.RegexTest(
    textToTest as text,
    regexPattern as text,
    optional caseInsensitive as logical,
    optional multiline as logical
) as logical
```

### Parameters

- `textToTest`: The input text to be tested against the regex pattern.
- `regexPattern`: The regular expression pattern to test.
- `caseInsensitive` (_optional_): A logical value indicating whether the regex matching should be case insensitive. Default is `false`.
- `multiline` (_optional_): A logical value indicating whether to treat the input text as multiline. Default is `false`.

### Return Value

Returns `true` if the input text matches the regex pattern, `false` otherwise.

### Remarks

- Uses .NET regular expressions for pattern matching.
- Due to Power Query's JavaScript parser limitations, some advanced regex features like lookbehind '(?<=pattern)' and negative lookbehind '(?<!pattern)' and certain flags (`s`, `u`, `v`, `d`, `y`) are not supported.
- Only the flags `i`, `m` are available.

### Examples

```fs
Text.RegexTest("Hello World", "World") // -> true
Text.RegexTest("abc 123", "^\d+$") // -> false
Text.RegexTest("Hello\nWorld", "^W.*d", false, true) // -> true
```

<br>

## [`Text.RemoveAccents`](/Power%20Query/Text.RemoveAccents.pq)

Removes accents from characters in a text string.

### Syntax

```fs
Text.RemoveAccents(
    inputText as text
) as text
```

### Parameters

- `inputText`: The text string from which to remove accents.

### Return Value

Returns the input text with all accented characters replaced by their unaccented equivalents.

### Examples

```fs
Text.RemoveAccents("Café") // -> "Cafe"
Text.RemoveAccents("naïve") // -> "naive"
```

<br>

## [`Text.RemoveDoubleSpaces`](/Power%20Query/Text.RemoveDoubleSpaces.pq)

Removes consecutive double spaces from a text string, replacing them with single spaces.

### Syntax

```fs
Text.RemoveDoubleSpaces(
    inputText as text
) as text
```

### Parameters

- `inputText`: The text string from which to remove double spaces.

### Return Value

Returns the input text with all consecutive double spaces replaced by single spaces.

### Example

```fs
Text.RemoveDoubleSpaces("This  is   a    test.")
```

**Result**

```fs
"This is a test."
```

<br>

## [`Text.RemoveLetters`](/Power%20Query/Text.RemoveLetters.pq)

Removes all alphabetic characters from a text string, leaving only non-letter characters.

### Syntax

```fs
Text.RemoveLetters(
    textToModify as text
) as text
```

### Parameters

- `textToModify`: The text string from which to remove alphabetic characters.

### Return Value

Returns the input text with all alphabetic characters removed.

### Example

```fs
Text.RemoveLetters("Hello123 World!")
```

**Result**

```fs
"123 !"
```

<br>

## [`Text.RemoveNumerals`](/Power%20Query/Text.RemoveNumerals.pq)

Removes all numeric characters from a text string, with an option to also remove Roman numerals.

### Syntax

```fs
Text.RemoveNumerals(
    textToRemove as text,
    optional removeRomanNumerals as logical
) as text
```

### Parameters

- `textToRemove`: The text string from which to remove numeric characters.
- `removeRomanNumerals` (_optional_): A logical value indicating whether to also remove Roman numeral characters (I, V, X, L, C, D, M). Default is `false`.

### Return Value

Returns the input text with all numeric characters (and optionally Roman numerals) removed.

### Example

**Example 1**: Removing numerals from a text.

```fs
Text.RemoveNumerals("Room 101 IV")
```

**Result**

```fs
"Room  IV"
```

**Example 2**: Removing numerals from a text, including Roman numerals.

```fs
Text.RemoveNumerals("Room 101 IV", true) 
```

**Result**

```fs
"Room  "
```

<br>

## [`Text.RemovePunctuations`](/Power%20Query/Text.RemovePunctuations.pq)

Removes all punctuation characters from a text string.

### Syntax

```fs
Text.RemovePunctuations(
    textToRemove as nullable text,
    optional replacer as text
) as text
```

### Parameters

- `textToRemove`: The text string from which to remove punctuation characters.
- `replacer` (_optional_): A text string to replace punctuation characters with. If omitted, punctuation characters are removed without replacement.

### Return Value

Returns the input text with all punctuation characters removed or replaced by the specified replacer.

### Examples

```fs
Text.RemovePunctuations("Hello, World!") // -> "Hello World"
Text.RemovePunctuations("Hello, World!", " ") // -> "Hello  World "
```

<br>

## [`Text.RemoveStopwords`](/Power%20Query/Text.RemoveStopwords.pq)

Removes common Portuguese stopwords from a text string to enhance text analysis.

### Syntax

```fs
Text.RemoveStopwords(
    textToModify as nullable text,
    optional undesirableWords as list
) as text
```

### Parameters

- `textToModify`: The text string from which to remove stopwords.
- `undesirableWords` (_optional_): A list of additional words to remove from the text. Default is an empty list.

### Return Value

Returns the input text with all Portuguese stopwords and any additional specified words removed.

### Examples

```fs
Text.RemoveStopwords("Este é um exemplo de texto para remover palavras comuns.")
// -> "exemplo texto remover palavras comuns."

Text.RemoveStopwords("Este é um exemplo de texto para remover palavras comuns.", {"exemplo", "texto"})
// -> "remover palavras comuns."
```

<br>

## [`Text.RemoveWeirdChars`](/Power%20Query/Text.RemoveWeirdChars.pq)

Removes special and non-printable characters from a text string, with an option to replace them with spaces.

### Syntax

```fs
Text.RemoveWeirdChars(
    textToClean as text,
    optional replacer as text
) as text
```

### Parameters

- `textToClean`: The text string to be cleaned.
- `replacer` (_optional_): A text string to replace special characters with. If omitted, special characters are replaced by an white space.

### Return Value

Returns the cleaned text with special characters either removed or replaced by the specified replacer.

### Examples

```fs
Text.RemoveWeirdChars("Hello" & Character.FromNumber(0) & "World!") // -> "Hello World!"
Text.RemoveWeirdChars("Hello" & Character.FromNumber(0) & "World!", "_") // -> "Hello_World!"
```

<br>

## VBA

<br>

## [`AreArraysEquals`](/VBA/AreArraysEqual.vba)

Compares two arrays to check if they are equal, meaning they have the same size and identical elements in the same order.

### Syntax

```vb
AreArraysEqual( _
    Array1 As Variant, _
    Array2 As Variant _
) As Boolean
```

### Parameters
- `Array1`: First array to compare
- `Array2`: Second array to compare

### Return Value

Returns `True` if both arrays are equal, `False` otherwise.

### Remarks

- Arrays must have the same upper and lower bounds
- Arrays must have identical elements in the same positions
- The function performs an element-by-element comparison
- Returns `False` if arrays have different sizes
- Can compare arrays of any type since parameters are declared as Variant

### Example

```vb
Dim arr1 As Variant
Dim arr2 As Variant
arr1 = Array(1, 2, 3)
arr2 = Array(1, 2, 3)

If AreArraysEqual(arr1, arr2) Then
    Debug.Print "Arrays are equal"
Else
    Debug.Print "Arrays are different"
End If
```

<br>

## [`AutoFillFormulas`](/VBA/AutoFillFormulas.vba)

Automatically fills formulas across a range using a reference cell's formula. The reference cell can be either the first or last cell containing a formula in the range.

### Syntax

```vb
AutoFillFormulas( _
    rng As Range, _
    Optional UseLastCellAsRef As Boolean = False _
)
```

### Parameters

- `rng`: The range where formulas will be filled
- `UseLastCellAsRef`: (_optional_) Boolean flag to determine which cell to use as reference
    - `False` (Default): Uses the first cell with formula as reference
    - `True`: Uses the last cell with formula as reference

### Remarks

- Does nothing if the range is empty (Nothing) or contains only one cell
- Only works if the range contains at least one formula
- Uses R1C1 formula notation to ensure proper relative references when filling
- Only fills formulas in cells that are part of the specified range
- Requires the helper function [`RangeHasAnyFormula`](#rangehasanyformula) to check for formulas in the range

### Example

```vb
Dim rng As Range
Set rng = Range("A1:A10")
AutoFillFormulas rng 'Uses first formula cell as reference

'Or using the last cell as reference:
AutoFillFormulas rng, True
```

<br>

## [`CleanString`](/VBA/CleanString.vba)

Cleans a string by removing or replacing special characters and control characters with spaces.

### Syntax

```vb
CleanString( _
    myString As String, _
    Optional ReplaceBySpace As Boolean = True, _
    Optional ConvertNonBreakingSpace As Boolean = True _
) As String
```

### Parameters

- `myString`: The input string to be cleaned
- `ReplaceBySpace`: (_optional_) Boolean flag that determines if special characters should be replaced by spaces
    - `True` (Default): Replaces special characters with spaces
    - `False`: Removes special characters without replacement
- `ConvertNonBreakingSpace`: (_optional_) Boolean flag to handle non-breaking spaces
    - `True` (Default): Converts non-breaking spaces (ASCII 160) to regular spaces
    - `False`: Leaves non-breaking spaces unchanged

### Return Value

Returns the cleaned string with special characters either removed or replaced by spaces.

### Remarks
- Removes ASCII control characters (0-31)
- Handles special characters like ASCII 127, 129, 141, 143, 144, and 157
- Converts non-breaking spaces to regular spaces (when enabled)
- Trims leading and trailing spaces from the final result
- Preserves all other printable characters

### Example

```vb
Dim cleanedStr As String

' Replace special characters with spaces
cleanedStr = CleanString("Hello" & Chr(0) & "World")
Debug.Print cleanedStr ' Result: "Hello World"

' Remove special characters
cleanedStr = CleanString("Hello" & Chr(0) & "World", False)
Debug.Print cleanedStr ' Result: "HelloWorld"

' Keep non-breaking spaces
cleanedStr = CleanString("Hello" & Chr(160) & "World", True, False)
Debug.Print cleanedStr ' Result: Original string unchanged
```

<br>

## [`DisableRefreshAll`](/VBA/DisableRefreshAll.vba)

Disables the "Refresh All" functionality for OLEDB connections in a specified workbook.

### Syntax

```vb
DisableRefreshAll( _
    ByRef wb As Workbook _
)
```

### Parameters

- `wb`: Reference to the workbook where OLEDB connections will be modified

### **Use Cases**

- Improve performance by preventing unnecessary data refreshes
- Control which connections should be updated during a "Refresh All" operation
- Selectively manage data refresh behavior in workbooks with multiple connections

### Remarks

- Only affects OLEDB type connections
- Does not modify PowerPivot or other connection types
- Changes are applied to each connection individually
- The connections will still be refreshable individually, just not through "Refresh All" option
- Changes are made directly to the workbook passed as parameter

### Example

```vb
Dim wb As Workbook
Set wb = ThisWorkbook
DisableRefreshAll wb
```

<br>

## [`EnableRefreshAll`](/VBA/EnableRefreshAll.vba)

Enables the "Refresh All" functionality for OLEDB connections in a specified workbook.

### Syntax

```vb
EnableRefreshAll( _
    ByRef wb As Workbook _
)
```

### Parameters

- `wb`: Reference to the workbook where OLEDB connections will be modified

### **Use Cases**

- Restore default refresh behavior for OLEDB connections
- Enable batch updates of multiple connections
- Ensure all OLEDB connections are included in "Refresh All" operations
- Manage data refresh settings after temporary disablement

### Remarks

- Only affects OLEDB type connections
- Does not modify PowerPivot or other connection types
- Changes are applied to each connection individually
- Allows connections to be updated when using "Refresh All" command
- Changes are made directly to the workbook passed as parameter

### Example

```vb
Dim wb As Workbook
Set wb = ThisWorkbook
EnableRefreshAll wb
```

<br>

## [`FileExists`](/VBA/FileExists.vba)

Checks if a file exists at the specified file path.

### Syntax

```vb
FileExists( _
    FilePath As String _
) As Boolean
```

### Parameters

- `FilePath`: The complete path to the file being checked

### Return Value

Returns `True` if the file exists, `False` otherwise.

### Remarks

- Uses VBA's `Dir` function to test file existence
- Works with any file type
- Path must be accessible from the current environment
- Case-insensitive file path checking

### Example

```vb
Dim exists As Boolean
exists = FileExists("C:\Documents\myfile.xlsx")

If exists Then
    Debug.Print "File exists"
Else
    Debug.Print "File not found"
End If
```

### **Credits**

- Original source: [www.TheSpreadsheetGuru.com/The-Code-Vault](www.TheSpreadsheetGuru.com/The-Code-Vault)
- Resource: [http://www.rondebruin.nl/win/s9/win003.htm](http://www.rondebruin.nl/win/s9/win003.htm)

<br>

## [`FileNameIsValid`](/VBA/FileNameIsValid.vba)

Validates if a given string can be used as a valid file name by checking for illegal characters.

### Syntax

```vb
FileNameIsValid( _
    FileName As String _
) As Boolean
```

### Parameters

- `FileName`: The string to be validated as a file name

### Return Value

Returns `True` if the file name is valid, `False` if it contains illegal characters or is empty.

### Remarks

- Checks for the following illegal characters: `\ / : * ? < > | [ ] "`
- Returns `False` for empty strings
- Case-sensitive validation
- Does not check file name length restrictions
- Does not validate against reserved Windows file names

### Example

```vb
Dim isValid As Boolean

isValid = FileNameIsValid("my_file.txt")
Debug.Print isValid ' True

isValid = FileNameIsValid("file*.txt") 
Debug.Print isValid ' False

isValid = FileNameIsValid("folder/file.txt")
Debug.Print isValid  ' False
```

### **Credits**

- Author: Jon Peltier
- Source: [www.TheSpreadsheetGuru.com/the-code-vault](www.TheSpreadsheetGuru.com/the-code-vault)

<br>

## [`GetAllFileNames`](/VBA/GetAllFileNames.vba)

Retrieves an array of all file names from a specified folder and its subfolders, with optional file extension filtering.

### Syntax

```vb
GetAllFileNames( _
    FolderPath As String, _
    Optional fileExt As String _
) As String()
```

### Parameters

- `FolderPath`: The path to the folder to search in
- `fileExt`: (_optional_) File extension to filter results. If omitted, returns all files

### Return Value

Returns a zero-based string array containing all matching file names.

### Remarks

- Recursively searches through all subfolders
- Case-insensitive file extension matching
- Uses `FileSystemObject` for file system operations
- Returns only file names, not full paths
- Extension filter doesn't require the dot prefix
- Empty array if no files are found
- Requires reference to Microsoft Scripting Runtime (or late binding)

### **Dependencies**

- `Scripting.FileSystemObject` reference

### Example

```vb
Dim files() As String
Dim i As Long

' Get all Excel files
files = GetAllFiles("C:\Documents", "xlsx")

' Get all files regardless of extension
files = GetAllFiles("C:\Documents")

' Print all found files
For i = 0 To UBound(files)
    Debug.Print files(i)
Next i
```

<br>

## [`GetLettersOnly`](/VBA/GetLettersOnly.vba)

Extracts only ASCII letters (a–z) from a string and returns them in lowercase.

### Syntax

```vb
GetLettersOnly( _
    Text As String _
) As String
```

### Parameters

- `Text`: The input string to process.

### Return Value

Returns a string containing only the letters a–z (converted to lowercase). Returns an empty string if no ASCII letters are found.

### Remarks

- Filters characters using ASCII range 97–122 (letters a–z).
- Converts characters to lowercase before testing and Result.
- Does not preserve original letter case.
- Does not include accented letters, non-Latin characters, or other alphabetic Unicode letters.
- Useful for normalizing or sanitizing input to ASCII letters only.

### Example

```vb
Dim result As String

result = GetLettersOnly("Hello, World! 123")   
Debug.Print result ' "helloworld"

result = GetLettersOnly("Ábç Def")
Debug.Print result ' "def" (accented letters removed)
```

<br>

## [`GetMonthNumberFromName`](/VBA/GetMonthNumberFromName.vba)

Converts a month name to its corresponding numeric value (1-12).

### Syntax

```vb
GetMonthNumberFromName( _
    MonthName As String _
) As Integer
```

### Parameters

- `MonthName`: The name of the month (full or abbreviated, in any language supported by Excel)

### Return Value

Returns an integer from 1 to 12 representing the month number.

### Remarks

- Works with month names in any language supported by Excel
Accepts both full month names and abbreviated forms
- Case-insensitive
- Returns error if month name is invalid

### Example

```vb
Dim monthNum As Integer

monthNum = GetMonthNumberFromName("January")   
Debug.Print monthNum ' Prints 1

monthNum = GetMonthNumberFromName("Jan")
Debug.Print monthNum ' Prints 1

monthNum = GetMonthNumberFromName("Janeiro")
Debug.Print monthNum ' Prints 1 (Portuguese)

monthNum = GetMonthNumberFromName("Janvier") 
Debug.Print monthNum ' Returns 1 (French)
```

<br>

## [`GetStringBetween`](/VBA/GetStringBetween.vba)

Extracts a substring between two specified delimiter strings.

### Syntax

```vb
GetStringBetween( _
    str As String, _
    startStr As String, _
    endStr As String _
) As String
```

### Parameters

- `str`: The input string to search in
- `startStr`: The starting delimiter string
- `endStr`: The ending delimiter string

### Return Value

Returns the text found between the start and end strings. Returns an empty string if no match is found.

### Remarks

- Uses VBScript RegExp for pattern matching
- Creates RegExp object using late binding to avoid explicit reference requirement
- Case-insensitive search
- Non-greedy matching (returns shortest match)
- Returns only the first match if multiple exist
- Removes the delimiter strings from the result

### Example

```vb
Dim result As String

result = GetStringBetween("Hello [World] Test", "[", "]")
Debug.Print result ' Returns "World"

result = GetStringBetween("<tag>Content</tag>", "<tag>", "</tag>")
Debug.Print result ' Returns "Content"

result = GetStringBetween("No delimiters here", "[", "]")
Debug.Print result  ' Returns ""
```

<br>

## [`GetStringWithSubstringInArray`](/VBA/GetStringWithSubstringInArray.vba)

Searches through an array of strings and returns the first string that contains a specified substring.

### Syntax

```vb
GetStringWithSubstringInArray( _
    SubString As String, _ 
    SourceArray As Variant, _
    Optional CaseSensitive As Boolean = False _
) As String
```

### Parameters

- `SubString`: The text to search for within each array element
- `SourceArray`: Array containing strings to search through
- `CaseSensitive`: (_optional_) Boolean flag to enable case-sensitive search. Default is False

### Return Value

Returns the first string from the array containing the substring. Returns an empty string if no match is found.

### Remarks

- Only processes elements that are strings (type `vbString`)
- Ignores non-string elements in the array
- Case-insensitive by default
- Returns first match found and exits
- Works with arrays of any dimension

### **Dependencies**

- Requires [`StringContains`]() function

### Example

```vb
Dim testArray As Variant
Dim result    As String

testArray = Array("Hello World", "Test String", "Another Text")

result = GetStringWithSubstringInArray("World", testArray)
Debug.Print result ' Returns "Hello World"

result = GetStringWithSubstringInArray("text", testArray)
Debug.Print result ' Returns "Another Text"

result = GetStringWithSubstringInArray("none", testArray)
Debug.Print result  ' Returns ""
```

<br>

## [`GetTableColumnNames`](/VBA/GetTableColumnNames.vba)

Returns the header names of an Excel ListObject (table) as a zero-based string array.

### Syntax

```vb
GetTableColumnNames( _
    lo As ListObject _
) As String()
```

### Parameters

- `lo`: The ListObject (Excel table) to read column headers from

### Return Value

Returns a zero-based array of strings containing the table column header values in left-to-right order.

### Remarks

- Includes hidden columns and preserves the table column order.

### Example

```vb
Dim colNames() As String
Dim i          As Long

Set tbl = ThisWorkbook.Worksheets("Sheet1").ListObjects("Table1")
colNames = GetTableColumnNames(tbl)

For i = 0 To UBound(colNames)
    Debug.Print colNames(i)
Next i
```

<br>

## [`IsAllTrue`](/VBA/IsAllTrue.vba)

Checks if all elements in a boolean array are `True`.

### Syntax

```vb
IsAllTrue( _
    blnArray As Variant _
) As Boolean
```

### Parameters

- `blnArray`: Array containing boolean values to be checked

### Return Value

Returns `True` if all elements in the array are boolean `True`, otherwise returns `False`.

### **Use Cases**

- Validating that multiple conditions are all met
- Checking status of multiple boolean flags
- Quality control checks where all criteria must be true

### Remarks

- Returns `False` if any element is not a boolean type
- Returns `False` if any element is `False`
- Early exit when first non-true value is found
- Can handle arrays of any dimension
- Array must be passed as `Variant` type

### Example

```vb
Dim testArray As Variant

testArray = Array(True, True, True)
Debug.Print IsAllTrue(testArray) ' Returns True

testArray = Array(True, False, True)
Debug.Print IsAllTrue(testArray) ' Returns False

testArray = Array(True, "True", True)
Debug.Print IsAllTrue(testArray) ' Returns False (non-boolean element)
```

<br>

## [`IsInArray`](/VBA/IsInArray.vba)

Checks whether a value exists in a one-dimensional array.

### Syntax

```vb
IsInArray( _
    ValueToBeFound As Variant, _
    SourceArray As Variant _
) As Boolean
```

### Parameters

- `ValueToBeFound`: The value to search for (any Variant).
- `SourceArray`: The one-dimensional array to search (Variant).

### Return Value

Returns `True` if the value is found in the array, otherwise returns `False`.

### Remarks

- Expects a one-dimensional array; passing an uninitialized or multi-dimensional array may cause errors.

### Example

```vb
Dim arr As Variant
arr = Array("apple", "banana", "cherry")

If IsInArray("banana", arr) Then
    Debug.Print "Found"
Else
    Debug.Print "Not found"
End If
```

<br>

## [`ListObjectExists`](/VBA/ListObjectExists.vba)

Checks whether a ListObject (Excel table) with a given name exists in a workbook.

### Syntax

```vb
ListObjectExists( _
    ByRef wb As Workbook, _
    loName As String _
) As Boolean
```

### Parameters

- `wb`: Workbook to search.
- `loName`: Name of the table (`ListObject`) to find.

### Return Value

Returns `True` if a ListObject with the specified name is found in any worksheet of the workbook; otherwise returns `False`.

### Remarks

- Performs a direct name comparison (behavior may be affected by the project's Option Compare setting).

### Example

```vb
Dim exists As Boolean
exists = ListObjectExists(ThisWorkbook, "Table1")

If exists Then
    Debug.Print "Table exists"
Else
    Debug.Print "Table not found"
End If
```

<br>

## [`PreviousMonthNumber`](/VBA/PreviousMonthNumber.vba)

Returns the numeric month (1–12) that precedes the month of a given date.

### Syntax

```vb
PreviousMonthNumber( _
    dt As Date _
) As Integer
```

### Parameters

- `dt`: Date value used to determine the previous month

### Return Value

Returns an Integer from 1 to 12 representing the previous month. For dates in January, returns 12 (December).

### Example

```vb
Dim prev As Integer

prev = PreviousMonthNumber(DateSerial(2025, 3, 15))
Debug.Print prev ' returns 2 (February)

prev = PreviousMonthNumber(DateSerial(2025, 1, 10))
Debug.Print prev ' returns 12 (December)
```

<br>

## [`RangeHasAnyFormula`](/VBA/RangeHasAnyFormula.vba)

Checks if a given range contains any cells with formulas.

### Syntax

```vb
RangeHasAnyFormula( _
    ByVal rng As Range _
) As Boolean
```

### Parameters

- `rng`: The range to be checked for formulas

### Return Value

Returns `True` if the range contains at least one formula, `False` otherwise.

### Remarks

- Returns `False` if the range is Nothing
- Uses error handling to detect the presence of formulas
- Shows an error message if any unexpected error occurs during execution
- Uses Excel's `SpecialCells` method with `xlCellTypeFormulas` to perform the check

### Example

```vb
Dim rng As Range
Set rng = Range("A1:D10")

If RangeHasAnyFormula(rng) Then
    Debug.Print "Range contains at least one formula"
Else
    Debug.Print "Range contains no formulas"
End If
```

### **Error Handling**

- Displays a message box with error details if an unexpected error occurs
- Properly handles the "No cells were found" error which indicates no formulas are present

<br>

## [`RangeHasConstantValues`](/VBA/RangeHasConstantValues.vba)

Checks whether a given range contains any constant (non-formula) cells.

### Syntax

```vb
RangeHasConstantValues( _
    rng As Range _
) As Boolean
```

### Parameters

- `rng`: Range to check for constant values.

### Return Value

Returns `True` if the range contains at least one constant cell; otherwise returns False. If `rng` is `Nothing` the function returns `False`.

### Example

```vb
Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Sheet1").Range("A1:A10")

If RangeHasConstantValues(rng) Then
    Debug.Print "Range contains constants"
Else
    Debug.Print "Range contains no constants or is invalid"
End If
```

<br>

## [`RangeIsHidden`](/VBA/RangeIsHidden.vba)

Determines whether a given range is entirely hidden (no visible cells).

### Syntax

```vb
RangeIsHidden( _
    rng As Range _
) As Boolean
```

### Parameters

- `rng`: The Range to check for visibility.

### Return Value

Returns `True` if the range contains no visible cells (i.e., is hidden). Returns `False` if at least one cell in the range is visible or if `rng` is `Nothing`.

### Example

```vb
Dim rng As Range
Set rng = ThisWorkbook.Worksheets("Sheet1").Range("A1:A10")

If RangeIsHidden(rng) Then
    Debug.Print "Range is hidden (no visible cells)."
Else
    Debug.Print "Range has visible cells."
End If
```

<br>

## [`RangeToHtml`](/VBA/RangeToHtml.vba)

Converts an Excel Range into an HTML string by copying the range to a temporary workbook, publishing that sheet as an HTML file, and returning the file contents.

### Syntax

```vb
RangeToHtml( _
    rng As Range _
) As String
```

### Parameters

- `rng`: The Range to convert to HTML.

### Return Value

Returns a string containing the HTML representation of the provided range. Returns an empty string if an error occurs.

### Remarks

- Creates a temporary workbook, pastes the range (values and formats) and removes drawing objects before publishing.
- Uses the system temporary folder (Environ$("temp")) to create an intermediate .htm file.
- Reads the generated HTML file into memory and deletes the temporary file and workbook.
- Replaces `align=center` with `align=left` in the resulting HTML.
- Images/drawing objects are deleted in the temporary workbook to avoid embedding them in the HTML.

### Example

```vba
Dim html As String
html = RangeToHtml(ThisWorkbook.Worksheets("Sheet1").Range("A1:D10"))
' html now contains the HTML representation of the range
```

<br>

## [`SendEmail`](/VBA/SendEmail.vba)

Sends an HTML email using CDO (Collaboration Data Objects) with NTLM authentication, typically used in corporate environments with Exchange Server.

### Syntax

```vb
SendEmail( _
    Sender As String, _
    Recipient As String, _
    Subject As String, _
    Message As String, _
    Optional CarbonCopy As String, _
    Optional BlindCarbonCopy As String _
)
```

### Parameters

- `Sender`: Email address of the sender
- `Recipient`: Email address(es) of the recipient(s)
- `Subject`: Subject line of the email
- `Message`: HTML-formatted body of the email
- `CarbonCopy`: (_optional_) Email address(es) for CC recipients
- `BlindCarbonCopy`: (_optional_) Email address(es) for BCC recipients

### Remarks

- Uses CDO with NTLM authentication (Windows Authentication)
- Configured for SMTP with STARTTLS (port 587)
- Supports HTML formatting in the message body
- Multiple recipients can be specified using semicolon (;) as separator
- No explicit error handling is implemented

### **Configuration Constants**

- `CDO_DEFAULT_SETTINGS`: -1 (Use system default settings)
- `CDO_NTLM_AUTHENTICATION`: 2 (Windows Authentication)
- `CDO_SEND_USING_PORT`: 2 (Direct SMTP)
- `CDO_SERVER_PORT`: 587 (STARTTLS port)
- `CDO_SMTP_SERVER`: "mailhost.yourdomain.net" (SMTP server address)

### **Dependencies**

- Requires CDO to be available on the system
- Requires proper SMTP server configuration
- Requires appropriate network/firewall access

### Example

```vb
Call SendEmail( _
    "sender@company.com", _
    "recipient@company.com", _
    "Test Subject", _
    "<h1>Hello</h1><p>This is a test email.</p>", _
    "cc@company.com", _
    "bcc@company.com")
```

<br>

## [`SetQueryFormula`](/VBA/SetQueryFormula.vba)

Modifies a Power Query formula in the current workbook based on a given value, handling different data types appropriately.

### Syntax

```vb
SetQueryFormula( _
    queryName As String, _
    value As Variant _
)
```

### Parameters

- `queryName`: Name of the Power Query to modify
- `value`: Value to set in the query formula (supports `String`, `Date`, and `Byte Array`)

### **Dependencies**

- Requires Excel version that supports Power Query

### Example

```vb
' Set a string value
SetQueryFormula "MyQuery", "Hello ""World"""  ' Results in: "Hello ""World"""

' Set a date value
SetQueryFormula "MyQuery", DateSerial(2023, 10, 17)  ' Results in: #date(2023,10,17)

' Set a byte array
Dim byteArr() As Byte
byteArr = Array(1, 2, 3)
SetQueryFormula "MyQuery", byteArr  ' Results in: {1,2,3}
```

<br>

## [`StringContains`](/VBA/StringContains.vba)

Checks if a string contains another string as a substring, with optional case sensitivity.

### Syntax

```vb
StringContains( _
    str1 As String, _
    str2 As String, _
    Optional caseSensitive As Boolean = False _
) As Boolean
```

### Parameters

- `str1`: The main string to search in
- `str2`: The substring to search for
- `caseSensitive`: (_optional_) Boolean flag to enable case-sensitive search. Default is `False`

### Return Value

Returns `True` if `str2` is found within `str1`, `False` otherwise.

### **Use Cases**

- Text validation
- String searching
- Pattern matching without regular expressions
- Case-insensitive text comparisons

### Example

```vb
Dim result As Boolean

result = StringContains("Hello World", "world")
Debug.Print result ' Returns True

result = StringContains("Hello World", "WORLD")
Debug.Print result ' Returns True

result = StringContains("Hello World", "world", True)
Debug.Print result ' Returns False

result = StringContains("Test", "xyz")
Debug.Print result ' Returns False
```

<br>

## [`StringEndsWith`](/VBA/StringEndsWith.vba)

Checks if a string ends with another string, with optional case sensitivity.

### Syntax

```vb
StringEndsWith( _
    str1 As String, _
    str2 As String, _
    Optional caseSensitive As Boolean = False _
) As Boolean
```

### Parameters

- `str1`: The main string to check
- `str2`: The ending string to look for
- `caseSensitive`: (_optional_) Boolean flag to enable case-sensitive comparison. Default is `False`

### Return Value

Returns `True` if `str1` ends with `str2`, `False` otherwise. Also returns `False` if `str2` is longer than `str1`.

### **Use Cases**

- File extension validation
- Text suffix checking
- String pattern matching
- Domain name validation

### Example

```vb
Dim result As Boolean

result = StringEndsWith("Hello World", "world")
Debug.Print result ' Returns True

result = StringEndsWith("Hello World", "WORLD")
Debug.Print result ' Returns True

result = StringEndsWith("Hello World", "World", True)
Debug.Print result ' Returns True

result = StringEndsWith("Test", "xyz")
Debug.Print result ' Returns False
```

<br>

## [`StringStartsWith`](/VBA/StringStartsWith.vba)

Checks whether a string starts with a specified substring, with optional case sensitivity.

### Syntax

```vb
StringStartsWith( _
    str1 As String, _
    str2 As String, _
    Optional caseSensitive As Boolean = False _
) As Boolean
```

### Parameters

- `str1`: The main string to check.
- `str2`: The prefix substring to look for.
- `caseSensitive`: (_optional_) If `True`, comparison is case-sensitive; if `False` (default), comparison is case-insensitive.

### Return Value

Returns `True` if `str1` starts with `str2`; otherwise returns `False`. Also returns `False` if `str2` is longer than `str1`.

### Example

```vb
Dim result As Boolean

result = StringStartsWith("Report.xlsx", "Report")
Debug.Print result ' True

result = StringStartsWith("Report.xlsx", "report")
Debug.Print result ' True (case-insensitive)

result = StringStartsWith("Report.xlsx", "report", True)
Debug.Print result ' False (case-sensitive)

result = StringStartsWith("Test", "LongPrefix")
Debug.Print result ' False
```

<br>

## [`SubstringIsInArray`](/VBA/SubstringIsInArray.vba)

Searches a one-dimensional array for any string element that contains a specified substring and returns `True` on the first match.

### Syntax

```vb
StringStartsWith( _
    str1 As String, _
    str2 As String, _
    Optional caseSensitive As Boolean = False _
) As Boolean
```

### Parameters

- `subStr`: The substring to search for.
- `srcArray`: One-dimensional array containing elements to search.
- `caseSensitive`: (_optional_) If `True`, performs a case-sensitive search; default is `False`.

### Return Value

Returns `True` if any string element in `srcArray` contains `subStr`; otherwise returns `False`.

### Remarks

- Only inspects elements typed as `String`; non-string elements are ignored.

### **Dependencies**

- Depends on the helper function [`StringContains`](#stringcontains) for substring checks.

### Example

```vb
Dim arr As Variant
arr = Array("Hello World", "Sample", "Test")

Debug.Print SubstringIsInArray("world", arr)        ' True (case-insensitive)
Debug.Print SubstringIsInArray("WORLD", arr, True) ' False (case-sensitive)
```

<br>

## [`Summation`](/VBA/Summation.vba)

Computes the numeric summation of a mathematical expression over an integer index range.

### Syntax

```vb
Summation( _
    Expression As String, _
    First As Long, _
    Last As Long _
) As Double
```

### Parameters

- `Expression`: A string representing the math expression in terms of a variable (e.g. `"2*n-1"` or `"1/x^2"`). The function extracts the variable name as the last alphabetical character found in the expression.
- `First`: Starting integer index.
- `Last`: Ending integer index.

### Return Value

Returns the summation's result from expression evaluated for the index running from `First` to `Last`.

### Remarks

- The variable used in Expression is determined by extracting letters from the expression and taking the last letter. Ensure your expression contains the intended variable and that it is the last letter in the expression if multiple letters appear
- Depends on the helper function [`GetLettersOnly`](#getlettersonly) in order to identify the variable in expression

### Examples

```vb
Debug.Print Summation("2*n-1", 1, 10) ' prints 100
Debug.Print Summation("1/x^2", 1, 1000000) ' ≈ 1.64 (approaches π²/6)
Debug.Print Summation("n^2", 1, 5) ' prints 55
```

<br>

## [`TableHasQuery`](/VBA/TableHasQuery.vba)

Checks whether a ListObject (Excel table) has an associated QueryTable.

### Syntax

```vb
TableHasQuery( _
    tbl As ListObject _
) As Boolean
```

### Parameters

- `tbl`: The ListObject (table) to check.

### Return Value

Returns `True` if the table has an associated `QueryTable`; otherwise returns `False`. If `tbl` is `Nothing`, the function returns `False`.

### Example

```vb
Dim tbl As ListObject
Dim hasQuery As Boolean

Set tbl = ThisWorkbook.Worksheets("Sheet1").ListObjects("Table1")
hasQuery = TableHasQuery(tbl)

If hasQuery Then
    Debug.Print "Table has a QueryTable"
Else
    Debug.Print "Table does not have a QueryTable"
End If
```

<br>

## [`WorksheetHasListObject`](/VBA/WorksheetHasListObject.vba)

Checks whether a worksheet contains at least one ListObject (table).

### Syntax

```vb
WorksheetHasListObject( _
    ws As Worksheet _
) As Boolean
```

### Parameters

- `ws`: Worksheet to check for ListObjects.

### Return Value

Returns `True` if the worksheet contains one or more `ListObjects`; otherwise returns `False`.

### Example

```vb
Dim hasTable As Boolean
hasTable = WorksheetHasListObject(ThisWorkbook.Worksheets("Sheet1"))

If hasTable Then
    Debug.Print "Sheet1 contains at least one table."
Else
    Debug.Print "Sheet1 contains no tables."
End If
```