## [`Number.ToRoman`](/src/power_query/Number.ToRoman.pq)

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

### Example

```fs
Number.ToRoman(12)
```

**Result**

```fs
"XII"
```

<br>