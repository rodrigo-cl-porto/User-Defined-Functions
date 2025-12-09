# [`Text.ExtractNumbers`](/src/m/Text.ExtractNumbers.pq)

Extracts all numeric values from a given text string and returns them as a list of numbers.

## Syntax

```fs
Text.ExtractNumbers(
    inputText as text
) as {number}
```

## Parameters

- `inputText`: The text string from which to extract numeric values.

## Return Value

Returns a list of numbers extracted from the input text. If no numbers are found, returns an empty list.

## Examples

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
