# [`Text.RemoveNumerals`](/src/m/Text.RemoveNumerals.pq)

Removes all numeric characters from a text string, with an option to also remove Roman numerals.

## Syntax

```fs
Text.RemoveNumerals(
    textToRemove as text,
    optional removeRomanNumerals as logical
) as text
```

## Parameters

- `textToRemove`: The text string from which to remove numeric characters.
- `removeRomanNumerals` (_optional_): A logical value indicating whether to also remove Roman numeral characters (I, V, X, L, C, D, M). Default is `false`.

## Return Value

Returns the input text with all numeric characters (and optionally Roman numerals) removed.

## Examples

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