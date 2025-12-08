# [`Text.CountChar`](/src/power_query/Text.CountChar.pq)

Counts the occurrences of a specific character in a given text string.

## Syntax

```fs
Text.CountChar(
    textToCount as nullable text,
    charToCount as text
) as number
```

## Parameters

- `textToCount`: The text string in which to count occurrences of the character.
- `charToCount`: The character to count within the text string.

## Return Value

Returns a number representing the count of occurrences of the specified character in the input text. If the input text is null, returns 0.

## Remarks

- The function is case-sensitive; 'a' and 'A' are considered different characters.
- If `charToCount` is an empty string, the function returns 0.

## Examples

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
