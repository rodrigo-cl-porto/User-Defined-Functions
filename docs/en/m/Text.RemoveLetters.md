# [`Text.RemoveLetters`](/src/m/Text.RemoveLetters.pq)

Removes all alphabetic characters from a text, leaving only non-letter characters.

## Syntax

```fs
Text.RemoveLetters(
    textToModify as text
) as text
```

## Parameters

- `textToModify`: The text string from which to remove alphabetic characters.

## Return Value

Returns the input text with all alphabetic characters removed.

## Example

```fs
Text.RemoveLetters("Hello123 World!")
```

**Result**

```fs
"123 !"
```