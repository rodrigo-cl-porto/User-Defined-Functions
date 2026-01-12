# [`Text.RemoveAccents`](/src/m/Text.RemoveAccents.pq)

Removes accents from characters in a text string.

## Syntax

```fs
Text.RemoveAccents(
    inputText as text
) as text
```

## Parameters

- `inputText`: The text string from which to remove accents.

## Return Value

Returns the input text with all accented characters replaced by their unaccented equivalents.

## Example

```fs
Text.RemoveAccents("Café")
```

**Result**

```fs
"Cafe"
```
