# [`Text.RemoveDoubleSpaces`](/src/power_query/Text.RemoveDoubleSpaces.pq)

Removes consecutive double spaces from a text string, replacing them with single spaces.

## Syntax

```fs
Text.RemoveDoubleSpaces(
    inputText as text
) as text
```

## Parameters

- `inputText`: The text string from which to remove double spaces.

## Return Value

Returns the input text with all consecutive double spaces replaced by single spaces.

## Example

```fs
Text.RemoveDoubleSpaces("This  is   a    test.")
```

**Result**

```fs
"This is a test."
```