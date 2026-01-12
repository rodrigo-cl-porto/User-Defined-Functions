# [`Text.RemovePunctuations`](/src/m/Text.RemovePunctuations.pq)

Removes all punctuation characters from a text string.

## Syntax

```fs
Text.RemovePunctuations(
    textToRemove as nullable text,
    optional replacer as text
) as text
```

## Parameters

- `textToRemove`: The text string from which to remove punctuation characters.
- `replacer` (_optional_): A text string to replace punctuation characters with. If omitted, punctuation characters are removed without replacement.

## Return Value

Returns the input text with all punctuation characters removed or replaced by the specified replacer.

## Examples

**Example 1**: Removes all punctuations in a text.

```fs
Text.RemovePunctuations("Hello, World!") 
```

**Result**

```fs
"Hello World"
```

**Example 2**: Replaces all punctuations in a text by a specified character.

```fs
Text.RemovePunctuations("Hello, World!", " ")
```

**Result**

```fs
"Hello  World "
```
