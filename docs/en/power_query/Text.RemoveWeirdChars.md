# [`Text.RemoveWeirdChars`](/src/power_query/Text.RemoveWeirdChars.pq)

Removes special and non-printable characters from a text string, with an option to replace them with spaces.

## Syntax

```fs
Text.RemoveWeirdChars(
    textToClean as text,
    optional replacer as text
) as text
```

## Parameters

- `textToClean`: The text string to be cleaned.
- `replacer` (_optional_): A text string to replace special characters with. If omitted, special characters are replaced by an white space.

## Return Value

Returns the cleaned text with special characters either removed or replaced by the specified replacer.

## Examples

**Example 1**: Cleans text with special characters.

```fs
Text.RemoveWeirdChars("Hello" & Character.FromNumber(0) & "World!")
```

**Result**

```fs
"Hello World!"
```

**Example 2**: Replaces all weird characters if a replacer is specified.

```fs
Text.RemoveWeirdChars("Hello" & Character.FromNumber(0) & "World!", "_")
```

**Result**

```fs
"Hello_World!"
```
