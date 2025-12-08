# [`Text.RemoveStopwords`](/src/power_query/Text.RemoveStopwords.pq)

Removes common Portuguese stopwords from a text string to enhance text analysis.

## Syntax

```fs
Text.RemoveStopwords(
    textToModify as nullable text,
    optional undesirableWords as list
) as text
```

## Parameters

- `textToModify`: The text string from which to remove stopwords.
- `undesirableWords` (_optional_): A list of additional words to remove from the text. Default is an empty list.

## Return Value

Returns the input text with all Portuguese stopwords and any additional specified words removed.

## Examples

**Example 1**: Removes all stopwords.

```fs
Text.RemoveStopwords("This is an example of text to remove stopwords.")
```

**Result**

```fs
"example text remove stopwords."
```

**Example 2**: Also removes any undesireble words if specified.

```fs
Text.RemoveStopwords(
    "This is an example of text to remove stopwords.",
    {"exemple", "text"}
)
```

**Result**

```fs
"remove stopwords."
```
