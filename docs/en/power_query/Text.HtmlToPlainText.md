# [`Text.HtmlToPlainText.pq`](/src/power_query/Text.HtmlToPlainText.pq)

Converts HTML content to plain text by stripping HTML tags.

## Syntax

```fs
Text.HtmlToPlainText(
    htmlText as text
) as text
```

## Parameters

- `htmlText`: The HTML text to be converted to plain text.

## Return Value

Returns the plain text content extracted from the HTML input.

## Example

```fs
Text.HtmlToPlainText("<p>Hello <b>World</b>!</p>")
```

**Result**

```fs
"Hello World!"
```
