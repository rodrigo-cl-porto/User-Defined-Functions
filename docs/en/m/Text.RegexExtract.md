# [`Text.RegexExtract`](/src/m/Text.RegexExtract.pq)

Extracts a substring from a text by using a regular expression pattern.

## Syntax

```fs
Text.RegexExtract(
    textToExtract as text,
    regexPattern as text,
    optional global as logical,
    optional caseInsensitive as logical,
    optional multiline as logical
) as any
```

## Parameters

- `textToExtract`: The input text from which to extract the substring.
- `regexPattern`: The regular expression pattern to use for extraction.
- `global` (_optional_): A logical value indicating whether to extract all matches (`true`) or just the first match (`false`). Default is `false`.
- `caseInsensitive` (_optional_): A logical value indicating whether the regex matching should be case insensitive. Default is `false`.
- `multiline` (_optional_): A logical value indicating whether to treat the input text as multiline. Default is `false`.

## Return Value

Returns the extracted substring(s) based on the regex pattern. If `global` is `true`, returns a list of all matches; otherwise, returns the first match or `null` if no match is found.

## Remarks

- Uses .NET regular expressions for pattern matching.
- If `global` is `true`, returns a list of all matches; otherwise, returns the first match or `null` if no match is found.
- Due to Power Query's JavaScript parser limitations, some advanced regex features like lookbehind '(?<=pattern)' and negative lookbehind '(?<!pattern)' and certain flags (`s`, `u`, `v`, `d`, `y`) are not supported.
- Only the flags `g`, `i`, `m` are available.

## Examples

**Example 1**: Extract patterns which start with "W" and end with "d".

```fs
Text.RegexExtract("Hello World", "W.*d")
```

**Result**

```fs
"World"
```

**Example 2**: Extracts all numbers from a text by activating the `global` flag.

```fs
Text.RegexExtract("abc 123 def 456", "\d+", true)
```

**Result**

```fs
{"123", "456"}
```

**Example 3**: By activating the `multiline` flag, the character "^" and "$" comes to mean, respectively, "start of line" and "end of line" instead of "start of text" and "end of text".

```fs
Text.RegexExtract("Hello#(lf)World", "^W.*?d", false, false, true)
```

**Result**

```fs
"World"
```
