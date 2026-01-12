# [`Text.RegexReplace`](/src/m/Text.RegexReplace.pq)

Replaces substrings in a text that match a regular expression pattern with a specified replacement string.

## Syntax

```fs
Text.RegexReplace(
    textToModify as text,
    regexPattern as text,
    replacer as text,
    optional global as logical,
    optional caseInsensitive as logical,
    optional multiline as logical
) as nullable text
```

## Parameters

- `textToModify`: The input text in which to perform the replacements.
- `regexPattern`: The regular expression pattern to match substrings for replacement.
- `replacer`: The string to replace matched substrings with.
- `global` (_optional_): A logical value indicating whether to replace all occurrences (`true`) or just the first occurrence (`false`). Default is `false`.
- `caseInsensitive` (_optional_): A logical value indicating whether the regex matching should be case insensitive. Default is `false`.
- `multiline` (_optional_): A logical value indicating whether to treat the input text as multiline. Default is `false`.

## Return Value

Returns the modified text with the specified replacements. If no matches are found, returns the original text.

## Remarks

- Uses .NET regular expressions for pattern matching and replacement.
- If `global` is `true`, replaces all matches; otherwise, replaces only the first match.
- Due to Power Query's JavaScript parser limitations, some advanced regex features like lookbehind '(?<=pattern)' and negative lookbehind '(?<!pattern)' and certain flags (`s`, `u`, `v`, `d`, `y`) are not supported.
- Only the flags `g`, `i`, `m` are available.

## Examples

**Example 1**: Replacing text to another one.

```fs
Text.RegexReplace("Hello World", "World", "Universe")
```

**Result**

```fs
"Hello Universe"
```

**Example 2**: Replacing all numbers in text to word "number".

```fs
Text.RegexReplace("abc 123 def 456", "\d+", "number", true)
```

**Result**

```fs
"abc number def number"
```

**Example 3**: Replacing all words at start of a line which start with a "W" and end with a "d" by "Everyone".

```fs
Text.RegexReplace("Hello#(lf)World", "^W\w*?d", "Everyone", false, false, true)
```

**Result**

```fs
"Hello#(lf)Everyone"
```
