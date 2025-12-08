# [`Text.RegexSplit`](/src/power_query/Text.RegexSplit.pq)

Splits a text into a list of substrings based on a regular expression pattern.

## Syntax

```fs
Text.RegexSplit(
    textToSplit as text,
    regexPattern as text,
    optional caseInsensitive as logical,
    optional multiline as logical
) as list
```

## Parameters

- `textToSplit`: The input text to be split.
- `regexPattern`: The regular expression pattern to use as the delimiter for splitting.
- `caseInsensitive` (_optional_): A logical value indicating whether the regex matching should be case insensitive. Default is `false`.
- `multiline` (_optional_): A logical value indicating whether to treat the input text as multiline. Default is `false`.

## Return Value

Returns a list of substrings obtained by splitting the input text at each match of the regex pattern.

## Remarks

- Uses .NET regular expressions for pattern matching.
- Due to Power Query's JavaScript parser limitations, some advanced regex features like lookbehind '(?<=pattern)' and negative lookbehind '(?<!pattern)' and certain flags (`s`, `u`, `v`, `d`, `y`) are not supported.
- Only the flags `i`, `m` are available.

## Examples

**Example 1**: Splits text by comma.

```fs
Text.RegexSplit("apple,banana,cherry", ",")
```

**Result**

```fs
{"apple", "banana", "cherry"}
```

**Example 2**: Splits text by any digit.

```fs
Text.RegexSplit("one1two2three3", "\d")
```

**Result**

```fs
{"one", "two", "three", ""}
```

**Example 3**: Splits text by any line feed.

```fs
Text.RegexSplit("Hello\nWorld", "\n", false, true)
```

**Result**

```fs
{"Hello", "World"}
```
