# [`Text.RegexTest`](/src/m/Text.RegexTest.pq)

Tests whether a text matches a regular expression pattern.

## Syntax

```fs
Text.RegexTest(
    textToTest as text,
    regexPattern as text,
    optional caseInsensitive as logical,
    optional multiline as logical
) as logical
```

## Parameters

- `textToTest`: The input text to be tested against the regex pattern.
- `regexPattern`: The regular expression pattern to test.
- `caseInsensitive` (_optional_): A logical value indicating whether the regex matching should be case insensitive. Default is `false`.
- `multiline` (_optional_): A logical value indicating whether to treat the input text as multiline. Default is `false`.

## Return Value

Returns `true` if the input text matches the regex pattern, `false` otherwise.

## Remarks

- Uses .NET regular expressions for pattern matching.
- Due to Power Query's JavaScript parser limitations, some advanced regex features like lookbehind '(?<=pattern)' and negative lookbehind '(?<!pattern)' and certain flags (`s`, `u`, `v`, `d`, `y`) are not supported.
- Only the flags `i`, `m` are available.

## Examples

**Example 1**: Checks if text contains the word "World".

```fs
Text.RegexTest("Hello World", "World")
```

**Result**

```fs
true
```

**Example 2**: Checks if text contains any digit.

```fs
Text.RegexTest("abc 123", "^\d+$")
```

**Result**

```fs
false
```

**Example 3**: Checks if text contains any digit e-mail.

```fs
Text.RegexTest(
    "My e-mail is example.email@mail.com", 
    "[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}",
    false
)
```

**Result**

```fs
true
```
