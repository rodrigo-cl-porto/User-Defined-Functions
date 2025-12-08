# [`Number.IsPrime`](/src/power_query/Number.IsPrime.pq)

Checks if a given number is prime.

## Syntax

```fs
Number.IsPrime(
    value as Int64.Type
) as logical
```

## Parameters

- `value`: A number to check.

## Return Value

Returns `true` if the number is prime, `false` otherwise.

## Examples

**Example 1**: Check if 7 is prime

```fs
Number.IsPrime(7)
```

**Result**

```fs
true
```

**Example 2**: Check if 100 is prime

```fs
Number.IsPrime(100)
```

**Result**

```fs
false
```

## Credits

- Author: Abigail
- Source: [Abigail's regex to test for prime numbers](http://test.neilk.net/blog/2000/06/01/abigails-regex-to-test-for-prime-numbers/)
- YouTube Video: [How on Earth does ^.?$|^(..+?)\1+$ produce primes?](https://www.youtube.com/watch?v=5vbk0TwkokM)