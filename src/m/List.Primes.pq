let
    func = (n as number) as list => // Returns the list of prime numbers equals or less than n
    if n < 2 then
        {}
    else if n < 1000 then // Apply Eratostenes
        List.Accumulate(
            {3..n},
            {2} & List.Numbers(3, n/2 - 1, 2),
            (pool as list, candidate as number) as list =>
                List.Select(pool, each _ = candidate or Number.Mod(_, candidate) <> 0)
        )
    else let // Apply Dijkistra's algorithm
        Pool = List.Accumulate(
            {3..n},
            {[prime=2, multiple=4]},
            (pool as list, candidate as number) as list => let 
                multiples = List.Transform(pool, each [multiple])
                in if candidate < List.Min(multiples) then // It's a prime number!
                    pool & {[prime=candidate, multiple=Number.Power(candidate, 2)]}
                else // It's not a prime number :(
                    List.Transform(
                        pool,
                        each if [multiple] = candidate then 
                            [
                                prime=[prime], 
                                multiple = [multiple] + [prime]
                            ]
                        else 
                            _
                    )
        ),
        Primes = List.Transform(Pool, each [prime])
        in Primes,

    documentation = type function (
        n as (Int64.Type meta [
            Documentation.FieldCaption = "number",
            Documentation.FieldDescription = "A positive integer; if n < 2, the function returns an empty list."
        ])) as {number} meta [
            Documentation.Name = "List.Primes",
            Documentation.Description = "Returns a list of prime numbers less than or equal to a given number n.",
            Documentation.LongDescription = "Returns a list of prime numbers less than or equal to a given number n. It uses the Sieve of Eratosthenes for small values and a variation of Dijkstra's algorithm for larger values to efficiently generate prime numbers.",
            Documentation.Category = "List",
            Documentation.Version = "1.0.0",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Examples = {[
                Description = "Primes up to 10",
                Code = "List.Primes(10)",
                Result = "{2, 3, 5, 7}"
            ], [
                Description = "Primes up to 30",
                Code = "List.Primes(30)",
                Result = "{2, 3, 5, 7, 11, 13, 17, 19, 23, 29}"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in  
    documentedFunction