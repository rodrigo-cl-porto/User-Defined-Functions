(n as number, optional method as nullable text) as list => // Returns the list of prime numbers equals or less than n
if n < 2 then
  {}
else if Text.Upper(method) = "ERATOSTHENES" then

  List.Accumulate(
    {3..n},
    {2} & List.Numbers(3, n/2 - 1, 2),
    (pool as list, candidate as number) as list =>
    List.Select(pool, each _ = candidate or Number.Mod(_, candidate) <> 0)
  )

else if Text.Upper(method) = "TRIAL DIVISION" then

  {2}

else // Apply Dijkistra's approach
  let 
    Pool = List.Accumulate(
      {3..n},
      {[prime=2, multiple=4]},
      (pool as list, candidate as number) as list => 
      let 
        multiples = List.Transform(pool, each [multiple])
      in if candidate < List.Min(multiples) then // It's a prime number!
        pool & {[prime=candidate, multiple=Number.Power(candidate, 2)]}
      else // It's not a prime number :(
        List.Transform(
          pool,
          each if [multiple] = candidate then 
            [prime=[prime], multiple = [multiple] + [prime]]
          else _
        )
    ),
    Primes = List.Transform(Pool, each [prime])
  in
    Primes
