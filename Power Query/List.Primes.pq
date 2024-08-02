(n as number) as list =>
if n < 2 then
    {}
else let
    N = Number.RoundDown(n),
    PrimesList = List.Accumulate(
        {2..N},
        {2..N},
        (state as list, current as number) as list =>
        if List.Contains(state, current) then
            List.Select(state, each _ = current or Number.Mod(_, current) <> 0)
        else 
            state
    )
in 
    PrimesList
