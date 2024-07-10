(X as list, Y as list) as number =>
let 
    N = List.Count(X)
in if N > 0 and List.Count(Y) = N then
    let
        SumX = List.Sum(X),
        SumY = List.Sum(Y),
        SumXY = List.Accumulate({0..N-1}, 0, (sum, i) => sum + X{i} * Y{i}),
        SumX2 = List.Accumulate({0..N-1}, 0, (sum, i) => sum + Number.Power(X{i}, 2))
    in  
        (SumY * SumX2 - SumX * SumXY) / (N * SumX2 - Number.Power(SumX, 2))
else
    null
