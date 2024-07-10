(listNumbers as list) as nullable number =>
let  
    N = List.Count(listNumbers)
in if N < 1 then 
    null
else 
    let  
        ZippedNumbers = List.Zip({{1..N}, listNumbers}),
        Quotient = List.Accumulate(
            ZippedNumbers,
            0,
            (sum as number, current as list) as number =>
            let  
                i = current{0},
                value = current{1}
            in
                sum + i * value
        ),
        Divisor = (1 + N) * N / 2
    in
        Quotient / Divisor
