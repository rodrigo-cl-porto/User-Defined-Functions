(romanToConvert as text) as number =>
let
    Chars = Text.ToList(Text.Upper(romanToConvert))
in if not List.MatchesAll(Chars, each List.Contains({"I", "V", "X", "L", "C", "D", "M"}, _)) then
    null
else let
    ListNumbers = List.Transform(
        Chars,
        each if _ = "I" then 1
        else if _ = "V" then 5
        else if _ = "X" then 10
        else if _ = "L" then 50
        else if _ = "C" then 100
        else if _ = "D" then 500
        else if _ = "M" then 1000
        else 0
    ),
    ReversedListNumbers = List.Reverse(ListNumbers),
    N = List.Count(ReversedListNumbers),
    ConvertedToNumber = List.Accumulate(
        {0..N-1},
        0,
        (state as number, i as number) as number =>
        if i = 0 then
            state + ReversedListNumbers{0}
        else if ReversedListNumbers{i} >= ReversedListNumbers{i-1} then
            state + ReversedListNumbers{i}
        else
            state - ReversedListNumbers{i}
in
    ConvertedToNumber
