(numbersList as list, optional populational as nullable logical) as number =>
if populational = null or populational = false then // uses sampling variance formula, which takes (n-1) on the denominator.
    Number.Power(List.StandardDeviation(numbersList), 2)
else // uses populational variance formula, which takes (n) on the denominator
    List.Covariance(numbersList, numbersList)
