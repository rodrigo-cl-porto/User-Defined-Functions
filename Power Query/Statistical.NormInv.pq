let
    func = (probability as number, optional mean as number, optional sd as number) as number => let
        p = if probability < 0 then 0
            else if probability > 1 then 1
            else probability,
        RationalApproximation = (p as number) as number => let
            a = {
                -39.69683028665376,
                220.9460984245205,
                -275.9285104469687,
                138.3577518672690,
                -30.66479806614716,
                2.506628277459239
            },
            b = {
                -54.47609879822406,
                161.5858368580409,
                -155.6989798598866,
                66.80131188771972,
                -13.28068155288572
            },
            c = {
                -0.007784894002430293,
                -0.3223964580411365,
                -2.400758277161838,
                -2.549732539343734,
                4.374664141464968,
                2.938163982698783
            },
            d = {
                0.007784695709041462,
                0.3224671290700398,
                2.445134137142996,
                3.754408661907416
            },

            // Define break-points.
            p_low  = 0.02425,
            p_high = 0.97575,
            
            z = if p < p_low then let // Lower region.
                    q = Number.Sqrt(-2 * Number.Ln(p)),
                    z = (((((c{0}*q+c{1})*q+c{2})*q+c{3})*q+c{4})*q+c{5}) / ((((d{0}*q+d{1})*q+d{2})*q+d{3})*q+1)
                    in z
                else if p <= p_high then let // Central region.
                    q = p - 0.5,
                    r = q*q,
                    z = (((((a{0}*r+a{1})*r+a{2})*r+a{3})*r+a{4})*r+a{5})*q / (((((b{0}*r+b{1})*r+b{2})*r+b{3})*r+b{4})*r+1)
                    in z
                else let // Upper region.
                    q = Number.Sqrt(-2 * Number.Ln(1-p)),
                    z = -(((((c{0}*q+c{1})*q+c{2})*q+c{3})*q+c{4})*q+c{5}) / ((((d{0}*q+d{1})*q+d{2})*q+d{3})*q+1)
                    in z
            in z,

        CDF = (z as number) as number => let
            GaussianIntegral = (z as number) as number => let
                GaussianQuadratureTable = #table(
                    type table [i = number, w = number, t = number], {
                    { 1,  0.0640568928626056, 0.1279381953467522},
                    { 2, -0.0640568928626056, 0.1279381953467522},
                    { 3,  0.1911188674736163, 0.1258374563468283},
                    { 4, -0.1911188674736163, 0.1258374563468283},
                    { 5,  0.3150426796961634, 0.1216704729278034},
                    { 6, -0.3150426796961634, 0.1216704729278034},
                    { 7,  0.4337935076260451, 0.1155056680537256},
                    { 8, -0.4337935076260451, 0.1155056680537256},
                    { 9,  0.5454214713888396, 0.1074442701159656},
                    {10, -0.5454214713888396, 0.1074442701159656},
                    {11,  0.6480936519369755, 0.0976186521041139},
                    {12, -0.6480936519369755, 0.0976186521041139},
                    {13,  0.7401241915785544, 0.0861901615319533},
                    {14, -0.7401241915785544, 0.0861901615319533},
                    {15,  0.8200019859739029, 0.0733464814110803},
                    {16, -0.8200019859739029, 0.0733464814110803},
                    {17,  0.8864155270044011, 0.0592985849154368},
                    {18, -0.8864155270044011, 0.0592985849154368},
                    {19,  0.9382745520027328, 0.0442774388174198},
                    {20, -0.9382745520027328, 0.0442774388174198},
                    {21,  0.9747285559713095, 0.0285313886289337},
                    {22, -0.9747285559713095, 0.0285313886289337},
                    {23,  0.9951872199970213, 0.0123412297999872},
                    {24, -0.9951872199970213, 0.0123412297999872}
                }), // Source: https://pomax.github.io/bezierinfo/legendre-gauss.html
                AddedGaussianFunction = Table.AddColumn(GaussianQuadratureTable, "f(t)", each Number.Exp(-0.125 * Number.Power(z*([t]+1), 2)), Decimal.Type),
                AddedProduct = Table.AddColumn(AddedGaussianFunction, "w * f(t)", each [w] * [#"f(t)"], type number),
                Summation = List.Sum(AddedProduct[#"w * f(t)"], Precision.Decimal),
                GaussianQuadrature = Number.Sqrt(2/Number.PI) * (z/4) * Summation
                in GaussianQuadrature,
            AccumulatedProbability = 0.5 + GaussianIntegral(z)
            in AccumulatedProbability,
        
        FDP = (z as number) as number => Number.Exp(-0.5*z*z) / Number.Sqrt(2*Number.PI),

        // Calculating the Z-Score
        z = if p = 0 then
                Number.NegativeInfinity
            else if p = 1 then
                Number.PositiveInfinity
            else let
                z0 = RationalApproximation(p),
                z = List.Accumulate({1..2}, z0, (zi, i) => zi - (CDF(zi) - p) / FDP(zi))
                in z,
        
        // Calculating the x value
        mu = if mean is null then 0 else mean,
        sigma = if sd is null then 1 else sd,
        x = mu + z*sigma
        in x,

    documentation = type function (
        probability as (type number meta [
            Documentation.FieldCaption = "probability",
            Documentation.FieldDescription = ""
        ]),
        optional mean as (type number meta [
            Documentation.FieldCaption = "mean",
            Documentation.FieldDescription = "The mean of the distribution. Defaults to 0 if not provided."
        ]),
        optional sd as (type number meta [
            Documentation.FieldCaption = "standard deviation",
            Documentation.FieldDescription = "The standard deviation of the distribution. Defaults to 1 if not provided."
        ])) as number meta [
            Documentation.Name = "Statistical.NormInv",
            Documentation.Description = "Returns the inverse of the cumulative distribution function (CDF) of the normal distribution.",
            Documentation.LongDescription = "Returns the inverse of the cumulative distribution function (CDF) of the normal distribution.",
            Documentation.Category = "Statistical",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Source = "https://web.archive.org/web/20151030215612/http://home.online.no/~pjacklam/notes/invnorm/",
            Documentation.Version = "1.0.1",
            Documentation.Examples = {[
                Description= "Returns x such as P(X ≤ x) = p for a normal distribution with given mean and standard deviation.",
                Code= "Statistical.NormInv(0.9, 100, 15)",
                Result = "119,22327346210234"
            ], [
                Description= "If neither mean nor standard deviation are informed, returns the value z such that P(Z ≤ z) equals the probability under the standard normal distribution.",
                Code= "Statistical.NormInv(0.9)",
                Result = "1,2815515641401563"
            ]}
        ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction