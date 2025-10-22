let
    func = (x as number, optional mean as number, optional sd as number, optional cumulative as logical) as number => let
        mu = if mean is null then 0 else mean,
        sigma = if sd is null then 1 else sd,
        z = (x - mu)/sigma,
        result = if cumulative = false then
            1/Number.Sqrt(2*Number.PI) * Number.Exp(-Number.Power(z, 2)/2)
        else let
            GaussianIntegral = (z as number) as number => let
                GaussianQuadratureTable = #table(
                    type table [i = number, t = number, w = number], {
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
                AddedGaussianFunction = Table.AddColumn(GaussianQuadratureTable, "f(t)", each Number.Exp(-Number.Power(z*([t]+1), 2)/8), type number),
                AddedProduct = Table.AddColumn(AddedGaussianFunction, "wf(t)", each [w] * [#"f(t)"], type number),
                GaussianQuadrature = (Number.Sqrt(2)/4) * z * List.Sum(AddedProduct[#"wf(t)"])
                in GaussianQuadrature,
            AccumulatedGaussian = 0.5 + 1/Number.Sqrt(Number.PI) * GaussianIntegral(z)
            in AccumulatedGaussian 
        in result,
    
    documentation = type function (
        x as (type number meta [
            Documentation.FieldCaption = "x",
            Documentation.FieldDescription = "The value for which the normal distribution will be evaluated."
        ]),
        optional mean as (type number meta [
            Documentation.FieldCaption = "mean",
            Documentation.FieldDescription = "The mean of the distribution. Defaults to 0 if not provided."
        ]),
        optional sd as (type number meta [
            Documentation.FieldCaption = "standard deviation",
            Documentation.FieldDescription = "The standard deviation of the distribution. Defaults to 1 if not provided."
        ]),
        optional cumulative as (type logical meta [
            Documentation.FieldCaption = "cumulative",
            Documentation.FieldDescription = "Logical value indicating whether to return the cumulative distribution (true) or the probability density (false). The default is true",
            Documentation.AllowedValues = {true, false},
            Documentation.DefaultValue = true
        ])) as number meta [
            Documentation.Name = "Statistical.NormDist",
            Documentation.Description = "Calculates the value of the normal distribution for a given input x.",
            Documentation.LongDescription = "Calculates the value of the normal distribution (also known as Gaussian distribution) for a given input x. It supports both the probability density function (PDF) and the cumulative distribution function (CDF), depending on the cumulative parameter.",
            Documentation.Category = "Statistical",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Version = "1.0.0",
            Documentation.Examples = {[
                Description = "Calculating the cumulative probability for a value of x in a normal distribution with provided mean and standard deviation.",
                Code = "Statistical.NormDist(100, 80, 10)",
                Result = "0.9772498680518209"
            ], [
                Description = "In order to calculate the standard normal CDF, just don't input any mean nor standard deviation.",
                Code = "Statistical.NormDist(1.96)",
                Result = "0.97500210485177963"
            ], [
                Description = "Calculating the standard normal PDF",
                Code = "Statistical.NormDist(1.96, null, null, false)",
                Result = "0.058440944333451469"
            ]}
        ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction