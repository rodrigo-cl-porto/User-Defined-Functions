let
    func = (x as number, optional mean as number, optional sd as number, optional cumulative as logical) as number => let
        mu = if mean is null then 0 else mean,
        sigma = if sd is null then 1 else sd,
        z = (x - mu)/sigma,
        // Gaussian Probability Density Function
        PDF = (z as number) as number => Number.Exp(-0.5*z*z) / Number.Sqrt(2 * Number.PI),
        GaussianIntegral = (z as number) as number => let
            n = 100000,
            DeltaZ = z / n,
            Summation = List.Accumulate(
                {1..n},
                0,
                (sum as number, i as number) => let
                    t_i = (i-0.5) * DeltaZ
                    in sum + Number.Exp(-0.5*t_i*t_i)
            ),
            RiemannSum = (DeltaZ / Number.Sqrt(2 * Number.PI)) * Summation
            in RiemannSum,
        // Gaussian Cumulative Distribution Function
        CDF = (z as number) as number => 0.5 + GaussianIntegral(z),
        TailCut = 8,
        Result = if cumulative = false then
                PDF(z)
            else if z >= -TailCut and z <= TailCut then
                CDF(z)
            else if z > TailCut then
                1
            else
                0
        in Result,

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
                Result = "0.97724986805182079"
            ], [
                Description = "Calculating the normal PDF for given mean and standard deviation.",
                Code= "Statistical.NormDist(100, 80, 10, false)",
                Result= "0.0539909665131881"
            ], [
                Description = "In order to calculate the standard normal CDF, just don't input any mean nor standard deviation.",
                Code = "Statistical.NormDist(1.96)",
                Result = "0.97500210485177974"
            ], [
                Description = "Calculating the standard normal PDF",
                Code = "Statistical.NormDist(1.96, null, null, false)",
                Result = "0.058440944333451476"
            ]}
        ],
    
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction