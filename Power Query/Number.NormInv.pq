let 
    func = (probability as number, mean as number, std as number) as number => 
    let
        // Source for NormalCDFInverse: https://www.johndcook.com/blog/normal_cdf_inverse/
        p = if probability >= 0 and probability <= 1 then
            probability
        else
            error Error.Record(
                "Value Error",
                "Probability must be between 0 and 1.",
                "Received: " & Number.ToText(probability)
            ),
        RationalApproximation = (t as number) as number =>
        let 
            c = {2.515517, 0.802853, 0.010328},
            d = {1.432788, 0.189269, 0.001308},
            z = t - ((c{2} * t + c{1}) * t + c{0}) / (((d{2} * t + d{1}) * t + d{0}) * t + 1)
        in
            z,
        t = if p < 0.5 then
            Number.Sqrt(-2 * Number.Log(p)) 
        else
            Number.Sqrt(-2 * Number.Log(1 - p)),
        z = if p = 0.5 then 0 else RationalApproximation(t),
        DenormCDFInverse = mean + z * std
    in
        DenormCDFInverse,
    documentation = type function (
        probability as (type number meta [
            Documentation.FieldCaption = "Probability",
            Documentation.FieldDescription = "Probability value (between 0 and 1) to calculate the inverse of the normal distribution."
        ]),
        mean as (type number meta [
            Documentation.FieldCaption = "Mean",
            Documentation.FieldDescription = "The mean of normal distribution"
        ]),
        std as (type number meta [
            Documentation.FieldCaption = "Standar Deviation",
            Documentation.FieldDescription = "The standard deviation of normal distribution"
        ])
    ) as number meta [
        Documentation.Name =  "Number.NormInv",
        Documentation.Description = "Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.",
        Documentation.LongDescription = "This function calculates the value corresponding to the given probability in the normal distribution, 
            adjusted by the given mean and standard deviation.
            It is equivalent to the Excel NORM.INV function.
            For more details, see: https://support.microsoft.com/en-us/office/norm-inv-function-54b30935-fee7-493c-bedb-2278a9db7e13",
        Documentation.Category = "Statistical",
        Documentation.Source = "www.TheBIccountant.com",
        Documentation.Version = "1.0.1",
        Documentation.Author = "Imke Feldmann",
        Documentation.Examples = {[
            Description = "Calculating the value for accumulated distribution of 95%",
            Code = "Number.NormInv(0.95, 100, 15)",
            Result = "124.6782 (roughly)"
        ], [
            Description = "Calculating the Z Score for probability of 60%",
            Code = "Number.NormInv(0.6, 0, 1)",
            Result = "0.2429 (roughly)"
        ]}
    ],
    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
