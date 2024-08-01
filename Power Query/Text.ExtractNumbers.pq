(textToExtract as nullable text) as list =>
if textToExtract = null or textToExtract = "" then
    {}
else let
    Digits = {"0".."9", " "},
    ExtractedDigits = List.Select(Text.ToList(textToExtract), each List.Contains(Digits, _)),
    ExtractedNumbers = List.RemoveItems(Text.Split(Text.Combine(ExtractedDigits), " "), {"", null}),
    ConvertedToListNumbers = List.Transform(ExtractedNumbers, Number.FromText)
in 
    ConvertedToListNumbers
