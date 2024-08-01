(textToRemove as nullable text, optional replacer as nullable text) as text =>
if textToRemove = null or textToRemove = "" then
    ""
else let
    Punctuations = Text.ToList("!""#$%&'()*+,-./:;<=>?@[]\^_`´{|}~ºª°§"),
    Chars = Text.ToList(textToRemove),
    ListWithoutPunctuation = List.Transform(
        Chars,
        each if List.Contains(Punctuations, _) then
            if replacer = null then "" else replacer
        else 
            _
    ),
    TextWithoutPunctuation = Text.Combine(ListWithoutPunctuation),
    SplittedText = Text.Split(TextWithoutPunctuation, " "),
    RemovedEmpties = List.RemoveItems(SplittedText, {"", null}),
    RemovedDoubleSpaces = Text.Combine(RemovedEmpties, " ")
in
    RemovedDoubleSpaces
