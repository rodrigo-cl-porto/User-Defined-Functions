(myText as text, optional replacer as nullable text) as text =>
let
    Punctuations = Text.ToList("!""#$%&'()*+,-./:;<=>?@[]\^_`´{|}~ºª°§"),
    myListText = Text.ToList(myText),
    ListWithoutPunctuation = List.Transform(
        myListText,
        each if List.Contains(Punctuations, _) then
            if replacer = null then "" else replacer
        else 
            _
    ),
    TextWithoutPunctuation = Text.Combine(ListWithoutPunctuation),
    SplittedText = Text.Split(TextWithoutPunctuation, " "),
    RemovedEmpties = List.RemoveItems(SplittedText, {null, ""}),
    RemovedDoubleSpaces = Text.Combine(RemovedEmpties, " ")
in
    RemovedDoubleSpaces
