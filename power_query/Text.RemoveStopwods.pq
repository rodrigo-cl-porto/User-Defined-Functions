(textToRemove as text) as text =>
let
    LowerText = Text.Lower(textToRemove),
    SplittedText = Text.Split(LowerText, " "),
    Stopwords = {"de", "do", "da", "dos", "das", "a", "o", "as", "os", "à", "ao", "um", "uma", "uns", "umas", "em", "no", "na", "nos", "nas", "por", "pelo", "pela", "pelos", "pelas", "com", "sem", "para", "pra", "pro", "pras", "pros", "e", "ou", "é", "sobre"},
    RemovedStopwords = List.RemoveItems(SplittedText, Stopwords),
    CombinedText = Text.Combine(RemovedStopwords, " ")
in
    CombinedText
