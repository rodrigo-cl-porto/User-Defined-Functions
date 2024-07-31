(textToRemove as text, optional UndesirableWords as nullable list) as text =>
let
    LowerText = Text.Lower(textToRemove),
    SplittedText = Text.Split(LowerText, " "),
    Stopwords = {"de", "do", "da", "dos", "das", "a", "o", "as", "os", "à", "ao", "um", "uma", "uns", "umas", "em", "no", "na", "nos", "nas", "por", "pelo", "pela", "pelos", "pelas", "com", "sem", "para", "pra", "pro", "pras", "pros", "e", "ou", "é", "sobre", "seu", "sua", "seus", "suas", "nosso", "nossa", "nossos", "nossas", "haver", "há", "houve", "houveram", "hão"},
    RemovedWords = 
    if UndesirableWords = null or UndesirableWords = {} then
        List.RemoveItems(SplittedText, Stopwords)
    else
        List.RemoveItems(SplittedText, Stopwords & UndesirableWords)
    ,
    CombinedText = Text.Combine(RemovedWords, " ")
in
    CombinedText
