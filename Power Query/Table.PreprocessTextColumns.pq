(tbl as table, optional columns as nullable list, optional TextFormat as nullable text, optional wordsToRemove as nullable list, optional removeAccents as nullable logical, optional removePunctuation as logical, optional removeStopwords as nullable logical, optional removeDigits as nullable logical) as table =>
let 
    TextColumns = 
    if columns = null or columns = {} or not Value.Is(columns, List.Type) then
        Table.ColumnsOfType(tbl, {type text, type nullable text})
    else
        columns
    ,
    ListColumnsTransformations = List.Transform(
        TextColumns,
        each
        {
            _,
            each if _ <> null then
                let 
                    LineFeedReplaced = Text.Replace(_, "#(lf)", " "),
                    CleanedText = Text.Clean(LineFeedReplaced),
                    WeirdSpaceReplaced = Text.Replace(CleanedText, Character.FromNumber(160), " "),
                    TrimmedText = Text.Trim(WeirdSpaceReplaced),
                    RemovedDuplicatedSpaces = 
                    let
                        SplittedText = Text.Split(TrimmedText, " "),
                        RemovedEmpties = List.RemoveItems(SplittedText, {"", null}),
                        CombinedText = Text.Combine(RemovedEmpties, " ")
                    in
                        CombinedText
                    ,
                    RemovedAccents = 
                    if removeAccents = true then 
                        let
                            textBinary = Text.ToBinary (RemovedDuplicatedSpaces, 1251),
                            textASCII  = Text.FromBinary(textBinary, TextEncoding.Ascii)
                        in
                            textASCII
                    else 
                        RemovedDuplicatedSpaces
                    ,
                    RemovedDigits =
                    if removeDigits = true then
                        let
                            DigitsToRemove = Text.ToList("0123456789"),
                            ListChars = Text.ToList(RemovedAccents),
                            ListSelectedChars = List.Select(ListChars, each not List.Contains(DigitsToRemove, _)),
                            TextWithSelectedChars = Text.Combine(ListSelectedChars)
                        in 
                            TextWithSelectedChars
                    else 
                        RemovedAccents
                    ,
                    RemovedPunctuation = 
                    if removePunctuation = true then
                        let
                            PuncsToRemove = Text.ToList("!""#$%&'()*+,-./:;<=>?@[\]^_`{|}~"),
                            ListChars = Text.ToList(RemovedDigits),
                            ListSelectedChars = List.Select(ListChars, each not List.Contains(PuncsToRemove, _)),
                            TextWithSelectedChars = Text.Combine(ListSelectedChars)
                        in 
                            TextWithSelectedChars
                    else 
                        RemovedDigits
                    ,
                    RemovedWords = 
                    let
                        WordsToRemove = if wordsToRemove <> null then wordsToRemove else {},
                        AddedStopwordsToRemove = 
                        if removeStopwords = true then 
                            let   
                                Stopwords = {"de", "do", "da", "dos", "das", "a", "o", "as", "os", "ao", "à", "um", "uma", "uns", "umas", "em", "no", "na", "nos", "nas", "por", "pelo", "pela", "pelos", "pelas", "com", "sem", "para", "pra", "pro", "pras", "pros", "e", "ou", "é", "sobre"},
                                CombinedList = List.Combine(WordsToRemove, Stopwords)
                            in 
                                CombinedList
                        else
                            WordsToRemove
                        ,
                        RemovedUndesirableWords = 
                        if AddedStopwordsToRemove = {} then 
                            RemovedPunctuation
                        else let 
                            LowerText = Text.Lower(RemovedPunctuation),
                            SplittedText = Text.Split(LowerText, " "),
                            RemovedUndesirableWords = List.RemoveItems(SplittedText, AddedStopwordsToRemove),
                            CombinedText = Text.Combine(RemovedUndesirableWords, " ")
                        in 
                            CombinedText
                    in  
                        RemovedUndesirableWords
                    ,
                    CasedText = 
                    if Text.Upper(TextFormat) = "PROPER" then
                        Text.Proper(RemovedWords)
                    else if Text.Upper(TextFormat) = "LOWER" then
                        Text.Lower(RemovedWords)
                    else if Text.Upper(TextFormat) = "UPPER" then
                        Text.Upper(RemovedWords)
                    else
                        RemovedWords
                in 
                    CasedText
            else ""
            ,
            type text
        }
    ),
    TransformedTable = Table.TransformColumns(tbl, ListColumnsTransformations)
in 
    TransformedTable
