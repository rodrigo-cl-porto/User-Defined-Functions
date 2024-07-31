(textToRemove as text, optional UndesirableWords as nullable list) as text =>
let
    LowerText = Text.Lower(textToRemove),
    SplittedText = Text.Split(LowerText, " "),
    Stopwords = {
        "a", "o", "as", "os",
        "à", "às", "ao", "aos",
        "aqui", "ali", "acolá",
        "além", "aquém",
        "aquilo", "àquilo", "aquele", "aqueles", "àquele", "àqueles", "aquela", "aquelas", "àquela", "àquelas",
        "cada",
        "com", "sem",
        "de", "do", "da", "dos", "das",
        "e", "ou", "mas", "entretanto", "conquanto", "desde",
        "em", "no", "na", "nos", "nas",
        "era", "eras", "éramos", "éreis", "eram",
        "esse", "essa", "esses", "essas", "isso", "nisso",
        "este", "esta", "estes", "estas", "isto", "nisto",
        "estar",
        "estou", "estás", "está", "estamos", "estás", 
        "eu", "tu", "você", "ele", "ela", "nós", "vós", "vocês", "eles", "elas",
        "existir", "existe", "existem",
        "fui", "foste", "foi", "fomos", "fostes", "foram",
        "fora", "foras", "fôramos", "fôreis",
        "fosse", "fosses", "fôssemos", "fôsseis", "fossem",
        "for", "fores", "formos", "fordes", "forem",
        "mas", "porém", "entretanto", "todavia", "apesar", "mesmo",
        "para", "pro", "pra", "pros", "pras",
        "por", "pôr", "pelo", "pela", "pelos", "pelas",
        "seu", "sua", "seus", "suas",
        "nosso", "nossa", "nossos", "nossas",
        "qual", "quais", "qualquer", "quaisquer",
        "que", "quê", "quem", "como", "quando", "porque", "porquê",
        "haver", "há", "houve", "haverá", "hão",
        "se", "sê",
        "seja", "sejas", "sejamos", "sejais", "sejam",
        "ser", "seres", "sermos", "serdes", "serem",
        "serei", "serás", "será", "seremos", "sereis", "serão",
        "seria", "serias", "seríamos", "seríeis", "seriam",
        "sou", "és", "é", "somos", "sois", "são",
        "sobre", "sob",
        "ter", "tinha", "tem", "têm", "terá", "teria",
        "teu", "tua", "teus", "tuas",
        "um", "uma", "uns", "umas",
    },
    RemovedWords = 
    if UndesirableWords = null or UndesirableWords = {} then
        List.RemoveItems(SplittedText, Stopwords)
    else
        List.RemoveItems(SplittedText, Stopwords & UndesirableWords)
    ,
    CombinedText = Text.Combine(RemovedWords, " ")
in
    CombinedText
