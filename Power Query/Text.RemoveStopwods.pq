(textToRemove as nullable text, optional UndesirableWords as nullable list) as text =>
if textToRemove = null or textToRemove = "" then
    ""
else let
    LowerText = Text.Lower(textToRemove),
    SplittedText = Text.Split(LowerText, " "),
    Stopwords = {
        "a", "o", "as", "os",
        "à", "às", "ao", "aos",
        "aqui", "ali", "acolá", "assim",
        "além", "aquém",
        "aquilo", "àquilo", "aquele", "àquele", "aqueles", "àqueles", "aquela", "àquela", "aquelas", "àquelas",
        "cada", "coisa",
        "com", "sem",
        "de", "do", "da", "dos", "das",
        "e", "ou", "mas", "entretanto", "conquanto", "desde",
        "em", "no", "na", "nos", "nas",
        "era", "eras", "éramos", "éreis", "eram",
        "esse", "essa", "esses", "essas", "isso", "nisso",
        "este", "esta", "estes", "estas", "isto", "nisto",
        "estar", "estai", "estardes",
        "estou", "estás", "está", "estamos", "estás", "estão",
        "estava", "estavas", "estávamos", "estáveis", "estavam",
        "estive", "estiveste", "estivemos", "estivestes", "estiveram",
        "estivera", "estiveras", "estivéramos", "estivéreis", "estiveram",
        "estarei", "estarás", "estará", "estaremos", "estareis", "estarão",
        "estaria", "estarias", "estaríamos", "estaríeis", "estariam",
        "esteja", "estejas", "estejamos", "estejais", "estejam",
        "estivesse", "estivesses", "estivéssemos", "estivésses", "estivessem",
        "estiver", "estiveres", "estivermos", "estiverdes", "estiverem",
        "eu", "tu", "você", "ele", "ela", "nós", "vós", "vocês", "eles", "elas",
        "existir", "existe", "existem",
        "fui", "foste", "foi", "fomos", "fostes", "foram",
        "fora", "foras", "fôramos", "fôreis",
        "fosse", "fosses", "fôssemos", "fôsseis", "fossem",
        "for", "fores", "formos", "fordes", "forem",
        "mas", "porém", "contudo", "entretanto", "todavia", "apesar",
        "mesmo", "mesma", "mesmos", "mesmas",
        "para", "pro", "pra", "pros", "pras",
        "por", "pôr", "pelo", "pela", "pelos", "pelas",
        "seu", "sua", "seus", "suas",
        "nosso", "nossa", "nossos", "nossas",
        "qual", "quais", "qualquer", "quaisquer",
        "que", "quê", "quem", "como", "quando", "onde", "porque",
        "haver", "há", "houve", "haverá", "hão",
        "se", "sê", "então", "logo", "portanto",
        "seja", "sejas", "sejamos", "sejais", "sejam",
        "ser", "seres", "sermos", "serdes", "serem",
        "serei", "serás", "será", "seremos", "sereis", "serão",
        "seria", "serias", "seríamos", "seríeis", "seriam",
        "sou", "és", "é", "somos", "sois", "são",
        "sobre", "sob",
        "ter", "tinha", "tem", "têm", "terá", "teria", "teriam",
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
