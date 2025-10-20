let 
    func = (textToModify as nullable text, optional undesirableWords as list) as text =>
        if textToModify = null or textToModify = "" then
            ""
        else let
            LowerText = Text.Lower(textToModify),
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
                "um", "uma", "uns", "umas"
            },
            RemovedWords = if undesirableWords = null or undesirableWords = {} then
                    List.RemoveItems(SplittedText, Stopwords)
                else
                    List.RemoveItems(SplittedText, Stopwords & undesirableWords),
            CombinedText = Text.Combine(RemovedWords, " ")
        in
            CombinedText,
    
    documentation = type function (
        textToRemove as (type nullable text meta [
            Documentation.FieldCaption = "textToRemove",
            Documentation.FieldDescription = "The text string from which stopwords will be removed."
        ]),
        optional undesirableWords as (type {text} meta [
            Documentation.FieldCaption = "undesirableWords",
            Documentation.FieldDescription = "An optional list of additional words to be removed from the text string."
        ])) as text meta [
            Documentation.Name = "Text.RemoveStopwords",
            Documentation.Description = "Removes common Portuguese stopwords from a text string, with an option to specify additional words to remove.",
            Documentation.LongDescription = "This function takes a text string and removes common Portuguese stopwords from it. Users can also provide an optional list of additional undesirable words to be removed. The function processes the text by converting it to lowercase, splitting it into words, removing the specified stopwords and undesirable words, and then recombining the remaining words into a single text string.",
            Documentation.Category = "Text",
            Documentation.Examples = {[
                Description = "Removing stopwords from a simple text.",
                Code = "Text.RemoveStopwords(""Este é um exemplo de texto com algumas stopwords."")",
                Result = """exemplo texto algumas stopwords."""
            ], [
                Description = "Removing stopwords and additional undesirable words from a text.",
                Code = "Text.RemoveStopwords(""Este é um exemplo de texto com algumas stopwords."", {""exemplo"", ""texto""})",
                Result = """algumas stopwords."""
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
