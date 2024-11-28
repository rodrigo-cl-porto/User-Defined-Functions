(HtmlText as text) as text =>
let
    RemoveHtmlTags = 
        each let
            Tag = Text.BetweenDelimiters(_, "<", ">")
        in 
            if Tag = "" then _ else @RemoveHtmlTags(Text.Replace(_, "<" & Tag & ">", "")),
    ReplacedPunctuations = List.Accumulate(
        {
            {"&aacute;", "á"},
            {"&Aacute;", "Á"},
            {"&atilde;", "ã"},
            {"&Atilde;", "Ã"},
            {"&acirc;",	"â"},
            {"&Acirc;",	"Â"},
            {"&agrave;", "à"},
            {"&eacute;", "é"},
            {"&Eacute;", "É"},
            {"&ecirc;",	"ê"},
            {"&Ecirc;",	"Ê"},
            {"&iacute;", "í"},
            {"&Iacute;", "Í"},
            {"&oacute;", "ó"},
            {"&Oacute;", "Ó"},
            {"&otilde;", "õ"},
            {"&Otilde;", "Õ"},
            {"&ocirc;",	"ô"},
            {"&Ocirc;",	"Ô"},
            {"&uacute;", "ú"},
            {"&Uacute;", "Ú"},
            {"&ccedil;", "ç"},
            {"&Ccedil;", "Ç"},
            {"&ndash;", "–"},
            {"&mdash;", "—"},
            {"&nbsp;", Character.FromNumber(10)},
            {"&ordm;", "º"},
            {"&deg;", "°"},
            {"&le;", "<="},
            {"&iexcl;", "¡"},
            {"&iquest;", "¿"}
        },
        RemoveHtmlTags(HtmlText),
        (state as text, current as list) as text => Text.Replace(state, current{0}, current{1})
    )
in 
    ReplacedPunctuations
