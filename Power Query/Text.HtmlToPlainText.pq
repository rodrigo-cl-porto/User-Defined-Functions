let
    func = (htmlText as text) as text => Record.Field(Html.Table(htmlText, {{"text", ":root"}}){0}, "text"), // Html.Table() may not work on Excel Power Query

    documentation = type function (
        htmlText as (type text meta [
            Documentation.FieldCaption = "htmlText",
            Documentation.FieldDescription = "The HTML text to be converted to plain text"
        ])) as text meta [
            Documentation.Name = "Text.HtmlToPlainText",
            Documentation.Description = "Converts HTML content to plain text by stripping HTML tags.",
            Documentation.LongDescription = "Utilizes Power Query's Html.Table function to extract and return the plain text content from the provided HTML input, effectively removing all HTML tags.",
            Documentation.Category = "Text",
            Documentation.Author = "Rodrigo Celso de Lima Porto",
            Documentation.Version = "1.0.0",
            Documentation.Examples = {[
                Description = "Basic usage - converts simple HTML to plain text.",
                Code = "Text.HtmlToPlainText(""<p>Hello <b>World</b>!</p>"")",
                Result = "Hello World!"
            ]}
        ],

    documentedFunction = Value.ReplaceType(func, documentation)
in
    documentedFunction
