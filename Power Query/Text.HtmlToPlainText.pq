(textToConvert as text) as text => 
    Record.Field(Html.Table(textToConvert, {{"text", ":root"}}){0}, "text") // Html.Table() may not work on Excel Power Query
