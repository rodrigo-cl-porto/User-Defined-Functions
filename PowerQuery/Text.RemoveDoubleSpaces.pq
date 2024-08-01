(textToRemove as nullable text) as text =>
if textToRemove = null or textToRemove = "" then
    ""
else let
    SplittedText = Text.Split(textToRemove, " "),
    RemovedEmpties = List.RemoveItems(SplittedText, {"", null}),
    CombinedText = Text.Combine(RemovedEmpties, " ")
in
    CombinedText
