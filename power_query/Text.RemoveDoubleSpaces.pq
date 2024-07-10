(textToRemove as text) as text => 
let
    SplittedText = Text.Split(textToRemove, " "),
    RemovedEmpties = List.RemoveItems(SplittedText, {"", null}),
    CombinedText = Text.Combine(RemovedEmpties, " ")
in
    CombinedText
