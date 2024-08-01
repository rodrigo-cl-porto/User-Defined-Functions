(textToRemove as nullable text) as text =>
if textToRemove = null or textToRemove = "" then
    ""
else let
    textBinary = Text.ToBinary(textToRemove, 1251),
    textASCII  = Text.FromBinary(textBinary, TextEncoding.Ascii),
    RemovedLetters = List.RemoveItems(Text.ToList(textASCII), {"a".."z", "A".."Z"}),
    CombinedText = Text.Combine(RemovedLetters)
in 
    CombinedText
