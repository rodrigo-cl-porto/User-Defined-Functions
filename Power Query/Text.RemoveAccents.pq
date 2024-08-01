(textToConvert as nullable text) as text =>
if textToConvert = null or textToConvert = "" then
    ""
else let
    textBinary = Text.ToBinary(textToConvert, 1251),
    textASCII  = Text.FromBinary(textBinary, TextEncoding.Ascii)
in
    textASCII
