(textToConvert as text) as text =>
let
    textBinary = Text.ToBinary (textToConvert, 1251),
    textASCII  = Text.FromBinary(textBinary, TextEncoding.Ascii)
in
    textASCII
