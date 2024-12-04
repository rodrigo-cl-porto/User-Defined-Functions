(regex as text, str as text) as text =>
let
    JavaScript = "<script>" &
        "var regex = " & regex & ";" &
        "var str = """ & str & """;" &
        "var res = str.match(regex);" &
        "document.write(res);" &
    "</script>"
in 
    Web.Page(JavaScript)
