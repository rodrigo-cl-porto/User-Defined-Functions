(textToCount as nullable text, charToCount as nullable text) as nullable number =>
if textToCount = null or textToCount = "" or charToCount = null or charToCount = "" then
	0
else
	List.Count(Text.PositionOf(textToCount, charToCount, Occurrence.All))
