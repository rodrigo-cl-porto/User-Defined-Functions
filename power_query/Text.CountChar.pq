(textToCount as text, char as text) as any =>
	List.Count(Text.PositionOf(textToCount, char, Occurrence.All))
