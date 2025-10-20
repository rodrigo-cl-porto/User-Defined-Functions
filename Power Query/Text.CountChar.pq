let 
	func = (textToCount as nullable text, charToCount as text) as number =>
		if textToCount = null or textToCount = "" or charToCount = null or charToCount = "" then
			0
		else
			List.Count(Text.PositionOf(textToCount, charToCount, Occurrence.All)),

	documentation = type function (
		textToCount as (type nullable text meta [
			Documentation.FieldCaption = "Text to Count",
			Documentation.FieldDescription = "The input text in which the occurrences of the specified character will be counted."
		]),
		charToCount as (type text meta [
			Documentation.FieldCaption = "Character to Count",
			Documentation.FieldDescription = "The character whose occurrences will be counted in the input text."
		])) as number meta [
			Documentation.Name = "Text.CountChar",
			Documentation.Description = "Counts the occurrences of a specific character in a text string.",
			Documentation.LongDescription = "This function scans the input text and counts how many times the specified character appears. It is useful for text analysis tasks where character frequency is relevant.",
			Documentation.Category = "Text",
			Documentation.Author = "Rodrigo Celso de Lima Porto",
			Documentation.Version = "1.0.0",
			Documentation.Examples = {[
				Description = "Basic usage",
				Code = "Text.CountChar(""Hello World!"", ""o"")",
				Result = 2
			], [
				Description = "Character not found",
				Code = "Text.CountChar(""Hello World!"", ""x"")",
				Result = 0
			], [
				Description = "Empty text input",
				Code = "Text.CountChar("""", ""a"")",
				Result = 0
			]}
		],

	documentedFunction = Value.ReplaceType(func, documentation)
in
	documentedFunction