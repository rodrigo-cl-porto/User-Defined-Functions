# User Defined Functions

This repository brings together a set of user-defined functions (UDFs) developed for VBA and Power Query (M), with the goal of extending the native capabilities of Excel and Power BI.
Here you will find both functions created by me and useful functions developed by others — always properly organized, documented, and referenced.

## 🎯 Objective

- To centralize, organize, and facilitate access to a catalog of reusable functions, allowing:
- To accelerate the development of solutions in Excel, Power Query, and Power BI.
- To maintain a versioned and easily searchable repository.
- To reuse functions tested and validated in different contexts.

## Table of Contents

### M Code (Power Query)

|Function                                                                        |Documentation                                            |
|:-------------------------------------------------------------------------------|:-------------------------------------------------------:|
|[`Binary.Unzip`](/src/power_query/Binary.Unzip.pq)                              |[📄](/docs/en/power_query/Binary.Unzip.md)               |
|[`DateTime.ToUnixTime`](/src/power_query/DateTime.ToUnixTime.pq)                |[📄](/docs/en/power_query/DateTime.ToUnixTime.md)        |
|[`Decision.EntropyWeights`](/src/power_query/Decision.EntropyWeights.pq)        |[📄](/docs/en/power_query/Decision.EntropyWeights.md)    |
|[`Decision.TOPSIS`](/src/power_query/Decision.TOPSIS.pq)                        |[📄](/docs/en/power_query/Decision.TOPSIS.md)            |
|[`List.Correlation`](/src/power_query/List.Correlation.pq)                      |[📄](/docs/en/power_query/List.Correlation.md)           |
|[`List.Intercept`](/src/power_query/List.Intercept.pq)                          |[📄](/docs/en/power_query/List.Intercept.md)             |
|[`List.Outliers`](/src/power_query/List.Outliers.pq)                            |[📄](/docs/en/power_query/List.Outliers.md)              |
|[`List.PopulationStdDev`](/src/power_query/List.PopulationStdDev.pq)            |[📄](/docs/en/power_query/List.PopulationStdDev.md)      |
|[`List.Primes`](/src/power_query/List.Primes.pq)                                |[📄](/docs/en/power_query/List.Primes.md)                |
|[`List.Rank`](/src/power_query/List.Rank.pq)                                    |[📄](/docs/en/power_query/List.Rank.md)                  |
|[`List.Slope`](/src/power_query/List.Slope.pq)                                  |[📄](/docs/en/power_query/List.Slope.md)                 |
|[`List.Variance`](/src/power_query/List.Variance.pq)                            |[📄](/docs/en/power_query/List.Variance.md)              |
|[`List.WeightedAverage`](/src/power_query/List.WeightedAverage.pq)              |[📄](/docs/en/power_query/List.WeightedAverage.md)       |
|[`Number.FromRoman`](/src/power_query/Number.FromRoman.pq)                      |[📄](/docs/en/power_query/Number.FromRoman.md)           |
|[`Number.IsInteger`](/src/power_query/Number.IsInteger.pq)                      |[📄](/docs/en/power_query/Number.IsInteger.md)           |
|[`Number.IsPrime`](/src/power_query/Number.IsPrime.pq)                          |[📄](/docs/en/power_query/Number.IsPrime.md)             |
|[`Number.ToRoman`](/src/power_query/Number.ToRoman.pq)                          |[📄](/docs/en/power_query/Number.ToRoman.md)             |
|[`Statistical.NormDist`](/src/power_query/Statistical.NormDist.pq)              |[📄](/docs/en/power_query/Statistical.NormDist.md)       |
|[`Statistical.NormInv`](/src/power_query/Statistical.NormInv.pq)                |[📄](/docs/en/power_query/Statistical.NormInv.md)        |
|[`Table.AddColumnFromList`](/src/power_query/Table.AddColumnFromList.pq)        |[📄](/docs/en/power_query/Table.AddColumnFromList.md)    |
|[`Table.CorrelationMatrix`](/src/power_query/Table.CorrelationMatrix.pq)        |[📄](/docs/en/power_query/Table.CorrelationMatrix.md)    |
|[`Table.FixColumnNames`](/src/power_query/Table.FixColumnNames.pq)              |[📄](/docs/en/power_query/Table.FixColumnNames.md)       |
|[`Table.PreprocessTextColumns`](/src/power_query/Table.PreprocessTextColumns.pq)|[📄](/docs/en/power_query/Table.PreprocessTextColumns.md)|
|[`Table.RemoveBlankColumns`](/src/power_query/Table.RemoveBlankColumns.pq)      |[📄](/docs/en/power_query/Table.RemoveBlankColumns.md)   |
|[`Table.TransposeCorrectly`](/src/power_query/Table.TransposeCorrectly.pq)      |[📄](/docs/en/power_query/Table.TransposeCorrectly.md)   |
|[`Text.CountChar`](/src/power_query/Text.CountChar.pq)                          |[📄](/docs/en/power_query/Text.CountChar.md)             |
|[`Text.ExtractNumbers`](/src/power_query/Text.ExtractNumbers.pq)                |[📄](/docs/en/power_query/Text.ExtractNumbers.md)        |
|[`Text.HtmlToPlainText`](/src/power_query/Text.HtmlToPlainText.pq)              |[📄](/docs/en/power_query/Text.HtmlToPlainText.md)       |
|[`Text.RegexExtract`](/src/power_query/Text.RegexExtract.pq)                    |[📄](/docs/en/power_query/Text.RegexExtract.md)          |
|[`Text.RegexReplace`](/src/power_query/Text.RegexReplace.pq)                    |[📄](/docs/en/power_query/Text.RegexReplace.md)          |
|[`Text.RegexSplit`](/src/power_query/Text.RegexSplit.pq)                        |[📄](/docs/en/power_query/Text.RegexSplit.md)            |
|[`Text.RegexTest`](/src/power_query/Text.RegexTest.pq)                          |[📄](/docs/en/power_query/Text.RegexTest.md)             |
|[`Text.RemoveAccents`](/src/power_query/Text.RemoveAccents.pq)                  |[📄](/docs/en/power_query/Text.RemoveAccents.md)         |
|[`Text.RemoveDoubleSpaces`](/src/power_query/Text.RemoveDoubleSpaces.pq)        |[📄](/docs/en/power_query/Text.RemoveDoubleSpaces.md)    |
|[`Text.RemoveLetters`](/src/power_query/Text.RemoveLetters.pq)                  |[📄](/docs/en/power_query/Text.RemoveLetters.md)         |
|[`Text.RemoveNumerals`](/src/power_query/Text.RemoveNumerals.pq)                |[📄](/docs/en/power_query/Text.RemoveNumerals.md)        |
|[`Text.RemovePunctuations`](/src/power_query/Text.RemovePunctuations.pq)        |[📄](/docs/en/power_query/Text.RemovePunctuations.md)    |
|[`Text.RemoveStopwords`](/src/power_query/Text.RemoveStopwords.pq)              |[📄](/docs/en/power_query/Text.RemoveStopwords.md)       |
|[`Text.RemoveWeirdChars`](/src/power_query/Text.RemoveWeirdChars.pq)            |[📄](/docs/en/power_query/Text.RemoveWeirdChars.md)      |

### VBA (Visual Basic Application)

|Function                                                                     |Documentation                                      |
|:----------------------------------------------------------------------------|:-------------------------------------------------:|
|[`AreArraysEquals`](/src/vba/AreArraysEqual.vba)                             |[📄](/docs/en/vba/AreArraysEquals.md)              |
|[`AutoFillFormulas`](/src/vba/AutoFillFormulas.vba)                          |[📄](/docs/en/vba/AutoFillFormulas.md)             |
|[`CleanString`](/src/vba/CleanString.vba)                                    |[📄](/docs/en/vba/CleanString.md)                  |
|[`DisableRefreshAll`](/src/vba/DisableRefreshAll.vba)                        |[📄](/docs/en/vba/DisableRefreshAll.md)            |
|[`EnableRefreshAll`](/src/vba/EnableRefreshAll.vba)                          |[📄](/docs/en/vba/EnableRefreshAll.md)             |
|[`CleanString`](/src/vba/FileExists.vba)                                     |[📄](/docs/en/vba/FileExists.md)                   |
|[`FileNameIsValid`](/src/vba/FileNameIsValid.vba)                            |[📄](/docs/en/vba/FileNameIsValid.md)              |
|[`GetAllFileNames`](/src/vba/GetAllFileNames.vba)                            |[📄](/docs/en/vba/GetAllFileNames.md)              |
|[`GetLettersOnly`](/src/vba/GetLettersOnly.vba)                              |[📄](/docs/en/vba/GetLettersOnly.md)               |
|[`GetMonthNumberFromName`](/src/vba/GetMonthNumberFromName.vba)              |[📄](/docs/en/vba/GetMonthNumberFromName.md)       |
|[`GetStringBetween`](/src/vba/GetStringBetween.vba)                          |[📄](/docs/en/vba/GetStringBetween.md)             |
|[`GetStringWithSubstringInArray`](/src/vba/GetStringWithSubstringInArray.vba)|[📄](/docs/en/vba/GetStringWithSubstringInArray.md)|
|[`GetTableColumnNames`](/src/vba/GetTableColumnNames.vba)                    |[📄](/docs/en/vba/GetTableColumnNames.md)          |
|[`IsAllTrue`](/src/vba/IsAllTrue.vba)                                        |[📄](/docs/en/vba/IsAllTrue.md)                    |
|[`IsInArray`](/src/vba/IsInArray.vba)                                        |[📄](/docs/en/vba/IsInArray.md)                    |
|[`ListObjectExists`](/src/vba/ListObjectExists.vba)                          |[📄](/docs/en/vba/ListObjectExists.md)             |
|[`PreviousMonthNumber`](/src/vba/PreviousMonthNumber.vba)                    |[📄](/docs/en/vba/PreviousMonthNumber.md)          |
|[`RangeHasAnyFormula`](/src/vba/RangeHasAnyFormula.vba)                      |[📄](/docs/en/vba/RangeHasAnyFormula.md)           |
|[`RangeHasConstantValues`](/src/vba/RangeHasConstantValues.vba)              |[📄](/docs/en/vba/RangeHasConstantValues.md)       |
|[`RangeIsHidden`](/src/vba/RangeIsHidden.vba)                                |[📄](/docs/en/vba/RangeIsHidden.md)                |
|[`RangeToHtml`](/src/vba/RangeToHtml.vba)                                    |[📄](/docs/en/vba/RangeToHtml.md)                  |
|[`SendEmail`](/src/vba/SendEmail.vba)                                        |[📄](/docs/en/vba/SendEmail.md)                    |
|[`SetQueryFormula`](/src/vba/SetQueryFormula.vba)                            |[📄](/docs/en/vba/SetQueryFormula.md)              |
|[`StringContains`](/src/vba/StringContains.vba)                              |[📄](/docs/en/vba/StringContains.md)               |
|[`StringEndsWith`](/src/vba/StringEndsWith.vba)                              |[📄](/docs/en/vba/StringEndsWith.md)               |
|[`StringStartsWith`](/src/vba/StringStartsWith.vba)                          |[📄](/docs/en/vba/StringStartsWith.md)             |
|[`SubstringIsInArray`](/src/vba/SubstringIsInArray.vba)                      |[📄](/docs/en/vba/SubstringIsInArray.md)           |
|[`Summation`](/src/vba/Summation.vba)                                        |[📄](/docs/en/vba/Summation.md)                    |
|[`TableHasQuery`](/src/vba/TableHasQuery.vba)                                |[📄](/docs/en/vba/TableHasQuery.md)                |
|[`WorksheetHasListObject`](/src/vba/WorksheetHasListObject.vba)              |[📄](/docs/en/vba/WorksheetHasListObject.md)       |

## 🤝 Contributions

Contributions are welcome!
If you have an interesting feature, improvement, or fix for any function or documentation, feel free to open a pull request or an issue.

## 🗂️ Other UDF Repositories

Here is a list of very useful repos of user-defined functions:

- [M](https://github.com/ImkeF/M) by Imke Feldmann
- [M Custom Functions](https://github.com/tirnovar/m-custom-functions) by Štěpán Rešl
- [m-custom-functions](https://github.com/tirnovar/m-custom-functions) by Tirnovar
- [M-tools](https://github.com/acaprojects/m-tools/tree/master) by Kim Burgess
- [PowerBi-code](https://github.com/ibarrau/PowerBi-code/tree/master) by ibarrau
- [PowerQueryFunctions](https://github.com/OscarValerock/PowerQueryFunctions) by OscarValerock
- [PowerQueryLib](https://github.com/ninmonkey/Ninmonkey.PowerQueryLib) by NinMonkey
