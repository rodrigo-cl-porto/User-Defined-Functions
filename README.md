# User Defined Functions

This repository brings together a set of user-defined functions (UDFs) developed for VBA and Power Query (M), with the goal of extending the native capabilities of Excel and Power BI.
Here you will find both functions created by me and useful functions developed by others — always properly organized, documented, and referenced.

## 🎯 Objective

- To centralize, organize, and facilitate access to a catalog of reusable functions
- To accelerate the development of solutions in Excel, Power Query, and Power BI.
- To maintain a versioned and easily searchable repository.
- To reuse functions tested and validated in different contexts.

## Table of Contents

### M Code (Power Query)

<table>
    <thead>
        <tr>
            <th>Function</th>
            <th>Description</th>
            <th>Documentation</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td><a href="/src/m/Binary.Unzip.pq"><code>Binary.Unzip</code></a></td>
            <td>Extracts file from a compressed ZIP file</td>
            <td align="center">
                <a href="/docs/en/m/Binary.Unzip.md">📄</a>
                <img src="/docs/assets/usa_flag.svg" alt="Binary.Unzip">
            </td>
        </tr>
        <tr>
            <td><a href="/src/m/DateTime.ToUnixTime.pq"><code>DateTime.ToUnixTime</code></a></td>
            <td>Converts a datetime format to Unix Time Stamp</td>
            <td align="center"><a href="/docs/en/m/DateTime.ToUnixTime.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Decision.EntropyWeights.pq"><code>Decision.EntropyWeights</code></a></td>
            <td>Calculates the weights of a decision multicriteria using the <strong>entropy weighting method</strong></td>
            <td align="center"><a href="/docs/en/m/Decision.EntropyWeights.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Decision.TOPSIS.pq"><code>Decision.TOPSIS</code></a></td>
            <td>Applies the <abbr title="Technique for Order Preference by Similarity to Ideal Solution">TOPSIS</abbr> multicriteria method to a table in order to rank alternatives</td>
            <td align="center"><a href="/docs/en/m/Decision.TOPSIS.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/List.Correlation.pq"><code>List.Correlation</code></a></td>
            <td>Calculates the correlation coefficient between two lists of numeric values</td>
            <td align="center"><a href="/docs/en/m/List.Correlation.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/List.Intercept.pq"><code>List.Intercept</code></a></td>
            <td>Calculates the intercept <em>B</em> of the linear regression line <em>Y = AX + B</em> between two numerical lists X and Y</td>
            <td align="center"><a href="/docs/en/m/List.Intercept.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/List.Outliers.pq"><code>List.Outliers</code></a></td>
            <td>Returns a numerical list of outliers existing in a list</td>
            <td align="center"><a href="/docs/en/m/List.Outliers.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/List.PopulationStdDev.pq"><code>List.PopulationStdDev</code></a></td>
            <td>Calculates the population standard deviation of a numerical list</td>
            <td align="center"><a href="/docs/en/m/List.PopulationStdDev.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/List.Primes.pq"><code>List.Primes</code></a></td>
            <td>Returns a list of prime numbers less than or equal to a given number <em>n</em></td>
            <td align="center"><a href="/docs/en/m/List.Primes.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/List.Rank.pq"><code>List.Rank</code></a></td>
            <td>Returns a list of ranks for a given list of values</td>
            <td align="center"><a href="/docs/en/m/List.Rank.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/List.Slope.pq"><code>List.Slope</code></a></td>
            <td>Calculates the slope <em>A</em> of the linear regression <em>Y = AX + B</em> between two numerical lists X and Y</td>
            <td align="center"><a href="/docs/en/m/List.Slope.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/List.Variance.pq"><code>List.Variance</code></a></td>
            <td>Calculates the population variance of a numerical list</td>
            <td align="center"><a href="/docs/en/m/List.Variance.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/List.WeightedAverage.pq"><code>List.WeightedAverage</code></a></td>
            <td>Calculates the weighted average of a list of values given a corresponding list of weights</td>
            <td align="center"><a href="/docs/en/m/List.WeightedAverage.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Number.FromRoman.pq"><code>Number.FromRoman</code></a></td>
            <td>Converts a Roman numeral to a number</td>
            <td align="center"><a href="/docs/en/m/Number.FromRoman.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Number.IsInteger.pq"><code>Number.IsInteger</code></a></td>
            <td>Checks if a number is integer or not</td>
            <td align="center"><a href="/docs/en/m/Number.IsInteger.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Number.IsPrime.pq"><code>Number.IsPrime</code></a></td>
            <td>Checks if a number is prime or not</td>
            <td align="center"><a href="/docs/en/m/Number.IsPrime.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Number.ToRoman.pq"><code>Number.ToRoman</code></a></td>
            <td>Converts a number to Roman numeral, if it's possible</td>
            <td align="center"><a href="/docs/en/m/Number.ToRoman.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Statistical.NormDist.pq"><code>Statistical.NormDist</code></a></td>
            <td>Calculates the Normal distribution for a given input <em>x</em></td>
            <td align="center"><a href="/docs/en/m/Statistical.NormDist.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Statistical.NormInv.pq"><code>Statistical.NormInv</code></a></td>
            <td>Returns the inverse of the <abbr title="Cumulative Distribution Function">CDF</abbr> of the normal distribution</td>
            <td align="center"><a href="/docs/en/m/Statistical.NormInv.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Table.AddColumnFromList.pq"><code>Table.AddColumnFromList</code></a></td>
            <td>Adds a new column to a table using values from a provided list</td>
            <td align="center"><a href="/docs/en/m/Table.AddColumnFromList.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Table.CorrelationMatrix.pq"><code>Table.CorrelationMatrix</code></a></td>
            <td>Calculates the correlation matrix for a given table</td>
            <td align="center"><a href="/docs/en/m/Table.CorrelationMatrix.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Table.NormalizeColumnNames.pq"><code>Table.NormalizeColumnNames</code></a></td>
            <td>Cleans and standardizes column names in a table</td>
            <td align="center"><a href="/docs/en/m/Table.NormalizeColumnNames.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Table.NormalizeTextColumns.pq"><code>Table.NormalizeTextColumns</code></a></td>
            <td>Cleans, standardizes and formats column names in a table</td>
            <td align="center"><a href="/docs/en/m/Table.NormalizeTextColumns.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Table.NormalizeTextColumns.pq"><code>Table.NormalizeTextColumns</code></a></td>
            <td>Cleans, standardizes and formats text columns in a table</td>
            <td align="center"><a href="/docs/en/m/Table.NormalizeTextColumns.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Table.RemoveBlankColumns.pq"><code>Table.RemoveBlankColumns</code></a></td>
            <td>Removes columns from a table that contain only blank values</td>
            <td align="center"><a href="/docs/en/m/Table.RemoveBlankColumns.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Table.TransposeCorrectly.pq"><code>Table.TransposeCorrectly</code></a></td>
            <td>Transposes a table without losing the original column names</td>
            <td align="center"><a href="/docs/en/m/Table.TransposeCorrectly.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.CountChar.pq"><code>Text.CountChar</code></a></td>
            <td>Counts the occurrences of a specific character in a given text string</td>
            <td align="center"><a href="/docs/en/m/Text.CountChar.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.ExtractNumbers.pq"><code>Text.ExtractNumbers</code></a></td>
            <td>Returns all numeric values from a given text</td>
            <td align="center"><a href="/docs/en/m/Text.ExtractNumbers.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.HtmlToPlainText.pq"><code>Text.HtmlToPlainText</code></a></td>
            <td>Converts HTML content to plain text</td>
            <td align="center"><a href="/docs/en/m/Text.HtmlToPlainText.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.RegexExtract.pq"><code>Text.RegexExtract</code></a></td>
            <td>Extracts a pattern from a text by using a regular expression</td>
            <td align="center"><a href="/docs/en/m/Text.RegexExtract.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.RegexReplace.pq"><code>Text.RegexReplace</code></a></td>
            <td>Replaces a pattern in a text that match a given regular expression</td>
            <td align="center"><a href="/docs/en/m/Text.RegexReplace.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.RegexSplit.pq"><code>Text.RegexSplit</code></a></td>
            <td>Splits a text into a list of strings based on a regular expression pattern</td>
            <td align="center"><a href="/docs/en/m/Text.RegexSplit.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.RegexTest.pq"><code>Text.RegexTest</code></a></td>
            <td>Tests whether a text matches a regular expression pattern</td>
            <td align="center"><a href="/docs/en/m/Text.RegexTest.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.RemoveAccents.pq"><code>Text.RemoveAccents</code></a></td>
            <td>Removes any accent from characters in a text</td>
            <td align="center"><a href="/docs/en/m/Text.RemoveAccents.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.RemoveDoubleSpaces.pq"><code>Text.RemoveDoubleSpaces</code></a></td>
            <td>Replaces any sequence of multiple spaces in a text to sigle spaces</td>
            <td align="center"><a href="/docs/en/m/Text.RemoveDoubleSpaces.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.RemoveLetters.pq"><code>Text.RemoveLetters</code></a></td>
            <td>Removes all alphabetic characters from a text</td>
            <td align="center"><a href="/docs/en/m/Text.RemoveLetters.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.RemoveNumerals.pq"><code>Text.RemoveNumerals</code></a></td>
            <td>Removes all digits from a text</td>
            <td align="center"><a href="/docs/en/m/Text.RemoveNumerals.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.RemovePunctuations.pq"><code>Text.RemovePunctuations</code></a></td>
            <td>Removes all punctuations from a text</td>
            <td align="center"><a href="/docs/en/m/Text.RemovePunctuations.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.RemoveStopwords.pq"><code>Text.RemoveStopwords</code></a></td>
            <td>Removes common Portuguese stopwords from a text</td>
            <td align="center"><a href="/docs/en/m/Text.RemoveStopwords.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/m/Text.RemoveWeirdChars.pq"><code>Text.RemoveWeirdChars</code></a></td>
            <td>Removes special and non-printable characters from a text</td>
            <td align="center"><a href="/docs/en/m/Text.RemoveWeirdChars.md">📄</a></td>
        </tr>
    </tbody>
</table>

### VBA (Visual Basic Application)

<table>
    <thead>
        <tr>
            <th>Function</th>
            <th>Description</th>
            <th>Documentation</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td><a href="/src/vba/AreArraysEquals.pq"><code>AreArraysEquals</code></a></td>
            <td>Checks if two arrays are equal</td>
            <td align="center"><a href="/docs/en/vba/AreArraysEquals.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/AutoFillFormulas.pq"><code>AutoFillFormulas</code></a></td>
            <td>Automatically fills formulas across a range using a reference cell's formula</td>
            <td align="center"><a href="/docs/en/vba/AutoFillFormulas.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/CleanString.pq"><code>CleanString</code></a></td>
            <td>Cleans a string by removing or replacing special and control characters with spaces</td>
            <td align="center"><a href="/docs/en/vba/CleanString.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/DisableRefreshAll.pq"><code>DisableRefreshAll</code></a></td>
            <td>Disables the "Refresh All" option for all <abbr title="Object Linking and Embedding, Database">OLEDB</abbr> connections in a workbook</td>
            <td align="center"><a href="/docs/en/vba/DisableRefreshAll.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/EnableRefreshAll.pq"><code>EnableRefreshAll</code></a></td>
            <td>Enables the "Refresh All" option for all <abbr title="Object Linking and Embedding, Database">OLEDB</abbr> connections in a workbook</td>
            <td align="center"><a href="/docs/en/vba/EnableRefreshAll.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/FileExists.pq"><code>FileExists</code></a></td>
            <td>Checks if a file exists at the specified file path.</td>
            <td align="center"><a href="/docs/en/vba/FileExists.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/FileNameIsValid.pq"><code>FileNameIsValid</code></a></td>
            <td>Checks if a given name can be used as a valid file name</td>
            <td align="center"><a href="/docs/en/vba/FileNameIsValid.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/GetAllFileNames.pq"><code>GetAllFileNames</code></a></td>
            <td>Retrieves an array of all file names from a folder and its subfolders</td>
            <td align="center"><a href="/docs/en/vba/GetAllFileNames.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/GetLetters.pq"><code>GetLetters</code></a></td>
            <td>Extracts ASCII letters from a string, in lowercase/td>
            <td align="center"><a href="/docs/en/vba/GetLetters.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/GetMonthNumberFromName.pq"><code>GetMonthNumberFromName</code></a></td>
            <td>Converts a month name to its corresponding numeric value</td>
            <td align="center"><a href="/docs/en/vba/GetMonthNumberFromName.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/GetStringBetween.pq"><code>GetStringBetween</code></a></td>
            <td>Extracts a string between two specified delimiters</td>
            <td align="center"><a href="/docs/en/vba/GetStringBetween.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/GetStringWithSubstringInArray.pq"><code>GetStringWithSubstringInArray</code></a></td>
            <td>Searches through an array of strings and returns the first string that contains a specified substring</td>
            <td align="center"><a href="/docs/en/vba/GetStringWithSubstringInArray.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/GetTableColumnNames.pq"><code>GetTableColumnNames</code></a></td>
            <td>Returns the column names of an Excel table</td>
            <td align="center"><a href="/docs/en/vba/GetTableColumnNames.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/IsAllTrue.pq"><code>IsAllTrue</code></a></td>
            <td>Checks if all elements in a boolean array are <code>True</code></td>
            <td align="center"><a href="/docs/en/vba/IsAllTrue.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/IsInArray.pq"><code>IsInArray</code></a></td>
            <td>Checks whether a value exists in an array</td>
            <td align="center"><a href="/docs/en/vba/IsInArray.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/ListObjectExists.pq"><code>ListObjectExists</code></a></td>
            <td>Checks whether a ListObject (Excel table) exists in a workbook</td>
            <td align="center"><a href="/docs/en/vba/ListObjectExists.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/PreviousMonthNumber.pq"><code>PreviousMonthNumber</code></a></td>
            <td>Returns the month's number that precedes the month of a given date</td>
            <td align="center"><a href="/docs/en/vba/PreviousMonthNumber.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/RangeHasAnyFormula.pq"><code>RangeHasAnyFormula</code></a></td>
            <td>Checks if a range contains any cell with formulas</td>
            <td align="center"><a href="/docs/en/vba/RangeHasAnyFormula.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/RangeHasConstantValues.pq"><code>RangeHasConstantValues</code></a></td>
            <td>Checks if a range contains any constant (non-formula) cell</td>
            <td align="center"><a href="/docs/en/vba/RangeHasConstantValues.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/RangeIsHidden.pq"><code>RangeIsHidden</code></a></td>
            <td>Checks if a range has no visible cell</td>
            <td align="center"><a href="/docs/en/vba/RangeIsHidden.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/RangeToHtml.pq"><code>RangeToHtml</code></a></td>
            <td>Converts a range into an HTML string</td>
            <td align="center"><a href="/docs/en/vba/RangeToHtml.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/SendEmail.pq"><code>SendEmail</code></a></td>
            <td>Sends an HTML email using <abbr title="Collaboration Data Objects">CDO</abbr> with <abbr title="New Technology LAN Manager">NTLM</abbr> authentication</td>
            <td align="center"><a href="/docs/en/vba/SendEmail.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/SetQueryFormula.pq"><code>SetQueryFormula</code></a></td>
            <td>Sets a Power Query formula for a query in the current workbook</td>
            <td align="center"><a href="/docs/en/vba/SetQueryFormula.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/StringContains.pq"><code>StringContains</code></a></td>
            <td>Checks if a string contains a substring</td>
            <td align="center"><a href="/docs/en/vba/StringContains.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/StringEndsWith.pq"><code>StringEndsWith</code></a></td>
            <td>Checks if a string ends with a specified substring</td>
            <td align="center"><a href="/docs/en/vba/StringEndsWith.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/StringStartsWith.pq"><code>StringStartsWith</code></a></td>
            <td>Checks if a string starts with a specified substring</td>
            <td align="center"><a href="/docs/en/vba/StringStartsWith.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/SubstringIsInArray.pq"><code>SubstringIsInArray</code></a></td>
            <td>Searches a array for any string element that contains a specified substring and returns <code>True</code> on the first match</td>
            <td align="center"><a href="/docs/en/vba/SubstringIsInArray.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/Summation.pq"><code>Summation</code></a></td>
            <td>Computes the numeric summation of a mathematical expression over an integer index range</td>
            <td align="center"><a href="/docs/en/vba/Summation.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/TableHasQuery.pq"><code>TableHasQuery</code></a></td>
            <td>Checks if a ListObject (Excel table) is associated to a query</td>
            <td align="center"><a href="/docs/en/vba/TableHasQuery.md">📄</a></td>
        </tr>
        <tr>
            <td><a href="/src/vba/WorksheetHasListObject.pq"><code>WorksheetHasListObject</code></a></td>
            <td>Checks if a worksheet contains a ListObject (Excel table)</td>
            <td align="center"><a href="/docs/en/vba/WorksheetHasListObject.md">📄</a></td>
        </tr>
    </tbody>
</table>

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
