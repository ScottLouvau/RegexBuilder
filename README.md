# Regex Builder
Regex Builder helps you write, debug, and test regular expressions using the .NET syntax.
 - Your regex is run continuously on the text as you type, providing instant feedback.
 - Press F5 to run the selected subset of the regex to debug why regexes aren't matching.
 - See matches, groups, and captures in a tree to confirm they're matching what you intended.
 - Use the '>' menu for regex syntax reminders to quickly construct expressions.
 - Copy and paste C# or VB code with the regex properly string escaped instantly.
 - Copy all matches to the clipboard to gather matches from text quickly.
 - 'Double-click debug' regex non-matches from production with an xml file format that has the regex, options, and text.
 
I used Regex Builder to help me write and debug expressions used in Visual Studio editor automated tests. I copied several examples-to-match when writing regexes to instantly test them as I was writing them. I debugged failed matches by running parts of the expression to narrow down on the part that wasn't matching and then correct it. I wrote my test harness so that unexpected failures to match a regex wrote out the xml file format, so I could double-click it to immediately open it for debugging.
