# Regex Builder
Regular Expressions (regexes) are powerful but hard to write and **very** hard to debug. Regex Builder makes working with regexes easier.

 - Your regex runs continuously on the text as you edit it, providing instant feedback.
 - Run the selected subset of the regex with F5 to quickly narrow down why regexes aren't matching.
 - See matches, groups, and captures in a tree to confirm they're capturing what you intended.
 - Use the '>' menu for regex syntax reminders to quickly construct expressions.
 - Copy and paste C# or VB code with the regex properly string escaped.
 - Copy all matches to the clipboard to use elsewhere quickly.
 - 'Double-click debug' regex problems with an xml format that loads the regex, options, and text in one step.
 
I used Regex Builder to help write and debug regexes for Visual Studio editor automated tests. I copied several examples-to-match when first writing regexes to instantly test them as I was writing them. I debugged failed matches by running parts of the regex to narrow down what wasn't matching. My test harness wrote unexpected regex match failures to the xml file format, so I could double-click the file to immediately debug the problem.
