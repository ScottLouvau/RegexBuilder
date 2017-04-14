# RegexBuilder
RegexBuilder helps you write, debug, and test regular expressions using the .NET syntax.
 - Use the menus to help construct expressions.
 - Run all or part of the expression on your source text quickly.
 - See matches, groups, and captures in a tree to see how it matches.
 - Supports an Xml file format so you can have other tools output files for quick diagnosis in RegexBuilder.
 
 I used RegexBuilder to help me build and debug expressions used in Visual Studio editor automated tests. Test failures due to a regex failing to match would automatically write an Xml file which RegexBuilder could import for diagnosis.
