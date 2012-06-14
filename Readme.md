# ExcelRegex
====
ExcelRegex is a dll that provides various utility functions for Excel 2010 and above, written in C# 4.0

It uses [ExcelDna](http://exceldna.codeplex.com/), which in turn uses Excel's low-level 'C' API 
to inject functions into Excel.

# Installation
----

Easy-peasy. In Excel:

File->Options->Addins->Excel Addins->Go

Browse to ExcelRegex.xll

That's it! It's installed. You can check that it works, by creating a new worksheet and entering the following:

cell A1: "Hello ExcelRegex World!"

cell A2: =RegexExtract(A1, "\s\w+\s")

# Functions
----

### =RegexExtract(Input, Pattern, GroupNum)

*Usage:*

Input is string (cell) to be matched

Pattern is the Regex to match against

GroupNum is which numbered capture group to return. Default is 0 (i.e. the whole match)

*Example:*

Assuming worksheet cell A1 has the value "123 abc" then

=RegexExtract(A1, "\d+") 

will return "123"


### =MD5Hash(File URL)

### =TimespanToMinutes("3 minutes")

### =Shorten(Url)

Returns a shortened version of specified URL. Use Bitly to shorten.

