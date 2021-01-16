# busybox-vba for Windows

grep, sed, awk for VBA

# Setup

 - (1) Please copy whole files without sample.xlsm.
 - (2) And import busybox.vba in your VBA project.


# Functions

grep, sed, awk ...

## grep function

 - Public Function GrepSheet(ByVal RegExp As String, Options As String, ByRef InSheet As Worksheet, OutSheet As Worksheet) As Boolean
 - Public Function GrepText(ByVal RegExp As String, Options As String, InText As String) As String

## sed function

 - Public Function SedSheet(ByVal Commands As String, ByRef InSheet As Worksheet, OutSheet As Worksheet) As Boolean
 - Public Function SedText(ByVal Commands As String, ByVal InText As String) As String

## command function

 - Public Function ExecBatch(command As String, FailStr As String) As String
 - Public Sub ClearSheet(ByRef Sheet As Worksheet, ByVal TopRow As Integer)
 - Public Sub TSVToSheet(ByRef Sheet As Worksheet, ByVal tsv As String, TopRow As Integer)
 - Public Function ToTSV(ByRef Sheet As Worksheet) As String
 - Public Sub SaveToFile(Filename, Text, Charset)
 - Public Function ReadTextFile(Filename, Charset) As String


