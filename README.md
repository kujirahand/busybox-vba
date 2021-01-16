# busybox-vba for Windows

grep, sed, awk for VBA

# Setup

 - (1) Please copy whole files without sample.xlsm.
 - (2) And import busybox.vba in your VBA project.


# Functions

grep, sed, awk ...

# Sample

Please see Module1 in Sample.xlsm.

```
' Test grep Sheet
Sub TestGrepSheet()
    ClearSheet Sheet2, 1
    GrepSheet "^y", "-i", Sheet1, Sheet2
End Sub

' Test grep Text
Sub TestGrepText()
    Dim tsv
    tsv = ToTSV(Sheet1)
    Debug.Print GrepText("^y", "-i", tsv)
End Sub

' Test sed Sheet
Sub TestSedSheet()
    ClearSheet Sheet2, 1
    SedSheet "s/^Y2021/(new)/p", "-n", Sheet1, Sheet2
End Sub

' Test sed Text
Sub TestSedText()
    Dim tsv
    tsv = ToTSV(Sheet1)
    Debug.Print SedText("s/^Y2021/(new)/p", "-n", tsv)
End Sub

' Test awk Text
Sub TestAwkText()
    Dim tsv
    tsv = ToTSV(Sheet1)
    Debug.Print AwkText("{$3=int($3*1.1);print}", "-F""\t"" -vOFS=""\t""", tsv)
End Sub

' Test awk Sheet
Sub TestAwkSheet()
    ClearSheet Sheet2, 1
    AwkSheet "{$3=int($3*1.1);print}", "", Sheet1, Sheet2
End Sub
```

