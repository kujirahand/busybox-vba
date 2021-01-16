Attribute VB_Name = "busybox"
Option Explicit
' ============================================================
' busybox for vba (Windows)
' [URL] https://github.com/kujirahand/busybox-vba
' ============================================================

' Global
Dim BusyboxPath As String

' SetBusyboxPath
Public Sub SetBusyboxPath(Path As String)
    BusyboxPath = Path
End Sub

' GrepSheet
Public Function GrepSheet(ByVal RegExp As String, ByVal Options As String, ByRef InSheet As Worksheet, ByRef OutSheet As Worksheet) As Boolean
    ' �ΏۃV�[�g��TSV�ɕϊ�
    Dim tsv As String, TmpFile As String
    tsv = ToTSV(InSheet)
    TmpFile = GetTempPath(".tsv")
    SaveToFile TmpFile, tsv, "utf-8"
    
    ' grep�����s���Č��ʂ𓾂�
    Dim cmd As String, s As String
    cmd = "grep " & Options & " " & qq(RegExp) & " " & qq(TmpFile)
    s = ExecBatch(cmd, "__ERROR__")
    If s = "__ERROR__" Then
        GrepSheet = False
        Exit Function
    End If

    ' ���ʂ��V�[�g�ɒ���t����
    TSVToSheet OutSheet, s, 1
    GrepSheet = True
End Function

' GrepText
Public Function GrepText(ByVal RegExp As String, ByVal Options As String, ByVal InText As String) As String
    Dim TmpFile
    ' Save text to file
    TmpFile = GetTempPath(".tsv")
    SaveToFile TmpFile, InText, "utf-8"
    
    ' grep�����s���Č��ʂ𓾂�
    Dim cmd As String, s As String
    cmd = "grep " & Options & " " & qq(RegExp) & " " & qq(TmpFile)
    s = ExecBatch(cmd, "__ERROR__")
    If s = "__ERROR__" Then
        GrepText = False
        Exit Function
    End If
    GrepText = s
End Function

' SedSheet
Public Function SedSheet(ByVal Commands As String, ByRef InSheet As Worksheet, OutSheet As Worksheet) As Boolean
    ' Save text to file
    Dim TmpFile, tsv
    tsv = ToTSV(InSheet)
    TmpFile = GetTempPath(".tsv")
    SaveToFile TmpFile, tsv, "utf-8"
    
    ' sed�����s���Č��ʂ𓾂�
    Dim cmd As String, s As String
    cmd = "sed " & Commands & " " & qq(TmpFile)
    s = ExecBatch(cmd, "__ERROR__")
    If s = "__ERROR__" Then
        SedSheet = False
        Exit Function
    End If
    
    ' ���ʂ��V�[�g��
    TSVToSheet OutSheet, s, 1
    SedSheet = True
End Function


' SedText
Public Function SedText(ByVal Commands As String, ByVal InText As String) As String
    ' Save text to file
    Dim TmpFile
    TmpFile = GetTempPath(".tsv")
    SaveToFile TmpFile, InText, "utf-8"
    
    ' grep�����s���Č��ʂ𓾂�
    Dim cmd As String, s As String
    cmd = "sed " & Commands & " " & qq(TmpFile)
    s = ExecBatch(cmd, "__ERROR__")
    If s = "__ERROR__" Then
        SedText = False
        Exit Function
    End If
    SedText = s
End Function


' AwkText
Public Function AwkText(ByVal Commands As String, ByVal InText As String) As String
    ' Save text to file
    Dim TmpFile
    TmpFile = GetTempPath(".tsv")
    SaveToFile TmpFile, InText, "utf-8"
    
    ' awk�����s���Č��ʂ𓾂�
    Dim cmd As String, s As String
    cmd = "awk " & Commands & " " & qq(TmpFile)
    s = ExecBatch(cmd, "__ERROR__")
    If s = "__ERROR__" Then
        AwkText = False
        Exit Function
    End If
    AwkText = s
End Function

' AwkSheet
Public Function AwkSheet(ByVal Commands As String, ByRef InSheet As Worksheet, OutSheet As Worksheet) As Boolean
    ' Save text to file
    Dim TmpFile, tsv
    tsv = ToTSV(InSheet)
    TmpFile = GetTempPath(".tsv")
    SaveToFile TmpFile, tsv, "utf-8"
    
    ' sed�����s���Č��ʂ𓾂�
    Dim cmd As String, s As String
    cmd = "awk " & Commands & " " & qq(TmpFile)
    s = ExecBatch(cmd, "__ERROR__")
    If s = "__ERROR__" Then
        AwkSheet = False
        Exit Function
    End If
    
    ' ���ʂ��V�[�g��
    TSVToSheet OutSheet, s, 1
    AwkSheet = True
End Function


' Initalize busybox
Private Sub BusyboxInit()
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' Check Busybox path
    If BusyboxPath = "" Then
        BusyboxPath = ThisWorkbook.Path & "\busybox.exe"
        If Not FSO.FileExists(BusyboxPath) Then
            BusyboxPath = ThisWorkbook.Path & "\bin\busybox.exe"
            If Not FSO.FileExists(BusyboxPath) Then
            BusyboxPath = ThisWorkbook.Path & "\lib\busybox.exe"
            End If
        End If
    End If
    ' Show Error Message
    If Not FSO.FileExists(BusyboxPath) Then
        MsgBox "busybox.exe not found", vbCritical
    End If
End Sub

' ShellWait is Execute command and wait
Public Function ShellWait(command As String) As Boolean
    On Error GoTo SHELL_ERROR
    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    Dim res As Integer
    res = wsh.Run(command, 7, True) ' minimize not focus
    ShellWait = (res = 0)
    Exit Function
SHELL_ERROR:
    ShellWait = False
End Function

Private Function GetTempPath(Ext As String) As String
    Dim FSO As Object, Tmp As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Tmp = FSO.GetSpecialFolder(2) & "\" & FSO.GetBaseName(FSO.GetTempName) & Ext
    GetTempPath = Tmp
End Function

' Execute Batch Command
Public Function ExecBatch(command As String, FailStr As String) As String
    Call BusyboxInit
    ' GetTempFile
    Dim FSO As Object, BatFile As String, OutFile As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    BatFile = GetTempPath(".bat")
    OutFile = GetTempPath(".txt")

    ' Save batfile
    Dim Src As String
    Src = qq(BusyboxPath) & " " & command & ">" & qq(OutFile) & vbCrLf
    SaveToFile BatFile, Src, "sjis"
    Debug.Print Src
    
    ' execute batch
    Dim r As Boolean
    r = ShellWait(BatFile)
    If Not r Then
        Debug.Print "[Error] Batch command faild. Path=" & BatFile
        ExecBatch = FailStr
        Exit Function
    End If
    ' GetResult
    Dim res As String
    res = ReadTextFile(OutFile, "utf-8")
    ExecBatch = res
End Function

' �p�X�̑O��Ƀ_�u���N�H�[�g������
Private Function qq(Path) As String
    qq = """" & Path & """"
End Function

' �����̃V�[�g�̃Z�����󔒂ɂ���
Public Sub ClearSheet(ByRef Sheet As Worksheet, ByVal TopRow As Integer)
    Dim EndCol, EndRow, Row, Col
    With Sheet.UsedRange
        EndRow = .Rows(.Rows.Count).Row
        EndCol = .Columns(.Columns.Count).Column
    End With
    For Row = TopRow To EndRow
        For Col = 1 To EndCol
            Sheet.Cells(Row, Col) = ""
        Next
    Next
End Sub

' TSV�̓��e���V�[�g�ɏ�������
Public Sub TSVToSheet(ByRef Sheet As Worksheet, ByVal tsv As String, TopRow As Integer)
    Dim Rows As Variant, Cols As Variant
    Dim i, j
    Rows = Split(tsv, Chr(10))
    For i = 0 To UBound(Rows)
        Cols = Split(Rows(i), Chr(9))
        For j = 0 To UBound(Cols)
            Dim v
            v = Cols(j)
            v = Replace(v, "��", vbCrLf)
            Sheet.Cells(i + TopRow, j + 1) = v
        Next
    Next
End Sub


' �V�[�g��TSV�ɕϊ�
Public Function ToTSV(ByRef Sheet As Worksheet) As String
    Dim s As String
    s = ""
    ' �V�[�g�͈̔͂��擾
    Dim BottomRow As Integer, RightCol As Integer
    BottomRow = Sheet.Range("A1").End(xlDown).Row
    RightCol = Sheet.Range("A1").End(xlToRight).Column
    ' �V�[�g�͈͂����ォ�珇�Ɏ擾
    Dim y, x, v
    For y = 1 To BottomRow
        For x = 1 To RightCol
            v = Sheet.Cells(y, x)
            ' �Z�����̉��s�����͒u�����Ă���
            v = Replace(v, vbCrLf, "��")
            s = s & v & Chr(9)
        Next
        s = s & vbCrLf
    Next
    ToTSV = s
End Function

' �e�L�X�g���t�@�C���ɕۑ�
Public Sub SaveToFile(Filename, Text, Charset)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = Charset
    stream.Open
    stream.WriteText Text
    stream.SaveToFile Filename, 2
    stream.Close
End Sub

' �C�ӂ̕����G���R�[�f�B���O���w�肵�ăe�L�X�g�t�@�C����ǂ�
Public Function ReadTextFile(Filename, Charset) As String
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' text
    stream.Charset = Charset
    stream.Open
    stream.LoadFromFile Filename
    ReadTextFile = stream.ReadText
    stream.Close
End Function



