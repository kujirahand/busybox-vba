Attribute VB_Name = "busybox"
Option Explicit
' ============================================================
' busybox for vba (Windows)
' [URL] https://github.com/kujirahand/busybox-vba
' ============================================================

' Global
Dim BusyboxPath As String
Dim BusyboxCharset As String
Const BusyboxCharsetDefault = "UTF-8"

' SetBusyboxPath
Public Sub SetBusyboxPath(Path As String)
    BusyboxPath = Path
End Sub

' SetBusyboxPath
Public Sub SetBusyboxCharset(ByVal Charset As String)
    BusyboxCharset = Charset
End Sub

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


Public Function ExecSheet(ByVal Command As String, ByVal Pattern As String, ByVal Options As String, ByRef InSheet As Worksheet, ByRef OutSheet As Worksheet) As Boolean
    ' �ΏۃV�[�g��TSV�ɕϊ�
    Dim tsv As String, tmpfile As String
    tsv = ToTSV(InSheet)
    tmpfile = GetTempPath(".tsv")
    SaveText tmpfile, tsv
        
    ' �R�}���h���\�z
    Dim cmd As String, s As String
    cmd = Command & " " & Options & " " & sq(Pattern) & " " & qq(tmpfile)
    
    ' grep�����s���Č��ʂ𓾂�
    s = ExecBatch(cmd, "__ERROR__")
    If s = "__ERROR__" Then
        ExecSheet = False
        Exit Function
    End If

    ' ���ʂ��V�[�g�ɒ���t����
    TSVToSheet OutSheet, s, 1
    ExecSheet = True
End Function


' GrepSheet
Public Function GrepSheet(ByVal Pattern As String, ByVal Options As String, ByRef InSheet As Worksheet, ByRef OutSheet As Worksheet) As Boolean
    GrepSheet = ExecSheet("grep", Pattern, Options, InSheet, OutSheet)
End Function

' SedSheet
Public Function SedSheet(ByVal Script As String, ByVal Options As String, ByRef InSheet As Worksheet, OutSheet As Worksheet) As Boolean
    SedSheet = ExecSheet("sed", Script, Options, InSheet, OutSheet)
End Function

' AwkSheet
Public Function AwkSheet(ByVal Script As String, ByVal Options As String, ByRef InSheet As Worksheet, OutSheet As Worksheet) As Boolean
    ' �I�v�V�������w�肳��Ă��Ȃ��Ƃ��A�^�u�L�����w�肷��
    If Options = "" Then
        Options = "-F""\t"" -vOFS=""\t"""
    End If
    
    ' �R�}���h�����s
    AwkSheet = ExecSheet("awk", Script, Options, InSheet, OutSheet)
End Function


' GrepText
Public Function ExecText(ByVal Command As String, ByVal Pattern As String, ByVal Options As String, ByVal InText As String) As String
    Dim tmpfile
    ' Save text to file
    tmpfile = GetTempPath(".tsv")
    SaveText tmpfile, InText
    
    ' grep�����s���Č��ʂ𓾂�
    Dim cmd As String, s As String
    cmd = Command & " " & Options & " " & sq(Pattern) & " " & qq(tmpfile)
    s = ExecBatch(cmd, "__ERROR__")
    If s = "__ERROR__" Then
        ExecText = False
        Exit Function
    End If
    ExecText = s
End Function

' GrepText
Public Function GrepText(ByVal RegExp As String, ByVal Options As String, ByVal InText As String) As String
    GrepText = ExecText("grep", RegExp, Options, InText)
End Function

' SedText
Public Function SedText(ByVal Script As String, ByVal Options As String, ByVal InText As String) As String
    SedText = ExecText("sed", Script, Options, InText)
End Function

' AwkText
Public Function AwkText(ByVal Script As String, ByVal Options As String, ByVal InText As String) As String
    AwkText = ExecText("awk", Script, Options, InText)
End Function


' Execute Batch Command
Public Function ExecBatch(ByVal Commands As String, ByVal FailStr As String) As String
    Call BusyboxInit
    ' GetTempFile
    Dim FSO As Object, BatFile As String, OutFile As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    BatFile = GetTempPath(".bash")
    OutFile = GetTempPath(".txt")

    ' Save batfile
    Dim Src As String
    Src = Commands & " >" & qq(OutFile) & vbCrLf
    ' Src = qq(BusyboxPath) & Commands
    ' Src = Src & ">" & qq(OutFile) & vbCrLf
    ' Src = "chcp 65001" & vbCrLf & Src ' Set Charset UTF-8
    ' Src = Src & vbCrLf & "pause" & vbCrLf
    SaveText BatFile, Src
    Debug.Print Src
    
    ' execute batch
    Dim r As Boolean, sh
    sh = qq(BusyboxPath) & " bash " & qq(BatFile)
    r = ShellWait(sh)
    If Not r Then
        Debug.Print "[Error] Batch command faild. Path=" & BatFile
        ExecBatch = FailStr
        Exit Function
    End If
    ' GetResult
    Dim res As String
    res = LoadText(OutFile)
    ExecBatch = res
End Function

' �ȉ��͉������֐�

' ShellWait is Execute command and wait
Public Function ShellWait(ByVal Command As String) As Boolean
    On Error GoTo SHELL_ERROR
    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    Dim res As Integer
    res = wsh.Run(Command, 7, True) ' minimize not focus
    ShellWait = (res = 0)
    Exit Function
SHELL_ERROR:
    ShellWait = False
End Function

Private Function GetTempPath(Ext As String) As String
    Dim FSO As Object, tmp As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    tmp = FSO.GetSpecialFolder(2) & "\" & FSO.GetBaseName(FSO.GetTempName) & Ext
    GetTempPath = tmp
End Function


' quote path
Private Function qq(Path) As String
    qq = """" & Path & """"
End Function

' quote string
Private Function sq(ss) As String
    Dim s
    s = Replace(ss, "'", "'\''")
    sq = "'" & s & "'"
End Function

' Clear sheet
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

' TSV to Sheet
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


' Convert Sheet to TSV
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

Public Sub SaveText(ByVal Filename As String, ByVal Text As String)
    If BusyboxCharset = "" Then BusyboxCharset = BusyboxCharsetDefault
    SaveToFile Filename, Text, BusyboxCharset
End Sub

Public Function LoadText(Filename) As String
    If BusyboxCharset = "" Then BusyboxCharset = BusyboxCharsetDefault
    LoadText = LoadFromFile(Filename, BusyboxCharset)
End Function


' �e�L�X�g���w�蕶���R�[�h�Ńt�@�C���ɕۑ�
Public Sub SaveToFile(ByVal Filename, ByVal Text, ByVal Charset)
    ' UTF-8 �̏ꍇ BOM�͕s�v
    If LCase(Charset) = "utf-8" Or LCase(Charset) = "utf-8n" Or LCase(Charset) = "utf8" Then
        Call SaveToFileUTF8N(Filename, Text)
        Exit Sub
    End If
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = Charset
    stream.Open
    stream.WriteText Text
    stream.SaveToFile Filename, 2
    stream.Close
End Sub

' BOM�Ȃ���UTF-8�Ńt�@�C���Ƀe�L�X�g����������
Public Sub SaveToFileUTF8N(Filename, Text)
    Dim stream, buf
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2 ' �e�L�X�g���[�h���w�� --- (*1)
        .Charset = "UTF-8"
        .Open
        .WriteText Text ' �e�L�X�g����������
        .Position = 0 ' �J�[�\�����t�@�C���擪�� --- (*2)
        .Type = 1 ' �o�C�i�����[�h�ɕύX
        .Position = 3 ' BOM(3�o�C�g)���΂�
        buf = .Read() ' ���e��ǂݍ���
        .Position = 0 ' �J�[�\����擪�� --- (*3)
        .Write buf ' BOM�Ȃ��̃e�L�X�g����������
        .SetEOS
        .SaveToFile Filename, 2
        .Close
    End With
End Sub


' �C�ӂ̕����G���R�[�f�B���O���w�肵�ăe�L�X�g�t�@�C����ǂ�
Public Function LoadFromFile(Filename, Charset) As String
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' text
    stream.Charset = Charset
    stream.Open
    stream.LoadFromFile Filename
    LoadFromFile = stream.ReadText
    stream.Close
End Function



