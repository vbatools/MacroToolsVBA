Attribute VB_Name = "AA_SwapEgual"
Option Explicit
Option Private Module

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SwapEgual - Поменять местами левую часть и правую часть относительно знака =
'* Created    : 05-10-2020 14:00
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub SwapEgual()
11:    Dim nStartLine  As Long
12:    Dim nStartColumn As Long
13:    Dim nEndline    As Long
14:    Dim nEndColumn  As Long
15:    Dim arrStr As Variant
16:
17:    Dim sLine       As String
18:    Dim sCode       As String
19:
20:    Dim nLine       As Integer
21:    Dim sToken1     As String
22:    Dim sToken2     As String
23:
24:    Dim sNew        As String
25:
26:    Dim nI          As Integer
27:
28:    Dim prjProject  As VBProject
29:    Dim cpCodePane  As CodePane
30:
31:    Dim nPos        As Integer
32:
33:
34:    On Error Resume Next
35:
36:    Set prjProject = Application.VBE.ActiveVBProject
37:
38:    If prjProject Is Nothing Then
39:        Debug.Print "Нет активного проекта VBA"
40:        Exit Sub
41:    End If
42:
43:    Set cpCodePane = Application.VBE.ActiveCodePane
44:
45:    If cpCodePane Is Nothing Then
46:        Debug.Print "Нет активного модуля кода VBA"
47:        Exit Sub
48:    End If
49:
50:    cpCodePane.GetSelection nStartLine, nStartColumn, nEndline, nEndColumn
51:    If nEndColumn > 1 Then nEndline = nEndline + 1
52:    sCode = cpCodePane.CodeModule.Lines(nStartLine, IIf(nEndline - nStartLine = 0, 1, nEndline - nStartLine))
53:
54:    If (sCode = vbNullString) Then
55:        Debug.Print "Код VBA не выделен"
56:        Exit Sub
57:    End If
58:
59:    sNew = vbNullString
60:    nLine = nStartLine
61:    arrStr = VBA.Split(sCode, vbNewLine)
62:    For nI = 0 To UBound(arrStr)
63:        sLine = arrStr(nI)
64:        nPos = InStr(sLine, " = ")
65:        If nPos > 0 Then
66:            sToken1 = RTrim(Left(sLine, nPos - 1))
67:            sToken2 = Right$(sLine, Len(sLine) - nPos - 2)
68:            sNew = Space(Len(sToken1) - Len(LTrim(sToken1))) & Trim$(sToken2) & " = " & Trim$(sToken1)
69:
70:            cpCodePane.CodeModule.ReplaceLine nLine, sNew
71:        End If
72:
73:        nLine = nLine + 1
74:    Next
75:    cpCodePane.SetSelection nStartLine, nStartColumn, nStartLine, nStartColumn
76: End Sub
