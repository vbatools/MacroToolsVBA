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
12:    Dim nStartLine  As Long
13:    Dim nStartColumn As Long
14:    Dim nEndline    As Long
15:    Dim nEndColumn  As Long
16:    Dim arrStr As Variant
17:
18:    Dim sLine       As String
19:    Dim sCode       As String
20:
21:    Dim nLine       As Integer
22:    Dim sToken1     As String
23:    Dim sToken2     As String
24:
25:    Dim sNew        As String
26:
27:    Dim nI          As Integer
28:
29:    Dim prjProject  As VBProject
30:    Dim cpCodePane  As CodePane
31:
32:    Dim nPos        As Integer
33:
34:
35:    On Error Resume Next
36:
37:    Set prjProject = Application.VBE.ActiveVBProject
38:
39:    If prjProject Is Nothing Then
40:        Debug.Print "No active VBA project"
41:        Exit Sub
42:    End If
43:
44:    Set cpCodePane = Application.VBE.ActiveCodePane
45:
46:    If cpCodePane Is Nothing Then
47:        Debug.Print "No active VBA code module"
48:        Exit Sub
49:    End If
50:
51:    cpCodePane.GetSelection nStartLine, nStartColumn, nEndline, nEndColumn
52:    If nEndColumn > 1 Then nEndline = nEndline + 1
53:    sCode = cpCodePane.CodeModule.Lines(nStartLine, IIf(nEndline - nStartLine = 0, 1, nEndline - nStartLine))
54:
55:    If (sCode = vbNullString) Then
56:        Debug.Print "VBA code is not allocated"
57:        Exit Sub
58:    End If
59:
60:    sNew = vbNullString
61:    nLine = nStartLine
62:    arrStr = VBA.Split(sCode, vbNewLine)
63:    For nI = 0 To UBound(arrStr)
64:        sLine = arrStr(nI)
65:        nPos = InStr(sLine, " = ")
66:        If nPos > 0 Then
67:            sToken1 = RTrim(Left(sLine, nPos - 1))
68:            sToken2 = Right$(sLine, Len(sLine) - nPos - 2)
69:            sNew = Space(Len(sToken1) - Len(LTrim(sToken1))) & Trim$(sToken2) & " = " & Trim$(sToken1)
70:
71:            cpCodePane.CodeModule.ReplaceLine nLine, sNew
72:        End If
73:
74:        nLine = nLine + 1
75:    Next
76:    cpCodePane.SetSelection nStartLine, nStartColumn, nStartLine, nStartColumn
End Sub
