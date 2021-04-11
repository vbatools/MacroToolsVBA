Attribute VB_Name = "U_UpperAndLowerCase"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : U_UpperAndLowerCase - изменение регистра выделенных строк
'* Created    : 18-02-2020 09:05
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : toUpperCase - перевод выделенного кода в ВЕРХНИЙ регистр
'* Created    : 18-02-2020 09:05
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub toUpperCase()
18:    On Error GoTo ErrorHandler
19:    Dim i      As Long
20:    Dim newText As String
21:    Dim lineText As String
22:    Dim sL     As Long
23:    Dim eL     As Long
24:    Dim sC     As Long
25:    Dim eC     As Long
26:
27:    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)
28:
29:    If sL = eL Then
30:        lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1)
31:        newText = VBA.Mid(lineText, 1, sC - 1) & VBA.UCase$(VBA.Mid(lineText, sC, eC - sC)) & VBA.Mid(lineText, eC)
32:        If newText <> vbNullString Then Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(sL, newText)
33:    Else
34:        For i = sL To eL
35:            newText = ""
36:            lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1)
37:            If i = sL Then
38:                newText = VBA.Mid(lineText, 1, sC - 1) & VBA.UCase$(VBA.Mid(lineText, sC))
39:            ElseIf i = eL Then
40:                newText = VBA.UCase$(VBA.Mid(lineText, 1, eC - 1)) & VBA.Mid(lineText, eC)
41:            Else
42:                newText = VBA.UCase$(lineText)
43:            End If
44:            If newText <> vbNullString Then Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(i, newText)
45:        Next i
46:    End If
47:    Call Application.VBE.ActiveCodePane.SetSelection(sL, sC, eL, eC)
ErrorHandler:
49:    Select Case Err.Number
        Case 0
51:        Case Else
52:            Debug.Print "Error in U_UpperAndLowerCase.toUpperCase" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
53:            Call WriteErrorLog("U_UpperAndLowerCase.toUpperCase")
54:    End Select
55: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : toLowerCase - перевод выделенного кода в нижний регистр
'* Created    : 18-02-2020 09:06
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub toLowerCase()
65:    On Error GoTo ErrorHandler
66:    Dim i      As Long
67:    Dim newText As String
68:    Dim lineText As String
69:    Dim sL     As Long
70:    Dim eL     As Long
71:    Dim sC     As Long
72:    Dim eC     As Long
73:
74:    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)
75:
76:    If sL = eL Then
77:        lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1)
78:        newText = VBA.Mid(lineText, 1, sC - 1) & VBA.LCase(VBA.Mid(lineText, sC, eC - sC)) & VBA.Mid(lineText, eC)
79:        If newText <> vbNullString Then Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(sL, newText)
80:    Else
81:        For i = sL To eL
82:            newText = ""
83:            lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1)
84:            If i = sL Then
85:                newText = VBA.Mid(lineText, 1, sC - 1) & VBA.LCase(VBA.Mid(lineText, sC))
86:            ElseIf i = eL Then
87:                newText = VBA.LCase(VBA.Mid(lineText, 1, eC - 1)) & VBA.Mid(lineText, eC)
88:            Else
89:                newText = VBA.LCase(lineText)
90:            End If
91:            If newText <> vbNullString Then Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(i, newText)
92:        Next i
93:    End If
94:
95:    Call Application.VBE.ActiveCodePane.SetSelection(sL, sC, eL, eC)
96:
ErrorHandler:
98:    Select Case Err.Number
        Case 0
100:        Case Else
101:            Debug.Print "Error in U_Upper AndLowerCase.toLowerCase" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
102:            Call WriteErrorLog("U_UpperAndLowerCase.toLowerCase")
103:    End Select
End Sub
