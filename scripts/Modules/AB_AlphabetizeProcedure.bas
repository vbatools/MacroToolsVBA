Attribute VB_Name = "AB_AlphabetizeProcedure"
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : AlphabetizeProcedure - сортировка процедур и функций по алфавиту
'* Created    : 05-10-2020 14:01
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Const ProcUnderscore = "2"
Const ProcNoUnderscore = "1"

    Public Sub AlphabetizeProcedure()
15:
16:   Dim modCode          As CodeModule
17:   Dim cpCodePane       As CodePane
18:   Dim sProcName        As String
19:   Dim nProcKind        As Long
20:   Dim nSelectLine      As Long
21:   Dim nStartLine       As Long
22:   Dim nStartColumn     As Long
23:   Dim nEndline         As Long
24:   Dim nEndColumn       As Long
25:   Dim nCountOfLines    As Long
26:   Dim sProcText        As String
27:   Dim CollectedProcs   As Object
28:   Dim CollectedKeys()  As String
29:   Dim sKey             As String
30:   Dim nI               As Integer
31:   Dim nIndex           As Integer
32:
33:   If MsgBox("Sort procedures and functions alphabetically?", vbQuestion + vbYesNo + vbDefaultButton1, "Sorting:") = vbNo Then
34:      Exit Sub
35:   End If
36:
37:   On Error Resume Next
38:
39:   Set CollectedProcs = New Collection
40:   ReDim CollectedKeys(0) As String
41:
42:   If Application.VBE.ActiveVBProject Is Nothing Then
43:      Debug.Print "No active VBA project"
44:      Exit Sub
45:   End If
46:
47:   Set cpCodePane = Application.VBE.ActiveCodePane
48:
49:   If cpCodePane Is Nothing Then
50:      Debug.Print "No active VBA code module"
51:      Exit Sub
52:   End If
53:
54:   Set modCode = cpCodePane.CodeModule
55:
56:   Do While modCode.CountOfLines > modCode.CountOfDeclarationLines
57:      nStartLine = modCode.CountOfDeclarationLines + 1
58:      cpCodePane.SetSelection modCode.CountOfDeclarationLines + 1, 1, modCode.CountOfDeclarationLines + 1, 1
59:      cpCodePane.GetSelection nSelectLine, nStartColumn, nEndline, nEndColumn
60:      sProcName = modCode.ProcOfLine(nSelectLine, nProcKind)
61:      nCountOfLines = modCode.ProcCountLines(sProcName, nProcKind)
62:      sProcText = modCode.Lines(nStartLine, nCountOfLines)
63:      sKey = IIf(InStr(sProcName, "_"), ProcUnderscore, ProcNoUnderscore)
64:      sKey = sKey & sProcName & StringProcKind(nProcKind)
65:      CollectedProcs.Add sProcText, sKey
66:      ReDim Preserve CollectedKeys(0 To UBound(CollectedKeys) + 1) As String
67:      CollectedKeys(UBound(CollectedKeys)) = sKey
68:      modCode.DeleteLines nStartLine, nCountOfLines
69:   Loop
70:
71:   Do
72:      nIndex = 0
73:      sKey = " "
74:
75:      For nI = 1 To UBound(CollectedKeys)
76:         If LCase$(CollectedKeys(nI)) > LCase$(sKey) Then
77:            sKey = CollectedKeys(nI)
78:            nIndex = nI
79:         End If
80:      Next
81:
82:      If nIndex > 0 Then
83:         sProcText = CollectedProcs(sKey)
84:         CollectedKeys(nIndex) = " "
85:         modCode.AddFromString sProcText
86:      End If
87:
88:   Loop Until nIndex = 0
89:
90:   cpCodePane.Window.setFocus
91:   cpCodePane.SetSelection 1, 1, 1, 1
92:
93: End Sub

     Private Function StringProcKind(ByVal kind As Long) As String
96:   Select Case kind
      Case vbext_pk_Get
98:         StringProcKind = " Get"
99:      Case vbext_pk_Let
100:         StringProcKind = " Let"
101:      Case vbext_pk_Set
102:         StringProcKind = " Set"
103:      Case vbext_pk_Proc
104:         StringProcKind = " "
105:   End Select
106: End Function


