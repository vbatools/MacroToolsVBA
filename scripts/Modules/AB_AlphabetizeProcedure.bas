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
14:
15:   Dim modCode          As CodeModule
16:   Dim cpCodePane       As CodePane
17:   Dim sProcName        As String
18:   Dim nProcKind        As Long
19:   Dim nSelectLine      As Long
20:   Dim nStartLine       As Long
21:   Dim nStartColumn     As Long
22:   Dim nEndline         As Long
23:   Dim nEndColumn       As Long
24:   Dim nCountOfLines    As Long
25:   Dim sProcText        As String
26:   Dim CollectedProcs   As Object
27:   Dim CollectedKeys()  As String
28:   Dim sKey             As String
29:   Dim nI               As Integer
30:   Dim nIndex           As Integer
31:
32:   If MsgBox("Sort procedures and functions alphabetically?", vbQuestion + vbYesNo + vbDefaultButton1, "Sorting:") = vbNo Then
33:      Exit Sub
34:   End If
35:
36:   On Error Resume Next
37:
38:   Set CollectedProcs = New Collection
39:   ReDim CollectedKeys(0) As String
40:
41:   If Application.VBE.ActiveVBProject Is Nothing Then
42:      Debug.Print "There is no active VBA project"
43:      Exit Sub
44:   End If
45:
46:   Set cpCodePane = Application.VBE.ActiveCodePane
47:
48:   If cpCodePane Is Nothing Then
49:      Debug.Print "There is no active VBA code module"
50:      Exit Sub
51:   End If
52:
53:   Set modCode = cpCodePane.CodeModule
54:
55:   Do While modCode.CountOfLines > modCode.CountOfDeclarationLines
56:      nStartLine = modCode.CountOfDeclarationLines + 1
57:      cpCodePane.SetSelection modCode.CountOfDeclarationLines + 1, 1, modCode.CountOfDeclarationLines + 1, 1
58:      cpCodePane.GetSelection nSelectLine, nStartColumn, nEndline, nEndColumn
59:      sProcName = modCode.ProcOfLine(nSelectLine, nProcKind)
60:      nCountOfLines = modCode.ProcCountLines(sProcName, nProcKind)
61:      sProcText = modCode.Lines(nStartLine, nCountOfLines)
62:      sKey = IIf(InStr(sProcName, "_"), ProcUnderscore, ProcNoUnderscore)
63:      sKey = sKey & sProcName & StringProcKind(nProcKind)
64:      CollectedProcs.Add sProcText, sKey
65:      ReDim Preserve CollectedKeys(0 To UBound(CollectedKeys) + 1) As String
66:      CollectedKeys(UBound(CollectedKeys)) = sKey
67:      modCode.DeleteLines nStartLine, nCountOfLines
68:   Loop
69:
70:   Do
71:      nIndex = 0
72:      sKey = " "
73:
74:      For nI = 1 To UBound(CollectedKeys)
75:         If LCase$(CollectedKeys(nI)) > LCase$(sKey) Then
76:            sKey = CollectedKeys(nI)
77:            nIndex = nI
78:         End If
79:      Next
80:
81:      If nIndex > 0 Then
82:         sProcText = CollectedProcs(sKey)
83:         CollectedKeys(nIndex) = " "
84:         modCode.AddFromString sProcText
85:      End If
86:
87:   Loop Until nIndex = 0
88:
89:   cpCodePane.Window.setFocus
90:   cpCodePane.SetSelection 1, 1, 1, 1
91:
92: End Sub

     Private Function StringProcKind(ByVal kind As Long) As String
95:   Select Case kind
      Case vbext_pk_Get
97:         StringProcKind = " Get"
98:      Case vbext_pk_Let
99:         StringProcKind = " Let"
100:      Case vbext_pk_Set
101:         StringProcKind = " Set"
102:      Case vbext_pk_Proc
103:         StringProcKind = " "
104:   End Select
105: End Function


