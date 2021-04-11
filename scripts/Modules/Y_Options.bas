Attribute VB_Name = "Y_Options"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : Y_Options - модуль создание Options
'* Created    : 17-09-2020 14:35
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module

    Public Sub subOptions()
13:    Dim sOptions    As String
14:    Dim moCM        As CodeModule
15:    Dim vbComp      As VBIDE.VBComponent
16:    Dim objForm     As AddOptions
17:    Dim sActiveVBProject As String
18:
19:    On Error Resume Next
20:    sActiveVBProject = Application.VBE.ActiveVBProject.Filename
21:    On Error GoTo 0
22:
23:    On Error GoTo ErrorHandler
24:    Set objForm = New AddOptions
25:
26:    With objForm
27:
28:        If sActiveVBProject <> vbNullString Then .lbNameProject.Caption = sGetFileName(sActiveVBProject)
29:        .Show
30:        If .chOptionExplicit.Value Then
31:            sOptions = "Option Explicit" & vbNewLine
32:        End If
33:        If .chOptionPrivate.Value Then
34:            sOptions = sOptions & "Option Private Module" & vbNewLine
35:        End If
36:        If .chOptionCompare.Value Then
37:            sOptions = sOptions & "Option Compare Text" & vbNewLine
38:        End If
39:        If .chOptionBase.Value Then
40:            sOptions = sOptions & "Option Base 1" & vbNewLine
41:        End If
42:        If sOptions = vbNullString Then Exit Sub
43:        sOptions = VBA.Left$(sOptions, VBA.Len(sOptions) - 2)
44:        If sOptions = vbNullString Then Exit Sub
45:
46:        If .obtnModule Then
47:            Set moCM = Application.VBE.ActiveCodePane.CodeModule
48:            Call addString(moCM, sOptions)
49:        Else
50:            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
51:                Set moCM = vbComp.CodeModule
52:                Call addString(moCM, sOptions)
53:            Next vbComp
54:        End If
55:    End With
56:    Set objForm = Nothing
57:    Exit Sub
ErrorHandler:
59:    Select Case Err.Number
        Case 91:
61:            Exit Sub
62:        Case Else:
63:            Debug.Print "Error in add Options" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
64:            Call WriteErrorLog("addOptions")
65:    End Select
66:    Err.Clear
67: End Sub

Private Sub addString(ByRef moCM As CodeModule, ByVal sOptions As String)
70:    Dim i           As Long
71:    Dim sLines      As String
72:    With moCM
73:        i = .CountOfDeclarationLines
74:        If i > 0 Then
75:            sLines = .Lines(1, i)
76:            Call .DeleteLines(1, i)
77:        End If
78:        sLines = VBA.Replace(sLines, "Option Explicit", vbNullString)
79:        sLines = VBA.Replace(sLines, "Option Private Module", vbNullString)
80:        sLines = VBA.Replace(sLines, "Option Base 1", vbNullString)
81:        sLines = VBA.Replace(sLines, "Option Base 0", vbNullString)
82:        sLines = VBA.Replace(sLines, "Option Compare Text", vbNullString)
83:        sLines = VBA.Replace(sLines, "Option Compare Binary", vbNullString)
84:
85:        If .Parent.Type <> vbext_ct_StdModule Then
86:            sOptions = VBA.Replace(sOptions, "Option Private Module" & vbNewLine, vbNullString)
87:            sOptions = VBA.Replace(sOptions, "Option Private Module", vbNullString)
88:        End If
89:
90:        sLines = VBA.Replace(sLines, vbNewLine & vbNewLine, "||")
91:        sLines = VBA.Replace(sLines, vbNewLine, "||")
92:        If sLines = vbNullString Then
93:            sLines = sOptions
94:        Else
95:            sLines = sOptions & vbNewLine & VBA.Replace(sLines, "||", vbNewLine)
96:        End If
97:        Call .InsertLines(1, sLines)
98:    End With
End Sub
