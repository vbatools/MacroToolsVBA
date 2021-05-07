Attribute VB_Name = "G_AddCodeViewForm"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : G_AddCodeViewForm - управление формой создания сниппетов
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Private Module
Option Explicit
    Public Sub AddCode(ByVal X As Long)
12:    Dim frm    As AddEditCode
13:    Set frm = New AddEditCode
14:    With frm
15:        .Caption = "CREATE SNIPPET:"
16:        .lbOK.Caption = "CREATE"
17:        .txtRow = X + 1
18:        .Show
19:    End With
20: End Sub
    Public Sub EditCode(ByVal X As Long, ByRef ListB As MSForms.ListBox)
22:    Dim frm    As AddEditCode
23:    Dim snippets As ListObject
24:    Dim st()   As String
25:    X = X + 1
26:    If X <= 0 Then
27:        Call MsgBox("Nothing is selected in the table!", vbCritical, "Nothing selected:")
28:        Exit Sub
29:    End If
30:    Set frm = New AddEditCode
31:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
32:    With frm
33:        .Caption = "CHANGE SNIPPET:"
34:        .lbOK.Caption = "CHANGE"
35:        st = Split(snippets.ListColumns(3).Range(X, 1), ".")
36:        .cmbENUM.Style = fmStyleDropDownCombo
37:        .cmbENUM.Text = st(0)
38:        .txtENUMBack.Text = .cmbENUM.Text
39:        .txtSNIP.Text = snippets.ListColumns(2).Range(X, 1)
40:        .txtSNIPBack.Text = .txtSNIP.Text
41:        .txtCode.Text = snippets.ListColumns(4).Range(X, 1)
42:        .txtCodeBack.Text = .txtCode.Text
43:        .cmbOBJ.Value = snippets.ListColumns(5).Range(X, 1)
44:        .txtRow = X
45:        .Show
46:    End With
47: End Sub
    Public Sub DeletRow(ByVal X As Long, ByRef objList As MSForms.ListBox)
49:    Dim snippets As ListObject
50:    On Error GoTo errmsg
51:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
52:    If X <= 0 Then
53:        Call MsgBox("Nothing is selected in the table!", vbCritical, "Nothing selected:")
54:        Exit Sub
55:    End If
56:    If MsgBox("Remove SNIPPET: [" & snippets.DataBodyRange.Cells(X, 2).Value & " ] ?", vbYesNo, "Deleting SNIPPET:") = vbYes Then
57:        snippets.ListRows.Item(X).Delete
58:        Call G_AddCodeViewForm.TbAdd(objList)
59:    End If
60:    Exit Sub
errmsg:
62:    If Err.Number = 91 Then
63:        Err.Clear
64:    Else
65:        Debug.Print "Error in Deleterow!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
66:        Call WriteErrorLog("DeletRow")
67:    End If
68: End Sub
Public Sub TbAdd(ByRef objList As MSForms.ListBox)
70:    Dim snippets As ListObject
71:    Dim i      As Long
72:    On Error GoTo errmsg
73:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
74:    With objList
75:        .Clear
76:        For i = 1 To snippets.DataBodyRange.Rows.Count
77:            .AddItem snippets.ListColumns(5).Range(i + 1, 1).Value
78:            .List(i - 1, 1) = snippets.ListColumns(1).Range(i + 1, 1).Value
79:            .List(i - 1, 2) = snippets.ListColumns(2).Range(i + 1, 1).Value
80:            .List(i - 1, 3) = snippets.ListColumns(3).Range(i + 1, 1).Value
81:        Next i
82:    End With
83:    Exit Sub
errmsg:
85:    If Err.Number = 91 Then
86:        Err.Clear
87:    Else
88:        Debug.Print "Error in TbAdd!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
89:        Call WriteErrorLog("TbAdd")
90:    End If
End Sub
