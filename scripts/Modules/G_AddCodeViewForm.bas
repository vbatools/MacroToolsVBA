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
4:    Dim frm    As AddEditCode
5:    Set frm = New AddEditCode
6:    With frm
7:        .Caption = "СОЗДАТЬ SNIPPET:"
8:        .lbOK.Caption = "СОЗДАТЬ"
9:        .txtRow = X + 1
10:        .Show
11:    End With
12: End Sub
    Public Sub EditCode(ByVal X As Long, ByRef ListB As MSForms.ListBox)
14:    Dim frm    As AddEditCode
15:    Dim snippets As ListObject
16:    Dim st()   As String
17:    X = X + 1
18:    If X <= 0 Then
19:        Call MsgBox("Ничего не выбрано в таблице!", vbCritical, "Ничего не выбрано:")
20:        Exit Sub
21:    End If
22:    Set frm = New AddEditCode
23:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
24:    With frm
25:        .Caption = "ИЗМЕНИТЬ SNIPPET:"
26:        .lbOK.Caption = "ИЗМЕНИТЬ"
27:        st = Split(snippets.ListColumns(3).Range(X, 1), ".")
28:        .cmbENUM.Style = fmStyleDropDownCombo
29:        .cmbENUM.Text = st(0)
30:        .txtENUMBack.Text = .cmbENUM.Text
31:        .txtSNIP.Text = snippets.ListColumns(2).Range(X, 1)
32:        .txtSNIPBack.Text = .txtSNIP.Text
33:        .txtCode.Text = snippets.ListColumns(4).Range(X, 1)
34:        .txtCodeBack.Text = .txtCode.Text
35:        .cmbOBJ.Value = snippets.ListColumns(5).Range(X, 1)
36:        .txtRow = X
37:        .Show
38:    End With
39: End Sub
    Public Sub DeletRow(ByVal X As Long, ByRef objList As MSForms.ListBox)
41:    Dim snippets As ListObject
42:    On Error GoTo errMsg
43:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
44:    If X <= 0 Then
45:        Call MsgBox("Ничего не выбрано в таблице!", vbCritical, "Ничего не выбрано:")
46:        Exit Sub
47:    End If
48:    If MsgBox("Удалить SNIPPET: [ " & snippets.DataBodyRange.Cells(X, 2).Value & " ] ?", vbYesNo, "Удаление SNIPPET:") = vbYes Then
49:        snippets.ListRows.Item(X).Delete
50:        Call G_AddCodeViewForm.TbAdd(objList)
51:    End If
52:    Exit Sub
errMsg:
54:    If Err.Number = 91 Then
55:        Err.Clear
56:    Else
57:        Debug.Print "Ошибка в DeletRow!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
58:        Call WriteErrorLog("DeletRow")
59:    End If
60: End Sub
    Public Sub TbAdd(ByRef objList As MSForms.ListBox)
62:    Dim snippets As ListObject
63:    Dim i      As Long
64:    On Error GoTo errMsg
65:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
66:    With objList
67:        .Clear
68:        For i = 1 To snippets.DataBodyRange.Rows.Count
69:            .AddItem snippets.ListColumns(5).Range(i + 1, 1).Value
70:            .List(i - 1, 1) = snippets.ListColumns(1).Range(i + 1, 1).Value
71:            .List(i - 1, 2) = snippets.ListColumns(2).Range(i + 1, 1).Value
72:            .List(i - 1, 3) = snippets.ListColumns(3).Range(i + 1, 1).Value
73:        Next i
74:    End With
75:    Exit Sub
errMsg:
77:    If Err.Number = 91 Then
78:        Err.Clear
79:    Else
80:        Debug.Print "Ошибка в TbAdd!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
81:        Call WriteErrorLog("TbAdd")
82:    End If
83: End Sub
