VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsAddCommentsProc 
   Caption         =   "Наcтройки Комментарий Кода:"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   OleObjectBlob   =   "SettingsAddCommentsProc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsAddCommentsProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module   :   SettingsAddCommentsProc - настройка автодокументирования кода
'* Created  :   13-01-2020 14:33
'* Author   :   amkorobchanu
'* Contacts :   http://vbatools.ru/ https://vk.com/vbatools
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

    Private Sub btnCancel_Click()
10:    Unload Me
11: End Sub
    Private Sub lbCancel_Click()
13:    Call btnCancel_Click
14: End Sub
    Private Sub UserForm_Activate()
16:    Me.StartUpPosition = 0
17:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
18:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
19:
20:    Dim TBComment As ListObject
21:    Set TBComment = SHSNIPPETS.ListObjects(C_Const.TB_COMMENT)
22:    With TBComment.ListColumns(2)
23:        txtName.Value = .Range(2, 1).Value
24:        txtName1.Value = txtName.Value
25:        txtContacts.Value = .Range(3, 1).Value
26:        txtContacts1.Value = txtContacts.Value
27:        txtCopyright.Value = .Range(4, 1).Value
28:        txtCopyright1.Value = txtCopyright.Value
29:        txtOther.Value = .Range(5, 1).Value
30:        txtOther1.Value = txtOther.Value
31:    End With
32:    lbOk.Enabled = False
33: End Sub
    Private Sub txtName_Change()
35:    Call TakeSave
36: End Sub
    Private Sub txtContacts_Change()
38:    Call TakeSave
39: End Sub
    Private Sub txtCopyright_Change()
41:    Call TakeSave
42: End Sub
    Private Sub txtOther_Change()
44:    Call TakeSave
45: End Sub
    Private Sub TakeSave()
47:    Dim Flag As Boolean
48:    If txtName1.Value <> txtName.Value Then Flag = True
49:    If txtContacts1.Value <> txtContacts.Value Then Flag = True
50:    If txtCopyright1.Value <> txtCopyright.Value Then Flag = True
51:    If txtOther1.Value <> txtOther.Value Then Flag = True
52:    lbOk.Enabled = Flag
53: End Sub
    Private Sub lbOK_Click()
55:    Dim TBComment As ListObject
56:    Set TBComment = SHSNIPPETS.ListObjects(C_Const.TB_COMMENT)
57:    With TBComment.ListColumns(2)
58:        .Range(2, 1).Value = txtName.Value
59:        txtName1.Value = txtName.Value
60:        .Range(3, 1).Value = txtContacts.Value
61:        txtContacts1.Value = txtContacts.Value
62:        .Range(4, 1).Value = txtCopyright.Value
63:        txtCopyright1.Value = txtCopyright.Value
64:        .Range(5, 1).Value = txtOther.Value
65:        txtOther1.Value = txtOther.Value
66:    End With
67:    lbOk.Enabled = False
68: End Sub

