VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCodeView 
   Caption         =   "База SNIPPET's:"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13995
   OleObjectBlob   =   "AddCodeView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddCodeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : AddCodeView - редактирование снипетов
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Private m_colContextMenus As Collection
Private m_clsAnchors As CAnchors
    Private Sub btnCancel_Click()
12:    Unload Me
13: End Sub
    Private Sub lbCancel_Click()
15:    Call btnCancel_Click
16: End Sub

    Private Sub lbHelp_Click()
19:    Call URLLinks(C_Const.URL_STYLE_SNNIP)
20: End Sub

    Private Sub ListBoxMain_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
23:    Call G_AddCodeViewForm.EditCode(ListBoxMain.ListIndex + 1, ListBoxMain)
24: End Sub
    Private Sub UserForm_Terminate()
26:    Do While m_colContextMenus.Count > 0
27:        m_colContextMenus.Remove m_colContextMenus.Count
28:    Loop
29:    Set m_colContextMenus = Nothing
30:    Set m_clsAnchors = Nothing
31: End Sub
    Private Sub UserForm_Activate()
33:    Call G_AddCodeViewForm.TbAdd(ListBoxMain)
34:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
35: End Sub
    Private Sub UserForm_Initialize()
37:    Me.StartUpPosition = 0
38:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
39:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
40:    Call AddContextMenu
41:    Call AddCAnchors
42: End Sub
    Private Sub AddContextMenu()
44:    Dim clsContextMenu As CTextBox_ContextMenu
45:    Set m_colContextMenus = New Collection
46:    Set clsContextMenu = New CTextBox_ContextMenu
47:    With clsContextMenu
48:        Set .TBox = AddCodeView.ListBoxMain
49:        Set .prParent = Me
50:    End With
51:    m_colContextMenus.Add clsContextMenu, CStr(m_colContextMenus.Count + 1)
52: End Sub
    Private Sub AddCAnchors()
54:    Set m_clsAnchors = New CAnchors
55:    Set m_clsAnchors.objParent = Me
56:    ' restrict minimum size of userform
57:    m_clsAnchors.MinimumWidth = 705
58:    m_clsAnchors.MinimumHeight = 378
59:    With m_clsAnchors
60:        .funAnchor("ListBoxMain").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleBottom
61:        .funAnchor("txtCode").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleRight Or enumAnchorStyleBottom Or enumAnchorStyleLeft
62:        .funAnchor("lbOK").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleLeft
63:        .funAnchor("lbCancel").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
64:        .funAnchor("lbHelp").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
65:    End With
66: End Sub
    Private Sub ListBoxMain_Click()
68:    Dim row_i       As Long
69:    Dim snippets    As ListObject
70:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
71:    With ListBoxMain
72:        If .ListIndex > -1 Then
73:            row_i = snippets.ListColumns(2).DataBodyRange.Find(What:=.List(.ListIndex, 2), LookIn:=xlValues, LookAt:=xlWhole).Row
74:            If txtCode.Text <> txtCodeBackCap.Text And txtCodeBackCap.Text <> vbNullString Then
75:                If MsgBox("Сохранить изменения Кода ?", vbYesNo, "Сохранение кода:") = vbYes Then
76:                    snippets.ListColumns(4).Range(CLng(txtRow.Text), 1).Value = txtCode.Text
77:                End If
78:            End If
79:            txtCode.Text = snippets.ListColumns(4).Range(row_i, 1).Value
80:            txtCodeBackCap.Text = txtCode.Text
81:            txtRow.Text = row_i
82:            lbOk.Enabled = False
83:        End If
84:    End With
85: End Sub
    Private Sub lbOK_Click()
87:    Dim row_i       As Long
88:    Dim snippets    As ListObject
89:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
90:    With ListBoxMain
91:        If .ListIndex > -1 Then
92:            row_i = snippets.ListColumns(2).DataBodyRange.Find(What:=.List(.ListIndex, 2), LookIn:=xlValues, LookAt:=xlWhole).Row
93:            snippets.ListColumns(4).Range(row_i, 1).Value = txtCode.Text
94:            txtCodeBackCap.Text = txtCode.Text
95:            lbOk.Enabled = False
96:        End If
97:    End With
98: End Sub
     Private Sub txtCode_Change()
100:    If txtCode.Text <> txtCodeBackCap.Text Then
101:        lbOk.Enabled = True
102:    End If
103: End Sub
     Private Sub txtSerch_Change()
105:    Dim strVar      As String
106:    strVar = txtSerch.Text
107:    TB_Result.visible = False
108:    If strVar <> vbNullString Then
109:        LB_ClearList.visible = True
110:    Else
111:        LB_ClearList.visible = False
112:    End If
113:    Call G_AddCodeViewForm.TbAdd(ListBoxMain)
114:    Call SerchSnippet
115: End Sub
     Private Sub LB_ClearList_Click()
117:    LB_ClearList.visible = False
118:    txtSerch.Text = vbNullString
119:    TB_Result.visible = False
120:    Call G_AddCodeViewForm.TbAdd(ListBoxMain)
121: End Sub
Private Sub SerchSnippet()
123:    On Error GoTo MyBtnSerch_Err
124:    Dim Flag        As Boolean
125:    Dim X           As Long
126:    Dim strVar      As String
127:    strVar = txtSerch.Text
128:    Flag = True
129:    If strVar = vbNullString Then    'введено пусто
130:        Call G_AddCodeViewForm.TbAdd(ListBoxMain)
131:        Exit Sub
132:    End If
133:    If strVar = " " Then    'введено пусто
134:        Call G_AddCodeViewForm.TbAdd(ListBoxMain)
135:        TB_Result.visible = True
136:        LB_ClearList.visible = True
137:        Exit Sub
138:    End If
139:    For X = ListBoxMain.ListCount - 1 To 0 Step -1
140:        If StrConv(ListBoxMain.List(X, 2), vbLowerCase) Like StrConv("*" & strVar & "*", vbLowerCase) Then
141:            Flag = False
142:        Else
143:            ListBoxMain.RemoveItem X
144:            LB_ClearList.visible = True
145:        End If
146:    Next X
147:    If Flag Then
148:        LB_ClearList.visible = True
149:        TB_Result.visible = True
150:    End If
151:    Exit Sub
MyBtnSerch_Err:
153:    Unload Me
154:    MsgBox Err.Description & vbCrLf & "в AddCodeView.SerchSnippet " & vbCrLf & "в строке " & Erl, vbExclamation + vbOKOnly, "Ошибка:"
155:    Call WriteErrorLog("AddCodeView.SerchSnippet")
End Sub
