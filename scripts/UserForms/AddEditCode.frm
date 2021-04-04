VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddEditCode 
   Caption         =   "UserForm1"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11175
   OleObjectBlob   =   "AddEditCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddEditCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : AddEditCode - изменене снипетов
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Private m_clsAnchorsEditAdd As CAnchors
    Private Sub btnCancel_Click()
11:    Unload Me
12: End Sub
    Private Sub lbCancel_Click()
14:    Call btnCancel_Click
15: End Sub
    Private Sub lbClearTextbox_Click()
17:    txtCode.Text = vbNullString
18: End Sub

    Private Sub lbHelp_Click()
21:    Call URLLinks(C_Const.URL_STYLE_SNNIP)
22: End Sub

    Private Sub txtCode_Change()
25:    Call BorderColorCntr(txtCode)
26:    Call SaveBtn
27: End Sub
    Private Sub txtCode_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
29:    lbEn.visible = False
30:    If KeyAscii > 127 Then
31:        KeyAscii = 0
32:        lbEn.visible = True
33:    End If
34: End Sub
    Private Sub cmbENUM_Change()
36:    Dim snippets    As ListObject
37:    Dim i_row       As Long
38:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_DESCRIPTION)
39:    lbPreView.Caption = cmbENUM.Text & txtSNIP.Text
40:    On Error GoTo errmsg
41:    i_row = snippets.ListColumns(1).DataBodyRange.Find(What:=cmbENUM.Text, LookIn:=xlValues, LookAt:=xlWhole).Row
42:    txtDescription.Text = snippets.Range(i_row, 3)
43:    Call BorderColorCntr(cmbENUM)
44:    Call SaveBtn
45:    Exit Sub
errmsg:
47:    If Err.Number = 91 Then
48:        Err.Clear
49:    End If
50: End Sub
    Private Sub txtSNIP_Change()
52:    lbPreView.Caption = cmbENUM.Text & txtSNIP.Text
53:    Call BorderColorCntr(txtSNIP)
54:    Call SaveBtn
55: End Sub
    Private Sub txtSNIP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
57:    Call InterKeyAscii(KeyAscii)
58: End Sub
    Private Sub InterKeyAscii(ByVal KeyAscii As MSForms.ReturnInteger)
60:    Select Case KeyAscii
        Case 65 To 90:
62:        Case 97 To 122:
63:        Case Else: KeyAscii = 0
64:    End Select
65: End Sub
    Private Sub UserForm_Activate()
67:    lbOK.Enabled = False
68:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
69: End Sub
    Private Sub UserForm_Terminate()
71:    Set m_clsAnchorsEditAdd = Nothing
72: End Sub
    Private Sub UserForm_Initialize()
74:    Me.StartUpPosition = 0
75:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
76:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
77:
78:    Dim snippets    As ListObject
79:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_DESCRIPTION)
80:    Me.cmbOBJ.AddItem "VBA"
81:    Me.cmbOBJ.AddItem "EXCEL"
82:    Me.cmbENUM.List = GetUniqueValueFromRange(snippets.ListColumns(1).Range)
83:    Me.cmbENUM.RemoveItem (0)
84:    lbClearTextbox.Picture = Application.CommandBars.GetImageMso("ExcludeSelectedRecord", 18, 18)
85:    Call AddCAnchors
86: End Sub
     Private Sub AddCAnchors()
88:    Set m_clsAnchorsEditAdd = New CAnchors
89:    Set m_clsAnchorsEditAdd.objParent = Me
90:    ' restrict minimum size of userform
91:    m_clsAnchorsEditAdd.MinimumWidth = 560
92:    m_clsAnchorsEditAdd.MinimumHeight = 405
93:    With m_clsAnchorsEditAdd
94:        .funAnchor("txtCode").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleRight Or enumAnchorStyleBottom Or enumAnchorStyleLeft
95:        .funAnchor("cmbOBJ").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleRight Or enumAnchorStyleLeft
96:        .funAnchor("txtSNIP").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleRight Or enumAnchorStyleLeft
97:        .funAnchor("cmbENUM").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleLeft
98:        .funAnchor("lbCancel").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
99:        .funAnchor("lbOK").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
100:        .funAnchor("lbClearTextbox").AnchorStyle = enumAnchorStyleRight
101:        .funAnchor("lbHelp").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
102:        .funAnchor("txtDescription").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleLeft Or enumAnchorStyleRight
103:        .funAnchor("Label5").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleLeft
104:    End With
105: End Sub
     Private Sub BorderColorCntr(ByRef cntr As MSForms.control)
107:    If Trim$(Replace(cntr.Text, Chr$(13), vbNullString)) <> vbNullString Then
108:        cntr.BorderColor = &H8000000D
109:    Else
110:        cntr.BorderColor = &HC0C0FF
111:    End If
112: End Sub
     Private Sub SaveBtn()
114:    lbOK.Enabled = False
115:    If txtCode.Text <> vbNullString And cmbENUM.Text <> vbNullString And txtSNIP.Text <> vbNullString Then
116:        If txtCode.Text <> txtCodeBack.Text Or cmbENUM.Text <> txtENUMBack.Text Or txtSNIP.Text <> txtSNIPBack.Text Then
117:            lbOK.Enabled = True
118:        End If
119:    End If
120: End Sub
     Private Sub lbOK_Click()
122:    Dim snippets    As ListObject
123:    Dim row_i       As Long
124:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
125:    With snippets
126:        row_i = CLng(txtRow.Text)
127:        If MsgBox(lbOK.Caption & " SNIPPET: [ " & lbPreView.Caption & " ] ?", vbYesNo, lbOK.Caption & " SNIPPET:") = vbYes Then
128:            If lbOK.Caption = "СОЗДАТЬ" Then
129:                .ListRows.Add Position:=row_i, AlwaysInsert:=True
130:                row_i = row_i + 1
131:            ElseIf lbOK.Caption = "ИЗМЕНИТЬ" Then
132:                'ничего не делаем
133:            Else
134:                Debug.Print "Ошибка в AddEditCode!" & vbLf & "Кнопка [lbOK] не содкржит подписи"
135:                Exit Sub
136:            End If
137:        End If
138:        .ListColumns(6).Range(row_i, 1) = cmbENUM.Text
139:        .ListColumns(2).Range(row_i, 1) = txtSNIP.Text
140:        '.ListColumns(3).Range(row_i, 1) = lbPreView.Caption
141:        .ListColumns(4).Range(row_i, 1) = txtCode.Text
142:        .ListColumns(5).Range(row_i, 1) = cmbOBJ.Value
143:    End With
144:    Call SortTb(snippets)
145:    Unload Me
146: End Sub
     Private Sub SortTb(ByVal snippets As ListObject)
148:    With snippets
149:        .Sort.SortFields.Clear
150:        .Sort.SortFields.Add Key:=Range("C:C"), SortOn:=xlSortOnValues, Order:= _
                        xlAscending, DataOption:=xlSortNormal
152:        With .Sort
153:            .Header = xlYes
154:            .MatchCase = False
155:            .Orientation = xlTopToBottom
156:            .SortMethod = xlPinYin
157:            .Apply
158:        End With
159:    End With
160: End Sub

Private Function GetUniqueValueFromRange(ByVal Arr As Variant) As String()
163:    Dim vItem, li   As Long
164:    Dim avArr()     As String
165:    li = 0
166:    With New Collection
167:        On Error Resume Next
168:        For Each vItem In Arr
169:            If Len(CStr(vItem)) Then
170:                .Add vItem, CStr(vItem)
171:                If Err = 0 Then
172:                    ReDim Preserve avArr(0 To li)
173:                    avArr(li) = CStr(vItem)
174:                    li = li + 1
175:                Else
176:                    Err.Clear
177:                End If
178:            End If
179:        Next
180:    End With
181:    If li = 0 Then
182:        ReDim Preserve avArr(0 To 0)
183:        avArr(0) = 0
184:    End If
185:    GetUniqueValueFromRange = avArr
End Function
