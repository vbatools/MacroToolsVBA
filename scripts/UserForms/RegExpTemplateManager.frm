VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegExpTemplateManager 
   Caption         =   "Template Manager:"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12885
   OleObjectBlob   =   "RegExpTemplateManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegExpTemplateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : RegExpTemplateManager - шаблоны регул€рных выражений
'* Created    : 08-10-2020 14:24
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Private Const sADD  As String = "To create"
Private Const sDEL  As String = "Remove"
Private Const sEDI  As String = "To change"
Private Const sSHNAME As String = "SHSNIPPETS"
Private Const sTBGRUPPA As String = "tbGrupa"
Private Const sTBPATTERN As String = "tbPattern"
Private Const sSHNAMETEST As String = "TestRegExpVBATools"
    Private Sub btnCancel_Click()
11:    Me.Hide
12: End Sub
    Private Sub lbCancel_Click()
14:    Call btnCancel_Click
15: End Sub
    Private Sub UserForm_Activate()
17:    Dim objList     As ListObject
18:    Me.StartUpPosition = 0
19:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
20:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
21:
22:    Set objList = ThisWorkbook.Worksheets(sSHNAME).ListObjects(sTBGRUPPA)
23:    With objList
24:        ListGrup.List = .ListColumns(1).DataBodyRange.Value2
25:        cmbItemGrupa.List = .ListColumns(1).DataBodyRange.Value2
26:    End With
27: End Sub
    Private Sub lbInsertPattern_Click()
29:    If ListTemplete.ListIndex >= 0 Then
30:        With ActiveSheet
31:            If .Name = sSHNAMETEST Then
32:                Me.Hide
33:                .Cells(2, 3).Value = ListTemplete.List(ListTemplete.ListIndex, 0)
34:            Else
35:                Me.Hide
36:                Dim objRng As Range
37:                Set objRng = GetAddressCell()
38:                If Not objRng Is Nothing Then
39:                    objRng.Resize(1, 1).Value = ListTemplete.List(ListTemplete.ListIndex, 0)
40:                End If
41:            End If
42:        End With
43:    Else
44:        Call MsgBox("Nothing is selected!", vbCritical, "Inserting a pattern:")
45:    End If
46: End Sub

    Private Function GetAddressCell(Optional sMsg As String = "Select the cell to insert the pattern in:") As Range
49:    On Error GoTo Canceled
50:    Dim sDefault    As String
51:
52:    If TypeName(Selection) = "Range" Then
53:        sDefault = Selection.Address
54:    Else
55:        sDefault = vbNullString
56:    End If
57:    Set GetAddressCell = Application.InputBox(Prompt:=sMsg, Type:=8, Default:=sDefault)
58:    Exit Function
Canceled:
60:    Set GetAddressCell = Nothing
61: End Function

    Private Sub ListGrup_Change()
64:    Dim sSelected   As String
65:
66:    sSelected = SelectedItemList(ListGrup)
67:    Call UpdateListPattern(sSelected)
68:    cmbItemGrupa.Value = sSelected
69:    If lbAddItem.Caption <> sADD Then
70:        txtItemGrupa.Value = sSelected
71:        txtItemPattern.Value = SelectedItemList(ListTemplete)
72:        txtItemDiscript.Value = SelectedItemList(ListTemplete, 1)
73:    End If
74: End Sub
    Private Sub UpdateListPattern(ByVal sSelected As String)
76:    Dim arrVal      As Variant
77:    Dim i           As Integer
78:    Dim objList     As ListObject
79:
80:    Set objList = ThisWorkbook.Worksheets(sSHNAME).ListObjects(sTBPATTERN)
81:    arrVal = objList.DataBodyRange.Value2
82:    With ListTemplete
83:        .Clear
84:        For i = 1 To UBound(arrVal)
85:            If arrVal(i, 1) = sSelected Then
86:                .AddItem arrVal(i, 2)
87:                .List(.ListCount - 1, 1) = arrVal(i, 3)
88:            End If
89:        Next i
90:        If ListTemplete.ListCount > 0 Then .Selected(0) = True
91:    End With
92: End Sub

    Private Sub cmbItemGrupa_Change()
95:    If cmbItemGrupa.ListIndex <> -1 Then ListGrup.Selected(cmbItemGrupa.ListIndex) = True
96: End Sub
     Private Sub ListTemplete_Click()
98:    LbDiscription.Caption = SelectedItemList(ListTemplete, 1)
99:    txtItemPattern.Value = SelectedItemList(ListTemplete)
100:    txtItemDiscript.Value = SelectedItemList(ListTemplete, 1)
101: End Sub
     Private Function SelectedItemList(ByRef objList As MSForms.ListBox, Optional byItem As Byte = 0) As String
103:    With objList
104:        If .ListIndex >= 0 Then SelectedItemList = .List(.ListIndex, byItem)
105:    End With
106: End Function

     Private Sub LbAddNew_Click()
109:    Call FrameVisibleChange(True, LbAddNew.Caption)
110: End Sub
     Private Sub lbDelNew_Click()
112:    Call FrameVisibleChange(True, lbDelNew.Caption)
113: End Sub
     Private Sub lbEditNew_Click()
115:    Call FrameVisibleChange(True, lbEditNew.Caption)
116: End Sub
     Private Sub lbAddItem_Click()
118:    Dim sDoing      As String
119:    Dim objListGrup As ListObject
120:    Dim objListPatt As ListObject
121:    Dim objFinde    As Range
122:
123:    With ThisWorkbook.Worksheets(sSHNAME)
124:        Set objListPatt = .ListObjects(sTBPATTERN)
125:        Set objListGrup = .ListObjects(sTBGRUPPA)
126:    End With
127:
128:    sDoing = lbAddItem.Caption
129:    If optGrupa Then
130:        'группа
131:        If txtItemGrupa.Value <> vbNullString Then
132:            Set objFinde = objListGrup.ListColumns(1).DataBodyRange.Find(txtItemGrupa.Value)
133:            Select Case sDoing
                Case sADD:
135:                    If objFinde Is Nothing Then
136:                        With objListGrup.ListRows.Add
137:                            .Range.Value2 = txtItemGrupa.Value
138:                            Call UpdateTBGruppa
139:                            Call MsgBox("Group [" & txtItemGrupa.Value & "]" & vbNewLine & "Created", vbInformation + vbOKOnly, "Creating a group:")
140:                        End With
141:                    Else
142:                        Call MsgBox("Group [" & txtItemGrupa.Value & "]" & vbNewLine & "Already created", vbCritical + vbOKOnly, "Creating a group:")
143:                        Exit Sub
144:                    End If
145:                Case sDEL:
146:                    If Not objFinde Is Nothing Then
147:                        If objFinde.Row > 0 Then
148:                            If MsgBox("Are you sure you want to delete the group [" & txtItemGrupa.Value & "] ?", vbQuestion + vbYesNo, "Deleting a group:") = vbYes Then
149:                                objListGrup.ListRows(objFinde.Row - 1).Delete
150:                                Call UpdateTBGruppa
151:                                Call MsgBox("Group [" & txtItemGrupa.Value & "]" & vbNewLine & "Deleted", vbInformation + vbOKOnly, "Deleting a group:")
152:                            End If
153:                        End If
154:                    End If
155:                Case sEDI:
156:                    Set objFinde = objListGrup.ListColumns(1).DataBodyRange.Find(SelectedItemList(ListGrup))
157:                    If Not objFinde Is Nothing Then
158:                        If objFinde.Row > 0 Then
159:                            If SelectedItemList(ListGrup) = txtItemGrupa.Value Then
160:                                Call MsgBox("Group [" & txtItemGrupa.Value & "]" & vbNewLine & "You haven't renamed it!", vbCritical + vbOKOnly, "Changing the group:")
161:                            Else
162:                                If MsgBox("Are you sure you want to name the group [" & SelectedItemList(ListGrup) & "] on [" & txtItemGrupa.Value & "] ?", vbQuestion + vbYesNo, "Changing the group:") = vbYes Then
163:                                    objListGrup.ListRows(objFinde.Row - 1).Range.Value = txtItemGrupa.Value
164:                                    Call UpdateTBGruppa
165:                                    Call MsgBox("Group [" & SelectedItemList(ListGrup) & "] changed to [" & txtItemGrupa.Value & "]", vbInformation + vbOKOnly, "Changing the group:")
166:                                End If
167:                            End If
168:                        End If
169:                    End If
170:            End Select
171:        Else
172:            Call MsgBox("The input field is not filled in!", vbCritical + vbOKOnly, "The input field is not filled in:")
173:            Exit Sub
174:        End If
175:    Else
176:        'шаблон
177:        If txtItemPattern.Value <> vbNullString Then
178:            Set objFinde = objListPatt.ListColumns(2).DataBodyRange.Find(txtItemPattern.Value)
179:            Select Case sDoing
                Case sADD:
181:                    If objFinde Is Nothing Then
182:                        With objListPatt.ListRows.Add
183:                            .Range(1, 1).Value2 = cmbItemGrupa.Value
184:                            .Range(1, 2).Value2 = txtItemPattern.Value
185:                            .Range(1, 3).Value2 = txtItemDiscript.Value
186:                            Call UpdateTBPattern
187:                            Call MsgBox("Template [" & txtItemPattern.Value & "] in the group [" & cmbItemGrupa.Value & "]" & vbNewLine & "Generated", vbInformation + vbOKOnly, "Creating a template:")
188:                        End With
189:                    Else
190:                        Call MsgBox("Template [" & txtItemPattern.Value & "] in the group [" & cmbItemGrupa.Value & "]" & vbNewLine & "Already created", vbCritical + vbOKOnly, "Creating a template:")
191:                        Exit Sub
192:                    End If
193:                Case sDEL:
194:                    If Not objFinde Is Nothing Then
195:                        If objFinde.Row > 0 Then
196:                            If MsgBox("Are you sure you want to delete the template [" & txtItemPattern.Value & "] ?", vbQuestion + vbYesNo, "Deleting a template:") = vbYes Then
197:                                objListPatt.ListRows(objFinde.Row - 1).Delete
198:                                Call UpdateTBPattern
199:                                Call MsgBox("Template [" & txtItemPattern.Value & "]" & vbNewLine & "Deleted", vbInformation + vbOKOnly, "Deleting a template:")
200:                            End If
201:                        End If
202:                    End If
203:                Case sEDI
204:                    Set objFinde = objListPatt.ListColumns(2).DataBodyRange.Find(SelectedItemList(ListTemplete))
205:                    If Not objFinde Is Nothing Then
206:                        If objFinde.Row > 0 Then
207:                            If SelectedItemList(ListTemplete) = txtItemPattern.Value And SelectedItemList(ListTemplete, 1) = txtItemDiscript.Value Then
208:                                Call MsgBox("Template [" & txtItemPattern.Value & "]" & vbNewLine & "You haven't renamed it!", vbCritical + vbOKOnly, "Changing the template:")
209:                            Else
210:                                If MsgBox("Are you sure you want to name the group [" & SelectedItemList(ListTemplete) & "] on [" & txtItemPattern.Value & "] ?", vbQuestion + vbYesNo, "Changing the template:") = vbYes Then
211:                                    objListPatt.ListRows(objFinde.Row - 1).Range(1, 2).Value2 = txtItemPattern.Value
212:                                    objListPatt.ListRows(objFinde.Row - 1).Range(1, 3).Value2 = txtItemDiscript.Value
213:                                    Call UpdateTBPattern
214:                                    Call MsgBox("Template [" & SelectedItemList(ListTemplete) & "] changed to [" & txtItemPattern.Value & "]", vbInformation + vbOKOnly, "Changing the template:")
215:                                End If
216:                            End If
217:                        End If
218:                    End If
219:            End Select
220:        Else
221:            Call MsgBox("The input fields are not filled in!", vbCritical + vbOKOnly, "Input fields are not filled in:")
222:            Exit Sub
223:        End If
224:    End If
225:
226:    Call FrameVisibleChange(False)
227: End Sub
     Private Sub UpdateTBPattern()
229:    Dim objListPattern As ListObject
230:    Set objListPattern = ThisWorkbook.Worksheets(sSHNAME).ListObjects(sTBPATTERN)
231:
232:    Call SortTableListObject(objListPattern, "tbPattern[[#All],[ѕатеррн]]")
233:    Call SortTableListObject(objListPattern, "tbPattern[[#All],[√руппа]]")
234:
235:    Call UpdateListPattern(SelectedItemList(ListGrup))
236: End Sub
     Private Sub UpdateTBGruppa()
238:    Dim objListGrup As ListObject
239:    Set objListGrup = ThisWorkbook.Worksheets(sSHNAME).ListObjects(sTBGRUPPA)
240:
241:    Call SortTableListObject(objListGrup, "tbGrupa[[#All],[√руппа]]")
242:
243:    ListGrup.Clear
244:    ListGrup.List = objListGrup.ListColumns(1).DataBodyRange.Value2
245:    cmbItemGrupa.Clear
246:    cmbItemGrupa.List = objListGrup.ListColumns(1).DataBodyRange.Value2
247: End Sub
     Private Sub SortTableListObject(ByRef objList As ListObject, ByVal sKey As String)
249:    With objList.Sort
250:        .SortFields.Clear
251:        .SortFields.Add Key:=Range(sKey), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
252:        .Orientation = xlTopToBottom
253:        .SortMethod = xlPinYin
254:        .Apply
255:    End With
256: End Sub
     Private Sub lbDelItem_Click()
258:    Call FrameVisibleChange(False)
259: End Sub
     Private Sub FrameVisibleChange(ByVal bFlagVisible As Boolean, Optional sCaptionTxt As String = vbNullString)
261:    frmAddItem.visible = bFlagVisible
262:    frmMainAddNew.visible = Not bFlagVisible
263:    txtItemDiscript.visible = Not bFlagVisible
264:    txtItemPattern.visible = Not optGrupa.Value
265:    txtItemDiscript.visible = Not optGrupa.Value
266:    lbDisc.visible = Not optGrupa.Value
267:    cmbItemGrupa.visible = Not optGrupa.Value
268:    lbAddItem.Caption = sCaptionTxt
269:    txtItemGrupa.visible = optGrupa.Value
270:    txtItemGrupa.Value = vbNullString
271:    txtItemPattern.Value = vbNullString
272:    txtItemDiscript.Value = vbNullString
273:    If bFlagVisible And lbAddItem.Caption <> sADD Then
274:        txtItemGrupa.Value = SelectedItemList(ListGrup)
275:        txtItemPattern.Value = SelectedItemList(ListTemplete)
276:        txtItemDiscript.Value = SelectedItemList(ListTemplete, 1)
277:    End If
278: End Sub


