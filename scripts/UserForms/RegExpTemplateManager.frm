VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegExpTemplateManager 
   Caption         =   "Template Manager:"
   ClientHeight    =   7845
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
17:    Me.Hide
18: End Sub
    Private Sub lbCancel_Click()
20:    Call btnCancel_Click
21: End Sub
    Private Sub UserForm_Activate()
23:    Dim objList     As ListObject
24:    Me.StartUpPosition = 0
25:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
26:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
27:
28:    Set objList = ThisWorkbook.Worksheets(sSHNAME).ListObjects(sTBGRUPPA)
29:    With objList
30:        ListGrup.List = .ListColumns(1).DataBodyRange.Value2
31:        cmbItemGrupa.List = .ListColumns(1).DataBodyRange.Value2
32:    End With
33: End Sub
    Private Sub lbInsertPattern_Click()
35:    If ListTemplete.ListIndex >= 0 Then
36:        With ActiveSheet
37:            If .Name = sSHNAMETEST Then
38:                Me.Hide
39:                .Cells(2, 3).Value = ListTemplete.List(ListTemplete.ListIndex, 0)
40:            Else
41:                Me.Hide
42:                Dim objRng As Range
43:                Set objRng = GetAddressCell()
44:                If Not objRng Is Nothing Then
45:                    objRng.Resize(1, 1).Value = ListTemplete.List(ListTemplete.ListIndex, 0)
46:                End If
47:            End If
48:        End With
49:    Else
50:        Call MsgBox("Nothing is selected!", vbCritical, "Insert pattern:")
51:    End If
52: End Sub

    Private Function GetAddressCell(Optional sMsg As String = "Select the cell to insert the pattern in:") As Range
55:    On Error GoTo Canceled
56:    Dim sDefault    As String
57:
58:    If TypeName(Selection) = "Range" Then
59:        sDefault = Selection.Address
60:    Else
61:        sDefault = vbNullString
62:    End If
63:    Set GetAddressCell = Application.InputBox(Prompt:=sMsg, Type:=8, Default:=sDefault)
64:    Exit Function
Canceled:
66:    Set GetAddressCell = Nothing
67: End Function

    Private Sub ListGrup_Change()
70:    Dim sSelected   As String
71:
72:    sSelected = SelectedItemList(ListGrup)
73:    Call UpdateListPattern(sSelected)
74:    cmbItemGrupa.Value = sSelected
75:    If lbAddItem.Caption <> sADD Then
76:        txtItemGrupa.Value = sSelected
77:        txtItemPattern.Value = SelectedItemList(ListTemplete)
78:        txtItemDiscript.Value = SelectedItemList(ListTemplete, 1)
79:    End If
80: End Sub
    Private Sub UpdateListPattern(ByVal sSelected As String)
82:    Dim arrVal      As Variant
83:    Dim i           As Integer
84:    Dim objList     As ListObject
85:
86:    Set objList = ThisWorkbook.Worksheets(sSHNAME).ListObjects(sTBPATTERN)
87:    arrVal = objList.DataBodyRange.Value2
88:    With ListTemplete
89:        .Clear
90:        For i = 1 To UBound(arrVal)
91:            If arrVal(i, 1) = sSelected Then
92:                .AddItem arrVal(i, 2)
93:                .List(.ListCount - 1, 1) = arrVal(i, 3)
94:            End If
95:        Next i
96:        If ListTemplete.ListCount > 0 Then .Selected(0) = True
97:    End With
98: End Sub

     Private Sub cmbItemGrupa_Change()
101:    If cmbItemGrupa.ListIndex <> -1 Then ListGrup.Selected(cmbItemGrupa.ListIndex) = True
102: End Sub
     Private Sub ListTemplete_Click()
104:    LbDiscription.Caption = SelectedItemList(ListTemplete, 1)
105:    txtItemPattern.Value = SelectedItemList(ListTemplete)
106:    txtItemDiscript.Value = SelectedItemList(ListTemplete, 1)
107: End Sub
     Private Function SelectedItemList(ByRef objList As MSForms.ListBox, Optional byItem As Byte = 0) As String
109:    With objList
110:        If .ListIndex >= 0 Then SelectedItemList = .List(.ListIndex, byItem)
111:    End With
112: End Function

     Private Sub LbAddNew_Click()
115:    Call FrameVisibleChange(True, LbAddNew.Caption)
116: End Sub
     Private Sub lbDelNew_Click()
118:    Call FrameVisibleChange(True, lbDelNew.Caption)
119: End Sub
     Private Sub lbEditNew_Click()
121:    Call FrameVisibleChange(True, lbEditNew.Caption)
122: End Sub
     Private Sub lbAddItem_Click()
124:    Dim sDoing      As String
125:    Dim objListGrup As ListObject
126:    Dim objListPatt As ListObject
127:    Dim objFinde    As Range
128:
129:    With ThisWorkbook.Worksheets(sSHNAME)
130:        Set objListPatt = .ListObjects(sTBPATTERN)
131:        Set objListGrup = .ListObjects(sTBGRUPPA)
132:    End With
133:
134:    sDoing = lbAddItem.Caption
135:    If optGrupa Then
136:        'группа
137:        If txtItemGrupa.Value <> vbNullString Then
138:            Set objFinde = objListGrup.ListColumns(1).DataBodyRange.Find(txtItemGrupa.Value)
139:            Select Case sDoing
                Case sADD:
141:                    If objFinde Is Nothing Then
142:                        With objListGrup.ListRows.Add
143:                            .Range.Value2 = txtItemGrupa.Value
144:                            Call UpdateTBGruppa
145:                            Call MsgBox("Group [" & txtItemGrupa.Value & "]" & vbNewLine & "Created", vbInformation + vbOKOnly, "Creating a group:")
146:                        End With
147:                    Else
148:                        Call MsgBox("Group [" & txtItemGrupa.Value & "]" & vbNewLine & "Already created", vbCritical + vbOKOnly, "Creating a group:")
149:                        Exit Sub
150:                    End If
151:                Case sDEL:
152:                    If Not objFinde Is Nothing Then
153:                        If objFinde.Row > 0 Then
154:                            If MsgBox("Are you sure you want to delete the group [" & txtItemGrupa.Value & "] ?", vbQuestion + vbYesNo, "Deleting a group:") = vbYes Then
155:                                objListGrup.ListRows(objFinde.Row - 1).Delete
156:                                Call UpdateTBGruppa
157:                                Call MsgBox("Group [" & txtItemGrupa.Value & "]" & vbNewLine & "Deleted", vbInformation + vbOKOnly, "Deleting a group:")
158:                            End If
159:                        End If
160:                    End If
161:                Case sEDI:
162:                    Set objFinde = objListGrup.ListColumns(1).DataBodyRange.Find(SelectedItemList(ListGrup))
163:                    If Not objFinde Is Nothing Then
164:                        If objFinde.Row > 0 Then
165:                            If SelectedItemList(ListGrup) = txtItemGrupa.Value Then
166:                                Call MsgBox("Group [" & txtItemGrupa.Value & "]" & vbNewLine & "You didn't rename it!", vbCritical + vbOKOnly, "Changing a group:")
167:                            Else
168:                                If MsgBox("Are you sure you want to change the group [" & SelectedItemList(ListGrup) & "] on [" & txtItemGrupa.Value & "] ?", vbQuestion + vbYesNo, "Changing a group:") = vbYes Then
169:                                    objListGrup.ListRows(objFinde.Row - 1).Range.Value = txtItemGrupa.Value
170:                                    Call UpdateTBGruppa
171:                                    Call MsgBox("Group [" & SelectedItemList(ListGrup) & "] changed to [" & txtItemGrupa.Value & "]", vbInformation + vbOKOnly, "Changing a group:")
172:                                End If
173:                            End If
174:                        End If
175:                    End If
176:            End Select
177:        Else
178:            Call MsgBox("The input field is not filled in!", vbCritical + vbOKOnly, "The input field is not filled in:")
179:            Exit Sub
180:        End If
181:    Else
182:        'шаблон
183:        If txtItemPattern.Value <> vbNullString Then
184:            Set objFinde = objListPatt.ListColumns(2).DataBodyRange.Find(txtItemPattern.Value)
185:            Select Case sDoing
                Case sADD:
187:                    If objFinde Is Nothing Then
188:                        With objListPatt.ListRows.Add
189:                            .Range(1, 1).Value2 = cmbItemGrupa.Value
190:                            .Range(1, 2).Value2 = txtItemPattern.Value
191:                            .Range(1, 3).Value2 = txtItemDiscript.Value
192:                            Call UpdateTBPattern
193:                            Call MsgBox("Template [" & txtItemPattern.Value & "] in the group [" & cmbItemGrupa.Value & "]" & vbNewLine & "Generated", vbInformation + vbOKOnly, "Creating a template:")
194:                        End With
195:                    Else
196:                        Call MsgBox("Template [" & txtItemPattern.Value & "] in the group [" & cmbItemGrupa.Value & "]" & vbNewLine & "Already created", vbCritical + vbOKOnly, "Creating a template:")
197:                        Exit Sub
198:                    End If
199:                Case sDEL:
200:                    If Not objFinde Is Nothing Then
201:                        If objFinde.Row > 0 Then
202:                            If MsgBox("Are you sure you want to delete a template [" & txtItemPattern.Value & "] ?", vbQuestion + vbYesNo, "Deleting a template:") = vbYes Then
203:                                objListPatt.ListRows(objFinde.Row - 1).Delete
204:                                Call UpdateTBPattern
205:                                Call MsgBox("Template [" & txtItemPattern.Value & "]" & vbNewLine & "Deleted", vbInformation + vbOKOnly, "Deleting a template:")
206:                            End If
207:                        End If
208:                    End If
209:                Case sEDI
210:                    Set objFinde = objListPatt.ListColumns(2).DataBodyRange.Find(SelectedItemList(ListTemplete))
211:                    If Not objFinde Is Nothing Then
212:                        If objFinde.Row > 0 Then
213:                            If SelectedItemList(ListTemplete) = txtItemPattern.Value And SelectedItemList(ListTemplete, 1) = txtItemDiscript.Value Then
214:                                Call MsgBox("Template [" & txtItemPattern.Value & "]" & vbNewLine & "You didn't rename it!", vbCritical + vbOKOnly, "Changing the template:")
215:                            Else
216:                                If MsgBox("Are you sure you want to change the group [" & SelectedItemList(ListTemplete) & "] on [" & txtItemPattern.Value & "] ?", vbQuestion + vbYesNo, "Changing the template:") = vbYes Then
217:                                    objListPatt.ListRows(objFinde.Row - 1).Range(1, 2).Value2 = txtItemPattern.Value
218:                                    objListPatt.ListRows(objFinde.Row - 1).Range(1, 3).Value2 = txtItemDiscript.Value
219:                                    Call UpdateTBPattern
220:                                    Call MsgBox("Template [" & SelectedItemList(ListTemplete) & "] changed to [" & txtItemPattern.Value & "]", vbInformation + vbOKOnly, "Changing the template:")
221:                                End If
222:                            End If
223:                        End If
224:                    End If
225:            End Select
226:        Else
227:            Call MsgBox("The input fields are not filled in!", vbCritical + vbOKOnly, "Input fields are not filled in:")
228:            Exit Sub
229:        End If
230:    End If
231:
232:    Call FrameVisibleChange(False)
233: End Sub
     Private Sub UpdateTBPattern()
235:    Dim objListPattern As ListObject
236:    Set objListPattern = ThisWorkbook.Worksheets(sSHNAME).ListObjects(sTBPATTERN)
237:
238:    Call SortTableListObject(objListPattern, "tbPattern[[#All],[Pattern]]")
239:    Call SortTableListObject(objListPattern, "tbPattern[[#All],[Group]]")
240:
241:    Call UpdateListPattern(SelectedItemList(ListGrup))
242: End Sub
     Private Sub UpdateTBGruppa()
244:    Dim objListGrup As ListObject
245:    Set objListGrup = ThisWorkbook.Worksheets(sSHNAME).ListObjects(sTBGRUPPA)
246:
247:    Call SortTableListObject(objListGrup, "tbGrupa[[#All],[Group]]")
248:
249:    ListGrup.Clear
250:    ListGrup.List = objListGrup.ListColumns(1).DataBodyRange.Value2
251:    cmbItemGrupa.Clear
252:    cmbItemGrupa.List = objListGrup.ListColumns(1).DataBodyRange.Value2
253: End Sub
     Private Sub SortTableListObject(ByRef objList As ListObject, ByVal sKey As String)
255:    With objList.Sort
256:        .SortFields.Clear
257:        .SortFields.Add Key:=Range(sKey), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
258:        .Orientation = xlTopToBottom
259:        .SortMethod = xlPinYin
260:        .Apply
261:    End With
262: End Sub
     Private Sub lbDelItem_Click()
264:    Call FrameVisibleChange(False)
265: End Sub
     Private Sub FrameVisibleChange(ByVal bFlagVisible As Boolean, Optional sCaptionTxt As String = vbNullString)
267:    frmAddItem.visible = bFlagVisible
268:    frmMainAddNew.visible = Not bFlagVisible
269:    txtItemDiscript.visible = Not bFlagVisible
270:    txtItemPattern.visible = Not optGrupa.Value
271:    txtItemDiscript.visible = Not optGrupa.Value
272:    lbDisc.visible = Not optGrupa.Value
273:    cmbItemGrupa.visible = Not optGrupa.Value
274:    lbAddItem.Caption = sCaptionTxt
275:    txtItemGrupa.visible = optGrupa.Value
276:    txtItemGrupa.Value = vbNullString
277:    txtItemPattern.Value = vbNullString
278:    txtItemDiscript.Value = vbNullString
279:    If bFlagVisible And lbAddItem.Caption <> sADD Then
280:        txtItemGrupa.Value = SelectedItemList(ListGrup)
281:        txtItemPattern.Value = SelectedItemList(ListTemplete)
282:        txtItemDiscript.Value = SelectedItemList(ListTemplete, 1)
283:    End If
284: End Sub


