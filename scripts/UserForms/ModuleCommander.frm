VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModuleCommander 
   Caption         =   "VBA Project Manager:"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10830
   OleObjectBlob   =   "ModuleCommander.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModuleCommander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : ModuleCommander - управление модулями VBA вставка копирование удаление
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Private m_clsAnchors As CAnchors
    Private Sub UserForm_Initialize()
13:    Me.StartUpPosition = 0
14:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
15:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
16:
17:    Set m_clsAnchors = New CAnchors
18:    Set m_clsAnchors.objParent = Me
19:    ' restrict minimum size of userform
20:    m_clsAnchors.MinimumWidth = 546
21:    m_clsAnchors.MinimumHeight = 441
22:    With m_clsAnchors
23:        .funAnchor("cmbMain").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft
24:        .funAnchor("ListCode").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft Or enumAnchorStyleBottom
25:        .funAnchor("lbRemoveModule").AnchorStyle = enumAnchorStyleBottom
26:        .funAnchor("lbExportModule").AnchorStyle = enumAnchorStyleBottom
27:        .funAnchor("lbImportModule").AnchorStyle = enumAnchorStyleBottom
28:        .funAnchor("lbCopytModule").AnchorStyle = enumAnchorStyleBottom
29:        .funAnchor("lbCancel").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
30:        .funAnchor("Label2").AnchorStyle = enumAnchorStyleRight
31:        .funAnchor("ListFilter1").AnchorStyle = enumAnchorStyleRight
32:        .funAnchor("CheckAll").AnchorStyle = enumAnchorStyleRight
33:        .funAnchor("ListFilter2").AnchorStyle = enumAnchorStyleRight
34:        .funAnchor("lbMsg").AnchorStyle = enumAnchorStyleRight
35:        .funAnchor("Label3").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
36:        .funAnchor("cmbMainCopy").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
37:    End With
38:    With ListFilter1
39:        .AddItem "All"
40:        .AddItem "Empty"
41:        .AddItem "Not empty"
42:        .AddItem "Reset everything"
43:    End With
44:    With ListFilter2
45:        .AddItem "Code Module"
46:        .AddItem "UserForm"
47:        .AddItem "Document Module"
48:        .AddItem "Class Module"
49:        .AddItem "ActiveX Designer"
50:    End With
51:    ListFilter1.Selected(0) = True
52: End Sub
    Private Sub UserForm_Activate()
54:    Dim vbProj      As VBIDE.VBProject
55:    If Workbooks.Count = 0 Then
56:        Unload Me
57:        Call MsgBox("No open" & Chr(34) & "Excel Files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Mistake:")
58:        Exit Sub
59:    End If
60:    With Me.cmbMain
61:        .Clear
62:        cmbMainCopy.Clear
63:        On Error Resume Next
64:        For Each vbProj In Application.VBE.VBProjects
65:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
66:        Next
67:        .Value = ActiveWorkbook.Name
68:        On Error GoTo 0
69:        On Error GoTo ErrorHandler
70:        cmbMainCopy.Value = .List(0)
71:    End With
72:    lbMsg.visible = True
73:    Call CheckAll_Change
74:    Exit Sub
ErrorHandler:
76:    Unload Me
77:    Select Case Err.Number
        Case Else:
79:            Call MsgBox("Mistake! in ModuleCommander.UserForm_Activate" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbExclamation, "Mistake:")
80:            Call WriteErrorLog("ModuleCommander.UserForm_Activate")
81:    End Select
82:    Err.Clear
83: End Sub
    Private Sub UserForm_Terminate()
85:    Set m_clsAnchors = Nothing
86: End Sub
    Private Sub lbImportModule_Click()
88:    Application.ScreenUpdating = False
89:    If cmbMain.Value <> vbNullString Then
90:        Call S_ModuleCommander.ImportAllModules(Workbooks(cmbMain.Value))
91:    End If
92:    Call AddListCode
93:    Call FilterRun
94:    Application.ScreenUpdating = True
95: End Sub
     Private Sub lbExportModule_Click()
97:    Application.ScreenUpdating = False
98:    If cmbMain.Value <> vbNullString And lbMsg.visible = False Then
99:        Call S_ModuleCommander.ExportAllModules(Workbooks(cmbMain.Value), SelectedListItems)
100:    Else
101:        Call MsgBox("Nothing is selected!", vbInformation, "Exporting a project:")
102:    End If
103:    Call AddListCode
104:    Call FilterRun
105:    Application.ScreenUpdating = True
106: End Sub
     Private Sub lbRemoveModule_Click()
108:    Application.ScreenUpdating = False
109:    If cmbMain.Value <> vbNullString And lbMsg.visible = False Then
110:        Call S_ModuleCommander.DeleteAllModulesInActiveProject(Workbooks(cmbMain.Value), SelectedListItems)
111:    Else
112:        Call MsgBox("Nothing is selected!", vbInformation, "Deleting a project:")
113:    End If
114:    Call AddListCode
115:    Call FilterRun
116:    Application.ScreenUpdating = True
117: End Sub
     Private Sub lbCopytModule_Click()
119:    If MsgBox("Copy VBA modules from [" & cmbMain.Value & "] in [" & cmbMainCopy.Value & "] ?", vbYesNo + vbQuestion, "Copying VBA modules:") = vbNo Then Exit Sub
120:    Application.ScreenUpdating = False
121:    If cmbMain.Value <> vbNullString And cmbMainCopy.Value <> vbNullString And lbMsg.visible = False Then
122:        Dim i       As Integer
123:        Dim vbCompObj As VBIDE.VBComponent
124:        With ListCode
125:            For i = 0 To .ListCount - 1
126:                If .Selected(i) Then
127:                    Set vbCompObj = Workbooks(cmbMain.Value).VBProject.VBComponents(.List(i, 2))
128:                    If vbCompObj.CodeModule.CountOfLines <> 0 Then
129:                        Call E_AddEnum.AddModuleToProject(.List(i, 2), vbCompObj.Type, vbCompObj.CodeModule.Lines(1, vbCompObj.CodeModule.CountOfLines), Workbooks(cmbMainCopy.Value).VBProject)
130:                    Else
131:                        Call E_AddEnum.AddModuleToProject(.List(i, 2), vbCompObj.Type, vbNullString, Workbooks(cmbMainCopy.Value).VBProject)
132:                    End If
133:                End If
134:            Next i
135:        End With
136:        Call MsgBox("Copying VBA modules from [" & cmbMain.Value & "] in [" & cmbMainCopy.Value & "] completed", vbOKOnly + vbInformation, "Copying VBA modules:")
137:    Else
138:        Call MsgBox("Nothing is selected!", vbInformation, "Copying a project:")
139:    End If
140:    Call AddListCode
141:    Call FilterRun
142:    Application.ScreenUpdating = True
143: End Sub
     Private Sub cmbCancel_Click()
145:    Unload Me
146: End Sub
     Private Sub lbCancel_Click()
148:    Call cmbCancel_Click
149: End Sub
     Private Sub ListFilter2_Change()
151:    Call FilterRun
152: End Sub
     Private Sub ListFilter1_Change()
154:    Call FilterRun
155: End Sub
     Private Sub CheckAll_Change()
157:    Dim i           As Integer
158:    With ListFilter2
159:        For i = 0 To .ListCount - 1
160:            .Selected(i) = CheckAll.Value
161:        Next i
162:    End With
163: End Sub
     Private Sub FilterRun()
165:    Application.ScreenUpdating = False
166:    Select Case SelectedListItem(ListFilter1)
        Case "All":
168:            Call FilterAllOrEmpty(True)
169:        Case "Reset everything":
170:            Call FilterAllOrEmpty(False)
171:        Case "Empty":
172:            Call FilterTypeModule(False)
173:        Case "Not empty":
174:            Call FilterTypeModule(True)
175:    End Select
176:    Application.ScreenUpdating = True
177: End Sub
     Private Function SelectedListItem(ByRef objList As MSForms.ListBox) As String
179:    Dim i           As Integer
180:    With objList
181:        For i = 0 To .ListCount - 1
182:            If .Selected(i) = True Then
183:                SelectedListItem = .List(i)
184:                Exit Function
185:            End If
186:        Next i
187:    End With
188: End Function
     Private Sub FilterAllOrEmpty(ByVal bFlag As Boolean)
190:    Dim i           As Integer
191:    With ListCode
192:        For i = 0 To .ListCount - 1
193:            .Selected(i) = bFlag
194:        Next i
195:        CheckAll.Value = bFlag
196:    End With
197: End Sub
     Private Sub FilterTypeModule(ByVal bFlag As Boolean)
199:    Dim i           As Integer
200:    Dim strVal      As String
201:    With ListFilter2
202:        For i = 0 To .ListCount - 1
203:            If .Selected(i) Then
204:                strVal = strVal & .List(i)
205:            End If
206:        Next i
207:    End With
208:    With ListCode
209:        For i = 0 To .ListCount - 1
210:            If strVal = vbNullString Then
211:                If .List(i, 3) = "empty" Then
212:                    .Selected(i) = Not bFlag
213:                Else
214:                    .Selected(i) = bFlag
215:                End If
216:            Else
217:                If strVal Like "*" & .List(i, 1) & "*" Then
218:                    .Selected(i) = True
219:                Else
220:                    .Selected(i) = False
221:                End If
222:                If .List(i, 3) = "empty" And .Selected(i) = True Then
223:                    .Selected(i) = Not bFlag
224:                ElseIf .Selected(i) = True Then
225:                    .Selected(i) = bFlag
226:                End If
227:            End If
228:        Next i
229:    End With
230: End Sub

     Private Sub CheckNotEmpty_Click()
233:    Dim i           As Integer
234:    With ListCode
235:        For i = 0 To .ListCount - 1
236:            If .List(i, 3) = "empty" Then
237:                .Selected(i) = True
238:            End If
239:        Next i
240:    End With
241: End Sub
     Private Sub cmbMain_Change()
243:    If cmbMain.Value <> vbNullString Then
244:        Call AddListCode
245:        Call FilterRun
246:    End If
247: End Sub
     Private Sub ListCode_Change()
249:    Dim i           As Integer
250:    Dim Flag        As Boolean
251:    With ListCode
252:        For i = 0 To .ListCount - 1
253:            If .Selected(i) Then
254:                lbMsg.visible = False
255:                Exit Sub
256:            End If
257:        Next i
258:    End With
259:    lbMsg.visible = True
260: End Sub
     Private Sub AddListCode()
262:    Application.ScreenUpdating = False
263:    Dim wb          As Workbook
264:    Dim iFile       As Integer
265:    Dim Arr()       As Variant
266:    Dim sLineCount  As String
267:    On Error GoTo ErrorHandler
268:    Set wb = Workbooks(cmbMain.Value)
269:    If wb.VBProject.Protection = vbext_pp_none Then
270:        With ListCode
271:            .Clear
272:            For iFile = 1 To wb.VBProject.VBComponents.Count
273:                .AddItem iFile
274:                .List(iFile - 1, 1) = ComponentTypeToString(wb.VBProject.VBComponents(iFile).Type)
275:                .List(iFile - 1, 2) = wb.VBProject.VBComponents(iFile).Name
276:                sLineCount = ModuleLineCount(wb.VBProject.VBComponents(iFile))
277:                If sLineCount = 0 Then sLineCount = "empty"
278:                .List(iFile - 1, 3) = sLineCount
279:            Next iFile
280:            Arr = .List
281:            Call Sort2_asc(Arr, 1)
282:            .List = Arr
283:            For iFile = 0 To .ListCount - 1
284:                .List(iFile, 0) = iFile + 1
285:            Next iFile
286:        End With
287:    Else
288:        ListCode.Clear
289:        Call MsgBox("VBA project in the book -" & wb.Name & "password protected!" & vbCrLf & "Remove the password!", vbCritical, "Mistake:")
290:    End If
291:    Application.ScreenUpdating = True
292:    Exit Sub
ErrorHandler:
294:    Select Case Err.Number
        Case 4160:
296:            ListCode.Clear
297:            Call MsgBox("Mistake! There is no access to the VBA project!", vbOKOnly + vbExclamation, "Mistake:")
298:        Case Else:
299:            Call MsgBox("Mistake! in AddListCode.AddListCode" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbExclamation, "Mistake:")
300:            Call WriteErrorLog("AddListCode.AddListCode")
301:    End Select
302:    Err.Clear
303:    Application.ScreenUpdating = True
304: End Sub
     Private Function SelectedListItems() As String()
306:    Dim i           As Integer
307:    Dim j           As Integer
308:    Dim ArrList()   As String
309:    With ListCode
310:        j = 0
311:        For i = 0 To .ListCount - 1
312:            If .Selected(i) Then
313:                ReDim Preserve ArrList(0 To j)
314:                ArrList(j) = .List(i, 2)
315:                j = j + 1
316:            End If
317:        Next i
318:    End With
319:    SelectedListItems = ArrList
320: End Function
'сортировка массива
     Private Sub Sort2_asc(Arr(), col As Long)
323:    Dim arrTemp()   As Variant
324:    Dim lb2 As Long, ub2 As Long, lTop As Long, bot As Long
325:
326:    lTop = LBound(Arr, 1)
327:    bot = UBound(Arr, 1)
328:    lb2 = LBound(Arr, 2)
329:    ub2 = UBound(Arr, 2)
330:    ReDim arrTemp(lb2 To ub2)
331:
332:    Call QSort2_asc(Arr(), col, lTop, bot, arrTemp(), lb2, ub2)
333: End Sub
Private Sub QSort2_asc(Arr(), C As Long, ByVal lTop As Long, ByVal bot As Long, temp(), lb2 As Long, ub2 As Long)
335:    Dim t As Long, LB As Long, MidItem, j As Long
336:    MidItem = Arr((lTop + bot) \ 2, C)
337:    t = lTop: LB = bot
338:    Do
339:        Do While Arr(t, C) < MidItem: t = t + 1: Loop
340:        Do While Arr(LB, C) > MidItem: LB = LB - 1: Loop
341:        If t < LB Then
342:            For j = lb2 To ub2: temp(j) = Arr(t, j): Next j
343:            For j = lb2 To ub2: Arr(t, j) = Arr(LB, j): Next j
344:            For j = lb2 To ub2: Arr(LB, j) = temp(j): Next j
345:            t = t + 1: LB = LB - 1
346:        ElseIf t = LB Then
347:            t = t + 1: LB = LB - 1
348:        End If
349:    Loop While t <= LB
350:
351:    If t < bot Then QSort2_asc Arr(), C, t, bot, temp(), lb2, ub2
352:    If lTop < LB Then QSort2_asc Arr(), C, lTop, LB, temp(), lb2, ub2
End Sub
