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
11:    Me.StartUpPosition = 0
12:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
13:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
14:
15:    Set m_clsAnchors = New CAnchors
16:    Set m_clsAnchors.objParent = Me
17:    ' restrict minimum size of userform
18:    m_clsAnchors.MinimumWidth = 546
19:    m_clsAnchors.MinimumHeight = 441
20:    With m_clsAnchors
21:        .funAnchor("cmbMain").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft
22:        .funAnchor("ListCode").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft Or enumAnchorStyleBottom
23:        .funAnchor("lbRemoveModule").AnchorStyle = enumAnchorStyleBottom
24:        .funAnchor("lbExportModule").AnchorStyle = enumAnchorStyleBottom
25:        .funAnchor("lbImportModule").AnchorStyle = enumAnchorStyleBottom
26:        .funAnchor("lbCopytModule").AnchorStyle = enumAnchorStyleBottom
27:        .funAnchor("lbCancel").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
28:        .funAnchor("Label2").AnchorStyle = enumAnchorStyleRight
29:        .funAnchor("ListFilter1").AnchorStyle = enumAnchorStyleRight
30:        .funAnchor("CheckAll").AnchorStyle = enumAnchorStyleRight
31:        .funAnchor("ListFilter2").AnchorStyle = enumAnchorStyleRight
32:        .funAnchor("lbMsg").AnchorStyle = enumAnchorStyleRight
33:        .funAnchor("Label3").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
34:        .funAnchor("cmbMainCopy").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
35:    End With
36:    With ListFilter1
37:        .AddItem "All"
38:        .AddItem "Empty ones"
39:        .AddItem "Not empty"
40:        .AddItem "Reset all"
41:    End With
42:    With ListFilter2
43:        .AddItem "Code Module"
44:        .AddItem "UserForm"
45:        .AddItem "Document Module"
46:        .AddItem "Class Module"
47:        .AddItem "ActiveX Designer"
48:    End With
49:    ListFilter1.Selected(0) = True
50: End Sub
    Private Sub UserForm_Activate()
52:    Dim vbProj      As VBIDE.VBProject
53:    If Workbooks.Count = 0 Then
54:        Unload Me
55:        Call MsgBox("No open ones" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
56:        Exit Sub
57:    End If
58:    With Me.cmbMain
59:        .Clear
60:        cmbMainCopy.Clear
61:        On Error Resume Next
62:        For Each vbProj In Application.VBE.VBProjects
63:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
64:        Next
65:        .Value = ActiveWorkbook.Name
66:        On Error GoTo 0
67:        On Error GoTo ErrorHandler
68:        cmbMainCopy.Value = .List(0)
69:    End With
70:    lbMsg.visible = True
71:    Call CheckAll_Change
72:    Exit Sub
ErrorHandler:
74:    Unload Me
75:    Select Case Err.Number
        Case Else:
77:            Call MsgBox("Error in Module Commander.UserForm_Activate" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbExclamation, "Error:")
78:            Call WriteErrorLog("ModuleCommander.UserForm_Activate")
79:    End Select
80:    Err.Clear
81: End Sub
    Private Sub UserForm_Terminate()
83:    Set m_clsAnchors = Nothing
84: End Sub
    Private Sub lbImportModule_Click()
86:    Application.ScreenUpdating = False
87:    If cmbMain.Value <> vbNullString Then
88:        Call S_ModuleCommander.ImportAllModules(Workbooks(cmbMain.Value))
89:    End If
90:    Call AddListCode
91:    Call FilterRun
92:    Application.ScreenUpdating = True
93: End Sub
     Private Sub lbExportModule_Click()
95:    Application.ScreenUpdating = False
96:    If cmbMain.Value <> vbNullString And lbMsg.visible = False Then
97:        Call S_ModuleCommander.ExportAllModules(Workbooks(cmbMain.Value), SelectedListItems)
98:    Else
99:        Call MsgBox("Nothing is selected!", vbInformation, "Exporting a project:")
100:    End If
101:    Call AddListCode
102:    Call FilterRun
103:    Application.ScreenUpdating = True
104: End Sub
     Private Sub lbRemoveModule_Click()
106:    Application.ScreenUpdating = False
107:    If cmbMain.Value <> vbNullString And lbMsg.visible = False Then
108:        Call S_ModuleCommander.DeleteAllModulesInActiveProject(Workbooks(cmbMain.Value), SelectedListItems)
109:    Else
110:        Call MsgBox("Nothing is selected!", vbInformation, "Deleting a project:")
111:    End If
112:    Call AddListCode
113:    Call FilterRun
114:    Application.ScreenUpdating = True
115: End Sub
     Private Sub lbCopytModule_Click()
117:    If MsgBox("Copy VBA modules from [" & cmbMain.Value & "] to [" & cmbMainCopy.Value & "] ?", vbYesNo + vbQuestion, "Copying VBA modules:") = vbNo Then Exit Sub
118:    Application.ScreenUpdating = False
119:    If cmbMain.Value <> vbNullString And cmbMainCopy.Value <> vbNullString And lbMsg.visible = False Then
120:        Dim i       As Integer
121:        Dim vbCompObj As VBIDE.VBComponent
122:        With ListCode
123:            For i = 0 To .ListCount - 1
124:                If .Selected(i) Then
125:                    Set vbCompObj = Workbooks(cmbMain.Value).VBProject.VBComponents(.List(i, 2))
126:                    If vbCompObj.CodeModule.CountOfLines <> 0 Then
127:                        Call E_AddEnum.AddModuleToProject(.List(i, 2), vbCompObj.Type, vbCompObj.CodeModule.Lines(1, vbCompObj.CodeModule.CountOfLines), Workbooks(cmbMainCopy.Value).VBProject)
128:                    Else
129:                        Call E_AddEnum.AddModuleToProject(.List(i, 2), vbCompObj.Type, vbNullString, Workbooks(cmbMainCopy.Value).VBProject)
130:                    End If
131:                End If
132:            Next i
133:        End With
134:        Call MsgBox("Copying VBA modules from [" & cmbMain.Value & "] to [" & cmbMainCopy.Value & "] completed", vbOKOnly + vbInformation, "Copying VBA modules:")
135:    Else
136:        Call MsgBox("Nothing is selected!", vbInformation, "Copy Project:")
137:    End If
138:    Call AddListCode
139:    Call FilterRun
140:    Application.ScreenUpdating = True
141: End Sub
     Private Sub cmbCancel_Click()
143:    Unload Me
144: End Sub
     Private Sub lbCancel_Click()
146:    Call cmbCancel_Click
147: End Sub
     Private Sub ListFilter2_Change()
149:    Call FilterRun
150: End Sub
     Private Sub ListFilter1_Change()
152:    Call FilterRun
153: End Sub
     Private Sub CheckAll_Change()
155:    Dim i           As Integer
156:    With ListFilter2
157:        For i = 0 To .ListCount - 1
158:            .Selected(i) = CheckAll.Value
159:        Next i
160:    End With
161: End Sub
     Private Sub FilterRun()
163:    Application.ScreenUpdating = False
164:    Select Case SelectedListItem(ListFilter1)
        Case "All":
166:            Call FilterAllOrEmpty(True)
167:        Case "Reset all":
168:            Call FilterAllOrEmpty(False)
169:        Case "Empty ones":
170:            Call FilterTypeModule(False)
171:        Case "Not empty":
172:            Call FilterTypeModule(True)
173:    End Select
174:    Application.ScreenUpdating = True
175: End Sub
     Private Function SelectedListItem(ByRef objList As MSForms.ListBox) As String
177:    Dim i           As Integer
178:    With objList
179:        For i = 0 To .ListCount - 1
180:            If .Selected(i) = True Then
181:                SelectedListItem = .List(i)
182:                Exit Function
183:            End If
184:        Next i
185:    End With
186: End Function
     Private Sub FilterAllOrEmpty(ByVal bFlag As Boolean)
188:    Dim i           As Integer
189:    With ListCode
190:        For i = 0 To .ListCount - 1
191:            .Selected(i) = bFlag
192:        Next i
193:        CheckAll.Value = bFlag
194:    End With
195: End Sub
     Private Sub FilterTypeModule(ByVal bFlag As Boolean)
197:    Dim i           As Integer
198:    Dim strVal      As String
199:    With ListFilter2
200:        For i = 0 To .ListCount - 1
201:            If .Selected(i) Then
202:                strVal = strVal & .List(i)
203:            End If
204:        Next i
205:    End With
206:    With ListCode
207:        For i = 0 To .ListCount - 1
208:            If strVal = vbNullString Then
209:                If .List(i, 3) = "empty empty" Then
210:                    .Selected(i) = Not bFlag
211:                Else
212:                    .Selected(i) = bFlag
213:                End If
214:            Else
215:                If strVal Like "*" & .List(i, 1) & "*" Then
216:                    .Selected(i) = True
217:                Else
218:                    .Selected(i) = False
219:                End If
220:                If .List(i, 3) = "empty empty" And .Selected(i) = True Then
221:                    .Selected(i) = Not bFlag
222:                ElseIf .Selected(i) = True Then
223:                    .Selected(i) = bFlag
224:                End If
225:            End If
226:        Next i
227:    End With
228: End Sub

     Private Sub CheckNotEmpty_Click()
231:    Dim i           As Integer
232:    With ListCode
233:        For i = 0 To .ListCount - 1
234:            If .List(i, 3) = "empty empty" Then
235:                .Selected(i) = True
236:            End If
237:        Next i
238:    End With
239: End Sub
     Private Sub cmbMain_Change()
241:    If cmbMain.Value <> vbNullString Then
242:        Call AddListCode
243:        Call FilterRun
244:    End If
245: End Sub
     Private Sub ListCode_Change()
247:    Dim i           As Integer
248:    Dim Flag        As Boolean
249:    With ListCode
250:        For i = 0 To .ListCount - 1
251:            If .Selected(i) Then
252:                lbMsg.visible = False
253:                Exit Sub
254:            End If
255:        Next i
256:    End With
257:    lbMsg.visible = True
258: End Sub
     Private Sub AddListCode()
260:    Application.ScreenUpdating = False
261:    Dim WB          As Workbook
262:    Dim iFile       As Integer
263:    Dim Arr()       As Variant
264:    Dim sLineCount  As String
265:    On Error GoTo ErrorHandler
266:    Set WB = Workbooks(cmbMain.Value)
267:    If WB.VBProject.Protection = vbext_pp_none Then
268:        With ListCode
269:            .Clear
270:            For iFile = 1 To WB.VBProject.VBComponents.Count
271:                .AddItem iFile
272:                .List(iFile - 1, 1) = ComponentTypeToString(WB.VBProject.VBComponents(iFile).Type)
273:                .List(iFile - 1, 2) = WB.VBProject.VBComponents(iFile).Name
274:                sLineCount = ModuleLineCount(WB.VBProject.VBComponents(iFile))
275:                If sLineCount = 0 Then sLineCount = "empty empty"
276:                .List(iFile - 1, 3) = sLineCount
277:            Next iFile
278:            Arr = .List
279:            Call Sort2_asc(Arr, 1)
280:            .List = Arr
281:            For iFile = 0 To .ListCount - 1
282:                .List(iFile, 0) = iFile + 1
283:            Next iFile
284:        End With
285:    Else
286:        ListCode.Clear
287:        Call MsgBox("VBA project in the book -" & WB.Name & "password protected!" & vbCrLf & "Remove the password!", vbCritical, "Error:")
288:    End If
289:    Application.ScreenUpdating = True
290:    Exit Sub
ErrorHandler:
292:    Select Case Err.Number
        Case 4160:
294:            ListCode.Clear
295:            Call MsgBox("Error No access to the VBA project!", vbOKOnly + vbExclamation, "Error:")
296:        Case Else:
297:            Call MsgBox("Error in the Add List Code.AddListCode" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbExclamation, "Error:")
298:            Call WriteErrorLog("AddListCode.AddListCode")
299:    End Select
300:    Err.Clear
301:    Application.ScreenUpdating = True
302: End Sub
     Private Function SelectedListItems() As String()
304:    Dim i           As Integer
305:    Dim j           As Integer
306:    Dim ArrList()   As String
307:    With ListCode
308:        j = 0
309:        For i = 0 To .ListCount - 1
310:            If .Selected(i) Then
311:                ReDim Preserve ArrList(0 To j)
312:                ArrList(j) = .List(i, 2)
313:                j = j + 1
314:            End If
315:        Next i
316:    End With
317:    SelectedListItems = ArrList
318: End Function
'сортировка массива
     Private Sub Sort2_asc(Arr(), col As Long)
321:    Dim arrTemp()   As Variant
322:    Dim lb2 As Long, ub2 As Long, lTop As Long, bot As Long
323:
324:    lTop = LBound(Arr, 1)
325:    bot = UBound(Arr, 1)
326:    lb2 = LBound(Arr, 2)
327:    ub2 = UBound(Arr, 2)
328:    ReDim arrTemp(lb2 To ub2)
329:
330:    Call QSort2_asc(Arr(), col, lTop, bot, arrTemp(), lb2, ub2)
331: End Sub
Private Sub QSort2_asc(Arr(), C As Long, ByVal lTop As Long, ByVal bot As Long, temp(), lb2 As Long, ub2 As Long)
333:    Dim t As Long, LB As Long, MidItem, j As Long
334:    MidItem = Arr((lTop + bot) \ 2, C)
335:    t = lTop: LB = bot
336:    Do
337:        Do While Arr(t, C) < MidItem: t = t + 1: Loop
338:        Do While Arr(LB, C) > MidItem: LB = LB - 1: Loop
339:        If t < LB Then
340:            For j = lb2 To ub2: temp(j) = Arr(t, j): Next j
341:            For j = lb2 To ub2: Arr(t, j) = Arr(LB, j): Next j
342:            For j = lb2 To ub2: Arr(LB, j) = temp(j): Next j
343:            t = t + 1: LB = LB - 1
344:        ElseIf t = LB Then
345:            t = t + 1: LB = LB - 1
346:        End If
347:    Loop While t <= LB
348:
349:    If t < bot Then QSort2_asc Arr(), C, t, bot, temp(), lb2, ub2
350:    If lTop < LB Then QSort2_asc Arr(), C, lTop, LB, temp(), lb2, ub2
End Sub
