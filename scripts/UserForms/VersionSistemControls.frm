VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VersionSistemControls 
   Caption         =   "Version Control:"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15690
   OleObjectBlob   =   "VersionSistemControls.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VersionSistemControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : VersionSistemControls - система контрол€ версий файлов
'* Created    : 10-03-2020 09:30
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private m_clsAnchors As CAnchors

'* глобальна€ перемена€ содержаща€ информацию из файла Config.cvs
Private sVersionInfo As String

Private Const myRedColor As Long = &HC0&
Private Const myGreenColor As Long = &HC000&
Private Const CONFIG As String = "Config.cvs"

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : VersionList; GetTypeCoomment; ParserCommet - парсер Config.cvs нужно обновл€ть две функции и перечисление
'* Created    : 10-03-2020 09:30
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal iItem As VersionList : перечисление
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Enum VersionList
    verNameFileVer = 0
    verVersion = 1
    verDateAdd = 2
    verOldVersion = 3
    verModuleNames = 4
    verComment = 5
End Enum

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetTypeCoomment - получение названи€ по перечислению
'* Created    : 10-03-2020 09:33
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal iItem As VersionList :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Function GetTypeCoomment(ByVal iItem As VersionList) As String
51:    Select Case iItem
        Case 0: GetTypeCoomment = "NameFile:"
53:        Case 1: GetTypeCoomment = "Version:"
54:        Case 2: GetTypeCoomment = "DateAdd:"
55:        Case 3: GetTypeCoomment = "OldVersion:"
56:        Case 4: GetTypeCoomment = "ModuleNames:"
57:        Case 5: GetTypeCoomment = "Comment:"
58:    End Select
59: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ParserCommet - парсер Config.cvs
'* Created    : 10-03-2020 09:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                     Description
'*
'* ByVal lSelectedVersion As Long : номер версии
'* ByVal iItem As VersionList     : номер искомого из перечислени€
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Function ParserCommet(ByVal lSelectedVersion As Long, ByVal iItem As VersionList) As String
74:    Dim sSTR   As String
75:    sSTR = Split(sVersionInfo, vbNewLine)(lSelectedVersion)
76:    sSTR = Split(sSTR, ";")(iItem)
77:    sSTR = VBA.Replace(sSTR, GetTypeCoomment(iItem), vbNullString)
78:    sSTR = VBA.Replace(sSTR, " vbNewLine ", vbNewLine)
79:    ParserCommet = sSTR
80: End Function
    Private Sub btnCancel_Click()
82:    Unload Me
83: End Sub
    Private Sub lbCancel_Click()
85:    Call btnCancel_Click
86: End Sub
     Private Sub lbAddSource_Click()
88:    Dim sPath  As String
89:    Dim sBaseNameFile As String
90:    Dim sExtensionFile As String
91:    Dim sNewDirForFile As String
92:    Dim sCongig As String
93:    Dim sDate  As String
94:    Dim sVersion As String
95:    Dim sCommentariy As String
96:    Dim sOldVer As String
97:
98:    On Error GoTo ErrorHandler
99:
100:    If Me.txtCommentariyByfer.Text <> vbNullString Then
101:        Me.txtCommentariy.Text = vbNullString
102:        Me.txtCommentariyByfer.Text = vbNullString
103:    End If
104:
105:    If Me.txtCommentariy.Text = vbNullString Then
106:        Me.txtMsg.Text = "Add a comment to the version before creating it"
107:        Me.txtMsg.ForeColor = myRedColor
108:        Exit Sub
109:    End If
110:
111:    sBaseNameFile = Me.txtBaseNameFile.Text
112:    sExtensionFile = Me.txtExtensionFile.Text
113:    Workbooks(sBaseNameFile & sExtensionFile).Save
114:
115:    sPath = Me.txtPath.Text
116:
117:    'если нет директории то создаем
118:    If Not C_PublicFunctions.FileHave(sPath, vbDirectory) Then
119:        MkDir (sPath)
120:    End If
121:
122:    sDate = VBA.Replace(VBA.Replace(VBA.Now(), ":", "."), " ", "_")
123:    sVersion = "v" & Me.ListVersion.ListCount + 1
124:    sNewDirForFile = sBaseNameFile & "_" & sDate & "_" & sVersion & sExtensionFile
125:    sOldVer = GetPathFormCode(GetCodeFromModule(Me.cmbFile.Value), 2, "Version    : ")
126:    'создание версии файла
127:    Call C_PublicFunctions.CopyFileFSO(Workbooks(Me.cmbFile.Value).FullName, sPath & sNewDirForFile)
128:    'формировани€е комментари€
129:    sCommentariy = VBA.Replace(VBA.Trim(Me.txtCommentariy.Text), vbNewLine, " vbNewLine ")
130:    sCongig = GetTypeCoomment(VersionList.verNameFileVer) & sNewDirForFile & ";" & _
                                GetTypeCoomment(VersionList.verVersion) & sVersion & ";" & _
                                GetTypeCoomment(VersionList.verDateAdd) & sDate & ";" & _
                                GetTypeCoomment(VersionList.verOldVersion) & sOldVer & ";" & _
                                GetTypeCoomment(VersionList.verModuleNames) & AddListModuleName(Me.cmbFile.Value) & ";" & _
                                GetTypeCoomment(VersionList.verComment) & sCommentariy & vbNewLine
136:    'обновление или создание файла Config.cvs
137:    Call C_PublicFunctions.TXTAddIntoTXTFile(sPath & CONFIG, sCongig)
138:    sVersionInfo = C_PublicFunctions.TXTReadALLFile(sPath & Application.PathSeparator & CONFIG, False)
139:
140:    'обновление или создание комментари€
141:    Call AddCommentVSC(AddCommentForModule(sPath, sVersion, sDate, sOldVer), Me.cmbFile.Value)
142:
143:    Call CheckSourcePath
144:    'Me.ListVersion.Selected(Me.ListVersion.ListCount - 1) = True
145:    Me.txtMsg.Text = "File version added to storage:" & vbNewLine & "[ " & sNewDirForFile & " ]"
146:    Me.txtMsg.ForeColor = myGreenColor
147:    Me.txtCommentariy.Text = vbNullString
148:
149:    Exit Sub
ErrorHandler:
151:    Debug.Print "Error in VersionSistemControls. lbAddSource_Click" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
152:    Call WriteErrorLog("VersionSistemControls.lbAddSource_Click")
153: End Sub
     Private Sub lbChoseFile_Click()
155:    Dim sCommentariy As String
156:
157:    Me.txtPath.Text = C_PublicFunctions.DirLoadFiles(Workbooks(Me.cmbFile.Value).Path)
158:    If C_PublicFunctions.FileHave(Me.txtPath.Text, vbDirectory) Then
159:        Me.lbAddSource.Enabled = True
160:    End If
161:    sVersionInfo = C_PublicFunctions.TXTReadALLFile(Me.txtPath.Text & CONFIG, False)
162:        Call ChangeColor
163: End Sub
     Private Sub lbOpenFileVersion_Click()
165:    Dim i      As Long
166:    Dim sFileName As String
167:
168:    On Error GoTo ErrorHandler
169:
170:    i = GetNomerItemSelectedList()
171:
172:    If i = -1 Then
173:        Me.txtMsg.Text = "No file version selected to open!"
174:        Me.txtMsg.ForeColor = myRedColor
175:        Exit Sub
176:    End If
177:
178:    sFileName = Me.txtPath & Me.ListVersion.List(i, 1)
179:    If C_PublicFunctions.FileHave(sFileName, Normal) Then
180:        Workbooks.Open Filename:=Me.txtPath & Me.ListVersion.List(i, 1)
181:        Me.txtMsg.Text = "File: [" & Me.ListVersion.List(i, 1) & "] open!"
182:        Me.txtMsg.ForeColor = myGreenColor
183:    Else
184:        Me.txtMsg.Text = "File: [" & Me.ListVersion.List(i, 1) & "] not found, in the storage!"
185:        Me.txtMsg.ForeColor = myRedColor
186:    End If
187:
188:    Exit Sub
ErrorHandler:
190:    Debug.Print "Error in Version SystemControls.lbOpenFileVersion_Click" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
191:    Call WriteErrorLog("VersionSistemControls.lbOpenFileVersion_Click")
192: End Sub
     Private Sub lbLoadFileVersion_Click()
194:    Dim i      As Long
195:    Dim sFileName As String
196:    Dim sMainPath As String
197:    Dim sLoadPath As String
198:    Dim SelectFileName As String
199:
200:    On Error GoTo ErrorHandler
201:
202:    i = GetNomerItemSelectedList()
203:
204:    If i = -1 Then
205:        Me.txtMsg.Text = "No file version selected to download!"
206:        Me.txtMsg.ForeColor = myRedColor
207:        Exit Sub
208:    End If
209:
210:    SelectFileName = Me.ListVersion.List(i, 1)
211:
212:    If Not C_PublicFunctions.FileHave(Me.txtPath & SelectFileName, Normal) Then
213:
214:        Me.txtMsg.Text = "File: [" & SelectFileName & "] not found, in the storage!"
215:        Me.txtMsg.ForeColor = myRedColor
216:
217:    ElseIf MsgBox("Upload a file: [" & SelectFileName & " ]" & vbNewLine & vbNewLine & _
                                "Attention, the current file [" & Workbooks(Me.cmbFile.Value).Name & "] will be overwritten!", _
                                vbYesNo + vbQuestion, "Download version:") = vbYes Then
220:        sMainPath = Workbooks(Me.cmbFile.Value).FullName
221:        sLoadPath = Me.txtPath & SelectFileName
222:
223:        Application.DisplayAlerts = False
224:        Workbooks(Me.cmbFile.Value).Close
225:        Application.DisplayAlerts = True
226:        Call C_PublicFunctions.CopyFileFSO(sLoadPath, sMainPath)
227:        Workbooks.Open Filename:=sMainPath
228:
229:        Me.txtMsg.Text = "Uploading a file: [" & SelectFileName & " ]" & vbNewLine & "Completed"
230:        Me.txtMsg.ForeColor = myGreenColor
231:
232:    End If
233:
234:    Exit Sub
ErrorHandler:
236:    Debug.Print "Error in VersionSistemControls.lbLoadFileVersion_Click" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
237:    Call WriteErrorLog("VersionSistemControls.lbLoadFileVersion_Click")
238: End Sub
     Private Sub ListVersion_Click()
240:    Dim i      As Long
241:
242:    On Error GoTo ErrorHandler
243:
244:    i = GetNomerItemSelectedList()
245:    If i = -1 Then Exit Sub
246:
247:    Me.txtCommentariy.Text = "Previous version:" & ParserCommet(i, VersionList.verOldVersion) & vbNewLine & "Comment:" & vbNewLine & ParserCommet(i, VersionList.verComment)
248:    Me.txtCommentariyByfer.Text = Me.txtCommentariy.Text
249:
250:    Exit Sub
ErrorHandler:
252:    Debug.Print "Error d Version SystemControls.ListVersion_Click" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
253:    Call WriteErrorLog("VersionSistemControls.ListVersion_Click")
254: End Sub
     Private Sub cmbFile_Change()
256:    Call CheckSourcePath
257: End Sub

     Private Sub txtPath_Change()
260:    Call ChangeColor
261: End Sub

     Private Sub ChangeColor()
264:    If txtPath.Text = vbNullString Then
265:        txtPath.BorderColor = myRedColor
266:        lbChoseFile.BorderColor = myRedColor
267:        lbChoseFile.ForeColor = myRedColor
268:    Else
269:        txtPath.BorderColor = &H8000000D
270:        lbChoseFile.BorderColor = &H8000000D
271:        lbChoseFile.ForeColor = &H8000000D
272:    End If
273: End Sub

     Private Sub UserForm_Initialize()
276:    Me.StartUpPosition = 0
277:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
278:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
279:
280:
281:    On Error GoTo ErrorHandler
282:
283:    Set m_clsAnchors = New CAnchors
284:    Set m_clsAnchors.objParent = Me
285:    ' restrict minimum size of userform
286:    m_clsAnchors.MinimumWidth = 789
287:    m_clsAnchors.MinimumHeight = 429
288:    With m_clsAnchors
289:        .funAnchor("cmbFile").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft
290:        .funAnchor("txtPath").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft
291:        .funAnchor("lbChoseFile").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop
292:        .funAnchor("ListVersion").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft Or enumAnchorStyleBottom
293:        .funAnchor("txtMsg").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleRight
294:        .funAnchor("txtCommentariy").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleRight Or enumAnchorStyleBottom
295:        .funAnchor("lbLoadFileVersion").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
296:        .funAnchor("lbOpenFileVersion").AnchorStyle = enumAnchorStyleBottom
297:        .funAnchor("lbAddSource").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
298:        .funAnchor("lbCancel").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
299:    End With
300:
301:
302:    Dim vbProj      As VBIDE.VBProject
303:    With Me.cmbFile
304:        On Error Resume Next
305:        For Each vbProj In Application.VBE.VBProjects
306:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
307:        Next
308:        On Error GoTo 0
309:        On Error GoTo ErrorHandler
310:        .Value = ActiveWorkbook.Name
311:    End With
312:    Call ChangeColor
313:    Exit Sub
ErrorHandler:
315:    Debug.Print "Error in VersionSistemControls.UserForm_Initialize" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
316:    Call WriteErrorLog("VersionSistemControls.UserForm_Initialize")
317: End Sub
     Private Sub UserForm_Terminate()
319:    Set m_clsAnchors = Nothing
320: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CheckSourcePath - главна€ процедура
'* Created    : 06-03-2020 10:18
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub CheckSourcePath()
329:    Dim sPath  As String
330:    Dim sMsg   As String
331:
332:    On Error GoTo ErrorHandler
333:
334:    sPath = GetPathFormCode(GetCodeFromModule(Me.cmbFile.Value), 1, "Path       : ")
335:
336:    If sPath = vbNullString Then
337:        Me.lbAddSource.Enabled = False
338:    Else
339:        Me.lbAddSource.Enabled = True
340:    End If
341:
342:    Me.txtBaseNameFile.Text = C_PublicFunctions.sGetBaseName(Me.cmbFile.Value)
343:    Me.txtExtensionFile.Text = "." & C_PublicFunctions.sGetExtensionName(Me.cmbFile.Value)
344:    Me.txtMsg.ForeColor = myRedColor
345:    Me.ListVersion.Clear
346:    Me.txtPath.Text = vbNullString
347:
348:    If C_PublicFunctions.FileHave(sPath, vbDirectory) Then
349:        Me.txtPath.Text = sPath
350:    Else
351:        sMsg = "No storage created for the file:" & vbNewLine & "[ " & Me.cmbFile.Value & " ]"
352:        Me.txtCommentariy.Text = "Creating the first version"
353:        Me.txtMsg.Text = sMsg
354:        Exit Sub
355:    End If
356:
357:    If Not C_PublicFunctions.FileHave(sPath & CONFIG, vbDirectory) Then
358:        sMsg = "File not found: Config.sys"
359:        Me.txtCommentariy.Text = "Creating the first version"
360:    End If
361:
362:    Me.txtMsg.Text = sMsg
363:    sVersionInfo = C_PublicFunctions.TXTReadALLFile(sPath & Application.PathSeparator & CONFIG, False)
364:    Call RefrashListVersion(sPath)
365:
366:    Exit Sub
ErrorHandler:
368:    Debug.Print "Error in Version SystemControls.CheckSourcePath" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
369:    Call WriteErrorLog("VersionSistemControls.CheckSourcePath")
370: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RefrashListVersion обновление ListBox, названи€ файлов версий
'* Created    : 06-03-2020 10:17
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sPATH As String : директори€ хранилища
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub RefrashListVersion(ByVal sPath As String)
383:    Dim sConfigString As String
384:    Dim vVar   As Variant
385:    Dim i      As Long
386:    Dim j      As Long
387:
388:    On Error GoTo ErrorHandler
389:
390:    sConfigString = C_PublicFunctions.TXTReadALLFile(sPath & Application.PathSeparator & CONFIG, False)
391:    If sConfigString = vbNullString Then Exit Sub
392:    vVar = VBA.Split(sConfigString, vbNewLine)
393:    For i = 0 To UBound(vVar)
394:        If vVar(i) <> vbNullString Then
395:            With Me.ListVersion
396:                .AddItem j + 1
397:                .List(j, 1) = ParserCommet(i, VersionList.verNameFileVer)
398:                j = j + 1
399:            End With
400:        End If
401:    Next i
402:
403:    Exit Sub
ErrorHandler:
405:    Debug.Print "Error in VersionSistemControls.RefrashListVersion" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
406:    Call WriteErrorLog("VersionSistemControls.RefrashListVersion")
407: End Sub

'* * * * * * * * * * * * * * * * * * * * * * FUNCTION's * * * * * * * * * * * * * * * * * * * * * *

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetNomerItemSelectedList - номер выделеного значени€ в листбоксе
'* Created    : 10-03-2020 09:27
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function GetNomerItemSelectedList() As Long
419:    Dim i      As Long
420:
421:    On Error GoTo ErrorHandler
422:
423:    For i = 0 To Me.ListVersion.ListCount
424:        If Me.ListVersion.Selected(i) = True Then
425:            GetNomerItemSelectedList = i
426:            Exit Function
427:        End If
428:    Next i
429:    GetNomerItemSelectedList = -1
430:    Exit Function
ErrorHandler:
432:    GetNomerItemSelectedList = -1
433:    Debug.Print "Error in VersionSistemControls.Get NumberItemSelectedList" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
434:    Call WriteErrorLog("VersionSistemControls.GetNomerItemSelectedList")
435: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetCodeFromModule - строкова€ перемена€, возращающает текст файла модул€ "Ёта нига"
'* Created    : 06-03-2020 10:09
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function GetCodeFromModule(ByVal sWBName As String) As String
445:    Dim nLine  As Long
446:    Dim objModuleWB As VBIDE.CodeModule
447:
448:    On Error GoTo ErrorHandler
449:
450:    Set objModuleWB = Workbooks(sWBName).VBProject.VBComponents(1).CodeModule
451:    nLine = 0
452:    With objModuleWB
453:        GetCodeFromModule = .Lines(1, .CountOfLines)
454:    End With
455:    Set objModuleWB = Nothing
456:
457:    Exit Function
ErrorHandler:
459:    Err.Clear
460:    GetCodeFromModule = vbNullString
461:    Set objModuleWB = Nothing
462: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetPathFormCode - парсет код переменой, и выдает путь к файлу записанный в комментарий
'* Created    : 06-03-2020 10:12
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sCode As String : строкова€ перемена€, содержащи€ код модул€
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function GetPathFormCode(ByVal sCode As String, ByVal byItem As Byte, ByVal sReplace As String) As String
475:    Dim i      As Long
476:    Dim j      As Long
477:    Dim k      As Long
478:    Dim sStrSerch As String
479:    Dim sNewString As String
480:    Dim vTemp  As Variant
481:
482:    On Error GoTo errmsg
483:
484:    sStrSerch = VBA.Trim$(addString())
485:    k = VBA.Len(sStrSerch)
486:    i = VBA.InStr(1, sCode, sStrSerch)
487:    j = VBA.InStrRev(sCode, sStrSerch, -1)
488:    If i = 0 And j = 0 Then GoTo errmsg
489:    sNewString = VBA.Left$(sCode, j + k)
490:    sNewString = VBA.Right$(sNewString, VBA.Len(sNewString) - i + 1)
491:    sNewString = VBA.Replace(sNewString, sStrSerch, vbNullString)
492:    sNewString = VBA.Replace(sNewString, vbNewLine, vbNullString)
493:    vTemp = VBA.Split(sNewString, "'* ")
494:    sNewString = VBA.Replace(VBA.Trim$(vTemp(byItem)), sReplace, vbNullString)
495:
496:    GetPathFormCode = VBA.Trim$(sNewString)
497:    Exit Function
errmsg:
499:    Err.Clear
500:    GetPathFormCode = vbNullString
501: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddListModuleName - строкова€ перемена€, содержащи€ имена всех модулей выбранной книги
'* Created    : 06-03-2020 10:13
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function AddListModuleName(ByVal sWBName As String) As String
510:    Dim objVBComp As VBIDE.VBComponent
511:    Dim sTemp  As String
512:
513:    On Error GoTo ErrorHandler
514:
515:    For Each objVBComp In Workbooks(sWBName).VBProject.VBComponents
516:        sTemp = sTemp & objVBComp.Name & ","
517:    Next objVBComp
518:    sTemp = VBA.Left$(sTemp, VBA.Len(sTemp) - 1)
519:    AddListModuleName = sTemp
520:    Exit Function
ErrorHandler:
522:    Debug.Print "Error in VersionSistemControls.AddListModuleName" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
523:    Call WriteErrorLog("VersionSistemControls.AddListModuleName")
524: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddCommentForModule - создание комментари€
'* Created    : 06-03-2020 10:14
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sPATH As String : путь к файлу
'* sVersion As String    : верси€
'* sCreated As String    : дата создани€
'* sOldVersion As String : предыдуща€ верси€
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function AddCommentForModule(ByVal sPath As String, ByVal sVersion As String, ByVal sCreated As String, ByVal sOldVersion As String) As String
540:    Dim sComm  As String
541:    Dim sComm1 As String
542:    Const sChar As String = "'* "
543:
544:    On Error GoTo ErrorHandler
545:
546:    sComm1 = addString()
547:
548:    sComm = sComm1 & vbNewLine
549:    sComm = sComm & sChar & "Path       : " & sPath & vbNewLine
550:    sComm = sComm & sChar & "Version    : " & sVersion & vbNewLine
551:    sComm = sComm & sChar & "Created    : " & sCreated & vbNewLine
552:    sComm = sComm & sChar & "OldVersion : " & sOldVersion & vbNewLine
553:    sComm = sComm & sComm1
554:    AddCommentForModule = sComm
555:
556:    Exit Function
ErrorHandler:
558:    Debug.Print "Error in VersionSistemControls.Add CommentForModule" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
559:    Call WriteErrorLog("VersionSistemControls.AddCommentForModule")
560: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddCommentVSC - процедура внедрени€ комментраи€ в код модул€
'* Created    : 06-03-2020 10:11
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal sComment As String : строкова€ перемена€, помещаема€ в код
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub AddCommentVSC(ByVal sComment As String, ByVal sWBName As String)
573:    Dim sPatch1 As String
574:    Dim sPatch2 As String
575:    Dim sCode  As String
576:    Dim sStrSerch As String
577:    Dim i      As Long
578:    Dim j      As Long
579:    Dim objModuleWB As VBIDE.CodeModule
580:
581:    On Error GoTo ErrorHandler
582:
583:    Set objModuleWB = Workbooks(sWBName).VBProject.VBComponents(1).CodeModule
584:
585:    sStrSerch = VBA.Trim$(addString())
586:    With objModuleWB
587:        sCode = .Lines(1, .CountOfLines)
588:        .DeleteLines 1, .CountOfLines
589:        i = VBA.InStr(1, sCode, sStrSerch) - 1
590:        j = VBA.InStrRev(sCode, sStrSerch, -1)
591:        If i <= 0 And j = 0 Then
592:            sCode = sComment & vbNewLine & sCode
593:        Else
594:            sPatch1 = VBA.Left$(sCode, i)
595:            sPatch2 = VBA.Right(sCode, VBA.Len(sCode) - j - VBA.Len(sStrSerch))
596:            sCode = sPatch1 & sComment & sPatch2
597:        End If
598:        .InsertLines 1, sCode
599:    End With
600:
601:    Set objModuleWB = Nothing
602:    Exit Sub
ErrorHandler:
604:    Debug.Print "Error in VersionSistemControls.AddCommentVSC" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
605:    Call WriteErrorLog("VersionSistemControls.AddCommentVSC")
606: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddString - создает строку -> "'* * * * * * * VSC VBATools * * * * * * *
'* Created    : 06-03-2020 10:08
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function addString() As String
615:    Dim strTemp As String
616:    On Error GoTo ErrorHandler
617:    strTemp = VBA.Replace(VBA.String(20, "*"), "*", "* ")
618:    addString = "'" & strTemp & " VSC VBATools " & strTemp
619:    Exit Function
ErrorHandler:
621:    Debug.Print "Error in VersionSistemControls.AddString" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
622:    Call WriteErrorLog("VersionSistemControls.AddString")
End Function

