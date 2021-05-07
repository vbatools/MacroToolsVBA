VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VariableUnUsed 
   Caption         =   "Unused Variables:"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14445
   OleObjectBlob   =   "VariableUnUsed.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VariableUnUsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : VariableUnUsed - модуль поиска не используемых переменных в выбраном проекте
'* Created    : 23-01-2020 12:23
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
'словарь глобальных переменных
Private dicCollGlobalVariables As Scripting.Dictionary
'словарь Enum и Type
Private dicCollEnumType As Scripting.Dictionary
'словарь типов переменных
Private dicCollAsTypeVariable As Scripting.Dictionary
'словарь процедур и функций
Private dicCollSubsAndFunctions As Scripting.Dictionary

Private m_clsAnchors As CAnchors

    Private Sub UserForm_Initialize()
21:    Set m_clsAnchors = New CAnchors
22:    Set m_clsAnchors.objParent = Me
23:    ' restrict minimum size of userform
24:    m_clsAnchors.MinimumWidth = 727
25:    m_clsAnchors.MinimumHeight = 405.75
26:    With m_clsAnchors
27:        .funAnchor("cmbMain").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft
28:        .funAnchor("ListCode").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft Or enumAnchorStyleBottom
29:        .funAnchor("lbAnaliz").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleRight
30:        .funAnchor("btnCopyCode").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
31:        .funAnchor("lbCancel").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
32:        .funAnchor("lbLoad").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
33:        .funAnchor("lbLoadTxtFile").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
34:    End With
35: End Sub
    Private Sub UserForm_Activate()
37:    Dim vbProj      As VBIDE.VBProject
38:
39:    If Workbooks.Count = 0 Then
40:        Unload Me
41:        Call MsgBox("No open ones" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
42:        Exit Sub
43:    End If
44:    With Me.cmbMain
45:        .Clear
46:        On Error Resume Next
47:        For Each vbProj In Application.VBE.VBProjects
48:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
49:        Next
50:        On Error GoTo 0
51:        .Value = ActiveWorkbook.Name
52:    End With
53: End Sub
    Private Sub UserForm_Terminate()
55:    Set m_clsAnchors = Nothing
56: End Sub
    Private Sub btnCancel_Click()
58:    Unload Me
59: End Sub
    Private Sub lbCancel_Click()
61:    Unload Me
62: End Sub
    Private Sub lbLoadTxtFile_Click()
64:    Dim strVar      As String
65:    Dim strFileName As String
66:
67:    If cmbMain.Value = vbNullString Then Exit Sub
68:
69:    strVar = AddListImmediate()
70:    If strVar = vbNullString Then Exit Sub
71:
72:    With Workbooks(cmbMain.Value)
73:        strFileName = .Path & Application.PathSeparator & sGetFileName(.FullName) & ".txt"
74:    End With
75:
76:    If SaveTXTfile(strFileName, strVar) Then
77:        Call MsgBox("The data is copied to a txt file!", vbInformation, "Copying data:")
78:    Else
79:        Call MsgBox("Couldn't copy data to txt file!", vbCritical, "Copying data:")
80:    End If
81: End Sub
    Private Function SaveTXTfile(ByVal sFileName As String, ByVal Txt As String) As Boolean
83:    Dim FSO         As Object
84:    Dim ts          As Object
85:    On Error Resume Next: Err.Clear
86:    Set FSO = CreateObject("scripting.filesystemobject")
87:    Set ts = FSO.CreateTextFile(sFileName, True)
88:    ts.Write Txt: ts.Close
89:    SaveTXTfile = Err = 0
90:    Set ts = Nothing: Set FSO = Nothing
91: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : lbLoad_Click - выгрузка в окно immediate
'* Created    : 28-01-2020 14:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub lbLoad_Click()
101:    Debug.Print AddListImmediate()
102:    Call MsgBox("The data is copied to the Immediate window!", vbInformation, "Copying data:")
103: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : btnCopyCode_Click - сохранение в буффер обмена
'* Created    : 28-01-2020 14:35
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub btnCopyCode_Click()
112:    Dim strVar      As String
113:    strVar = AddListImmediate()
114:    If strVar = vbNullString Then Exit Sub
115:    Call C_PublicFunctions.SetTextIntoClipboard(strVar)
116:
117:    Call MsgBox("The data is copied to the clipboard!", vbInformation, "Copying data:")
118: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddListImmediate - создание текста не используемых переменых
'* Created    : 28-01-2020 14:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function AddListImmediate() As String
127:    Dim strData     As String
128:    Dim i           As Long
129:    Dim Max1        As Byte
130:    Dim Max2        As Byte
131:    Dim Max3        As Byte
132:    Dim Max4        As Byte
133:    Dim Max5        As Byte
134:
135:    Max1 = 0: Max2 = 0: Max3 = 0: Max4 = 0: Max5 = 0
136:
137:    With ListCode
138:        'поиск текста с максимальным кол-вом символов
139:        For i = 0 To .ListCount - 1
140:            If VBA.Len(.List(i, 1)) > Max1 Then Max1 = VBA.Len(.List(i, 1))
141:            If VBA.Len(.List(i, 2)) > Max2 Then Max2 = VBA.Len(.List(i, 2))
142:            If VBA.Len(.List(i, 3)) > Max3 Then Max3 = VBA.Len(.List(i, 3))
143:            If VBA.Len(.List(i, 4)) > Max4 Then Max4 = VBA.Len(.List(i, 4))
144:            If VBA.Len(.List(i, 5)) > Max5 Then Max5 = VBA.Len(.List(i, 5))
145:        Next i
146:        'вормирование текста
147:        For i = 0 To .ListCount - 1
148:            strData = strData & _
                            .List(i, 1) & VBA.String$(Max1 - VBA.Len(.List(i, 1)), " ") & vbTab & _
                            .List(i, 2) & VBA.String$(Max2 - VBA.Len(.List(i, 2)), " ") & vbTab & _
                            .List(i, 3) & VBA.String$(Max3 - VBA.Len(.List(i, 3)), " ") & vbTab & _
                            .List(i, 4) & VBA.String$(Max4 - VBA.Len(.List(i, 4)), " ") & vbTab & _
                            .List(i, 5) & VBA.String$(Max5 - VBA.Len(.List(i, 5)), " ") & vbNewLine
154:        Next i
155:    End With
156:    AddListImmediate = strData
157: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ListCode_DblClick - переход в модуль VBA
'* Created    : 28-01-2020 14:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByVal Cancel As MSForms.ReturnBoolean :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ListCode_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
170:    Dim i           As Long
171:    Dim WB          As Workbook
172:    Dim VBC         As VBIDE.VBComponent
173:
174:    'On Error GoTo ErrorHandler
175:
176:    If cmbMain.Value = vbNullString Then Exit Sub
177:    Set WB = Workbooks(cmbMain.Value)
178:    For i = 0 To ListCode.ListCount
179:        If ListCode.Selected(i) = True Then
180:            Set VBC = WB.VBProject.VBComponents(ListCode.List(i, 1))
181:            If VBC.Type = vbext_ct_MSForm Then
182:                VBC.CodeModule.CodePane.Show
183:            Else
184:                VBC.Activate
185:            End If
186:            Exit Sub
187:        End If
188:    Next i
189:    Exit Sub
ErrorHandler:
191:    Unload Me
192:    Select Case Err.Number
        Case Else:
194:            Call MsgBox("Error! in VariableUnUsed. ListCode_DblClick" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbExclamation, "Error:")
195:            Call WriteErrorLog("VariableUnUsed.ListCode_DblClick")
196:    End Select
197:    Err.Clear
198: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : cmbMain_Change - выбор файла Exsel
'* Created    : 28-01-2020 14:49
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub lbAnaliz_Click()
207:    If cmbMain.Value = vbNullString Then Exit Sub
208:    ListCode.Clear
209:    Call MainSubAddUnUsed
210: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : MainSubAddUnUsed - главная процедура запуска поиска не используемых перемменых
'* Created    : 28-01-2020 14:50
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub MainSubAddUnUsed()
219:    Dim WBName      As String
220:    WBName = cmbMain.Value
221:    If WBName <> vbNullString Then
222:        Dim objVBP  As VBIDE.VBProject
223:        Set objVBP = Workbooks(WBName).VBProject
224:        If objVBP.Protection <> vbext_pp_none Then
225:            ListCode.Clear
226:            Call MsgBox("VBA project in the book -" & WBName & "password protected!" & vbCrLf & "Remove the password!", vbCritical, "Error:")
227:            Exit Sub
228:        End If
229:
230:        Dim dTime   As Date
231:        dTime = Now()
232:
233:        'процедуры создания словарей
234:        Call CreateGlobalVariableCollection(WBName)
235:        Debug.Print "All global variables:" & Format(Now() - dTime, "Long Time") & "completed!"
236:        Call ProcessUnusedVariable(WBName)
237:        Debug.Print "Procedure variables:" & Format(Now() - dTime, "Long Time") & "completed!"
238:        Call FillUnusedGlobalVariables(WBName)
239:        Debug.Print "Global variables unused:" & Format(Now() - dTime, "Long Time") & "completed!"
240:        Call AddListUnUsedEnumAndType
241:        Debug.Print "Types and enumerations:" & Format(Now() - dTime, "Long Time") & "completed!"
242:        Call AddSubAndFuncListUnUsed(WBName)
243:        Debug.Print "Procedures and functions:" & Format(Now() - dTime, "Long Time") & "completed!"
244:    End If
245:
246:    'удаляю все словари
247:    Set dicCollGlobalVariables = Nothing
248:    Set dicCollEnumType = Nothing
249:    Set dicCollAsTypeVariable = Nothing
250:    Set dicCollSubsAndFunctions = Nothing
251:
252: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ProcessUnusedVariable - главная процедура поиска не используемых переменных в процедурах и функциях
'* Created    : 23-01-2020 12:24
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal WBName As String : имя файла Excel
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ProcessUnusedVariable(ByVal WBName As String)
265:
266:    Dim objVBP      As VBIDE.VBProject
267:    Dim vbComp      As VBIDE.VBComponent
268:    Dim ProcKind    As VBIDE.vbext_ProcKind
269:    Dim CodeMod     As CodeModule
270:    Dim intLine     As Integer
271:    Dim strFunctionBody As String
272:    Dim strFunctionTypeAs As String
273:    Dim strProcName As String
274: Dim strFerstStringSub As String
275:    Dim g_wbkVBAExcel As Workbook
276:
277:    'On Error GoTo ErrorHandler
278:
279:    Set g_wbkVBAExcel = Workbooks(WBName)
280:    Set objVBP = g_wbkVBAExcel.VBProject
281:
282:    Set dicCollSubsAndFunctions = Nothing
283:    Set dicCollSubsAndFunctions = New Scripting.Dictionary
284:    For Each vbComp In objVBP.VBComponents
285:        If vbComp.Type <> vbext_ct_ClassModule Then
286:            Set CodeMod = vbComp.CodeModule
287:            For intLine = 1 To CodeMod.CountOfLines
288:                strProcName = CodeMod.ProcOfLine(intLine, ProcKind)
289:                If strProcName <> vbNullString Then
290:                    'создание словаря переменных процедур и функций
291:                    strFunctionBody = CodeMod.Lines(intLine, CodeMod.ProcCountLines(strProcName, ProcKind))
292: strFerstStringSub = CodeMod.Lines(CodeMod.ProcBodyLine(strProcName, ProcKind), 1)
293:
294: If strFerstStringSub Like "*) As *" And strFerstStringSub Like "Function *" Then
295:                        Dim arrAs As Variant
296:                        arrAs = Split(strFerstStringSub, ") As ")
297:                        strFunctionTypeAs = arrAs(UBound(arrAs))
298:                        arrAs = Empty
299:                    Else
300:                        strFunctionTypeAs = "-"
301:                    End If
302:
303:                    intLine = intLine + CodeMod.ProcCountLines(strProcName, ProcKind) - 1
304:                    Call FillUnusedLocalVariables(strFunctionBody, strProcName, vbComp.Name, ProcKind)
305:
306:                    'создание словаря названий процедур и функций
307:                    If dicCollSubsAndFunctions.Exists(vbComp.Name & "." & strProcName) = False Then
308:                        Dim sTypeSubFun As String
309:
310:                        If strFunctionBody Like "*Private *" Then
311:                            sTypeSubFun = "Private"
312:                        Else
313:                            sTypeSubFun = "Public"
314:                        End If
315:                        'не загружаю процедуры рибон понели, события листов и книги, UserForm в формах
316:                        If (Not strFunctionBody Like "*As IRibbonControl*") And _
                                        (Not WorkBookAndSheetsEvents(strFunctionBody, vbComp.Type)) And _
                                        (Not (strFunctionBody Like "* UserForm_*" And vbComp.Type = vbext_ct_MSForm)) Then
319:                            Call dicCollSubsAndFunctions.Add(vbComp.Name & "." & _
                                            strProcName, _
                                            sTypeSubFun & "." & _
                                            I_StatisticVBAProj.TypeProcedyre(strFunctionBody) & "." & _
                                            byTypeProc(strFerstStringSub) & "." & _
                                            strFunctionTypeAs)
325:                        End If
326:                    End If
327: strFerstStringSub = vbNullString
328:                    strFunctionBody = vbNullString
329:                    strProcName = vbNullString
330:                    strFunctionTypeAs = vbNullString
331:                End If
332:            Next
333:        End If
334:    Next
335:    Set objVBP = Nothing
336:    Set vbComp = Nothing
337:    Set CodeMod = Nothing
338:
339:    Exit Sub
ErrorHandler:
341:    Unload Me
342:    Select Case Err.Number
        Case Else:
344:            Call MsgBox("Error in VariableUnUsed.ProcessUnusedVariable" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbExclamation, "Error:")
345:            Call WriteErrorLog("VariableUnUsed.ProcessUnusedVariable")
346:    End Select
347:    Err.Clear
348: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : FillUnusedLocalVariables - поиск не использоемых переменых в выбраной процедуре или функции и дополнение словаря типов переменных
'* Created    : 23-01-2020 12:26
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                     Description
'*
'* FunctionBody As String        : код процедуры
'* FunctionName As String        : имя процедуры
'* ModuleName As String          : имя модуля
'* ByVal FunctionType As Integer : тип функции
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub FillUnusedLocalVariables(FunctionBody As String, FunctionName As String, ModuleName As String, ByVal FunctionType As Integer)
364:
365:    Dim strVrName   As String
366:    Dim strArrLine() As String
367:    Dim strLine     As String
368:    Dim strArrVariables() As String
369:    Dim intCounterLine As Integer
370:    Dim intCounterVariable As Integer
371:    Dim strDeclaration As String
372:
373:    'On Error GoTo ErrorHandler
374:
375:    strVrName = vbNullString
376:
377:    FunctionBody = Replace(FunctionBody, " _" & vbCrLf, " ")
378:    strArrLine = Split(FunctionBody, vbCrLf)
379:
380:    For intCounterLine = LBound(strArrLine) To UBound(strArrLine)
381:        strLine = Trim(strArrLine(intCounterLine))
382:        'переменные
383:        If strLine Like "Dim *" Or strLine Like "Static *" Or strLine Like "Const *" Then
384:            strArrVariables = Split(RemoveEnclosedStringAndComments(strLine), ",")
385:            For intCounterVariable = LBound(strArrVariables) To UBound(strArrVariables)
386:                strDeclaration = Trim(strArrVariables(intCounterVariable))
387:                strVrName = Mid(strArrVariables(intCounterVariable), InStr(1, strArrVariables(intCounterVariable), " "), 100)
388:                strVrName = Replace(strVrName, ",", vbNullString)
389:                If strVrName Like "*()" Then
390:                    strVrName = Left(strVrName, Len(strVrName) - 2)
391:                End If
392:                If IsVariableUsed(strVrName, strArrLine) = False Then
393:                    With ListCode
394:                        Dim lListRow As Long
395:                        Dim arrVar As Variant
396:                        Dim strVar As String
397:                        lListRow = .ListCount
398:                        .AddItem lListRow + 1
399:                        .List(lListRow, 1) = VBA.Trim$(ModuleName)
400:                        If strDeclaration Like "Const *" Then
401:                            .List(lListRow, 2) = "Const"
402:                        Else
403:                            .List(lListRow, 2) = "Dim"
404:                        End If
405:                        .List(lListRow, 3) = VBA.Trim$(FunctionName)
406:                        arrVar = Split(strVrName, " As ")
407:                        .List(lListRow, 4) = VBA.Trim$(VBA.Replace(arrVar(0), "=", vbNullString))
408:                        If UBound(arrVar) = 0 Then
409:                            strVar = "Variant"
410:                        Else
411:                            If strDeclaration Like "Const *" Then
412:                                strVar = VBA.Trim$(VBA.Left$(arrVar(1), VBA.InStr(1, arrVar(1), "=") - 1))
413:                            Else
414:                                strVar = VBA.Trim$(VBA.Replace(arrVar(1), "=", vbNullString))
415:                            End If
416:                        End If
417:                        .List(lListRow, 5) = strVar
418:                    End With
419:                End If
420:            Next intCounterVariable
421:        End If
422:        'типы переменых переменные дополнение словаря типов переменных
423:        If strLine Like "* As *" Then
424:            Call AddDictionariAsType(strLine, ModuleName)
425:        End If
426:    Next intCounterLine
427:    Exit Sub
ErrorHandler:
429:    Unload Me
430:    Select Case Err.Number
        Case Else:
432:            Call MsgBox("Error in VariableUnUsed.FillUnusedLocalVariables" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbExclamation, "Error:")
433:            Call WriteErrorLog("VariableUnUsed.FillUnusedLocalVariables")
434:    End Select
435:    Err.Clear
436: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : IsVariableUsed - определение использования переменой
'* Created    : 23-01-2020 12:27
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                     Description
'*
'* ByVal strVrName As String     : имя переменой
'* ByRef strArrofLine As Variant : массив слов кода
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function IsVariableUsed(ByVal strVrName As String, ByRef strArrofLine As Variant) As Boolean
450:    Dim intLoop     As Integer
451:    Dim strLine     As String
452:
453:    IsVariableUsed = False
454:
455:    ' Format Variablename
456:    strVrName = C_PublicFunctions.TrimSpace(strVrName)
457:    strVrName = Replace(strVrName, "(", " ")
458:    strVrName = Replace(strVrName, "%", vbNullString)     'Integer
459:    strVrName = Replace(strVrName, "&", vbNullString)     'Long
460:    strVrName = Replace(strVrName, "$", vbNullString)     'String
461:    strVrName = Replace(strVrName, "!", vbNullString)     'Single
462:    strVrName = Replace(strVrName, "#", vbNullString)     'Double
463:    strVrName = Replace(strVrName, "@", vbNullString)     'Currency
464:    If strVrName <> vbNullString Then strVrName = Split(strVrName, " ")(0)
465:    For intLoop = 0 To UBound(strArrofLine)
466:        strLine = C_PublicFunctions.TrimSpace(strArrofLine(intLoop))
467:        strLine = RemoveEnclosedStringAndComments(strLine)
468:
469:        If Not StrigLikeWord(strLine) Then
470:            strLine = Replace(strLine, "(", " ")
471:            strLine = Replace(strLine, ".", " ")
472:            strLine = Replace(strLine, ")", " ")
473:            strLine = Replace(strLine, ",", " ")
474:            strLine = " " & strLine & " "
475:
476:            If strLine Like "* " & strVrName & " *" Then
477:                IsVariableUsed = True
478:                Exit Function
479:            End If
480:        End If
481:    Next
482:
483: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : FillUnusedGlobalVariables - поиск глобальных не использоемых переменых в выбраном проекте VBA
'* Created    : 23-01-2020 12:28
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal WBName As String : имя файла Excel
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub FillUnusedGlobalVariables(ByVal WBName As String)
496:
497:    Dim objVBP      As VBIDE.VBProject
498:    Dim vbComp      As VBIDE.VBComponent
499:    Dim VBCompTemp  As VBIDE.VBComponent
500:    Dim objCodeMod  As CodeModule
501:    Dim strArrLine() As String
502:    Dim strContent  As String
503:    Dim strArrWord() As String
504:    Dim strWord     As String
505:    Dim intWordCount As Integer
506:    Dim strLine     As String
507:    Dim intCounter  As Integer
508:    Dim collUsedGlobal As New Dictionary
509:    Dim g_wbkVBAExcel As Workbook
510:
511:    'On Error GoTo ErrorHandler
512:
513:    Set g_wbkVBAExcel = Workbooks(WBName)
514:    Set objVBP = g_wbkVBAExcel.VBProject
515:    'удаление из словаря используемых переменных
516:    For Each vbComp In objVBP.VBComponents
517:        If vbComp.Type <> vbext_ct_ClassModule Then     '!!!
518:            Set objCodeMod = vbComp.CodeModule
519:            If objCodeMod.CountOfLines > 0 Then
520:                strContent = objCodeMod.Lines(1, objCodeMod.CountOfLines)
521:                strContent = Replace(strContent, " _" & vbCrLf, " ")
522:                strArrLine = Split(strContent, vbCrLf)
523:
524:                For intCounter = LBound(strArrLine) To UBound(strArrLine)
525:                    strLine = RemoveEnclosedStringAndComments(strArrLine(intCounter))
526:                    'глобальная переменная с присвоением, проверяем вторую часть
527:                    If StrigLikeWord(strLine) And strLine Like "*=*" Then strLine = Split(strLine, "=")(1)
528:                    If Not StrigLikeWord(strLine) Then
529:                        strLine = Replace(strLine, ",", " ")
530:                        strLine = Replace(strLine, ")", " ")
531:                        strLine = Replace(strLine, "(", " ")
532:                        strArrWord = Split(Replace(strLine, ".", " "), " ")
533:                        For intWordCount = LBound(strArrWord) To UBound(strArrWord)
534:                            If strArrWord(intWordCount) <> "" Then
535:                                strWord = Split(strArrWord(intWordCount), "(")(0)
536:                                For Each VBCompTemp In objVBP.VBComponents
537:                                    If dicCollGlobalVariables.Exists(VBCompTemp.Name & "." & strWord) Then
538:                                        Call collUsedGlobal.Add(VBCompTemp.Name & "." & strWord, dicCollGlobalVariables.Item(VBCompTemp.Name & "." & strWord))
539:                                        dicCollGlobalVariables.Remove (VBCompTemp.Name & "." & strWord)
540:                                        Exit For
541:                                    End If
542:                                Next VBCompTemp
543:                            End If
544:                        Next intWordCount
545:                    End If
546:                Next intCounter
547:            End If
548:        End If
549:    Next vbComp
550:    'глабальные переменные которые не используются
551:    For intCounter = 0 To dicCollGlobalVariables.Count - 1
552:        With ListCode
553:            Dim lListRow As Long
554:            Dim strarr As Variant
555:            Dim strarr1 As Variant
556:            Dim strarr2 As Variant
557:            Dim strVal As String
558:            Dim strVariable As String
559:            lListRow = .ListCount
560:            .AddItem lListRow
561:            .List(lListRow, 1) = VBA.Trim$(Split(dicCollGlobalVariables.Keys(intCounter), ".")(0))
562:            .List(lListRow, 3) = "-"
563:
564:            strVariable = VBA.Trim$(dicCollGlobalVariables.Items(intCounter))
565:            strarr = Split(strVariable, "As")
566:            strarr2 = Split(Trim$(strarr(0)), " ")
567:            If UBound(strarr) = 0 Then
568:                .List(lListRow, 4) = strarr2(UBound(strarr2))
569:                .List(lListRow, 5) = "Variant"
570:                .List(lListRow, 2) = Replace(strarr(0), " " & .List(lListRow, 4), vbNullString)
571:            Else
572:                strarr1 = Split(Trim$(strarr(1)), " ")
573:                If UBound(strarr1) = 0 Then
574:                    strVal = strarr(1)
575:                Else
576:                    strVal = strarr1(0)
577:                End If
578:                .List(lListRow, 4) = strarr2(UBound(strarr2))
579:                .List(lListRow, 5) = VBA.Trim$(strVal)
580:                .List(lListRow, 2) = Replace(strarr(0), " " & .List(lListRow, 4), vbNullString)
581:            End If
582:        End With
583:    Next
584:
585:    Exit Sub
ErrorHandler:
587:    Unload Me
588:    Select Case Err.Number
        Case Else:
590:            Call MsgBox("Error in VariableUnUsed.FillUnusedGlobalVariables" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbExclamation, "Error:")
591:            Call WriteErrorLog("VariableUnUsed.FillUnusedGlobalVariables")
592:    End Select
593:    Err.Clear
594: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CreateGlobalVariableCollection - создание словарей с глобальными переменными и их типами выбранного проекта VBA
'* Created    : 23-01-2020 12:30
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal WBName As String : имя файла Excel
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub CreateGlobalVariableCollection(ByVal WBName As String)
607:
608:    Dim vbComp      As VBIDE.VBComponent
609:    Dim vbProj      As VBIDE.VBProject
610:    Dim CodeMod     As VBIDE.CodeModule
611:    Dim strContent  As String
612:    Dim strArrLine() As String
613:    Dim strArrVariables() As String
614:    Dim intLineCounter As Integer
615:    Dim intVarCounter As Integer
616:    Dim strVarName  As String
617:    Dim strVarNameEnumType As String
618:    Dim strVarDeclaration As String
619:    Dim strLine     As String
620:    Dim g_wbkVBAExcel As Workbook
621:
622:    'On Error GoTo ErrorHandler
623:
624:    Set g_wbkVBAExcel = Workbooks(WBName)
625:    Set vbProj = g_wbkVBAExcel.VBProject
626:
627:    Set dicCollGlobalVariables = Nothing
628:    Set dicCollGlobalVariables = New Scripting.Dictionary
629:
630:    Set dicCollEnumType = Nothing
631:    Set dicCollEnumType = New Scripting.Dictionary
632:
633:    Set dicCollAsTypeVariable = Nothing
634:    Set dicCollAsTypeVariable = New Scripting.Dictionary
635:
636:    For Each vbComp In vbProj.VBComponents
637:        If vbComp.Type <> vbext_ct_ClassModule Then
638:            Set CodeMod = vbComp.CodeModule
639:            If CodeMod.CountOfDeclarationLines > 0 Then
640:                strContent = CodeMod.Lines(1, CodeMod.CountOfDeclarationLines)
641:                strContent = Replace(strContent, " _" & vbCrLf, " ")
642:
643:                strArrLine = Split(strContent, vbCrLf)
644:                For intLineCounter = LBound(strArrLine) To UBound(strArrLine)
645:
646:                    strLine = RemoveEnclosedStringAndComments(strArrLine(intLineCounter))
647:                    'собираю типы глабальных переменных
648:                    If strLine Like "* As *" Then
649:                        Call AddDictionariAsType(strLine, CodeMod.Name)
650:                    End If
651:                    If StrigLikeWord(strLine) Then
652:
653:                        strLine = Replace(strLine, "=", vbNullString)
654:
655:                        strArrVariables = Split(strLine, ",")
656:                        For intVarCounter = LBound(strArrVariables) To UBound(strArrVariables)
657:                            strVarDeclaration = strArrVariables(intVarCounter)
658:                            strVarName = Trim(Split(strArrVariables(intVarCounter), " ")(1))
659:                            If strVarName <> vbNullString Then
660:                                If strVarName = "Const" Then
661:                                    strVarName = Trim(Split(strArrVariables(intVarCounter), " ")(2))
662:                                End If
663:                                'словарь для Enum и Type
664:                                If strVarName = "Enum" Or strVarName = "Type" Then
665:                                    strVarNameEnumType = Trim(Split(strArrVariables(intVarCounter), " ")(2))
666:                                    If dicCollEnumType.Exists(vbComp.Name & "." & strVarNameEnumType) = False Then Call dicCollEnumType.Add(vbComp.Name & "." & strVarNameEnumType, strVarDeclaration)
667:                                End If
668:                                strVarName = Split(strVarName, "(")(0)
669:                                strVarName = Replace(strVarName, ",", vbNullString)
670:                                'словарь для переменных
671:                                If Not (strVarDeclaration Like "Enum*" Or strVarDeclaration Like "Type*") Then
672:                                    If dicCollGlobalVariables.Exists(vbComp.Name & "." & strVarName) = False Then Call dicCollGlobalVariables.Add(vbComp.Name & "." & strVarName, strVarDeclaration)
673:                                End If
674:                            End If
675:                        Next intVarCounter
676:
677:                    End If
678:                Next intLineCounter
679:            End If
680:        End If
681:    Next vbComp
682:
683:    Exit Sub
ErrorHandler:
685:    Unload Me
686:    Select Case Err.Number
        Case Else:
688:            Call MsgBox("Error in Variable UnUsed.CreateGlobalVariablesCollection" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbExclamation, "Error:")
689:            Call WriteErrorLog("VariableUnUsed.CreateGlobalVariableCollection")
690:    End Select
691:    Err.Clear
692: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddDictionariAsType - создание словаря всех типов данных используемых в файле
'* Created    : 28-01-2020 11:38
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal strLine As String    : текстовая строка для анализа
'* ByVal NameModule As String : название модуля VBA
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub AddDictionariAsType(ByVal strLine As String, ByVal NameModule As String)
706:    Dim strVar      As String
707:    Dim strVar1     As String
708:    Dim arrVar      As Variant
709:
710:    Dim i           As Long
711:
712:    arrVar = Split(strLine, " As ")
713:
714:    For i = 1 To UBound(arrVar)
715:
716:        strVar = VBA.Trim$(arrVar(i))
717:
718:        If strVar Like "*, *" Then
719:            strVar = VBA.Trim$(Split(strVar, ",")(0))
720:        End If
721:        If strVar Like "* *" Then
722:            strVar = VBA.Trim$(Split(strVar, " ")(0))
723:        End If
724:
725:        If strVar Like "*)" Then
726:            strVar = VBA.Trim$(Replace(strVar, ")", vbNullString))
727:        End If
728:        If strVar Like "*(" Then
729:            strVar = VBA.Trim$(Replace(strVar, "(", vbNullString))
730:        End If
731:        strVar = VBA.Trim$(Replace(strVar, Chr(34), vbNullString))
732:        strVar = VBA.Trim$(Replace(strVar, "*", vbNullString))
733:        strVar = VBA.Trim$(Replace(strVar, "#", "-"))
734:
735:        If Not strVar Like "-*" And strVar <> vbNullString Then
736:            strVar1 = strVar
737:            If strLine Like "Private *" Then
738:                strVar = VBA.Trim$(NameModule) & "." & strVar
739:            Else
740:                strVar = strVar
741:            End If
742:            If dicCollAsTypeVariable.Exists(strVar1) = False Then Call dicCollAsTypeVariable.Add(strVar1, strVar)
743:        End If
744:    Next i
745: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddListUnUsedEnumAndType - вывод не используемых типов и пречеслений
'* Created    : 28-01-2020 11:33
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub AddListUnUsedEnumAndType()
754:    Dim intCounter  As Integer
755:    Dim strKey      As String
756:    Dim strKey1     As String
757:    Dim lListRow    As Long
758:
759:    For intCounter = 0 To dicCollEnumType.Count - 1
760:        strKey = dicCollEnumType.Keys()(intCounter)
761:        strKey1 = Split(strKey, ".")(1)
762:        strKey = Split(strKey, ".")(0)
763:        If dicCollAsTypeVariable.Exists(strKey1) = False Then
764:            With ListCode
765:                lListRow = .ListCount
766:                .AddItem lListRow + 1
767:                .List(lListRow, 1) = strKey
768:                .List(lListRow, 2) = VBA.Trim$(Replace(dicCollEnumType.Items()(intCounter), strKey1, vbNullString))
769:                .List(lListRow, 3) = "-"
770:                .List(lListRow, 4) = strKey1
771:                .List(lListRow, 5) = "-"
772:                '                If .List(lListRow, 4) Like "*Enum*" Then
773:                '                    .List(lListRow, 4) = "перечисление"
774:                '                Else
775:                '                    .List(lListRow, 4) = "пользовательский тип"
776:                '                End If
777:            End With
778:        End If
779:    Next intCounter
780: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : StrigLikeWord - если в слове содержится часть другого слова то истина
'* Created    : 23-01-2020 10:53
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sTxt As String : исходная строка
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function StrigLikeWord(ByVal sTxt As String) As Boolean
793:    Dim Flag        As Boolean
794:    Flag = False
795:    Select Case True
        Case sTxt Like "* declare *": Flag = False
797:        Case sTxt Like "* Declare *": Flag = False
798:
799:
800:        Case sTxt Like vbNullString: Flag = True
801:        Case sTxt Like "Public *": Flag = True
802:        Case sTxt Like "Private *": Flag = True
803:        Case sTxt Like "Global *": Flag = True
804:        Case sTxt Like "Const *": Flag = True
805:        Case sTxt Like "Set*= Nothing": Flag = True
806:        Case sTxt Like "Dim *": Flag = True
807:        Case sTxt Like "*Enum *": Flag = True
808:        Case sTxt Like "*Type *": Flag = True
809:    End Select
810:    StrigLikeWord = Flag
811: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : RemoveEnclosedStringAndComments - очистка строки
'* Created    : 23-01-2020 12:28
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal strLine As String : строка входная
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function RemoveEnclosedStringAndComments(ByVal strLine As String) As String
824:
825:    Dim strArrParts() As String
826:    Dim intCntPart  As Integer
827:
828:    If strLine = vbNullString Then
829:        RemoveEnclosedStringAndComments = vbNullString
830:        Exit Function
831:    End If
832:
833:    strLine = Replace(strLine, vbTab, vbNullString)
834:    strArrParts = Split(strLine, """")
835:    strLine = vbNullString
836:
837:    For intCntPart = LBound(strArrParts) To UBound(strArrParts)
838:        If intCntPart Mod 2 = 0 Then
839:            strLine = strLine & " " & Trim(strArrParts(intCntPart))
840:        End If
841:    Next intCntPart
842:
843:    If strLine <> vbNullString Then
844:        strLine = Split(strLine, "'")(0)
845:    End If
846:    strLine = C_PublicFunctions.TrimSpace(strLine)
847:    RemoveEnclosedStringAndComments = strLine
848:
849: End Function

     Private Sub AddSubAndFuncListUnUsed(ByVal WBName As String)
852:    Dim vbProj      As VBIDE.VBProject
853:    Dim vbComp      As VBIDE.VBComponent
854:    Dim CodeMod     As VBIDE.CodeModule
855:    Dim intCounter  As Integer
856:    Dim strKey      As String
857:    Dim strKeyNameMod As String
858: Dim strKeyNameSub As String
859:    Dim strItemTypePub As String
860: Dim strItemTypeSub As String
861:    Dim strItemTypeProc As String
862:    Dim strFunctionBody As String
863:    Dim strModuleBody As String
864:
865:    Set vbProj = Workbooks(WBName).VBProject
866:
867:    For intCounter = dicCollSubsAndFunctions.Count - 1 To 0 Step -1
868:        strKey = dicCollSubsAndFunctions.Keys()(intCounter)
869:        strKeyNameMod = Split(dicCollSubsAndFunctions.Keys()(intCounter), ".")(0)
870: strKeyNameSub = Split(dicCollSubsAndFunctions.Keys()(intCounter), ".")(1)
871:        strItemTypePub = Split(dicCollSubsAndFunctions.Items()(intCounter), ".")(0)
872: strItemTypeSub = Split(dicCollSubsAndFunctions.Items()(intCounter), ".")(1)
873:        strItemTypeProc = CByte(Split(dicCollSubsAndFunctions.Items()(intCounter), ".")(2))
874:
875:        'если Private  то ищем в именно в этом модуле в других не надо
876:        If strItemTypePub = "Private" Then
877:            Set vbComp = vbProj.VBComponents(strKeyNameMod)
878:            Set CodeMod = vbComp.CodeModule
879:            With CodeMod
880:                strFunctionBody = .Lines(.ProcStartLine(strKeyNameSub, strItemTypeProc), .ProcCountLines(strKeyNameSub, strItemTypeProc))
881:                strModuleBody = .Lines(1, .CountOfLines)
882:            End With
883:            strModuleBody = VBA.Replace(strModuleBody, strFunctionBody, vbNullString)
884:            'если нашли то удаляем из словаря
885:            If VBA.InStr(1, strModuleBody, strKeyNameSub, vbTextCompare) <> 0 Then
886:                dicCollSubsAndFunctions.Remove (strKey)
887:                'если форма
888:            ElseIf vbComp.Type = vbext_ct_MSForm Then
889: If strKeyNameSub Like "*_*" Then
890:                    'если элемент формы
891:                    Dim strNameItemDesiner As String
892:                    Dim arrVar As Variant
893:                    arrVar = Split(strKeyNameSub, "_")
894:                    strNameItemDesiner = VBA.Replace(strKeyNameSub, "_" & arrVar(UBound(arrVar)), vbNullString)
895:                    On Error Resume Next
896:                    strNameItemDesiner = vbComp.Designer.Item(strNameItemDesiner).Name
897:                    If strNameItemDesiner <> vbNullString Then
898:                        dicCollSubsAndFunctions.Remove (strKey)
899:                    End If
900:                    strNameItemDesiner = vbNullString
901:                    On Error GoTo 0
902:                End If
903:            End If
904:            'если Public
905:        Else
906:            For Each vbComp In vbProj.VBComponents
907:                Set CodeMod = vbComp.CodeModule
908:                With CodeMod
909:                    If .CountOfLines <> 0 Then
910:                        'если ищем в родительском модуле
911:                        If .Name = strKeyNameMod Then
912:                            strFunctionBody = .Lines(.ProcStartLine(strKeyNameSub, strItemTypeProc), .ProcCountLines(strKeyNameSub, strItemTypeProc))
913:                            strModuleBody = .Lines(1, .CountOfLines)
914:                            strModuleBody = VBA.Replace(strModuleBody, strFunctionBody, vbNullString)
915:                        Else
916:                            strModuleBody = .Lines(1, .CountOfLines)
917:                        End If
918:
919:                        'если нашли то удаляем из словаря и выходим из цикла
920:                        If VBA.InStr(1, strModuleBody, strKeyNameSub, vbTextCompare) <> 0 Then
921:                            dicCollSubsAndFunctions.Remove (strKey)
922:                            Exit For
923:                        End If
924:                    End If
925:                End With
926:            Next vbComp
927:        End If
928:    Next intCounter
929:    For intCounter = 0 To dicCollSubsAndFunctions.Count - 1
930:        Dim lListRow As Integer
931:        With ListCode
932:            lListRow = .ListCount
933:            .AddItem lListRow
934:            .List(lListRow, 1) = Split(dicCollSubsAndFunctions.Keys()(intCounter), ".")(0)
935:            .List(lListRow, 2) = Split(dicCollSubsAndFunctions.Items()(intCounter), ".")(0) & " " & Split(dicCollSubsAndFunctions.Items()(intCounter), ".")(1)
936:            .List(lListRow, 3) = Split(dicCollSubsAndFunctions.Keys()(intCounter), ".")(1)
937:            .List(lListRow, 4) = "-"
938:            .List(lListRow, 5) = Split(dicCollSubsAndFunctions.Items()(intCounter), ".")(3)
939:        End With
940:    Next intCounter
941: End Sub
     Private Function byTypeProc(ByVal sTxt As String) As Byte
943:    Select Case True
        Case sTxt Like "*Property Get*": byTypeProc = 3
945:        Case sTxt Like "*Property Set*": byTypeProc = 2
946:        Case sTxt Like "*Property Let*": byTypeProc = 1
947:        Case Else: byTypeProc = 0
948:    End Select
949: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WorkBookAndSheetsEvents - определение процедур событий листов, книги и диаграммы
'* Created    : 29-01-2020 13:15
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sTxt As String : - анализируемая строка
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
      Private Function WorkBookAndSheetsEvents(ByVal sTxt As String, ByVal TypeModule As VBIDE.vbext_ComponentType) As Boolean
963:    Dim Flag        As Boolean
964:    Flag = False
965:    'только для модулей листов и книг
966:    If TypeModule = vbext_ct_Document Then
967:        Select Case True
            Case sTxt Like "*Worksheet_Activate(*": Flag = True
969:            Case sTxt Like "*Worksheet_BeforeDoubleClick(*": Flag = True
970:            Case sTxt Like "*Worksheet_BeforeRightClick(*": Flag = True
971:            Case sTxt Like "*Worksheet_Calculate(*": Flag = True
972:            Case sTxt Like "*Worksheet_Change(*": Flag = True
973:            Case sTxt Like "*Worksheet_Deactivate(*": Flag = True
974:            Case sTxt Like "*Worksheet_FollowHyperlink(*": Flag = True
975:            Case sTxt Like "*Worksheet_PivotTableAfterValueChange(*": Flag = True
976:            Case sTxt Like "*Worksheet_PivotTableBeforeAllocateChanges(*": Flag = True
977:            Case sTxt Like "*Worksheet_PivotTableBeforeCommitChanges(*": Flag = True
978:            Case sTxt Like "*Worksheet_PivotTableBeforeDiscardChanges(*": Flag = True
979:            Case sTxt Like "*Worksheet_PivotTableChangeSync(*": Flag = True
980:            Case sTxt Like "*Worksheet_PivotTableUpdate(*": Flag = True
981:            Case sTxt Like "*Worksheet_SelectionChange(*": Flag = True
982:            Case sTxt Like "*Chart_Activate(*": Flag = True
983:            Case sTxt Like "*Chart_BeforeDoubleClick(*": Flag = True
984:            Case sTxt Like "*Chart_BeforeRightClick(*": Flag = True
985:            Case sTxt Like "*Chart_Calculate(*": Flag = True
986:            Case sTxt Like "*Chart_Deactivate(*": Flag = True
987:            Case sTxt Like "*Chart_MouseDown(*": Flag = True
988:            Case sTxt Like "*Chart_MouseMove(*": Flag = True
989:            Case sTxt Like "*Chart_MouseUp(*": Flag = True
990:            Case sTxt Like "*Chart_Resize(*": Flag = True
991:            Case sTxt Like "*Chart_SeriesChange(*": Flag = True
992:            Case sTxt Like "*Workbook_Activate(*": Flag = True
993:            Case sTxt Like "*Workbook_AddinInstall(*": Flag = True
994:            Case sTxt Like "*Workbook_AddinUninstall(*": Flag = True
995:            Case sTxt Like "*Workbook_AfterSave(*": Flag = True
996:            Case sTxt Like "*Workbook_AfterXmlExport(*": Flag = True
997:            Case sTxt Like "*Workbook_AfterXmlImport(*": Flag = True
998:            Case sTxt Like "*Workbook_BeforeClose(*": Flag = True
999:            Case sTxt Like "*Workbook_BeforePrint(*": Flag = True
1000:            Case sTxt Like "*Workbook_BeforeSave(*": Flag = True
1001:            Case sTxt Like "*Workbook_BeforeXmlExport(*": Flag = True
1002:            Case sTxt Like "*Workbook_BeforeXmlImport(*": Flag = True
1003:            Case sTxt Like "*Workbook_Deactivate(*": Flag = True
1004:            Case sTxt Like "*Workbook_NewChart(*": Flag = True
1005:            Case sTxt Like "*Workbook_NewSheet(*": Flag = True
1006:            Case sTxt Like "*Workbook_Open(*": Flag = True
1007:            Case sTxt Like "*Workbook_PivotTableCloseConnection(*": Flag = True
1008:            Case sTxt Like "*Workbook_PivotTableOpenConnection(*": Flag = True
1009:            Case sTxt Like "*Workbook_RowsetComplete(*": Flag = True
1010:            Case sTxt Like "*Workbook_SheetActivate(*": Flag = True
1011:            Case sTxt Like "*Workbook_SheetBeforeDoubleClick(*": Flag = True
1012:            Case sTxt Like "*Workbook_SheetBeforeRightClick(*": Flag = True
1013:            Case sTxt Like "*Workbook_SheetCalculate(*": Flag = True
1014:            Case sTxt Like "*Workbook_SheetChange(*": Flag = True
1015:            Case sTxt Like "*Workbook_SheetDeactivate(*": Flag = True
1016:            Case sTxt Like "*Workbook_SheetFollowHyperlink(*": Flag = True
1017:            Case sTxt Like "*Workbook_SheetPivotTableAfterValueChange(*": Flag = True
1018:            Case sTxt Like "*Workbook_SheetPivotTableBeforeAllocateChanges(*": Flag = True
1019:            Case sTxt Like "*Workbook_SheetPivotTableBeforeCommitChanges(*": Flag = True
1020:            Case sTxt Like "*Workbook_SheetPivotTableBeforeDiscardChanges(*": Flag = True
1021:            Case sTxt Like "*Workbook_SheetPivotTableChangeSync(*": Flag = True
1022:            Case sTxt Like "*Workbook_SheetPivotTableUpdate(*": Flag = True
1023:            Case sTxt Like "*Workbook_SheetSelectionChange(*": Flag = True
1024:            Case sTxt Like "*Workbook_Sync(*": Flag = True
1025:            Case sTxt Like "*Workbook_WindowActivate(*": Flag = True
1026:            Case sTxt Like "*Workbook_WindowDeactivate(*": Flag = True
1027:            Case sTxt Like "*Workbook_WindowResize(*": Flag = True
1028:        End Select
1029:    End If
1030:    WorkBookAndSheetsEvents = Flag
1031: End Function

