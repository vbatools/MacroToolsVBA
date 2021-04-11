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
23:    Set m_clsAnchors = New CAnchors
24:    Set m_clsAnchors.objParent = Me
25:    ' restrict minimum size of userform
26:    m_clsAnchors.MinimumWidth = 727
27:    m_clsAnchors.MinimumHeight = 405.75
28:    With m_clsAnchors
29:        .funAnchor("cmbMain").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft
30:        .funAnchor("ListCode").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft Or enumAnchorStyleBottom
31:        .funAnchor("lbAnaliz").AnchorStyle = enumAnchorStyleTop Or enumAnchorStyleRight
32:        .funAnchor("btnCopyCode").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
33:        .funAnchor("lbCancel").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
34:        .funAnchor("lbLoad").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
35:        .funAnchor("lbLoadTxtFile").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
36:    End With
37: End Sub
    Private Sub UserForm_Activate()
39:    Dim vbProj      As VBIDE.VBProject
40:
41:    If Workbooks.Count = 0 Then
42:        Unload Me
43:        Call MsgBox("No open ones" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
44:        Exit Sub
45:    End If
46:    With Me.cmbMain
47:        .Clear
48:        On Error Resume Next
49:        For Each vbProj In Application.VBE.VBProjects
50:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
51:        Next
52:        On Error GoTo 0
53:        .Value = ActiveWorkbook.Name
54:    End With
55: End Sub
    Private Sub UserForm_Terminate()
57:    Set m_clsAnchors = Nothing
58: End Sub
    Private Sub btnCancel_Click()
60:    Unload Me
61: End Sub
    Private Sub lbCancel_Click()
63:    Unload Me
64: End Sub
    Private Sub lbLoadTxtFile_Click()
66:    Dim strVar      As String
67:    Dim strFileName As String
68:
69:    If cmbMain.Value = vbNullString Then Exit Sub
70:
71:    strVar = AddListImmediate()
72:    If strVar = vbNullString Then Exit Sub
73:
74:    With Workbooks(cmbMain.Value)
75:        strFileName = .Path & Application.PathSeparator & sGetFileName(.FullName) & ".txt"
76:    End With
77:
78:    If SaveTXTfile(strFileName, strVar) Then
79:        Call MsgBox("The data is copied to a txt file!", vbInformation, "Copying data:")
80:    Else
81:        Call MsgBox("Couldn't copy data to txt file!", vbCritical, "Copying data:")
82:    End If
83: End Sub
    Private Function SaveTXTfile(ByVal sFileName As String, ByVal Txt As String) As Boolean
85:    Dim FSO         As Object
86:    Dim ts          As Object
87:    On Error Resume Next: Err.Clear
88:    Set FSO = CreateObject("scripting.filesystemobject")
89:    Set ts = FSO.CreateTextFile(sFileName, True)
90:    ts.Write Txt: ts.Close
91:    SaveTXTfile = Err = 0
92:    Set ts = Nothing: Set FSO = Nothing
93: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : lbLoad_Click - выгрузка в окно immediate
'* Created    : 28-01-2020 14:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub lbLoad_Click()
103:    Debug.Print AddListImmediate()
104:    Call MsgBox("The data is copied to the Immediate window!", vbInformation, "Copying data:")
105: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : btnCopyCode_Click - сохранение в буффер обмена
'* Created    : 28-01-2020 14:35
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub btnCopyCode_Click()
114:    Dim strVar      As String
115:    strVar = AddListImmediate()
116:    If strVar = vbNullString Then Exit Sub
117:    Call C_PublicFunctions.SetTextIntoClipboard(strVar)
118:
119:    Call MsgBox("The data is copied to the clipboard!", vbInformation, "Copying data:")
120: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddListImmediate - создание текста не используемых переменых
'* Created    : 28-01-2020 14:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function AddListImmediate() As String
129:    Dim strData     As String
130:    Dim i           As Long
131:    Dim Max1        As Byte
132:    Dim Max2        As Byte
133:    Dim Max3        As Byte
134:    Dim Max4        As Byte
135:    Dim Max5        As Byte
136:
137:    Max1 = 0: Max2 = 0: Max3 = 0: Max4 = 0: Max5 = 0
138:
139:    With ListCode
140:        'поиск текста с максимальным кол-вом символов
141:        For i = 0 To .ListCount - 1
142:            If VBA.Len(.List(i, 1)) > Max1 Then Max1 = VBA.Len(.List(i, 1))
143:            If VBA.Len(.List(i, 2)) > Max2 Then Max2 = VBA.Len(.List(i, 2))
144:            If VBA.Len(.List(i, 3)) > Max3 Then Max3 = VBA.Len(.List(i, 3))
145:            If VBA.Len(.List(i, 4)) > Max4 Then Max4 = VBA.Len(.List(i, 4))
146:            If VBA.Len(.List(i, 5)) > Max5 Then Max5 = VBA.Len(.List(i, 5))
147:        Next i
148:        'вормирование текста
149:        For i = 0 To .ListCount - 1
150:            strData = strData & _
                        .List(i, 1) & VBA.String$(Max1 - VBA.Len(.List(i, 1)), " ") & vbTab & _
                        .List(i, 2) & VBA.String$(Max2 - VBA.Len(.List(i, 2)), " ") & vbTab & _
                        .List(i, 3) & VBA.String$(Max3 - VBA.Len(.List(i, 3)), " ") & vbTab & _
                        .List(i, 4) & VBA.String$(Max4 - VBA.Len(.List(i, 4)), " ") & vbTab & _
                        .List(i, 5) & VBA.String$(Max5 - VBA.Len(.List(i, 5)), " ") & vbNewLine
156:        Next i
157:    End With
158:    AddListImmediate = strData
159: End Function
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
172:    Dim i           As Long
173:    Dim WB          As Workbook
174:    Dim VBC         As VBIDE.VBComponent
175:
176:    'On Error GoTo ErrorHandler
177:
178:    If cmbMain.Value = vbNullString Then Exit Sub
179:    Set WB = Workbooks(cmbMain.Value)
180:    For i = 0 To ListCode.ListCount
181:        If ListCode.Selected(i) = True Then
182:            Set VBC = WB.VBProject.VBComponents(ListCode.List(i, 1))
183:            If VBC.Type = vbext_ct_MSForm Then
184:                VBC.CodeModule.CodePane.Show
185:            Else
186:                VBC.Activate
187:            End If
188:            Exit Sub
189:        End If
190:    Next i
191:    Exit Sub
ErrorHandler:
193:    Unload Me
194:    Select Case Err.Number
        Case Else:
196:            Call MsgBox("Error! in VariableUnUsed. ListCode_DblClick" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbExclamation, "Error:")
197:            Call WriteErrorLog("VariableUnUsed.ListCode_DblClick")
198:    End Select
199:    Err.Clear
200: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : cmbMain_Change - выбор файла Exsel
'* Created    : 28-01-2020 14:49
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub lbAnaliz_Click()
209:    If cmbMain.Value = vbNullString Then Exit Sub
210:    ListCode.Clear
211:    Call MainSubAddUnUsed
212: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : MainSubAddUnUsed - главная процедура запуска поиска не используемых перемменых
'* Created    : 28-01-2020 14:50
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub MainSubAddUnUsed()
221:    Dim WBName      As String
222:    WBName = cmbMain.Value
223:    If WBName <> vbNullString Then
224:        Dim objVBP  As VBIDE.VBProject
225:        Set objVBP = Workbooks(WBName).VBProject
226:        If objVBP.Protection <> vbext_pp_none Then
227:            ListCode.Clear
228:            Call MsgBox("VBA project in the book -" & WBName & "password protected!" & vbCrLf & "Remove the password!", vbCritical, "Error:")
229:            Exit Sub
230:        End If
231:
232:        Dim dTime   As Date
233:        dTime = Now()
234:
235:        'процедуры создания словарей
236:        Call CreateGlobalVariableCollection(WBName)
237:        Debug.Print "All global variables:" & Format(Now() - dTime, "Long Time") & "completed!"
238:        Call ProcessUnusedVariable(WBName)
239:        Debug.Print "Procedure variables:" & Format(Now() - dTime, "Long Time") & "completed!"
240:        Call FillUnusedGlobalVariables(WBName)
241:        Debug.Print "Global variables unused:" & Format(Now() - dTime, "Long Time") & "completed!"
242:        Call AddListUnUsedEnumAndType
243:        Debug.Print "Types and enumerations:" & Format(Now() - dTime, "Long Time") & "completed!"
244:        Call AddSubAndFuncListUnUsed(WBName)
245:        Debug.Print "Procedures and functions:" & Format(Now() - dTime, "Long Time") & "completed!"
246:    End If
247:
248:    'удаляю все словари
249:    Set dicCollGlobalVariables = Nothing
250:    Set dicCollEnumType = Nothing
251:    Set dicCollAsTypeVariable = Nothing
252:    Set dicCollSubsAndFunctions = Nothing
253:
254: End Sub
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
267:
268:    Dim objVBP      As VBIDE.VBProject
269:    Dim vbComp      As VBIDE.VBComponent
270:    Dim ProcKind    As VBIDE.vbext_ProcKind
271:    Dim CodeMod     As CodeModule
272:    Dim intLine     As Integer
273:    Dim strFunctionBody As String
274:    Dim strFunctionTypeAs As String
275:    Dim strProcName As String
276:    Dim strFerstStringSub As String
277:    Dim g_wbkVBAExcel As Workbook
278:
279:    'On Error GoTo ErrorHandler
280:
281:    Set g_wbkVBAExcel = Workbooks(WBName)
282:    Set objVBP = g_wbkVBAExcel.VBProject
283:
284:    Set dicCollSubsAndFunctions = Nothing
285:    Set dicCollSubsAndFunctions = New Scripting.Dictionary
286:    For Each vbComp In objVBP.VBComponents
287:        If vbComp.Type <> vbext_ct_ClassModule Then
288:            Set CodeMod = vbComp.CodeModule
289:            For intLine = 1 To CodeMod.CountOfLines
290:                strProcName = CodeMod.ProcOfLine(intLine, ProcKind)
291:                If strProcName <> vbNullString Then
292:                    'создание словаря переменных процедур и функций
293:                    strFunctionBody = CodeMod.Lines(intLine, CodeMod.ProcCountLines(strProcName, ProcKind))
294:                    strFerstStringSub = CodeMod.Lines(CodeMod.ProcBodyLine(strProcName, ProcKind), 1)
295:
296:                    If strFerstStringSub Like "*) As *" And strFerstStringSub Like "Function *" Then
297:                        Dim arrAs As Variant
298:                        arrAs = Split(strFerstStringSub, ") As ")
299:                        strFunctionTypeAs = arrAs(UBound(arrAs))
300:                        arrAs = Empty
301:                    Else
302:                        strFunctionTypeAs = "-"
303:                    End If
304:
305:                    intLine = intLine + CodeMod.ProcCountLines(strProcName, ProcKind) - 1
306:                    Call FillUnusedLocalVariables(strFunctionBody, strProcName, vbComp.Name, ProcKind)
307:
308:                    'создание словаря названий процедур и функций
309:                    If dicCollSubsAndFunctions.Exists(vbComp.Name & "." & strProcName) = False Then
310:                        Dim sTypeSubFun As String
311:
312:                        If strFunctionBody Like "*Private *" Then
313:                            sTypeSubFun = "Private"
314:                        Else
315:                            sTypeSubFun = "Public"
316:                        End If
317:                        'не загружаю процедуры рибон понели, события листов и книги, UserForm в формах
318:                        If (Not strFunctionBody Like "*As IRibbonControl*") And _
                                    (Not WorkBookAndSheetsEvents(strFunctionBody, vbComp.Type)) And _
                                    (Not (strFunctionBody Like "* UserForm_*" And vbComp.Type = vbext_ct_MSForm)) Then
321:                            Call dicCollSubsAndFunctions.Add(vbComp.Name & "." & _
                                        strProcName, _
                                        sTypeSubFun & "." & _
                                        I_StatisticVBAProj.TypeProcedyre(strFunctionBody) & "." & _
                                        byTypeProc(strFerstStringSub) & "." & _
                                        strFunctionTypeAs)
327:                        End If
328:                    End If
329:                    strFerstStringSub = vbNullString
330:                    strFunctionBody = vbNullString
331:                    strProcName = vbNullString
332:                    strFunctionTypeAs = vbNullString
333:                End If
334:            Next
335:        End If
336:    Next
337:    Set objVBP = Nothing
338:    Set vbComp = Nothing
339:    Set CodeMod = Nothing
340:
341:    Exit Sub
ErrorHandler:
343:    Unload Me
344:    Select Case Err.Number
        Case Else:
346:            Call MsgBox("Error in VariableUnUsed.ProcessUnusedVariable" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbExclamation, "Error:")
347:            Call WriteErrorLog("VariableUnUsed.ProcessUnusedVariable")
348:    End Select
349:    Err.Clear
350: End Sub
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
366:
367:    Dim strVrName   As String
368:    Dim strArrLine() As String
369:    Dim strLine     As String
370:    Dim strArrVariables() As String
371:    Dim intCounterLine As Integer
372:    Dim intCounterVariable As Integer
373:    Dim strDeclaration As String
374:
375:    'On Error GoTo ErrorHandler
376:
377:    strVrName = vbNullString
378:
379:    FunctionBody = Replace(FunctionBody, " _" & vbCrLf, " ")
380:    strArrLine = Split(FunctionBody, vbCrLf)
381:
382:    For intCounterLine = LBound(strArrLine) To UBound(strArrLine)
383:        strLine = Trim(strArrLine(intCounterLine))
384:        'переменные
385:        If strLine Like "Dim *" Or strLine Like "Static *" Or strLine Like "Const *" Then
386:            strArrVariables = Split(RemoveEnclosedStringAndComments(strLine), ",")
387:            For intCounterVariable = LBound(strArrVariables) To UBound(strArrVariables)
388:                strDeclaration = Trim(strArrVariables(intCounterVariable))
389:                strVrName = Mid(strArrVariables(intCounterVariable), InStr(1, strArrVariables(intCounterVariable), " "), 100)
390:                strVrName = Replace(strVrName, ",", vbNullString)
391:                If strVrName Like "*()" Then
392:                    strVrName = Left(strVrName, Len(strVrName) - 2)
393:                End If
394:                If IsVariableUsed(strVrName, strArrLine) = False Then
395:                    With ListCode
396:                        Dim lListRow As Long
397:                        Dim arrVar As Variant
398:                        Dim strVar As String
399:                        lListRow = .ListCount
400:                        .AddItem lListRow + 1
401:                        .List(lListRow, 1) = VBA.Trim$(ModuleName)
402:                        If strDeclaration Like "Const *" Then
403:                            .List(lListRow, 2) = "Const"
404:                        Else
405:                            .List(lListRow, 2) = "Dim"
406:                        End If
407:                        .List(lListRow, 3) = VBA.Trim$(FunctionName)
408:                        arrVar = Split(strVrName, " As ")
409:                        .List(lListRow, 4) = VBA.Trim$(VBA.Replace(arrVar(0), "=", vbNullString))
410:                        If UBound(arrVar) = 0 Then
411:                            strVar = "Variant"
412:                        Else
413:                            If strDeclaration Like "Const *" Then
414:                                strVar = VBA.Trim$(VBA.Left$(arrVar(1), VBA.InStr(1, arrVar(1), "=") - 1))
415:                            Else
416:                                strVar = VBA.Trim$(VBA.Replace(arrVar(1), "=", vbNullString))
417:                            End If
418:                        End If
419:                        .List(lListRow, 5) = strVar
420:                    End With
421:                End If
422:            Next intCounterVariable
423:        End If
424:        'типы переменых переменные дополнение словаря типов переменных
425:        If strLine Like "* As *" Then
426:            Call AddDictionariAsType(strLine, ModuleName)
427:        End If
428:    Next intCounterLine
429:    Exit Sub
ErrorHandler:
431:    Unload Me
432:    Select Case Err.Number
        Case Else:
434:            Call MsgBox("Error in VariableUnUsed.FillUnusedLocalVariables" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbExclamation, "Error:")
435:            Call WriteErrorLog("VariableUnUsed.FillUnusedLocalVariables")
436:    End Select
437:    Err.Clear
438: End Sub
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
452:    Dim intLoop     As Integer
453:    Dim strLine     As String
454:
455:    IsVariableUsed = False
456:
457:    ' Format Variablename
458:    strVrName = C_PublicFunctions.TrimSpace(strVrName)
459:    strVrName = Replace(strVrName, "(", " ")
460:    strVrName = Replace(strVrName, "%", vbNullString)     'Integer
461:    strVrName = Replace(strVrName, "&", vbNullString)     'Long
462:    strVrName = Replace(strVrName, "$", vbNullString)     'String
463:    strVrName = Replace(strVrName, "!", vbNullString)     'Single
464:    strVrName = Replace(strVrName, "#", vbNullString)     'Double
465:    strVrName = Replace(strVrName, "@", vbNullString)     'Currency
466:    If strVrName <> vbNullString Then strVrName = Split(strVrName, " ")(0)
467:    For intLoop = 0 To UBound(strArrofLine)
468:        strLine = C_PublicFunctions.TrimSpace(strArrofLine(intLoop))
469:        strLine = RemoveEnclosedStringAndComments(strLine)
470:
471:        If Not StrigLikeWord(strLine) Then
472:            strLine = Replace(strLine, "(", " ")
473:            strLine = Replace(strLine, ".", " ")
474:            strLine = Replace(strLine, ")", " ")
475:            strLine = Replace(strLine, ",", " ")
476:            strLine = " " & strLine & " "
477:
478:            If strLine Like "* " & strVrName & " *" Then
479:                IsVariableUsed = True
480:                Exit Function
481:            End If
482:        End If
483:    Next
484:
485: End Function
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
498:
499:    Dim objVBP      As VBIDE.VBProject
500:    Dim vbComp      As VBIDE.VBComponent
501:    Dim VBCompTemp  As VBIDE.VBComponent
502:    Dim objCodeMod  As CodeModule
503:    Dim strArrLine() As String
504:    Dim strContent  As String
505:    Dim strArrWord() As String
506:    Dim strWord     As String
507:    Dim intWordCount As Integer
508:    Dim strLine     As String
509:    Dim intCounter  As Integer
510:    Dim collUsedGlobal As New Dictionary
511:    Dim g_wbkVBAExcel As Workbook
512:
513:    'On Error GoTo ErrorHandler
514:
515:    Set g_wbkVBAExcel = Workbooks(WBName)
516:    Set objVBP = g_wbkVBAExcel.VBProject
517:    'удаление из словаря используемых переменных
518:    For Each vbComp In objVBP.VBComponents
519:        If vbComp.Type <> vbext_ct_ClassModule Then     '!!!
520:            Set objCodeMod = vbComp.CodeModule
521:            If objCodeMod.CountOfLines > 0 Then
522:                strContent = objCodeMod.Lines(1, objCodeMod.CountOfLines)
523:                strContent = Replace(strContent, " _" & vbCrLf, " ")
524:                strArrLine = Split(strContent, vbCrLf)
525:
526:                For intCounter = LBound(strArrLine) To UBound(strArrLine)
527:                    strLine = RemoveEnclosedStringAndComments(strArrLine(intCounter))
528:                    'глобальная переменная с присвоением, проверяем вторую часть
529:                    If StrigLikeWord(strLine) And strLine Like "*=*" Then strLine = Split(strLine, "=")(1)
530:                    If Not StrigLikeWord(strLine) Then
531:                        strLine = Replace(strLine, ",", " ")
532:                        strLine = Replace(strLine, ")", " ")
533:                        strLine = Replace(strLine, "(", " ")
534:                        strArrWord = Split(Replace(strLine, ".", " "), " ")
535:                        For intWordCount = LBound(strArrWord) To UBound(strArrWord)
536:                            If strArrWord(intWordCount) <> "" Then
537:                                strWord = Split(strArrWord(intWordCount), "(")(0)
538:                                For Each VBCompTemp In objVBP.VBComponents
539:                                    If dicCollGlobalVariables.Exists(VBCompTemp.Name & "." & strWord) Then
540:                                        Call collUsedGlobal.Add(VBCompTemp.Name & "." & strWord, dicCollGlobalVariables.Item(VBCompTemp.Name & "." & strWord))
541:                                        dicCollGlobalVariables.Remove (VBCompTemp.Name & "." & strWord)
542:                                        Exit For
543:                                    End If
544:                                Next VBCompTemp
545:                            End If
546:                        Next intWordCount
547:                    End If
548:                Next intCounter
549:            End If
550:        End If
551:    Next vbComp
552:    'глабальные переменные которые не используются
553:    For intCounter = 0 To dicCollGlobalVariables.Count - 1
554:        With ListCode
555:            Dim lListRow As Long
556:            Dim strarr As Variant
557:            Dim strarr1 As Variant
558:            Dim strarr2 As Variant
559:            Dim strVal As String
560:            Dim strVariable As String
561:            lListRow = .ListCount
562:            .AddItem lListRow
563:            .List(lListRow, 1) = VBA.Trim$(Split(dicCollGlobalVariables.Keys(intCounter), ".")(0))
564:            .List(lListRow, 3) = "-"
565:
566:            strVariable = VBA.Trim$(dicCollGlobalVariables.Items(intCounter))
567:            strarr = Split(strVariable, "As")
568:            strarr2 = Split(Trim$(strarr(0)), " ")
569:            If UBound(strarr) = 0 Then
570:                .List(lListRow, 4) = strarr2(UBound(strarr2))
571:                .List(lListRow, 5) = "Variant"
572:                .List(lListRow, 2) = Replace(strarr(0), " " & .List(lListRow, 4), vbNullString)
573:            Else
574:                strarr1 = Split(Trim$(strarr(1)), " ")
575:                If UBound(strarr1) = 0 Then
576:                    strVal = strarr(1)
577:                Else
578:                    strVal = strarr1(0)
579:                End If
580:                .List(lListRow, 4) = strarr2(UBound(strarr2))
581:                .List(lListRow, 5) = VBA.Trim$(strVal)
582:                .List(lListRow, 2) = Replace(strarr(0), " " & .List(lListRow, 4), vbNullString)
583:            End If
584:        End With
585:    Next
586:
587:    Exit Sub
ErrorHandler:
589:    Unload Me
590:    Select Case Err.Number
        Case Else:
592:            Call MsgBox("Error in VariableUnUsed.FillUnusedGlobalVariables" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbExclamation, "Error:")
593:            Call WriteErrorLog("VariableUnUsed.FillUnusedGlobalVariables")
594:    End Select
595:    Err.Clear
596: End Sub
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
609:
610:    Dim vbComp      As VBIDE.VBComponent
611:    Dim vbProj      As VBIDE.VBProject
612:    Dim CodeMod     As VBIDE.CodeModule
613:    Dim strContent  As String
614:    Dim strArrLine() As String
615:    Dim strArrVariables() As String
616:    Dim intLineCounter As Integer
617:    Dim intVarCounter As Integer
618:    Dim strVarName  As String
619:    Dim strVarNameEnumType As String
620:    Dim strVarDeclaration As String
621:    Dim strLine     As String
622:    Dim g_wbkVBAExcel As Workbook
623:
624:    'On Error GoTo ErrorHandler
625:
626:    Set g_wbkVBAExcel = Workbooks(WBName)
627:    Set vbProj = g_wbkVBAExcel.VBProject
628:
629:    Set dicCollGlobalVariables = Nothing
630:    Set dicCollGlobalVariables = New Scripting.Dictionary
631:
632:    Set dicCollEnumType = Nothing
633:    Set dicCollEnumType = New Scripting.Dictionary
634:
635:    Set dicCollAsTypeVariable = Nothing
636:    Set dicCollAsTypeVariable = New Scripting.Dictionary
637:
638:    For Each vbComp In vbProj.VBComponents
639:        If vbComp.Type <> vbext_ct_ClassModule Then
640:            Set CodeMod = vbComp.CodeModule
641:            If CodeMod.CountOfDeclarationLines > 0 Then
642:                strContent = CodeMod.Lines(1, CodeMod.CountOfDeclarationLines)
643:                strContent = Replace(strContent, " _" & vbCrLf, " ")
644:
645:                strArrLine = Split(strContent, vbCrLf)
646:                For intLineCounter = LBound(strArrLine) To UBound(strArrLine)
647:
648:                    strLine = RemoveEnclosedStringAndComments(strArrLine(intLineCounter))
649:                    'собираю типы глабальных переменных
650:                    If strLine Like "* As *" Then
651:                        Call AddDictionariAsType(strLine, CodeMod.Name)
652:                    End If
653:                    If StrigLikeWord(strLine) Then
654:
655:                        strLine = Replace(strLine, "=", vbNullString)
656:
657:                        strArrVariables = Split(strLine, ",")
658:                        For intVarCounter = LBound(strArrVariables) To UBound(strArrVariables)
659:                            strVarDeclaration = strArrVariables(intVarCounter)
660:                            strVarName = Trim(Split(strArrVariables(intVarCounter), " ")(1))
661:                            If strVarName <> vbNullString Then
662:                                If strVarName = "Const" Then
663:                                    strVarName = Trim(Split(strArrVariables(intVarCounter), " ")(2))
664:                                End If
665:                                'словарь для Enum и Type
666:                                If strVarName = "Enum" Or strVarName = "Type" Then
667:                                    strVarNameEnumType = Trim(Split(strArrVariables(intVarCounter), " ")(2))
668:                                    If dicCollEnumType.Exists(vbComp.Name & "." & strVarNameEnumType) = False Then Call dicCollEnumType.Add(vbComp.Name & "." & strVarNameEnumType, strVarDeclaration)
669:                                End If
670:                                strVarName = Split(strVarName, "(")(0)
671:                                strVarName = Replace(strVarName, ",", vbNullString)
672:                                'словарь для переменных
673:                                If Not (strVarDeclaration Like "Enum*" Or strVarDeclaration Like "Type*") Then
674:                                    If dicCollGlobalVariables.Exists(vbComp.Name & "." & strVarName) = False Then Call dicCollGlobalVariables.Add(vbComp.Name & "." & strVarName, strVarDeclaration)
675:                                End If
676:                            End If
677:                        Next intVarCounter
678:
679:                    End If
680:                Next intLineCounter
681:            End If
682:        End If
683:    Next vbComp
684:
685:    Exit Sub
ErrorHandler:
687:    Unload Me
688:    Select Case Err.Number
        Case Else:
690:            Call MsgBox("Error in Variable UnUsed.CreateGlobalVariablesCollection" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbExclamation, "Error:")
691:            Call WriteErrorLog("VariableUnUsed.CreateGlobalVariableCollection")
692:    End Select
693:    Err.Clear
694: End Sub
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
708:    Dim strVar      As String
709:    Dim strVar1     As String
710:    Dim arrVar      As Variant
711:
712:    Dim i           As Long
713:
714:    arrVar = Split(strLine, " As ")
715:
716:    For i = 1 To UBound(arrVar)
717:
718:        strVar = VBA.Trim$(arrVar(i))
719:
720:        If strVar Like "*, *" Then
721:            strVar = VBA.Trim$(Split(strVar, ",")(0))
722:        End If
723:        If strVar Like "* *" Then
724:            strVar = VBA.Trim$(Split(strVar, " ")(0))
725:        End If
726:
727:        If strVar Like "*)" Then
728:            strVar = VBA.Trim$(Replace(strVar, ")", vbNullString))
729:        End If
730:        If strVar Like "*(" Then
731:            strVar = VBA.Trim$(Replace(strVar, "(", vbNullString))
732:        End If
733:        strVar = VBA.Trim$(Replace(strVar, Chr(34), vbNullString))
734:        strVar = VBA.Trim$(Replace(strVar, "*", vbNullString))
735:        strVar = VBA.Trim$(Replace(strVar, "#", "-"))
736:
737:        If Not strVar Like "-*" And strVar <> vbNullString Then
738:            strVar1 = strVar
739:            If strLine Like "Private *" Then
740:                strVar = VBA.Trim$(NameModule) & "." & strVar
741:            Else
742:                strVar = strVar
743:            End If
744:            If dicCollAsTypeVariable.Exists(strVar1) = False Then Call dicCollAsTypeVariable.Add(strVar1, strVar)
745:        End If
746:    Next i
747: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddListUnUsedEnumAndType - вывод не используемых типов и пречеслений
'* Created    : 28-01-2020 11:33
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub AddListUnUsedEnumAndType()
756:    Dim intCounter  As Integer
757:    Dim strKey      As String
758:    Dim strKey1     As String
759:    Dim lListRow    As Long
760:
761:    For intCounter = 0 To dicCollEnumType.Count - 1
762:        strKey = dicCollEnumType.Keys()(intCounter)
763:        strKey1 = Split(strKey, ".")(1)
764:        strKey = Split(strKey, ".")(0)
765:        If dicCollAsTypeVariable.Exists(strKey1) = False Then
766:            With ListCode
767:                lListRow = .ListCount
768:                .AddItem lListRow + 1
769:                .List(lListRow, 1) = strKey
770:                .List(lListRow, 2) = VBA.Trim$(Replace(dicCollEnumType.Items()(intCounter), strKey1, vbNullString))
771:                .List(lListRow, 3) = "-"
772:                .List(lListRow, 4) = strKey1
773:                .List(lListRow, 5) = "-"
774:                '                If .List(lListRow, 4) Like "*Enum*" Then
775:                '                    .List(lListRow, 4) = "перечисление"
776:                '                Else
777:                '                    .List(lListRow, 4) = "пользовательский тип"
778:                '                End If
779:            End With
780:        End If
781:    Next intCounter
782: End Sub
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
     Private Function StrigLikeWord(ByVal stxt As String) As Boolean
795:    Dim Flag        As Boolean
796:    Flag = False
797:    Select Case True
        Case stxt Like "* declare *": Flag = False
799:        Case stxt Like "* Declare *": Flag = False
800:
801:
802:        Case stxt Like vbNullString: Flag = True
803:        Case stxt Like "Public *": Flag = True
804:        Case stxt Like "Private *": Flag = True
805:        Case stxt Like "Global *": Flag = True
806:        Case stxt Like "Const *": Flag = True
807:        Case stxt Like "Set*= Nothing": Flag = True
808:        Case stxt Like "Dim *": Flag = True
809:        Case stxt Like "*Enum *": Flag = True
810:        Case stxt Like "*Type *": Flag = True
811:    End Select
812:    StrigLikeWord = Flag
813: End Function
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
826:
827:    Dim strArrParts() As String
828:    Dim intCntPart  As Integer
829:
830:    If strLine = vbNullString Then
831:        RemoveEnclosedStringAndComments = vbNullString
832:        Exit Function
833:    End If
834:
835:    strLine = Replace(strLine, vbTab, vbNullString)
836:    strArrParts = Split(strLine, """")
837:    strLine = vbNullString
838:
839:    For intCntPart = LBound(strArrParts) To UBound(strArrParts)
840:        If intCntPart Mod 2 = 0 Then
841:            strLine = strLine & " " & Trim(strArrParts(intCntPart))
842:        End If
843:    Next intCntPart
844:
845:    If strLine <> vbNullString Then
846:        strLine = Split(strLine, "'")(0)
847:    End If
848:    strLine = C_PublicFunctions.TrimSpace(strLine)
849:    RemoveEnclosedStringAndComments = strLine
850:
851: End Function

     Private Sub AddSubAndFuncListUnUsed(ByVal WBName As String)
854:    Dim vbProj      As VBIDE.VBProject
855:    Dim vbComp      As VBIDE.VBComponent
856:    Dim CodeMod     As VBIDE.CodeModule
857:    Dim intCounter  As Integer
858:    Dim strKey      As String
859:    Dim strKeyNameMod As String
860:    Dim strKeyNameSub As String
861:    Dim strItemTypePub As String
862:    Dim strItemTypeSub As String
863:    Dim strItemTypeProc As String
864:    Dim strFunctionBody As String
865:    Dim strModuleBody As String
866:
867:    Set vbProj = Workbooks(WBName).VBProject
868:
869:    For intCounter = dicCollSubsAndFunctions.Count - 1 To 0 Step -1
870:        strKey = dicCollSubsAndFunctions.Keys()(intCounter)
871:        strKeyNameMod = Split(dicCollSubsAndFunctions.Keys()(intCounter), ".")(0)
872:        strKeyNameSub = Split(dicCollSubsAndFunctions.Keys()(intCounter), ".")(1)
873:        strItemTypePub = Split(dicCollSubsAndFunctions.Items()(intCounter), ".")(0)
874:        strItemTypeSub = Split(dicCollSubsAndFunctions.Items()(intCounter), ".")(1)
875:        strItemTypeProc = CByte(Split(dicCollSubsAndFunctions.Items()(intCounter), ".")(2))
876:
877:        'если Private  то ищем в именно в этом модуле в других не надо
878:        If strItemTypePub = "Private" Then
879:            Set vbComp = vbProj.VBComponents(strKeyNameMod)
880:            Set CodeMod = vbComp.CodeModule
881:            With CodeMod
882:                strFunctionBody = .Lines(.ProcStartLine(strKeyNameSub, strItemTypeProc), .ProcCountLines(strKeyNameSub, strItemTypeProc))
883:                strModuleBody = .Lines(1, .CountOfLines)
884:            End With
885:            strModuleBody = VBA.Replace(strModuleBody, strFunctionBody, vbNullString)
886:            'если нашли то удаляем из словаря
887:            If VBA.InStr(1, strModuleBody, strKeyNameSub, vbTextCompare) <> 0 Then
888:                dicCollSubsAndFunctions.Remove (strKey)
889:                'если форма
890:            ElseIf vbComp.Type = vbext_ct_MSForm Then
891:                If strKeyNameSub Like "*_*" Then
892:                    'если элемент формы
893:                    Dim strNameItemDesiner As String
894:                    Dim arrVar As Variant
895:                    arrVar = Split(strKeyNameSub, "_")
896:                    strNameItemDesiner = VBA.Replace(strKeyNameSub, "_" & arrVar(UBound(arrVar)), vbNullString)
897:                    On Error Resume Next
898:                    strNameItemDesiner = vbComp.Designer.Item(strNameItemDesiner).Name
899:                    If strNameItemDesiner <> vbNullString Then
900:                        dicCollSubsAndFunctions.Remove (strKey)
901:                    End If
902:                    strNameItemDesiner = vbNullString
903:                    On Error GoTo 0
904:                End If
905:            End If
906:            'если Public
907:        Else
908:            For Each vbComp In vbProj.VBComponents
909:                Set CodeMod = vbComp.CodeModule
910:                With CodeMod
911:                    If .CountOfLines <> 0 Then
912:                        'если ищем в родительском модуле
913:                        If .Name = strKeyNameMod Then
914:                            strFunctionBody = .Lines(.ProcStartLine(strKeyNameSub, strItemTypeProc), .ProcCountLines(strKeyNameSub, strItemTypeProc))
915:                            strModuleBody = .Lines(1, .CountOfLines)
916:                            strModuleBody = VBA.Replace(strModuleBody, strFunctionBody, vbNullString)
917:                        Else
918:                            strModuleBody = .Lines(1, .CountOfLines)
919:                        End If
920:
921:                        'если нашли то удаляем из словаря и выходим из цикла
922:                        If VBA.InStr(1, strModuleBody, strKeyNameSub, vbTextCompare) <> 0 Then
923:                            dicCollSubsAndFunctions.Remove (strKey)
924:                            Exit For
925:                        End If
926:                    End If
927:                End With
928:            Next vbComp
929:        End If
930:    Next intCounter
931:    For intCounter = 0 To dicCollSubsAndFunctions.Count - 1
932:        Dim lListRow As Integer
933:        With ListCode
934:            lListRow = .ListCount
935:            .AddItem lListRow
936:            .List(lListRow, 1) = Split(dicCollSubsAndFunctions.Keys()(intCounter), ".")(0)
937:            .List(lListRow, 2) = Split(dicCollSubsAndFunctions.Items()(intCounter), ".")(0) & " " & Split(dicCollSubsAndFunctions.Items()(intCounter), ".")(1)
938:            .List(lListRow, 3) = Split(dicCollSubsAndFunctions.Keys()(intCounter), ".")(1)
939:            .List(lListRow, 4) = "-"
940:            .List(lListRow, 5) = Split(dicCollSubsAndFunctions.Items()(intCounter), ".")(3)
941:        End With
942:    Next intCounter
943: End Sub
     Private Function byTypeProc(ByVal stxt As String) As Byte
945:    Select Case True
        Case stxt Like "*Property Get*": byTypeProc = 3
947:        Case stxt Like "*Property Set*": byTypeProc = 2
948:        Case stxt Like "*Property Let*": byTypeProc = 1
949:        Case Else: byTypeProc = 0
950:    End Select
951: End Function

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
Private Function WorkBookAndSheetsEvents(ByVal stxt As String, ByVal TypeModule As VBIDE.vbext_ComponentType) As Boolean
965:    Dim Flag        As Boolean
966:    Flag = False
967:    'только для модулей листов и книг
968:    If TypeModule = vbext_ct_Document Then
969:        Select Case True
            Case stxt Like "*Worksheet_Activate(*": Flag = True
971:            Case stxt Like "*Worksheet_BeforeDoubleClick(*": Flag = True
972:            Case stxt Like "*Worksheet_BeforeRightClick(*": Flag = True
973:            Case stxt Like "*Worksheet_Calculate(*": Flag = True
974:            Case stxt Like "*Worksheet_Change(*": Flag = True
975:            Case stxt Like "*Worksheet_Deactivate(*": Flag = True
976:            Case stxt Like "*Worksheet_FollowHyperlink(*": Flag = True
977:            Case stxt Like "*Worksheet_PivotTableAfterValueChange(*": Flag = True
978:            Case stxt Like "*Worksheet_PivotTableBeforeAllocateChanges(*": Flag = True
979:            Case stxt Like "*Worksheet_PivotTableBeforeCommitChanges(*": Flag = True
980:            Case stxt Like "*Worksheet_PivotTableBeforeDiscardChanges(*": Flag = True
981:            Case stxt Like "*Worksheet_PivotTableChangeSync(*": Flag = True
982:            Case stxt Like "*Worksheet_PivotTableUpdate(*": Flag = True
983:            Case stxt Like "*Worksheet_SelectionChange(*": Flag = True
984:            Case stxt Like "*Chart_Activate(*": Flag = True
985:            Case stxt Like "*Chart_BeforeDoubleClick(*": Flag = True
986:            Case stxt Like "*Chart_BeforeRightClick(*": Flag = True
987:            Case stxt Like "*Chart_Calculate(*": Flag = True
988:            Case stxt Like "*Chart_Deactivate(*": Flag = True
989:            Case stxt Like "*Chart_MouseDown(*": Flag = True
990:            Case stxt Like "*Chart_MouseMove(*": Flag = True
991:            Case stxt Like "*Chart_MouseUp(*": Flag = True
992:            Case stxt Like "*Chart_Resize(*": Flag = True
993:            Case stxt Like "*Chart_SeriesChange(*": Flag = True
994:            Case stxt Like "*Workbook_Activate(*": Flag = True
995:            Case stxt Like "*Workbook_AddinInstall(*": Flag = True
996:            Case stxt Like "*Workbook_AddinUninstall(*": Flag = True
997:            Case stxt Like "*Workbook_AfterSave(*": Flag = True
998:            Case stxt Like "*Workbook_AfterXmlExport(*": Flag = True
999:            Case stxt Like "*Workbook_AfterXmlImport(*": Flag = True
1000:            Case stxt Like "*Workbook_BeforeClose(*": Flag = True
1001:            Case stxt Like "*Workbook_BeforePrint(*": Flag = True
1002:            Case stxt Like "*Workbook_BeforeSave(*": Flag = True
1003:            Case stxt Like "*Workbook_BeforeXmlExport(*": Flag = True
1004:            Case stxt Like "*Workbook_BeforeXmlImport(*": Flag = True
1005:            Case stxt Like "*Workbook_Deactivate(*": Flag = True
1006:            Case stxt Like "*Workbook_NewChart(*": Flag = True
1007:            Case stxt Like "*Workbook_NewSheet(*": Flag = True
1008:            Case stxt Like "*Workbook_Open(*": Flag = True
1009:            Case stxt Like "*Workbook_PivotTableCloseConnection(*": Flag = True
1010:            Case stxt Like "*Workbook_PivotTableOpenConnection(*": Flag = True
1011:            Case stxt Like "*Workbook_RowsetComplete(*": Flag = True
1012:            Case stxt Like "*Workbook_SheetActivate(*": Flag = True
1013:            Case stxt Like "*Workbook_SheetBeforeDoubleClick(*": Flag = True
1014:            Case stxt Like "*Workbook_SheetBeforeRightClick(*": Flag = True
1015:            Case stxt Like "*Workbook_SheetCalculate(*": Flag = True
1016:            Case stxt Like "*Workbook_SheetChange(*": Flag = True
1017:            Case stxt Like "*Workbook_SheetDeactivate(*": Flag = True
1018:            Case stxt Like "*Workbook_SheetFollowHyperlink(*": Flag = True
1019:            Case stxt Like "*Workbook_SheetPivotTableAfterValueChange(*": Flag = True
1020:            Case stxt Like "*Workbook_SheetPivotTableBeforeAllocateChanges(*": Flag = True
1021:            Case stxt Like "*Workbook_SheetPivotTableBeforeCommitChanges(*": Flag = True
1022:            Case stxt Like "*Workbook_SheetPivotTableBeforeDiscardChanges(*": Flag = True
1023:            Case stxt Like "*Workbook_SheetPivotTableChangeSync(*": Flag = True
1024:            Case stxt Like "*Workbook_SheetPivotTableUpdate(*": Flag = True
1025:            Case stxt Like "*Workbook_SheetSelectionChange(*": Flag = True
1026:            Case stxt Like "*Workbook_Sync(*": Flag = True
1027:            Case stxt Like "*Workbook_WindowActivate(*": Flag = True
1028:            Case stxt Like "*Workbook_WindowDeactivate(*": Flag = True
1029:            Case stxt Like "*Workbook_WindowResize(*": Flag = True
1030:        End Select
1031:    End If
1032:    WorkBookAndSheetsEvents = Flag
End Function

