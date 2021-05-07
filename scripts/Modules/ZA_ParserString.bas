Attribute VB_Name = "ZA_ParserString"
Option Explicit
Option Private Module

Const SH_STRING     As String = "STRING_"
Const SH_NAME_SET   As String = SH_STRING & "SET"
Const SH_NAME_FORM  As String = SH_STRING & "FORM_CONTROLS"
Const SH_NAME_UI    As String = SH_STRING & "UI"
Const SH_NAME_UI14  As String = SH_STRING & "UI14"
Const SH_NAME_CODE  As String = SH_STRING & "CODE"

    Public Sub ParserStringWB()
12:    Dim Form        As AddStatistic
13:    Dim sNameWB     As String
14:    Dim objWB       As Workbook
15:
16:    'On Error GoTo ErrStartParser
17:    Set Form = New AddStatistic
18:    With Form
19:        .Caption = "Collecting string data:"
20:        .lbOK.Caption = "COLLECT"
21:        .chQuestion.visible = False
22:        .chQuestion.Value = False
23:        .Show
24:        sNameWB = .cmbMain.Value
25:    End With
26:    If sNameWB = vbNullString Then Exit Sub
27:    Set objWB = Workbooks(sNameWB)
28:    If Not objWB.FullName Like "*" & Application.PathSeparator & "*" Then
29:        Call MsgBox("Selected file [" & sNameWB & "] not saved, save the file to continue!", vbCritical, "Error:")
30:        Exit Sub
31:    ElseIf objWB.VBProject.Protection = vbext_pp_locked Then
32:        Call MsgBox("Project, file [" & sNameWB & "] protected, remove the password!", vbCritical, "Project:")
33:        Exit Sub
34:    End If
35:
36:    Call ParserStr(objWB, Workbooks.Add)
37:    Set Form = Nothing
38:    Exit Sub
ErrStartParser:
40:    Application.Calculation = xlCalculationAutomatic
41:    Application.ScreenUpdating = True
42:    Call MsgBox("Error in ZA_Parser String.Parser String From W B" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbCritical, "Error:")
43:    Call WriteErrorLog("ParserStringFromWB")
44: End Sub

    Private Sub ParserStr(ByRef WBString As Workbook, ByRef WBNew As Workbook)
47:    'On Error GoTo ErrStartParser
48:    Dim sNameFile   As String
49:    sNameFile = WBString.Name
50:
51:    Application.ScreenUpdating = False
52:    Application.Calculation = xlCalculationManual
53:    Application.EnableEvents = False
54:
55:    Call N_ObfParserVBA.AddShhetInWBook(SH_NAME_SET, WBNew)
56:    With WBNew.Worksheets(SH_NAME_SET)
57:        .Cells(1, 1).Value = "Full Name WB"
58:        .Cells(2, 1).Value = WBString.FullName
59:    End With
60:
61:    Call ParserStrForms(WBString, WBNew)
62:    Call ParserStringsInCodeAdd(WBString, WBNew)
63:    Call ParserStrUI(WBString, WBNew, False)
64:
65:    WBNew.Activate
66:    Application.EnableEvents = True
67:    Application.Calculation = xlCalculationAutomatic
68:    Application.ScreenUpdating = True
69:    Call MsgBox("The string data of the book [" & sNameFile & "] is collected!", vbInformation, "Data collection:")
70:
71:
72:    Exit Sub
ErrStartParser:
74:    Application.EnableEvents = True
75:    Application.Calculation = xlCalculationAutomatic
76:    Application.ScreenUpdating = True
77:    Call MsgBox("Error in ParserStringFromWB" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbCritical, "Error:")
78: End Sub

'* * * * * ParserStrForm START * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserStrForm - сбор строк UserForm
'* Created    : 30-03-2021 11:27
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ParserStrForms(ByRef WB As Workbook, ByRef WBNew As Workbook)
92:    Dim objVB       As VBIDE.VBProject
93:    Dim objVBComp   As VBIDE.VBComponent
94:    Dim objCont     As MSForms.control
95:    Dim strCapiton  As String
96:    Dim strValue    As String
97:    Dim i           As Long
98:    Dim arrStr()    As String
99:
100:    Debug.Print "Start-collecting rows of UserForms controls"
101:
102:    Call N_ObfParserVBA.AddShhetInWBook(SH_NAME_FORM, WBNew)
103:
104:    With WBNew.Worksheets(SH_NAME_FORM)
105:        .Cells(1, 1).Value = "MODULE NAME"
106:        .Cells(1, 2).Value = "TYPE FORM/CONTROL SYSTEM"
107:        .Cells(1, 3).Value = "CONTROL NAME"
108:        .Cells(1, 4).Value = "MEANING"
109:        .Cells(1, 5).Value = "SIGNATURE"
110:        .Cells(1, 6).Value = "CONTROLTIPTEXT"
111:        .Cells(1, 7).Value = "MEANING"
112:        .Cells(1, 8).Value = "SIGNATURE"
113:        .Cells(1, 9).Value = "CONTROLTIPTEXT"
114:        .Columns("A:I").EntireColumn.AutoFit
115:        .Cells.NumberFormat = "@"
116:    End With
117:
118:    For Each objVBComp In WB.VBProject.VBComponents
119:        If objVBComp.Type = vbext_ct_MSForm Then
120:            i = i + 1
121:            ReDim Preserve arrStr(1 To 6, 1 To i)
122:
123:            arrStr(1, i) = objVBComp.Name
124:            arrStr(2, i) = "FORMA"
125:            arrStr(3, i) = arrStr(1, i)
126:            arrStr(4, i) = vbNullString
127:            arrStr(5, i) = GetPropertisForm(objVBComp)
128:            arrStr(6, i) = vbNullString
129:
130:            For Each objCont In objVBComp.Designer.Controls
131:                With objCont
132:                    If PropertyIsCapiton(objCont, True) Then
133:                        If .Caption <> vbNullString Then
134:                            strCapiton = .Caption
135:                        End If
136:                    ElseIf PropertyIsCapiton(objCont, False) Then
137:                        If .Value <> vbNullString Then
138:                            strValue = .Value
139:                        End If
140:                    End If
141:                    If strValue & strCapiton <> vbNullString Then
142:                        i = i + 1
143:                        ReDim Preserve arrStr(1 To 6, 1 To i)
144:                        arrStr(1, i) = objVBComp.Name
145:                        arrStr(2, i) = "CONTROL"
146:                        arrStr(3, i) = objCont.Name
147:                        arrStr(4, i) = strValue
148:                        arrStr(5, i) = strCapiton
149:                        arrStr(6, i) = objCont.ControlTipText
150:                    End If
151:                    strValue = vbNullString: strCapiton = vbNullString
152:                End With
153:            Next objCont
154:        End If
155:    Next objVBComp
156:
157:    If (Not Not arrStr) <> 0 Then
158:        With WBNew.Worksheets(SH_NAME_FORM)
159:            .Cells(2, 1).Resize(UBound(arrStr, 2), UBound(arrStr, 1)).Value2 = WorksheetFunction.Transpose(arrStr)
160:            .Columns("A:C").EntireColumn.AutoFit
161:        End With
162:        Debug.Print "Completed-collecting UserForms control rows"
163:    Else
164:        Debug.Print "Completed-collecting UserForms strings, UserForms is not in the file"
165:    End If
166: End Sub
     Private Function GetPropertisForm(ByRef objVBComp As VBIDE.VBComponent) As String
168:    objVBComp.Activate
169:    GetPropertisForm = objVBComp.Properties("Caption")
170: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : PropertyIsCapiton - проверка существования свойтва Caption у контрола
'* Created    : 30-03-2021 11:28
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                         Description
'*
'* ByRef objCont As MSForms.Control :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function PropertyIsCapiton(ByRef objCont As MSForms.control, Optional bCapiton As Boolean = True) As Boolean
183:    On Error GoTo errEnd
184:    Dim s           As String
185:    PropertyIsCapiton = True
186:    If bCapiton Then
187:        s = objCont.Caption
188:    Else
189:        s = objCont.Text
190:    End If
191:    Exit Function
errEnd:
193:    On Error GoTo 0
194:    PropertyIsCapiton = False
195: End Function
'* * * * * ParserStrForm END * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserStrUI
'* Created    : 30-03-2021 15:39
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook    :
'* ByRef WBNew As Workbook :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ParserStrUI(ByRef WB As Workbook, ByRef WBNew As Workbook, Optional bRenameUI As Boolean = False)
211:
212:    If VBA.UCase$(WB.Name) Like "*.XLS" Then
213:        Debug.Print "UI string collection is not possible in files with the extension [*. xls]"
214:        Debug.Print "Resave the file to the new format"
215:        Exit Sub
216:    End If
217:
218:    Dim cEditOpenXML As clsEditOpenXML
219:    Dim sFullNameFile As String
220:    Dim sFullNameXML As String
221:    sFullNameFile = WB.FullName
222:    WB.Close savechanges:=True
223:    Set cEditOpenXML = New clsEditOpenXML
224:    With cEditOpenXML
225:        .CreateBackupXML = False
226:        .SourceFile = sFullNameFile
227:        .UnzipFile
228:        sFullNameXML = .XMLFolder(XMLFolder_customUI)
229:
230:        If FileHave(sFullNameXML & "customUI.xml") Then
231:            If Not bRenameUI Then
232:                Debug.Print "Start-collecting UI strings ribbon panel UI"
233:                Call ParserStrUIMain(WBNew, SH_STRING & "UI", sFullNameXML & "customUI.xml")
234:                Debug.Print "Completed-collecting rows of the ribbon panel UI"
235:            Else
236:                Debug.Print "Start-renaming the rows of the ribbon panel UI"
237:                Call ReNameStrUI(WBNew, SH_STRING & "UI", sFullNameXML & "customUI.xml")
238:                Debug.Print "Completed-renaming the rows of the ribbon panel UI"
239:            End If
240:        Else
241:            Debug.Print "customUI Ribbon Panel - No"
242:        End If
243:        If FileHave(sFullNameXML & "customUI14.xml") Then
244:            If Not bRenameUI Then
245:                Debug.Print "Start-collecting UI strings ribbon panel UI14"
246:                Call ParserStrUIMain(WBNew, SH_STRING & "UI14", sFullNameXML & "customUI14.xml")
247:                Debug.Print "Completed-collecting rows of the UI14 ribbon panel"
248:            Else
249:                Debug.Print "Start-renaming rows of ribbon panel UI14"
250:                Call ReNameStrUI(WBNew, SH_STRING & "UI14", sFullNameXML & "customUI14.xml")
251:                Debug.Print "Completed-renaming the rows of the UI14 ribbon panel"
252:            End If
253:        Else
254:            Debug.Print "Ribbon panel customUI14-no"
255:        End If
256:        .ZipAllFilesInFolder
257:    End With
258:    Set cEditOpenXML = Nothing
259:    Workbooks.Open sFullNameFile
260: End Sub

'* * * * * ParserStrUI Statr * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserStrUI - парсер строк рибон панели
'* Created    : 30-03-2021 11:35
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sPathUI As String : путь к файлу xml рибон панели
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ParserStrUIMain(ByRef WBNew As Workbook, ByVal SHName As String, ByVal sPathUI As String)
274:    Dim oXMLDoc     As MSXML2.DOMDocument
275:    Dim oXMLRelsList As MSXML2.IXMLDOMNodeList
276:    Dim arrStr()    As String
277:
278:    Call AddShhetInWBook(SHName, WBNew)
279:
280:    With WBNew.Worksheets(SHName)
281:        .Cells(1, 1).Value = "TYPE"
282:        .Cells(1, 2).Value = "ID"
283:        .Cells(1, 3).Value = "LABEL"
284:        .Cells(1, 4).Value = "SUPERTIP"
285:        .Cells(1, 5).Value = "SCREENTIP"
286:        .Cells(1, 6).Value = "TITLE"
287:        .Cells(1, 7).Value = "NEW " & .Cells(1, 3).Value
288:        .Cells(1, 8).Value = "NEW " & .Cells(1, 4).Value
289:        .Cells(1, 9).Value = "NEW " & .Cells(1, 5).Value
290:        .Cells(1, 10).Value = "NEW " & .Cells(1, 6).Value
291:        .Cells(1, 11).Value = "ERRORS"
292:        .Cells.NumberFormat = "@"
293:    End With
294:
295:    Set oXMLDoc = New MSXML2.DOMDocument
296:
297:    With oXMLDoc
298:        .Load sPathUI
299:        Set oXMLRelsList = .SelectNodes("customUI/ribbon/tabs")
300:        Call LookXML(arrStr, oXMLRelsList.Item(0))
301:    End With
302:
303:    With WBNew.Worksheets(SHName)
304:        .Cells(2, 1).Resize(UBound(arrStr, 2), UBound(arrStr, 1)).Value2 = WorksheetFunction.Transpose(arrStr)
305:        .Columns("A:C").EntireColumn.AutoFit
306:        .Columns("F:H").EntireColumn.AutoFit
307:    End With
308: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : LookXML - чтение xml поиск значений атрибутов "id", "label", "supertip", "screentip"
'* Created    : 30-03-2021 11:36
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub LookXML(ByRef arrStr() As String, ByRef oXMLElem As MSXML2.IXMLDOMElement)
317:    Dim i           As Long
318:    With oXMLElem
319:        If .ChildNodes.Length = 0 Then
320:            Exit Sub
321:        Else
322:            For i = 0 To .ChildNodes.Length - 1
323:                If Not .ChildNodes(i).Attributes Is Nothing Then
324:                    Call ReadAtributeValue(arrStr, .ChildNodes(i), Array("id", "label", "supertip", "screentip", "title"))
325:                End If
326:                If .ChildNodes(i).NodeType = NODE_ELEMENT Then
327:                    Call LookXML(arrStr, .ChildNodes(i))
328:                End If
329:            Next i
330:        End If
331:    End With
332: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ReadAtributeValue - считывание значения атрибута
'* Created    : 30-03-2021 11:37
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ReadAtributeValue(ByRef arrStr() As String, ByRef oXMLElem As MSXML2.IXMLDOMElement, ByVal arrNameAtributes As Variant)
342:    Dim i           As Long
343:    Dim iCount      As Long
344:    With oXMLElem.Attributes
345:
346:        If (Not Not arrStr) <> 0 Then
347:            iCount = UBound(arrStr, 2) + 1
348:        Else
349:            iCount = 1
350:        End If
351:
352:        ReDim Preserve arrStr(1 To 6, 1 To iCount)
353:        arrStr(1, iCount) = GetFullNodeName(oXMLElem, oXMLElem.BaseName)
354:        For i = 0 To UBound(arrNameAtributes)
355:            If Not .getNamedItem(arrNameAtributes(i)) Is Nothing Then
356:                arrStr(i + 2, iCount) = .getNamedItem(arrNameAtributes(i)).Text
357:            End If
358:        Next i
359:    End With
360: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetFullNodeName - получение полного дерева до узла xml
'* Created    : 30-03-2021 11:37
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef oXMLElem As MSXML2.IXMLDOMElement :
'* stxt As String                          :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function GetFullNodeName(ByRef oXMLElem As MSXML2.IXMLDOMElement, sTxt As String) As String
375:    With oXMLElem
376:
377:        If Not oXMLElem.ParentNode.NodeType = NODE_DOCUMENT Then
378:            sTxt = oXMLElem.ParentNode.BaseName & "/" & sTxt
379:            sTxt = GetFullNodeName(oXMLElem.ParentNode, sTxt)
380:        Else
381:            GetFullNodeName = sTxt
382:            Exit Function
383:        End If
384:    End With
385:    GetFullNodeName = sTxt
386: End Function

     Public Sub ReNameStr()
389:    If MsgBox("Continue executing [Rename String values] ?" & vbNewLine & "This operation cannot be canceled!", vbYesNo + vbQuestion, "Renaming rows:") = vbNo Then
390:        Exit Sub
391:    End If
392:    Dim WBNew       As Workbook
393:    Set WBNew = ActiveWorkbook
394:    With ActiveSheet
395:        If .Name = SH_NAME_SET Then
396:            Dim sPath As String
397:            sPath = .Cells(2, 1).Value
398:            If FileHave(sPath) Then
399:                Dim WBString As Workbook
400:                Dim sWBName As String
401:                sWBName = C_PublicFunctions.sGetFileName(sPath)
402:
403:                If C_PublicFunctions.WorkbookIsOpen(sWBName) Then
404:                    Set WBString = Workbooks(sWBName)
405:                Else
406:                    Set WBString = Workbooks.Open(sPath)
407:                End If
408:
409:                If WBString.VBProject.Protection = vbext_pp_locked Then
410:                    Call MsgBox("The project is protected, remove the password!", vbCritical, "Project:")
411:                Else
412:                    If HaveSheetInFile(WBNew, SH_NAME_FORM) Then
413:                        Call ReNameFormControls(WBString, WBNew)
414:                    End If
415:                    If HaveSheetInFile(WBNew, SH_NAME_CODE) Then
416:                        Call ReNameParserStringsInCodeAdd(WBString, WBNew)
417:                    End If
418:                    Call ParserStrUI(WBString, WBNew, True)
419:                End If
420:            Else
421:                Call MsgBox("File not found on the sheet: [" & SH_NAME_SET & "]", vbCritical, "Error:")
422:            End If
423:        Else
424:            Call MsgBox("Create or navigate to a sheet: [" & SH_NAME_SET & "]", vbCritical, "Search for settings:")
425:        End If
426:    End With
427: End Sub
     Private Function HaveSheetInFile(ByRef WB As Workbook, ByVal SHName As String) As Boolean
429:    Dim SH          As Worksheet
430:    On Error Resume Next
431:    Set SH = WB.Worksheets(SHName)
432:    If Err.Number = 0 Then
433:        HaveSheetInFile = True
434:    Else
435:        HaveSheetInFile = False
436:        Debug.Print "Sheet not found - [" & SHName & "]"
437:    End If
438:    Err.Clear
439: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ReNameFormControls - изменение свойств контроллов Value, Caption, ControlTipText
'* Created    : 30-03-2021 16:07
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook    :
'* ByRef WBNew As Workbook :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ReNameFormControls(ByRef WB As Workbook, ByRef WBNew As Workbook)
454:
455:    Dim arrData     As Variant
456:    Dim lLastRow    As Long
457:
458:    With WBNew.Worksheets(SH_NAME_FORM)
459:        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
460:        If lLastRow < 2 Then Exit Sub
461:        arrData = .Range(.Cells(2, 1), .Cells(lLastRow, 10)).Value2
462:    End With
463:
464:    Dim objVB       As VBIDE.VBProject
465:    Dim objVBCom    As VBIDE.VBComponent
466:    Dim objControl  As MSForms.control
467:    Dim i           As Long
468:
469:    Debug.Print "Start-renaming UserForms controls"
470:    Set objVB = WB.VBProject
471:
472:    For i = 1 To UBound(arrData)
473:        If CheckVBComponent(objVB, arrData(i, 1)) Then
474:            Set objVBCom = objVB.VBComponents(arrData(i, 1))
475:            If arrData(i, 2) = "FORMA" Then
476:                If arrData(i, 8) <> vbNullString Then Call SetPropertisForm(objVBCom, arrData(i, 8))
477:            Else
478:                If CheckControlOnForm(objVBCom, arrData(i, 3)) Then
479:                    Set objControl = objVBCom.Designer.Controls(arrData(i, 3))
480:                    With objControl
481:                        If arrData(i, 7) <> vbNullString Then
482:                            If PropertyIsCapiton(objControl, False) Then .Value = arrData(i, 7)
483:                        End If
484:                        If arrData(i, 8) <> vbNullString Then
485:                            If PropertyIsCapiton(objControl, True) Then .Caption = arrData(i, 8)
486:                        End If
487:                        .ControlTipText = arrData(i, 9)
488:                    End With
489:                Else
490:                    arrData(i, 10) = "Controller not found"
491:                End If
492:            End If
493:        Else
494:            arrData(i, 10) = "Module not found"
495:        End If
496:    Next i
497:
498:    With WBNew.Worksheets(SH_NAME_FORM)
499:        .Cells(1, 10).Value = "ERRORS"
500:        .Cells(2, 1).Resize(UBound(arrData, 1), UBound(arrData, 2)).Value2 = arrData
501:    End With
502:    Debug.Print "Completed-renaming of UserForms controls"
503:
504: End Sub
     Private Sub SetPropertisForm(ByRef objVBComp As VBIDE.VBComponent, ByVal sVal As String)
506:    objVBComp.Activate
507:    objVBComp.Properties("Caption") = sVal
508: End Sub

     Private Function CheckVBComponent(ByRef objVB As VBIDE.VBProject, ByVal sNameComponent As String) As Boolean
511:    Dim objVBCom    As VBIDE.VBComponent
512:    On Error GoTo endFun
513:    CheckVBComponent = True
514:    Set objVBCom = objVB.VBComponents(sNameComponent)
515:    Exit Function
endFun:
517:    Err.Clear
518:    CheckVBComponent = False
519: End Function

     Private Function CheckControlOnForm(ByRef objVBCom As VBIDE.VBComponent, ByVal sNameControl As String) As Boolean
522:    Dim objControl  As MSForms.control
523:    On Error GoTo endFun
524:    CheckControlOnForm = True
525:    Set objControl = objVBCom.Designer.Controls(sNameControl)
526:    Exit Function
endFun:
528:    Err.Clear
529:    CheckControlOnForm = False
530: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ReNameStrUI - переименование элементов риббон панелей "label", "supertip", "screentip"
'* Created    : 30-03-2021 15:42
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WBNew As Workbook :
'* ByVal sPathUI As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ReNameStrUI(ByRef WBNew As Workbook, ByVal SHName As String, ByVal sPathUI As String)
545:
546:    Dim arrData     As Variant
547:    Dim lLastRow    As Long
548:
549:    With WBNew.Worksheets(SHName)
550:        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
551:        If lLastRow < 2 Then Exit Sub
552:        arrData = .Range(.Cells(2, 1), .Cells(lLastRow, 11)).Value2
553:    End With
554:
555:    Dim oXMLDoc     As MSXML2.DOMDocument
556:    Dim oXMLRelsList As MSXML2.IXMLDOMNodeList
557:    Dim i           As Long
558:
559:    Set oXMLDoc = New MSXML2.DOMDocument
560:
561:    oXMLDoc.Load sPathUI
562:    For i = 1 To UBound(arrData)
563:        If arrData(i, 2) <> vbNullString Then
564:            Set oXMLRelsList = oXMLDoc.SelectNodes(arrData(i, 1) & "[@id='" & arrData(i, 2) & "']")
565:            With oXMLRelsList.Item(0)
566:                Call ChengeAtribute(.Attributes, "label", arrData(i, 7))
567:                Call ChengeAtribute(.Attributes, "supertip", arrData(i, 8))
568:                Call ChengeAtribute(.Attributes, "screentip", arrData(i, 9))
569:                Call ChengeAtribute(.Attributes, "title", arrData(i, 10))
570:            End With
571:        End If
572:    Next i
573:    Call oXMLDoc.Save(sPathUI)
574: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ChengeAtribute - изменение значений антрибутов xml
'* Created    : 30-03-2021 15:44
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                                     Description
'*
'* ByRef oNodeMap As MSXML2.IXMLDOMNamedNodeMap :
'* ByVal sNameAtr As String                     :
'* ByVal sVal As String                         :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ChengeAtribute(ByRef oNodeMap As MSXML2.IXMLDOMNamedNodeMap, ByVal sNameAtr As String, ByVal sVal As String)
590:    If sVal <> vbNullString Then
591:        With oNodeMap
592:            If Not .getNamedItem(sNameAtr) Is Nothing Then
593:                .getNamedItem(sNameAtr).Text = sVal
594:            End If
595:        End With
596:    End If
597: End Sub

     Private Sub ParserStringsInCodeAdd(ByRef WB As Workbook, ByRef WBNew As Workbook)
600:    Dim oVBP        As VBIDE.VBProject
601:    Dim oVBCom      As VBIDE.VBComponent
602:    Dim iLineCode   As Long
603:    Dim sCode       As String
604:    Dim arrString   As Variant
605:    Dim i           As Long
606:    Dim k           As Long
607:    Dim j           As Integer
608:    Dim sStrCode    As String
609:    Dim arrParser() As String
610:    Dim arrPartStr  As Variant
611:
612:    Debug.Print "Start - collecting rows in modules"
613:
614:    Set oVBP = WB.VBProject
615:    For Each oVBCom In oVBP.VBComponents
616:        With oVBCom.CodeModule
617:            iLineCode = .CountOfLines
618:            If iLineCode > 0 Then
619:                sCode = .Lines(1, iLineCode)
620:                If sCode <> vbNullString And sCode Like "*" & VBA.Chr$(34) & "*" Then
621:                    sCode = VBA.Replace(sCode, " _" & vbNewLine, vbNullString)
622:                    arrString = VBA.Split(sCode, vbNewLine)
623:                    For i = 0 To UBound(arrString)
624:                        sStrCode = arrString(i)
625:                        sStrCode = TrimSpace(sStrCode)
626:                        If sStrCode <> vbNullString And VBA.Left$(sStrCode, 1) <> "'" And sStrCode Like "*" & VBA.Chr$(34) & "*" Then
627:                            sStrCode = DeleteCommentString(sStrCode)
628:                            sStrCode = VBA.Replace(sStrCode, " " & VBA.Chr$(34) & VBA.Chr$(34) & " ", vbNullString)
629:                            sStrCode = VBA.Replace(sStrCode, " " & VBA.Chr$(34) & VBA.Chr$(34), vbNullString)
630:                            arrPartStr = VBA.Split(sStrCode, VBA.Chr$(34))
631:                            For j = 1 To UBound(arrPartStr) Step 2
632:                                If arrPartStr(j) <> vbNullString Then
633:                                    k = k + 1
634:                                    ReDim Preserve arrParser(1 To 2, 1 To k)
635:                                    arrParser(1, k) = oVBCom.Name
636:                                    arrParser(2, k) = arrPartStr(j)
637:                                End If
638:                            Next j
639:                        End If
640:                    Next i
641:                End If
642:            End If
643:        End With
644:    Next oVBCom
645:
646:    If (Not Not arrParser) <> 0 Then
647:        Call AddShhetInWBook(SH_NAME_CODE, WBNew)
648:        With WBNew.Worksheets(SH_NAME_CODE)
649:            .Cells(1, 1).Value = "NAME MODULE"
650:            .Cells(1, 2).Value = "STRING"
651:            .Cells(1, 3).Value = "NEW STRING"
652:            .Cells(1, 4).Value = "ERRORS"
653:            .Cells.NumberFormat = "@"
654:            .Cells(2, 1).Resize(UBound(arrParser, 2), UBound(arrParser, 1)).Value2 = WorksheetFunction.Transpose(arrParser)
655:            .Columns("A:D").EntireColumn.AutoFit
656:            Debug.Print "Completed-collecting rows in modules"
657:        End With
658:    Else
659:        Debug.Print "Completed - collection of rows in modules, no rows"
660:    End If
661: End Sub

     Private Sub ReNameParserStringsInCodeAdd(ByRef WB As Workbook, ByRef WBNew As Workbook)
664:    Dim oVBP        As VBIDE.VBProject
665:    Dim iLineCode   As Long
666:    Dim sCode       As String
667:    Dim sCodeNew    As String
668:    Dim iCount      As Long
669:    Dim arrData     As Variant
670:    Dim lLastRow    As Long
671:    Dim i           As Long
672:    Dim k           As Long
673:    Dim oVBCMod     As VBIDE.CodeModule
674:
675:
676:    With WBNew.Worksheets(SH_NAME_CODE)
677:        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
678:        If lLastRow < 2 Then Exit Sub
679:        arrData = .Range(.Cells(2, 1), .Cells(lLastRow, 4)).Value2
680:    End With
681:
682:    Set oVBP = WB.VBProject
683:    iCount = UBound(arrData)
684:    For i = 1 To iCount
685:        k = 1
686:        If i = iCount Then k = 0
687:        If arrData(i, 1) <> vbNullString Then
688:            If i = 1 Then
689:                Set oVBCMod = oVBP.VBComponents(arrData(i, 1)).CodeModule
690:                sCode = GetCodeFromModule(oVBCMod)
691:                sCodeNew = sCode
692:                If arrData(i, 3) <> vbNullString Then
693:                    sCodeNew = VBA.Replace(sCodeNew, VBA.Chr$(34) & arrData(i, 2) & VBA.Chr$(34), VBA.Chr$(34) & arrData(i, 3) & VBA.Chr$(34))
694:                End If
695:            End If
696:            'если в таблице всего одна запись
697:            If iCount = 1 Then
698:                Call SetCodeInModule(oVBCMod, sCode, sCodeNew)
699:            Else
700:                If arrData(i, 3) <> vbNullString Then
701:                    sCodeNew = VBA.Replace(sCodeNew, VBA.Chr$(34) & arrData(i, 2) & VBA.Chr$(34), VBA.Chr$(34) & arrData(i, 3) & VBA.Chr$(34))
702:                End If
703:                If arrData(i, 1) <> arrData(i + k, 1) Or i = iCount Then
704:                    Call SetCodeInModule(oVBCMod, sCode, sCodeNew)
705:                    Set oVBCMod = oVBP.VBComponents(arrData(i + k, 1)).CodeModule
706:                    sCode = GetCodeFromModule(oVBCMod)
707:                    sCodeNew = sCode
708:                End If
709:            End If
710:        End If
711:    Next i
712: End Sub

     Private Function GetCodeFromModule(ByRef oVBCMod As VBIDE.CodeModule) As String
715:    Dim iLineCode   As Long
716:    Dim sCode       As String
717:    With oVBCMod
718:        iLineCode = .CountOfLines
719:        If iLineCode > 0 Then
720:            sCode = .Lines(1, iLineCode)
721:            If sCode <> vbNullString And sCode Like "*" & VBA.Chr$(34) & "*" Then
722:                GetCodeFromModule = sCode
723:            End If
724:        End If
725:    End With
726: End Function

Private Sub SetCodeInModule(ByRef oVBCMod As VBIDE.CodeModule, ByVal sCode As String, ByVal sCodeNew As String)
729:    If sCode <> sCodeNew Then
730:        Dim iLineCode As Long
731:        With oVBCMod
732:            iLineCode = .CountOfLines
733:            If iLineCode > 0 Then
734:                Call .DeleteLines(1, iLineCode)
735:                Call .InsertLines(1, sCodeNew)
736:            End If
737:        End With
738:    End If
End Sub
