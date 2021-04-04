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
19:        .Caption = "Сбор строковых данных:"
20:        .lbOK.Caption = "СОБРАТЬ"
21:        .chQuestion.visible = False
22:        .chQuestion.Value = False
23:        .Show
24:        sNameWB = .cmbMain.Value
25:    End With
26:    If sNameWB = vbNullString Then Exit Sub
27:    Set objWB = Workbooks(sNameWB)
28:    If Not objWB.FullName Like "*" & Application.PathSeparator & "*" Then
29:        Call MsgBox("Выбранный файл не сохранен, для продолжения сохраните файл!", vbCritical, "Ошибка:")
30:        Exit Sub
31:    End If
32:
33:    Call ParserStr(objWB, Workbooks.Add)
34:    Set Form = Nothing
35:    Exit Sub
ErrStartParser:
37:    Application.Calculation = xlCalculationAutomatic
38:    Application.ScreenUpdating = True
39:    Call MsgBox("Ошибка в ZA_ParserString.ParserStringFromWB" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
40:    Call WriteErrorLog("ParserStringFromWB")
41: End Sub

    Private Sub ParserStr(ByRef WBString As Workbook, ByRef WBNew As Workbook)
44:    'On Error GoTo ErrStartParser
45:    Dim sNameFile   As String
46:    sNameFile = WBString.Name
47:    If WBString.VBProject.Protection = vbext_pp_locked Then
48:        Call MsgBox("Проект защищен, снимите пароль!", vbCritical, "Проект:")
49:    Else
50:
51:        Application.ScreenUpdating = False
52:        Application.Calculation = xlCalculationManual
53:        Application.EnableEvents = False
54:
55:        Call N_ObfParserVBA.AddShhetInWBook(SH_NAME_SET, WBNew)
56:        With WBNew.Worksheets(SH_NAME_SET)
57:            .Cells(1, 1).Value = "Full Name WB"
58:            .Cells(2, 1).Value = WBString.FullName
59:        End With
60:
61:        Call ParserStrForms(WBString, WBNew)
62:        Call ParserStringsInCodeAdd(WBString, WBNew)
63:        Call ParserStrUI(WBString, WBNew, False)
64:
65:        WBNew.Activate
66:        Application.EnableEvents = True
67:        Application.Calculation = xlCalculationAutomatic
68:        Application.ScreenUpdating = True
69:        Call MsgBox("Строковые данные книги [" & sNameFile & "] собраны!", vbInformation, "Сбор данных:")
70:
71:    End If
72:    Exit Sub
ErrStartParser:
74:    Application.EnableEvents = True
75:    Application.Calculation = xlCalculationAutomatic
76:    Application.ScreenUpdating = True
77:    Call MsgBox("Ошибка в ParserStringFromWB" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
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
100:    Debug.Print "Начало - сбора строк UserForms контролов"
101:
102:    Call N_ObfParserVBA.AddShhetInWBook(SH_NAME_FORM, WBNew)
103:
104:    With WBNew.Worksheets(SH_NAME_FORM)
105:        .Cells(1, 1).Value = "НАЗВАНИЕ МОДУЛЯ"
106:        .Cells(1, 2).Value = "ИМЯ КОНТРОЛА"
107:        .Cells(1, 3).Value = "ЗНАЧЕНИЕ"
108:        .Cells(1, 4).Value = "ПОДПИСЬ"
109:        .Cells(1, 5).Value = "CONTROLTIPTEXT"
110:        .Cells(1, 6).Value = "ЗНАЧЕНИЕ"
111:        .Cells(1, 7).Value = "ПОДПИСЬ"
112:        .Cells(1, 8).Value = "CONTROLTIPTEXT"
113:        .Cells.NumberFormat = "@"
114:    End With
115:
116:    For Each objVBComp In WB.VBProject.VBComponents
117:        If objVBComp.Type = vbext_ct_MSForm Then
118:            For Each objCont In objVBComp.Designer.Controls
119:                With objCont
120:                    If PropertyIsCapiton(objCont, True) Then
121:                        If .Caption <> vbNullString Then
122:                            strCapiton = .Caption
123:                        End If
124:                    ElseIf PropertyIsCapiton(objCont, False) Then
125:                        If .Value <> vbNullString Then
126:                            strValue = .Value
127:                        End If
128:                    End If
129:                    If strValue & strCapiton <> vbNullString Then
130:                        i = i + 1
131:                        ReDim Preserve arrStr(1 To 5, 1 To i)
132:                        arrStr(1, i) = objVBComp.Name
133:                        arrStr(2, i) = objCont.Name
134:                        arrStr(3, i) = strValue
135:                        arrStr(4, i) = strCapiton
136:                        arrStr(5, i) = objCont.ControlTipText
137:                    End If
138:                    strValue = vbNullString: strCapiton = vbNullString
139:                End With
140:            Next objCont
141:        End If
142:    Next objVBComp
143:
144:    If (Not Not arrStr) <> 0 Then
145:        With WBNew.Worksheets(SH_NAME_FORM)
146:            .Cells(2, 1).Resize(UBound(arrStr, 2), UBound(arrStr, 1)).Value2 = WorksheetFunction.Transpose(arrStr)
147:            .Columns("A:H").EntireColumn.AutoFit
148:        End With
149:        Debug.Print "Завершен - сбор строк UserForms контролов"
150:    Else
151:        Debug.Print "Завершен - сбор строк UserForms, UserForms нет в файле"
152:    End If
153: End Sub
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
     Private Function PropertyIsCapiton(ByRef objCont As MSForms.control, Optional bCapiton As Boolean = True) As Boolean
166:    On Error GoTo errEnd
167:    Dim s           As String
168:    PropertyIsCapiton = True
169:    If bCapiton Then
170:        s = objCont.Caption
171:    Else
172:        s = objCont.Text
173:    End If
174:    Exit Function
errEnd:
176:    On Error GoTo 0
177:    PropertyIsCapiton = False
178: End Function
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
194:
195:    If VBA.UCase$(WB.Name) Like "*.XLS" Then
196:        Debug.Print "Сбор строк UI не возможен в файлах с расширением [*.xls]"
197:        Debug.Print "Пресохраните файл в новый формат"
198:        Exit Sub
199:    End If
200:
201:    Debug.Print "Начало - сбора строк UI рибон панели"
202:    Dim cEditOpenXML As clsEditOpenXML
203:    Dim sFullNameFile As String
204:    Dim sFullNameXML As String
205:    sFullNameFile = WB.FullName
206:    WB.Close savechanges:=True
207:    Set cEditOpenXML = New clsEditOpenXML
208:    With cEditOpenXML
209:        .CreateBackupXML = False
210:        .SourceFile = sFullNameFile
211:        .UnzipFile
212:        sFullNameXML = .XMLFolder(XMLFolder_customUI)
213:
214:        If FileHave(sFullNameXML & "customUI.xml") Then
215:            If Not bRenameUI Then
216:                Call ParserStrUIMain(WBNew, SH_STRING & "UI", sFullNameXML & "customUI.xml")
217:            Else
218:                Call ReNameStrUI(WBNew, SH_STRING & "UI", sFullNameXML & "customUI.xml")
219:            End If
220:        Else
221:            Debug.Print "Рибон панели customUI - нет"
222:        End If
223:        If FileHave(sFullNameXML & "customUI14.xml") Then
224:            If Not bRenameUI Then
225:                Call ParserStrUIMain(WBNew, SH_STRING & "UI14", sFullNameXML & "customUI14.xml")
226:            Else
227:                Call ReNameStrUI(WBNew, SH_STRING & "UI14", sFullNameXML & "customUI14.xml")
228:            End If
229:        Else
230:            Debug.Print "Рибон панели customUI14 - нет"
231:        End If
232:        .ZipAllFilesInFolder
233:    End With
234:    Set cEditOpenXML = Nothing
235:
236:    Workbooks.Open sFullNameFile
237:    Debug.Print "Завершен - сбор строк UI рибон панели"
238:
239: End Sub

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
253:    Dim oXMLDoc     As MSXML2.DOMDocument
254:    Dim oXMLRelsList As MSXML2.IXMLDOMNodeList
255:    Dim arrStr()    As String
256:
257:    Call AddShhetInWBook(SHName, WBNew)
258:
259:    With WBNew.Worksheets(SHName)
260:        .Cells(1, 1).Value = "TYPE"
261:        .Cells(1, 2).Value = "ID"
262:        .Cells(1, 3).Value = "LABEL"
263:        .Cells(1, 4).Value = "SUPERTIP"
264:        .Cells(1, 5).Value = "SCREENTIP"
265:        .Cells(1, 6).Value = "TITLE"
266:        .Cells(1, 7).Value = "NEW " & .Cells(1, 3).Value
267:        .Cells(1, 8).Value = "NEW " & .Cells(1, 4).Value
268:        .Cells(1, 9).Value = "NEW " & .Cells(1, 5).Value
269:        .Cells(1, 10).Value = "NEW " & .Cells(1, 6).Value
270:        .Cells(1, 11).Value = "ERRORS"
271:        .Cells.NumberFormat = "@"
272:    End With
273:
274:    Set oXMLDoc = New MSXML2.DOMDocument
275:
276:    With oXMLDoc
277:        .Load sPathUI
278:        Set oXMLRelsList = .SelectNodes("customUI/ribbon/tabs")
279:        Call LookXML(arrStr, oXMLRelsList.Item(0))
280:    End With
281:
282:    With WBNew.Worksheets(SHName)
283:        .Cells(2, 1).Resize(UBound(arrStr, 2), UBound(arrStr, 1)).Value2 = WorksheetFunction.Transpose(arrStr)
284:        .Columns("A:C").EntireColumn.AutoFit
285:        .Columns("F:H").EntireColumn.AutoFit
286:    End With
287: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : LookXML - чтение xml поиск значений атрибутов "id", "label", "supertip", "screentip"
'* Created    : 30-03-2021 11:36
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub LookXML(ByRef arrStr() As String, ByRef oXMLElem As MSXML2.IXMLDOMElement)
296:    Dim i           As Long
297:    With oXMLElem
298:        If .ChildNodes.Length = 0 Then
299:            Exit Sub
300:        Else
301:            For i = 0 To .ChildNodes.Length - 1
302:                If Not .ChildNodes(i).Attributes Is Nothing Then
303:                    Call ReadAtributeValue(arrStr, .ChildNodes(i), Array("id", "label", "supertip", "screentip", "title"))
304:                End If
305:                If .ChildNodes(i).NodeType = NODE_ELEMENT Then
306:                    Call LookXML(arrStr, .ChildNodes(i))
307:                End If
308:            Next i
309:        End If
310:    End With
311: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ReadAtributeValue - считывание значения атрибута
'* Created    : 30-03-2021 11:37
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ReadAtributeValue(ByRef arrStr() As String, ByRef oXMLElem As MSXML2.IXMLDOMElement, ByVal arrNameAtributes As Variant)
321:    Dim i           As Long
322:    Dim iCount      As Long
323:    With oXMLElem.Attributes
324:
325:        If (Not Not arrStr) <> 0 Then
326:            iCount = UBound(arrStr, 2) + 1
327:        Else
328:            iCount = 1
329:        End If
330:
331:        ReDim Preserve arrStr(1 To 6, 1 To iCount)
332:        arrStr(1, iCount) = GetFullNodeName(oXMLElem, oXMLElem.BaseName)
333:        For i = 0 To UBound(arrNameAtributes)
334:            If Not .getNamedItem(arrNameAtributes(i)) Is Nothing Then
335:                arrStr(i + 2, iCount) = .getNamedItem(arrNameAtributes(i)).Text
336:            End If
337:        Next i
338:    End With
339: End Sub

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
     Private Function GetFullNodeName(ByRef oXMLElem As MSXML2.IXMLDOMElement, stxt As String) As String
354:    With oXMLElem
355:
356:        If Not oXMLElem.ParentNode.NodeType = NODE_DOCUMENT Then
357:            stxt = oXMLElem.ParentNode.BaseName & "/" & stxt
358:            stxt = GetFullNodeName(oXMLElem.ParentNode, stxt)
359:        Else
360:            GetFullNodeName = stxt
361:            Exit Function
362:        End If
363:    End With
364:    GetFullNodeName = stxt
365: End Function

     Public Sub ReNameStr()
368:    Dim WBNew       As Workbook
369:    Set WBNew = ActiveWorkbook
370:    With ActiveSheet
371:        If .Name = SH_NAME_SET Then
372:            Dim sPath As String
373:            sPath = .Cells(2, 1).Value
374:            If FileHave(sPath) Then
375:                Dim WBString As Workbook
376:                Set WBString = Workbooks.Open(sPath)
377:
378:                If WBString.VBProject.Protection = vbext_pp_locked Then
379:                    Call MsgBox("Проект защищен, снимите пароль!", vbCritical, "Проект:")
380:                Else
381:                    Call ReNameFormControls(WBString, WBNew)
382:                    Call ReNameParserStringsInCodeAdd(WBString, WBNew)
383:                    Call ParserStrUI(WBString, WBNew, True)
384:                End If
385:            Else
386:                Call MsgBox("Файл не найден на листе: [" & SH_NAME_SET & "]", vbCritical, "Ошибка:")
387:            End If
388:        Else
389:            Call MsgBox("Создайте или перейдите на лист: [" & SH_NAME_SET & "]", vbCritical, "Поиск настроек:")
390:        End If
391:    End With
392: End Sub


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
408:
409:    Dim arrData     As Variant
410:    Dim lLastRow    As Long
411:
412:    With WBNew.Worksheets(SH_NAME_FORM)
413:        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
414:        If lLastRow < 2 Then Exit Sub
415:        arrData = .Range(.Cells(2, 1), .Cells(lLastRow, 9)).Value2
416:    End With
417:
418:    Dim objVB       As VBIDE.VBProject
419:    Dim objVBCom    As VBIDE.VBComponent
420:    Dim objControl  As MSForms.control
421:    Dim i           As Long
422:
423:    Debug.Print "Начало - переименование UserForms контролов"
424:    Set objVB = WB.VBProject
425:
426:    For i = 1 To UBound(arrData)
427:        If CheckVBComponent(objVB, arrData(i, 1)) Then
428:            Set objVBCom = objVB.VBComponents(arrData(i, 1))
429:            If CheckControlOnForm(objVBCom, arrData(i, 2)) Then
430:                Set objControl = objVBCom.Designer.Controls(arrData(i, 2))
431:                With objControl
432:                    If arrData(i, 6) <> vbNullString Then
433:                        If PropertyIsCapiton(objControl, False) Then .Value = arrData(i, 6)
434:                    End If
435:                    If arrData(i, 7) <> vbNullString Then
436:                        If PropertyIsCapiton(objControl, True) Then .Caption = arrData(i, 7)
437:                    End If
438:                    .ControlTipText = arrData(i, 8)
439:                End With
440:            Else
441:                arrData(i, 9) = "Не найден контрол"
442:            End If
443:        Else
444:            arrData(i, 9) = "Не найден модуль"
445:        End If
446:    Next i
447:
448:    With WBNew.Worksheets(SH_NAME_FORM)
449:        .Cells(1, 9).Value = "ОШИБКИ"
450:        .Cells(2, 1).Resize(UBound(arrData, 1), UBound(arrData, 2)).Value2 = arrData
451:    End With
452:    Debug.Print "Завершено - переименование UserForms контролов"
453:
454: End Sub

     Private Function CheckVBComponent(ByRef objVB As VBIDE.VBProject, ByVal sNameComponent As String) As Boolean
457:    Dim objVBCom    As VBIDE.VBComponent
458:    On Error GoTo endFun
459:    CheckVBComponent = True
460:    Set objVBCom = objVB.VBComponents(sNameComponent)
461:    Exit Function
endFun:
463:    Err.Clear
464:    CheckVBComponent = False
465: End Function

     Private Function CheckControlOnForm(ByRef objVBCom As VBIDE.VBComponent, ByVal sNameControl As String) As Boolean
468:    Dim objControl  As MSForms.control
469:    On Error GoTo endFun
470:    CheckControlOnForm = True
471:    Set objControl = objVBCom.Designer.Controls(sNameControl)
472:    Exit Function
endFun:
474:    Err.Clear
475:    CheckControlOnForm = False
476: End Function

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
491:
492:    Dim arrData     As Variant
493:    Dim lLastRow    As Long
494:
495:    With WBNew.Worksheets(SHName)
496:        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
497:        If lLastRow < 2 Then Exit Sub
498:        arrData = .Range(.Cells(2, 1), .Cells(lLastRow, 11)).Value2
499:    End With
500:
501:    Dim oXMLDoc     As MSXML2.DOMDocument
502:    Dim oXMLRelsList As MSXML2.IXMLDOMNodeList
503:    Dim i           As Long
504:
505:    Set oXMLDoc = New MSXML2.DOMDocument
506:
507:    oXMLDoc.Load sPathUI
508:    For i = 1 To UBound(arrData)
509:        If arrData(i, 2) <> vbNullString Then
510:            Set oXMLRelsList = oXMLDoc.SelectNodes(arrData(i, 1) & "[@id='" & arrData(i, 2) & "']")
511:            With oXMLRelsList.Item(0)
512:                Call ChengeAtribute(.Attributes, "label", arrData(i, 7))
513:                Call ChengeAtribute(.Attributes, "supertip", arrData(i, 8))
514:                Call ChengeAtribute(.Attributes, "screentip", arrData(i, 9))
515:                Call ChengeAtribute(.Attributes, "title", arrData(i, 10))
516:            End With
517:        End If
518:    Next i
519:    Call oXMLDoc.Save(sPathUI)
520: End Sub

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
536:    If sVal <> vbNullString Then
537:        With oNodeMap
538:            If Not .getNamedItem(sNameAtr) Is Nothing Then
539:                .getNamedItem(sNameAtr).Text = sVal
540:            End If
541:        End With
542:    End If
543: End Sub

     Private Sub ParserStringsInCodeAdd(ByRef WB As Workbook, ByRef WBNew As Workbook)
546:    Dim oVBP        As VBIDE.VBProject
547:    Dim oVBCom      As VBIDE.VBComponent
548:    Dim iLineCode   As Long
549:    Dim sCode       As String
550:    Dim arrString   As Variant
551:    Dim i           As Long
552:    Dim k           As Long
553:    Dim j           As Integer
554:    Dim sStrCode    As String
555:    Dim arrParser() As String
556:    Dim arrPartStr  As Variant
557:
558:    Debug.Print "Начало - сбора строк в модулях"
559:
560:    Set oVBP = WB.VBProject
561:    For Each oVBCom In oVBP.VBComponents
562:        With oVBCom.CodeModule
563:            iLineCode = .CountOfLines
564:            If iLineCode > 0 Then
565:                sCode = .Lines(1, iLineCode)
566:                If sCode <> vbNullString And sCode Like "*" & VBA.Chr$(34) & "*" Then
567:                    sCode = VBA.Replace(sCode, " _" & vbNewLine, vbNullString)
568:                    arrString = VBA.Split(sCode, vbNewLine)
569:                    For i = 0 To UBound(arrString)
570:                        sStrCode = arrString(i)
571:                        sStrCode = TrimSpace(sStrCode)
572:                        If sStrCode <> vbNullString And VBA.Left$(sStrCode, 1) <> "'" And sStrCode Like "*" & VBA.Chr$(34) & "*" Then
573:                            sStrCode = DeleteCommentString(sStrCode)
574:                            sStrCode = VBA.Replace(sStrCode, " " & VBA.Chr$(34) & VBA.Chr$(34) & " ", vbNullString)
575:                            sStrCode = VBA.Replace(sStrCode, " " & VBA.Chr$(34) & VBA.Chr$(34), vbNullString)
576:                            arrPartStr = VBA.Split(sStrCode, VBA.Chr$(34))
577:                            For j = 1 To UBound(arrPartStr) Step 2
578:                                If arrPartStr(j) <> vbNullString Then
579:                                    k = k + 1
580:                                    ReDim Preserve arrParser(1 To 2, 1 To k)
581:                                    arrParser(1, k) = oVBCom.Name
582:                                    arrParser(2, k) = arrPartStr(j)
583:                                End If
584:                            Next j
585:                        End If
586:                    Next i
587:                End If
588:            End If
589:        End With
590:    Next oVBCom
591:
592:    If (Not Not arrParser) <> 0 Then
593:        Call AddShhetInWBook(SH_NAME_CODE, WBNew)
594:        With WBNew.Worksheets(SH_NAME_CODE)
595:            .Cells(1, 1).Value = "NAME MODULE"
596:            .Cells(1, 2).Value = "STRING"
597:            .Cells(1, 3).Value = "NEW STRING"
598:            .Cells(1, 4).Value = "ERRORS"
599:            .Cells.NumberFormat = "@"
600:            .Cells(2, 1).Resize(UBound(arrParser, 2), UBound(arrParser, 1)).Value2 = WorksheetFunction.Transpose(arrParser)
601:            .Columns("A:D").EntireColumn.AutoFit
602:            Debug.Print "Завершен - сбор строк в модулях"
603:        End With
604:    Else
605:        Debug.Print "Завершен - сбор строк в модулях, строк нет"
606:    End If
607: End Sub

     Private Sub ReNameParserStringsInCodeAdd(ByRef WB As Workbook, ByRef WBNew As Workbook)
610:    Dim oVBP        As VBIDE.VBProject
611:    Dim iLineCode   As Long
612:    Dim sCode       As String
613:    Dim sCodeNew    As String
614:    Dim iCount      As Long
615:    Dim arrData     As Variant
616:    Dim lLastRow    As Long
617:    Dim i           As Long
618:    Dim k           As Long
619:    Dim oVBCMod     As VBIDE.CodeModule
620:
621:
622:    With WBNew.Worksheets(SH_NAME_CODE)
623:        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
624:        If lLastRow < 2 Then Exit Sub
625:        arrData = .Range(.Cells(2, 1), .Cells(lLastRow, 4)).Value2
626:    End With
627:
628:    Set oVBP = WB.VBProject
629:    iCount = UBound(arrData)
630:    For i = 1 To iCount
631:        k = 1
632:        If i = iCount Then k = 0
633:        If arrData(i, 1) <> vbNullString Then
634:            If i = 1 Then
635:                Set oVBCMod = oVBP.VBComponents(arrData(i, 1)).CodeModule
636:                sCode = GetCodeFromModule(oVBCMod)
637:                sCodeNew = sCode
638:                If arrData(i, 3) <> vbNullString Then
639:                    sCodeNew = VBA.Replace(sCodeNew, VBA.Chr$(34) & arrData(i, 2) & VBA.Chr$(34), VBA.Chr$(34) & arrData(i, 3) & VBA.Chr$(34))
640:                End If
641:            End If
642:            'если в таблице всего одна запись
643:            If iCount = 1 Then
644:                Call SetCodeInModule(oVBCMod, sCode, sCodeNew)
645:            Else
646:                If arrData(i, 3) <> vbNullString Then
647:                    sCodeNew = VBA.Replace(sCodeNew, VBA.Chr$(34) & arrData(i, 2) & VBA.Chr$(34), VBA.Chr$(34) & arrData(i, 3) & VBA.Chr$(34))
648:                End If
649:                If arrData(i, 1) <> arrData(i + k, 1) Or i = iCount Then
650:                    Call SetCodeInModule(oVBCMod, sCode, sCodeNew)
651:                    Set oVBCMod = oVBP.VBComponents(arrData(i + k, 1)).CodeModule
652:                    sCode = GetCodeFromModule(oVBCMod)
653:                    sCodeNew = sCode
654:                End If
655:            End If
656:        End If
657:    Next i
658: End Sub

     Private Function GetCodeFromModule(ByRef oVBCMod As VBIDE.CodeModule) As String
661:    Dim iLineCode   As Long
662:    Dim sCode       As String
663:    With oVBCMod
664:        iLineCode = .CountOfLines
665:        If iLineCode > 0 Then
666:            sCode = .Lines(1, iLineCode)
667:            If sCode <> vbNullString And sCode Like "*" & VBA.Chr$(34) & "*" Then
668:                GetCodeFromModule = sCode
669:            End If
670:        End If
671:    End With
672: End Function

Private Sub SetCodeInModule(ByRef oVBCMod As VBIDE.CodeModule, ByVal sCode As String, ByVal sCodeNew As String)
675:    If sCode <> sCodeNew Then
676:        Dim iLineCode As Long
677:        With oVBCMod
678:            iLineCode = .CountOfLines
679:            If iLineCode > 0 Then
680:                Call .DeleteLines(1, iLineCode)
681:                Call .InsertLines(1, sCodeNew)
682:            End If
683:        End With
684:    End If
End Sub
