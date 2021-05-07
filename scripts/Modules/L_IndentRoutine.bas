Attribute VB_Name = "L_IndentRoutine"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : L_IndentRoutine - форматирование кода МИФ
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'***************************************************************************
'*
'* PROJECT NAME:    SMART INDENTER
'* AUTHOR:          STEPHEN BULLEN, Office Automation Ltd.
'*
'*                  COPYRIGHT В© 1999-2004 BY OFFICE AUTOMATION LTD
'*
'* CONTACT:         stephen@oaltd.co.uk
'* WEB SITE:        http://www.oaltd.co.uk
'*
'* DESCRIPTION:     Adds items to the VBE environment to recreate the indenting
'*                  for the current procedure, module or project.
'*
'* THIS MODULE:     Contains the main procedure to rebuild the code's indenting
'*
'* PROCEDURES:
'*   RebuildModule      Копирует модуль кода в массив для перестроения и резервная копия
'*   RebuildCodeArray   Основная процедура для форматирования кода
'*   fnFindFirstItem    Проверьте, содержит ли строка кода какие-либо специальные служебные слова
'*   CheckLine          Добавить или удалить отступ
'*   ArrayFromVariant   Преобразование массива variant в массив string для более быстрой обработки
'*   fnAlignFunction    Поиск где необходим отступ для продолжения
'*
'***************************************************************************
'*
'* CHANGE HISTORY
'*
'*  DATE        NAME                DESCRIPTION
'*  14/07/1999  Stephen Bullen      Initial version
'*  14/04/2000  Stephen Bullen      Improved algorithm, added options and split out module handling
'*  03/05/2000  Stephen Bullen      Added option to not indent Dims and handle line numbers
'*  24/05/2000  Stephen Bullen      Improved routine for aligning continued lines
'*  27/05/2000  Stephen Bullen      Fix comments with Type/Enum, Rem handling and brackets in strings
'*  04/07/2000  Stephen Bullen      Fix handling of aligned 'As' items and continued lines
'*  24/11/2000  Stephen Bullen      Added maintenance of Members' attributes for VB5 and 6
'*  07/10/2004  Stephen Bullen      Changed to Office Automation
'*  09/10/2004  Stephen Bullen      Bug fixes and more options
'*
'***************************************************************************
Option Explicit
Option Private Module
Option Compare Binary
Option Base 1
'UDT to store Undo information
Public Type uUndo
    oMod       As CodeModule
    sName      As String
    lStartLine As Long
    lEndLine   As Long
    asOriginal() As String
    asIndented() As String
End Type
Public pauUndo() As uUndo
Const miTAB    As Integer = 9
Public piUndoCount As Integer
'переменые массва to hold the code items to look for
'Variant arrays to hold the code items to look for
Dim masInProc() As String, masInCode() As String, masOutProc() As String, masOutCode() As String
Dim masDeclares() As String, masLookFor() As String, masFnAlign() As String
'Переменные для хранения наших вариантов отступов
Dim mbIndentProc As Boolean, mbIndentCmt As Boolean, mbIndentCase As Boolean, mbAlignCont As Boolean, mbIndentDim As Boolean
Dim mbIndentFirst As Boolean, mbAlignDim As Boolean, mbDebugCol1 As Boolean, mbEnableUndo As Boolean
Dim miIndentSpaces As Integer, miEOLAlignCol As Integer, miAlignDimCol As Integer, mbCompilerStuffCol1 As Boolean
Dim mbIndentCompilerStuff As Boolean, mbAlignIgnoreOps As Boolean
'Переменные для хранения оперативной информации
'Variables to hold operational information
Dim mbInitialised As Boolean, mbContinued As Boolean, mbInIf As Boolean, mbNoIndent As Boolean, mbFirstProcLine As Boolean
Dim msEOLComment As String
     Public Sub ReBild()
78:    Dim moCM   As CodeModule
79:    Dim cmb_txt As String
80:    Dim vbComp As VBIDE.VBComponent
81:    On Error GoTo ErrorHandler
82:    cmb_txt = B_CreateMenus.WhatIsTextInComboBoxHave
83:    Select Case cmb_txt
        Case C_Const.ALLVBAPROJECT:
85:            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
86:                Set moCM = vbComp.CodeModule
87:                Call RebuildModule(moCM, moCM.Parent.Name, 1, moCM.CountOfLines, 0)
88:            Next vbComp
89:        Case C_Const.SELECTEDMODULE:
90:            Set moCM = Application.VBE.ActiveCodePane.CodeModule
91:            Call RebuildModule(moCM, moCM.Parent.Name, 1, moCM.CountOfLines, 0)
92:    End Select
93:    Exit Sub
ErrorHandler:
95:    Select Case Err.Number
        Case 91:
97:            Exit Sub
98:        Case Else:
99:            Debug.Print "Error in Rebuild" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
100:            Call WriteErrorLog("ReBild")
101:    End Select
102:    Err.Clear
103: End Sub
''''''''''''''''''''''''''''''''''
' Function:   RebuildModule
'
' Comments:   This procedure goes through the lines in a module,
'             rebuilding the code's indenting.
'
' Arguments:  modCode    - The code module to indent
'             sName      - The display name of the item being indented
'             iStartLine - Value giving the line to start indenting from
'             iEndLine   - Value giving the line to end indenting at
'             iProgDone  - Value giving how much indenting has been done in total
'
     Private Sub RebuildModule( _
                  ByRef modCode As CodeModule, _
                  ByRef sName As String, _
                  ByRef iStartLine As Long, _
                  ByRef iEndline As Long, _
                  ByRef iProgDone As Long, _
                  Optional ByRef mbEnableUndo As Boolean = True)
123:    Dim asCode() As String, asOriginal() As String, i As Long
124: If iEndline = 0 Then Exit Sub   'On Error Resume Next
125:    ReDim asCode(0 To iEndline - iStartLine)
126:    ReDim asOriginal(0 To iEndline - iStartLine)
127:    'Это позволило отменить? Если это так, настройте наше хранилище
128:    If mbEnableUndo Then
129:        piUndoCount = piUndoCount + 1
130:        'Make some space in our undo array
131:        If piUndoCount = 1 Then
132:            ReDim pauUndo(1 To 1)
133:        Else
134:            ReDim Preserve pauUndo(1 To piUndoCount)
135:        End If
136:        'Store the undo information
137:        With pauUndo(piUndoCount)
138:            Set .oMod = modCode
139:            .sName = sName
140:            .lStartLine = iStartLine
141:            .lEndLine = iEndline
142:            ReDim .asIndented(0 To iEndline - iStartLine)
143:            ReDim .asOriginal(0 To iEndline - iStartLine)
144:        End With
145:    End If
146:    'Read code module into an array and store the original code in our undo array
147:    For i = 0 To iEndline - iStartLine
148:        asCode(i) = modCode.Lines(iStartLine + i, 1)
149:        asOriginal(i) = asCode(i)
150:        If mbEnableUndo Then pauUndo(piUndoCount).asOriginal(i) = asCode(i)
151:    Next
152:    'Indent the array, showing the progress
153:    RebuildCodeArray asCode, sName, iProgDone
154:    'Copy the changed code back into the module and store in our undo array
155:    For i = 0 To iEndline - iStartLine
156:        If asOriginal(i) <> asCode(i) Then
157:                On Error Resume Next
158:            modCode.ReplaceLine iStartLine + i, asCode(i)
159:                On Error GoTo 0
160:        End If
161:        If mbEnableUndo Then pauUndo(piUndoCount).asIndented(i) = asCode(i)
162:    Next
163: End Sub
     Public Sub RebuildCodeArray( _
                  ByRef asCodeLines() As String, _
                  ByRef sName As String, _
                  ByRef iProgDone As Long)
168:    'Переменные, используемые для кода отступа
169:    'Variables used for the indenting code
170:    Dim X As Integer, i As Integer, j As Integer, k As Integer, iGap As Integer, iLineAdjust As Integer
171:    Dim lLineCount As Long, iCommentStart As Long, iStart As Long, iScan As Long, iDebugAdjust As Integer
172:    Dim iIndents As Integer, iIndentNext As Integer, iIn As Integer, iOut As Integer
173:    Dim iFunctionStart As Long, iParamStart As Long
174:    Dim bInCmt As Boolean, bProcStart As Boolean, bAlign As Boolean, bFirstCont As Boolean
175:    Dim bAlreadyPadded As Boolean, bFirstDim As Boolean
176:    Dim sLine As String, sLeft As String, sRight As String, sMatch As String, sItem As String
177:    Dim vaScope As Variant, vaStatic As Variant, vaType As Variant, vaInProc As Variant
178:    Dim iCodeLineNum As Long, sCodeLineNum As String, sOrigLine As String
179:    Dim OptionsTb As ListObject
180:    Set OptionsTb = SHSNIPPETS.ListObjects(C_Const.TB_OPTIONSIDEDENT)
181:    On Error Resume Next
182:    With OptionsTb.ListColumns(2)
183:        mbNoIndent = False
184:        mbInIf = False
185:        'Read the indenting options from the registry
186:        miIndentSpaces = .Range(2, 1)    'Read VB's own setting for tab width
187:        mbIndentProc = .Range(3, 1)
188:        mbIndentFirst = .Range(4, 1)
189:        mbIndentDim = .Range(5, 1)
190:        mbIndentCmt = .Range(6, 1)
191:        mbIndentCase = .Range(7, 1)
192:        mbAlignCont = .Range(8, 1)
193:        mbAlignIgnoreOps = .Range(9, 1)
194:        mbDebugCol1 = .Range(10, 1)
195:        mbAlignDim = .Range(11, 1)
196:        miAlignDimCol = .Range(12, 1)
197:
198:        mbCompilerStuffCol1 = .Range(13, 1)
199:        mbIndentCompilerStuff = .Range(14, 1)
200:
201:        msEOLComment = .Range(15, 1)
202:        miEOLAlignCol = .Range(16, 1)
203:    End With
204:
205:    If mbCompilerStuffCol1 = True Or mbIndentCompilerStuff = True Then
206:        mbInitialised = False
207:    End If
208:
209:    ' Create the list of items to match for the indenting at procedure level
210:    If Not mbInitialised Then
211:        vaScope = Array(vbNullString, "Public ", "Private ", "Friend ")
212:        vaStatic = Array(vbNullString, "Static ")
213:        vaType = Array("Sub", "Function", "Property Let", "Property Get", "Property Set", "Type", "Enum")
214:        X = 1
215:        ReDim vaInProc(1)
216:        For i = 1 To UBound(vaScope)
217:            For j = 1 To UBound(vaStatic)
218:                For k = 1 To UBound(vaType)
219:                    ReDim Preserve vaInProc(X)
220:                    vaInProc(X) = vaScope(i) & vaStatic(j) & vaType(k)
221:                    X = X + 1
222:                Next
223:            Next
224:        Next
225:        ArrayFromVariant masInProc, vaInProc
226:        'Items to match when outdenting at procedure level
227:        ArrayFromVariant masOutProc, Array("End Sub", "End Function", "End Property", "End Type", "End Enum")
228:        If mbIndentCompilerStuff Then
229:            'Items to match when indenting within a procedure
230:            ArrayFromVariant masInCode, Array("If", "ElseIf", "Else", "#If", "#ElseIf", "#Else", "Select Case", "Case", "With", "For", "Do", "While")
231:            'Items to match when outdenting within a procedure
232:            ArrayFromVariant masOutCode, Array("ElseIf", "Else", "End If", "#ElseIf", "#Else", "#End If", "Case", "End Select", "End With", "Next", "Loop", "Wend")
233:        Else
234:            'Items to match when indenting within a procedure
235:            ArrayFromVariant masInCode, Array("If", "ElseIf", "Else", "Select Case", "Case", "With", "For", "Do", "While")
236:            'Items to match when outdenting within a procedure
237:            ArrayFromVariant masOutCode, Array("ElseIf", "Else", "End If", "Case", "End Select", "End With", "Next", "Loop", "Wend")
238:        End If
239:        'Items to match for declarations
240:        ArrayFromVariant masDeclares, Array("Dim", "Const", "Static", "Public", "Private", "#Const")
241:        'Things to look for within a line of code for special handling
242:        ArrayFromVariant masLookFor, Array("""", ": ", " As ", "'", "Rem ", "Stop ", "#If ", "#ElseIf ", "#Else ", "#End If ", "#Const ", "Debug.Print ", "Debug.Assert ")
243:        mbInitialised = True
244:    End If
245:    'Things to skip when finding the function start of a line
246:    ArrayFromVariant masFnAlign, Array("Set ", "Let ", "LSet ", "RSet ", "Declare Function", "Declare Sub", "Private Declare Function", "Private Declare Sub", "Public Declare Function", "Public Declare Sub")
247:    If masInCode(UBound(masInCode)) <> "Select Case" And mbIndentCase Then
248:        'If extra-indenting within Select Case, ensure that we have two items in the arrays
249:        ReDim Preserve masInCode(UBound(masInCode) + 1)
250:        masInCode(UBound(masInCode)) = "Select Case"
251:        ReDim Preserve masOutCode(UBound(masOutCode) + 1)
252:        masOutCode(UBound(masOutCode)) = "End Select"
253:    ElseIf masInCode(UBound(masInCode)) = "Select Case" And Not mbIndentCase Then
254:        'If not extra-indenting within Select Case, ensure that we have one item in the arrays
255:        ReDim Preserve masInCode(UBound(masInCode) - 1)
256:        ReDim Preserve masOutCode(UBound(masOutCode) - 1)
257:    End If
258:    'Flag if the lines are at the top of a procedure
259:    bProcStart = False
260:    bFirstDim = False
261:    bFirstCont = True
262:    'Loop through all the lines to indent
263:    For lLineCount = LBound(asCodeLines) To UBound(asCodeLines)
264:        iLineAdjust = 0
265:        bAlreadyPadded = False
266:        iCodeLineNum = -1
267:        sOrigLine = asCodeLines(lLineCount)
268:        'Read the line of code to indent
269:        sLine = Trim$(asCodeLines(lLineCount))
270:        'If we're not in a continued line, initialise some variables
271:        If Not (mbContinued Or bInCmt) Then
272:            mbFirstProcLine = False
273:            iIndentNext = 0
274:            iCommentStart = 0
275:            iIndents = iIndents + iDebugAdjust
276:            iDebugAdjust = 0
277:            iFunctionStart = 0
278:            iParamStart = 0
279:            i = InStr(1, sLine, " ")
280:            If i > 0 Then
281:                If IsNumeric(Left$(sLine, i - 1)) Then
282:                    iCodeLineNum = Val(Left$(sLine, i - 1))
283:                    sLine = Trim$(Mid$(sLine, i + 1))
284:                    sOrigLine = Space(i) & Mid$(sOrigLine, i + 1)
285:                End If
286:            End If
287:        End If
288:        'Is there anything on the line?
289:        If Len(sLine) > 0 Then
290:            ' Remove leading Tabs
291:            Do Until Left$(sLine, 1) <> Chr$(miTAB)
292:                sLine = Mid$(sLine, 2)
293:            Loop
294:            ' Add an extra space on the end
295:            sLine = sLine & " "
296:            If bInCmt Then
297:                'Within a multi-line comment - indent to line up the comment text
298:                sLine = Space$(iCommentStart) & sLine
299:                'Remember if we're in a continued comment line
300:                bInCmt = Right$(Trim$(sLine), 2) = " _"
301:                GoTo PTR_REPLACE_LINE
302:            End If
303:            'Remember the position of the line segment
304:            iStart = 1
305:            iScan = 0
306:            If mbContinued And mbAlignCont Then
307:                If mbAlignIgnoreOps And Left$(sLine, 2) = ", " Then iParamStart = iFunctionStart - 2
308:                If mbAlignIgnoreOps And (Mid$(sLine, 2, 1) = " " Or Left$(sLine, 2) = ":=") And Left$(sLine, 2) <> ", " Then
309:                    sLine = Space$(iParamStart - 3) & sLine
310:                    iLineAdjust = iLineAdjust + iParamStart - 3
311:                    iScan = iScan + iParamStart - 3
312:                Else
313:                    sLine = Space$(iParamStart - 1) & sLine
314:                    iLineAdjust = iLineAdjust + iParamStart - 1
315:                    iScan = iScan + iParamStart - 1
316:                End If
317:                bAlreadyPadded = True
318:            End If
319:            'Scan through the line, character by character, checking for
320:            'strings, multi-statement lines and comments
321:            Do
322:                iScan = iScan + 1
323:                sItem = fnFindFirstItem(sLine, iScan)
324:                Select Case sItem
                    Case vbNullString
326:                        iScan = iScan + 1
327:                        'Nothing found => Skip the rest of the line
328:                        GoTo PTR_NEXT_PART
329:                    Case """"
330:                        'Start of a string => Jump to the end of it
331:                        iScan = InStr(iScan + 1, sLine, """")
332:                        If iScan = 0 Then iScan = Len(sLine) + 1
333:                    Case ": "
334:                        'A multi-statement line separator => Tidy up and continue
335:                        If Right$(Left$(sLine, iScan), 6) <> " Then:" Then
336:                            sLine = Left$(sLine, iScan + 1) & Trim$(Mid$(sLine, iScan + 2))
337:                            'And check the indenting for the line segment
338:                            CheckLine Mid$(sLine, iStart, iScan - 1), iIn, iOut, bProcStart
339:                            If bProcStart Then bFirstDim = True
340:                            If iStart = 1 Then
341:                                iIndents = iIndents - iOut
342:                                If iIndents < 0 Then iIndents = 0
343:                                iIndentNext = iIndentNext + iIn
344:                            Else
345:                                iIndentNext = iIndentNext + iIn - iOut
346:                            End If
347:                        End If
348:                        'Update the pointer and continue
349:                        iStart = iScan + 2
350:                    Case " As "
351:                        'An " As " in a declaration => Line up to required column
352:                        If mbAlignDim Then
353:                            bAlign = mbNoIndent    'Don't need to check within Type
354:                            If Not bAlign Then
355:                                ' Check if we start with a declaration item
356:                                For i = LBound(masDeclares) To UBound(masDeclares)
357:                                    sMatch = masDeclares(i) & " "
358:                                    If Left$(sLine, Len(sMatch)) = sMatch Then
359:                                        bAlign = True
360:                                        Exit For
361:                                    End If
362:                                Next
363:                            End If
364:                            If bAlign Then
365:                                i = InStr(iScan + 3, sLine, " As ")
366:                                If i = 0 Then
367:                                    'OK to indent
368:                                    If mbIndentProc And bFirstDim And Not mbIndentDim And Not mbNoIndent Then
369:                                        iGap = miAlignDimCol - Len(RTrim$(Left$(sLine, iScan)))
370:                                        'Adjust for a line number at the start of the line
371:                                        If iCodeLineNum > -1 Then iGap = iGap - Len(CStr(iCodeLineNum)) - 1
372:                                    Else
373:                                        iGap = miAlignDimCol - Len(RTrim$(Left$(sLine, iScan))) - iIndents * miIndentSpaces
374:                                        'Adjust for a line number at the start of the line
375:                                        If iCodeLineNum > -1 Then
376:                                            If Len(CStr(iCodeLineNum)) >= iIndents * miIndentSpaces Then
377:                                                iGap = iGap - (Len(CStr(iCodeLineNum)) - iIndents * miIndentSpaces) - 1
378:                                            End If
379:                                        End If
380:                                    End If
381:                                    If iGap < 1 Then iGap = 1
382:                                Else
383:                                    'Multiple declarations on the line, so don't space out
384:                                    iGap = 1
385:                                End If
386:                                'Work out the new spacing
387:                                sLeft = RTrim$(Left$(sLine, iScan))
388:                                sLine = sLeft & Space$(iGap) & Mid$(sLine, iScan + 1)
389:                                'Update the counters
390:                                iLineAdjust = iLineAdjust + iGap + Len(sLeft) - iScan
391:                                iScan = Len(sLeft) + iGap + 3
392:                            End If
393:                        Else
394:                            'Not aligning Dims, so remove any existing spacing
395:                            iScan = Len(RTrim$(Left$(sLine, iScan)))
396:                            sLine = RTrim$(Left$(sLine, iScan)) & " " & Trim$(Mid$(sLine, iScan + 1))
397:                            iScan = iScan + 3
398:                        End If
399:                    Case "'", "Rem "
400:                        'The start of a comment => Handle end-of-line comments properly
401:                        If iScan = 1 Then
402:                            'New comment at start of line
403:                            If bProcStart And Not mbIndentFirst And Not mbNoIndent Then
404:                                'No indenting
405:                            ElseIf mbIndentCmt Or bProcStart Or mbNoIndent Then
406:                                'Inside the procedure, so indent to align with code
407:                                sLine = Space$(iIndents * miIndentSpaces) & sLine
408:                                iCommentStart = iScan + iIndents * miIndentSpaces
409:                            ElseIf iIndents > 0 And mbIndentProc And Not bProcStart Then
410:                                'At the top of the procedure, so indent once if required
411:                                sLine = Space$(miIndentSpaces) & sLine
412:                                iCommentStart = iScan + miIndentSpaces
413:                            End If
414:                        Else
415:                            'New comment at the end of a line
416:                            'Make sure it's a proper 'Rem'
417:                            If sItem = "Rem " And Mid$(sLine, iScan - 1, 1) <> " " And Mid$(sLine, iScan - 1, 1) <> ":" Then GoTo PTR_NEXT_PART
418:                            'Check the indenting of the previous code segment
419:                            CheckLine Mid$(sLine, iStart, iScan - 1), iIn, iOut, bProcStart
420:                            If bProcStart Then bFirstDim = True
421:                            If iStart = 1 Then
422:                                iIndents = iIndents - iOut
423:                                If iIndents < 0 Then iIndents = 0
424:                                iIndentNext = iIndentNext + iIn
425:                            Else
426:                                iIndentNext = iIndentNext + iIn - iOut
427:                            End If
428:                            'Get the text before the comment, and the comment text
429:                            sLeft = Trim$(Left$(sLine, iScan - 1))
430:                            sRight = Trim$(Mid$(sLine, iScan))
431:                            'Indent the code part of the line
432:                            If bAlreadyPadded Then
433:                                sLine = RTrim$(Left$(sLine, iScan - 1))
434:                            Else
435:                                If mbContinued Then
436:                                    sLine = Space$((iIndents + 2) * miIndentSpaces) & sLeft
437:                                Else
438:                                    If mbIndentProc And bFirstDim And Not mbIndentDim Then
439:                                        sLine = sLeft
440:                                    Else
441:                                        sLine = Space$(iIndents * miIndentSpaces) & sLeft
442:                                    End If
443:                                End If
444:                            End If
445:                            mbContinued = (Right$(Trim$(sLine), 2) = " _")
446:                            'How do we handle end-of-line comments?
447:                            Select Case msEOLComment
                                Case "Absolute"
449:                                    iScan = iScan - iLineAdjust + Len(sOrigLine) - Len(LTrim$(sOrigLine))
450:                                    iGap = iScan - Len(sLine) - 1
451:                                Case "SameGap"
452:                                    iScan = iScan - iLineAdjust + Len(sOrigLine) - Len(LTrim$(sOrigLine))
453:                                    iGap = iScan - Len(RTrim$(Left$(sOrigLine, iScan - 1))) - 1
454:                                Case "StandardGap"
455:                                    iGap = miIndentSpaces * 2
456:                                Case "AlignInCol"
457:                                    iGap = miEOLAlignCol - Len(sLine) - 1
458:                            End Select
459:                            'Adjust for a line number at the start of the line
460:                            If iCodeLineNum > -1 Then
461:                                Select Case msEOLComment
                                    Case "Absolute", "AlignInCol"
463:                                        If Len(CStr(iCodeLineNum)) >= iIndents * miIndentSpaces Then
464:                                            iGap = iGap - (Len(CStr(iCodeLineNum)) - iIndents * miIndentSpaces) - 1
465:                                        End If
466:                                End Select
467:                            End If
468:                            If iGap < 2 Then iGap = miIndentSpaces
469:                            iCommentStart = Len(sLine) + iGap
470:                            'Put the comment in the required column
471:                            sLine = sLine & Space$(iGap) & sRight
472:                        End If
473:                        'Work out where the text of the comment starts, to align the next line
474:                        If Mid$(sLine, iCommentStart, 4) = "Rem " Then iCommentStart = iCommentStart + 3
475:                        If Mid$(sLine, iCommentStart, 1) = "'" Then iCommentStart = iCommentStart + 1
476:                        Do Until Mid$(sLine, iCommentStart, 1) <> " "
477:                            iCommentStart = iCommentStart + 1
478:                        Loop
479:                        iCommentStart = iCommentStart - 1
480:                        'Adjust for a line number at the start of the line
481:                        If iCodeLineNum > -1 Then
482:                            If Len(CStr(iCodeLineNum)) >= iIndents * miIndentSpaces Then
483:                                iCommentStart = iCommentStart + (Len(CStr(iCodeLineNum)) - iIndents * miIndentSpaces) + 1
484:                            End If
485:                        End If
486:                        'Remember if we're in a continued comment line
487:                        bInCmt = Right$(Trim$(sLine), 2) = " _"
488:                        'Rest of line is comment, so no need to check any more
489:                        GoTo PTR_REPLACE_LINE
490:                    Case "Stop ", "Debug.Print ", "Debug.Assert "
491:                        'A debugging statement - do we want to force to column 1?
492:                        If mbDebugCol1 And iStart = 1 And iScan = 1 Then
493:                            iLineAdjust = iLineAdjust - (Len(sOrigLine) - LTrim$(Len(sOrigLine)))
494:                            iDebugAdjust = iIndents
495:                            iIndents = 0
496:                        End If
497:                    Case "#If ", "#ElseIf ", "#Else ", "#End If ", "#Const "
498:                        'Do we want to force compiler directives to column 1?
499:                        If mbCompilerStuffCol1 And iStart = 1 And iScan = 1 Then
500:                            iLineAdjust = iLineAdjust - (Len(sOrigLine) - LTrim$(Len(sOrigLine)))
501:                            iDebugAdjust = iIndents
502:                            iIndents = 0
503:                        End If
504:                End Select
PTR_NEXT_PART:
506:            Loop Until iScan > Len(sLine)    'Part of the line
507:            'Do we have some code left to check?
508:            '(i.e. a line without a comment or the last segment of a multi-statement line)
509:            If iStart < Len(sLine) Then
510:                If Not mbContinued Then bProcStart = False
511:                'Check the indenting of the remaining code segment
512:                CheckLine Mid$(sLine, iStart), iIn, iOut, bProcStart
513:                If bProcStart Then bFirstDim = True
514:                If iStart = 1 Then
515:                    iIndents = iIndents - iOut
516:                    If iIndents < 0 Then iIndents = 0
517:                    iIndentNext = iIndentNext + iIn
518:                Else
519:                    iIndentNext = iIndentNext + iIn - iOut
520:                End If
521:            End If
522:            'Start from the left at each procedure start
523:            If mbFirstProcLine Then iIndents = 0
524:            ' What about line continuations?  Here, I indent the continued line by
525:            ' two indents, and check for the end of the continuations.  Note
526:            ' that Excel won't allow comments in the middle of line continuations
527:            ' and that comments are treated differently above.
528:            If mbContinued Then
529:                If Not mbAlignCont Then
530:                    sLine = Space$((iIndents + 2) * miIndentSpaces) & sLine
531:                End If
532:            Else
533:                ' Check if we start with a declaration item
534:                bAlign = False
535:                If mbIndentProc And bFirstDim And Not mbIndentDim And Not bProcStart Then
536:                    For i = LBound(masDeclares) To UBound(masDeclares)
537:                        sMatch = masDeclares(i) & " "
538:                        If Left$(sLine, Len(sMatch)) = sMatch Then
539:                            bAlign = True
540:                            Exit For
541:                        End If
542:                    Next
543:                End If
544:                'Not a declaration item to left-align, so pad it out
545:                If Not bAlign Then
546:                    If Not bProcStart Then bFirstDim = False
547:                    sLine = Space$(iIndents * miIndentSpaces) & sLine
548:                End If
549:            End If
550:            mbContinued = (Right$(Trim$(sLine), 2) = " _")
551:        End If    'Anything there?
PTR_REPLACE_LINE:
553:        'Add the code line number back in
554:        If iCodeLineNum > -1 Then
555:            sCodeLineNum = CStr(iCodeLineNum)
556:            If Len(Trim$(Left$(sLine, Len(sCodeLineNum) + 1))) = 0 Then
557:                sLine = sCodeLineNum & Mid$(sLine, Len(sCodeLineNum) + 1)
558:            Else
559:                sLine = sCodeLineNum & " " & Trim$(sLine)
560:            End If
561:        End If
562:        asCodeLines(lLineCount) = RTrim$(sLine)
563:        'If it's not a continued line, update the indenting for the following lines
564:        If Not mbContinued Then
565:            iIndents = iIndents + iIndentNext
566:            iIndentNext = 0
567:            If iIndents < 0 Then iIndents = 0
568:        Else
569:            'A continued line, so if we're not in a comment and we want smart continuing,
570:            'work out which to continue from
571:            If mbAlignCont And Not bInCmt Then
572:                If Left$(Trim$(sLine), 2) = "& " Or Left$(Trim$(sLine), 2) = "+ " Then sLine = "  " & sLine
573:                iFunctionStart = fnAlignFunction(sLine, bFirstCont, iParamStart)
574:                If iFunctionStart = 0 Then
575:                    iFunctionStart = (iIndents + 2) * miIndentSpaces
576:                    iParamStart = iFunctionStart
577:                End If
578:            End If
579:        End If
580:        bFirstCont = Not mbContinued
581:    Next
582: End Sub
'
'  Find the first occurrence of one of our key items in the list
'
'    Returns the text of the item found
'    Updates the iFrom parameter to point to the location of the found item
'
     Private Function fnFindFirstItem(ByRef sLine As String, ByRef iFrom As Long) As String
590:    Dim sItem As String, iFirst As Long, iFound As Long, iItem As Integer
591:    On Error Resume Next
592:    'Assume we don't find anything
593:    iFirst = Len(sLine)
594:    'Loop through the items to find within the line
595:    For iItem = LBound(masLookFor) To UBound(masLookFor)
596:        'What to find?
597:        sItem = masLookFor(iItem)
598:        'Is it there?
599:        iFound = InStr(iFrom, sLine, sItem)
600:        'Is it before any other items?
601:        If iFound > 0 And iFound < iFirst Then
602:            iFirst = iFound
603:            fnFindFirstItem = sItem
604:        End If
605:    Next
606:    'Update the location of the found item
607:    iFrom = iFirst
608: End Function
'
'  Check the line (segment) to see if it needs in- or out-denting
'
     Private Function CheckLine( _
                  ByVal sLine As String, _
                  ByRef iIndentNext As Integer, _
                  ByRef iOutdentThis As Integer, _
                  ByRef bProcStart As Boolean)
617:    Dim i As Integer, j As Integer, sMatch As String
618:    On Error Resume Next
619:    'Assume we don't indent or outdent the code
620:    iIndentNext = 0
621:    iOutdentThis = 0
622:    'Tidy up the line
623:    sLine = Trim$(sLine) & " "
624:    'We don't check within Type and Enums
625:    If Not mbNoIndent Then
626:        ' Check for indenting within the code
627:        For i = LBound(masInCode) To UBound(masInCode)
628:            sMatch = masInCode(i)
629:            If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":")) Then
630:                iIndentNext = iIndentNext + 1
631:            End If
632:        Next
633:        ' Check for out-denting within the code
634:        For i = LBound(masOutCode) To UBound(masOutCode)
635:            sMatch = masOutCode(i)
636:            'Check at start of line for 'real' outdenting
637:            If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":" And Mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
638:                iOutdentThis = iOutdentThis + 1
639:            End If
640:        Next
641:    End If
642:    'Check procedure-level indenting
643:    For i = LBound(masInProc) To UBound(masInProc)
644:        sMatch = masInProc(i)
645:        If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":" And Mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
646:            bProcStart = True
647:            mbFirstProcLine = True
648:            'Don't indent within Type or Enum constructs
649:            If Right$(sMatch, 4) = "Type" Or Right$(sMatch, 4) = "Enum" Then
650:                iIndentNext = iIndentNext + 1
651:                mbNoIndent = True
652:            ElseIf mbIndentProc And Not mbNoIndent Then
653:                iIndentNext = iIndentNext + 1
654:            End If
655:            Exit For
656:        End If
657:    Next
658:    'Check procedure-level outdenting
659:    For i = LBound(masOutProc) To UBound(masOutProc)
660:        sMatch = masOutProc(i)
661:        If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":" And Mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
662:            'Don't indent within Type or Enum constructs
663:            If Right$(sMatch, 4) = "Type" Or Right$(sMatch, 4) = "Enum" Or mbIndentProc Then
664:                iOutdentThis = iOutdentThis + 1
665:                mbNoIndent = False
666:            End If
667:            Exit For
668:        End If
669:    Next
670:    'If we're not indenting, no need to consider the special cases
671:    If mbNoIndent Then Exit Function
672:    ' Treat If as a special case.  If anything other than a comment follows
673:    ' the Then, we don't indent
674:    If Left$(sLine, 3) = "If " Or Left$(sLine, 4) = "#If " Or mbInIf Then
675:        If mbInIf Then iIndentNext = 1
676:        'Strip any strings from the line
677:        i = InStr(1, sLine, """")
678:        Do Until i = 0
679:            j = InStr(i + 1, sLine, """")
680:            If j = 0 Then j = Len(sLine)
681:            sLine = Left$(sLine, i - 1) & Mid$(sLine, j + 1)
682:            i = InStr(1, sLine, """")
683:        Loop
684:        'And strip comments
685:        i = InStr(1, sLine, "'")
686:        If i > 0 Then sLine = Left$(sLine, i - 1)
687:        ' Do we have a Then statement in the line .  Adding a space on the
688:        ' end of the test means we can test for Then being both within or
689:        ' at the end of the line
690:        sLine = " " & sLine & " "
691:        i = InStr(1, sLine, " Then ")
692:        ' Allow for line continuations within the If statement
693:        mbInIf = (Right$(Trim$(sLine), 2) = " _")
694:        If i > 0 Then
695:            ' If there's something after the Then, we don't indent the If
696:            If Trim$(Mid$(sLine, i + 5)) <> vbNullString Then iIndentNext = 0
697:            ' No need to check next time around
698:            mbInIf = False
699:        End If
700:        If mbInIf Then iIndentNext = 0
701:    End If
702: End Function
'
' Convert a Variant array to a string array for faster comparisons
'
     Private Sub ArrayFromVariant(ByRef asString() As String, ByRef vaVariant As Variant)
707:    Dim iLow As Integer, iHigh As Integer, i As Integer
708:    On Error Resume Next
709:    iLow = LBound(vaVariant)
710:    iHigh = UBound(vaVariant)
711:    ReDim asString(iLow To iHigh)
712:    For i = iLow To iHigh
713:        asString(i) = vaVariant(i)
714:    Next
715: End Sub
'
' Locate the start of the first parameter on the line
'
     Private Function fnAlignFunction(ByVal sLine As String, ByRef bFirstLine As Boolean, ByRef iParamStart As Long) As Long
720:    Dim iLPad As Integer, iCheck As Long, iBrackets As Long, iChar As Long, sMatch As String, iSpace As Integer
721:    Dim vAlign As Variant, bFound As Boolean, iAlign As Integer
722:    Dim iFirstThisLine As Integer
723:    Static coBrackets As Collection
724:    On Error Resume Next
725:    ReDim vAlign(1 To 2)
726:    If bFirstLine Then Set coBrackets = New Collection
727:    'Convert and numbers at the start of the line to spaces
728:    iChar = InStr(1, sLine, " ")
729:    If iChar > 1 Then
730:        If IsNumeric(Left$(sLine, iChar - 1)) Then
731:            sLine = Mid$(sLine, iChar + 1)
732:            iLPad = iChar
733:        End If
734:    End If
735:    iLPad = iLPad + Len(sLine) - Len(LTrim$(sLine))
736:    iFirstThisLine = coBrackets.Count
737:    sLine = Trim$(sLine)
738:    iCheck = 1
739:    'Skip over stuff that we don't want to locate the start off
740:    For iChar = LBound(masFnAlign) To UBound(masFnAlign)
741:        sMatch = masFnAlign(iChar)
742:        If Left$(sLine, Len(sMatch)) = sMatch Then
743:            iCheck = iCheck + Len(sMatch) + 1
744:            Exit For
745:        End If
746:    Next
747:    iBrackets = 0
748:    iSpace = 999
749:    For iChar = iCheck To Len(sLine)
750:        Select Case Mid$(sLine, iChar, 1)
            Case """"
752:                'A String => jump to the end of it
753:                iChar = InStr(iChar + 1, sLine, """")
754:            Case "("
755:                'Start of another function => remember this position
756:                vAlign(1) = "("
757:                vAlign(2) = iChar + iLPad
758:                coBrackets.Add vAlign
759:                vAlign(1) = ","
760:                vAlign(2) = iChar + iLPad + 1
761:                coBrackets.Add vAlign
762:            Case ")"
763: 'Function finished => Remove back to the previous open bracket
764:                vAlign = coBrackets(coBrackets.Count)
765:                Do Until vAlign(1) = "(" Or coBrackets.Count = iFirstThisLine
766:                    coBrackets.Remove coBrackets.Count
767:                    vAlign = coBrackets(coBrackets.Count)
768:                Loop
769:                If coBrackets.Count > iFirstThisLine Then coBrackets.Remove coBrackets.Count
770:            Case " "
771:                If Mid$(sLine, iChar, 3) = " = " Then
772:                    'Space before an = sign => remember it to align to later
773:                    bFound = False
774:                    For iAlign = 1 To coBrackets.Count
775:                        vAlign = coBrackets(iAlign)
776:                        If vAlign(1) = "=" Or vAlign(1) = " " Then
777:                            bFound = True
778:                            Exit For
779:                        End If
780:                    Next
781:                    If Not bFound Then
782:                        vAlign(1) = "="
783:                        vAlign(2) = iChar + iLPad + 2
784:                        coBrackets.Add vAlign
785:                    End If
786:                ElseIf coBrackets.Count = 0 And iChar < Len(sLine) - 2 Then
787:                    'Space after a name before the end of the line => remember it for later
788:                    vAlign(1) = " "
789:                    vAlign(2) = iChar + iLPad
790:                    coBrackets.Add vAlign
791:                ElseIf iChar > 5 Then
792:                    'Clear the collection if we find a Then in an If...Then and set the
793:                    'indenting to align with the bit after the "If "
794:                    If Mid$(sLine, iChar - 5, 6) = " Then " Then
795:                        Do Until coBrackets.Count <= 1
796:                            coBrackets.Remove coBrackets.Count
797:                        Loop
798:                    End If
799:                End If
800:            Case ","
801:                'Start of a new parameter => remember it to align to
802:                vAlign(1) = ","
803:                vAlign(2) = iChar + iLPad + 2
804:                coBrackets.Add vAlign
805:            Case ":"
806:                If Mid$(sLine, iChar, 2) = ":=" Then
807:                    'A named paremeter => remember to align to after the name
808:                    vAlign(1) = ","
809:                    vAlign(2) = iChar + iLPad + 2
810:                    coBrackets.Add vAlign
811:                ElseIf Mid$(sLine, iChar, 2) = ": " Then
812:                    'A new line section, so clear the brackets
813:                    Set coBrackets = New Collection
814:                    iChar = iChar + 1
815:                End If
816:        End Select
817:    Next
818:    'If we end with a comma or a named parameter, get rid of all other comma alignments
819:    If Right$(Trim$(sLine), 3) = ", _" Or Right$(Trim$(sLine), 4) = ":= _" Then
820:        For iAlign = coBrackets.Count To 1 Step -1
821:            vAlign = coBrackets(iAlign)
822:            If vAlign(1) = "," Then
823:                coBrackets.Remove iAlign
824:            Else
825:                Exit For
826:            End If
827:        Next
828:    End If
829:    'If we end with a "( _", remove it and the space alignment after it
830:    If Right$(Trim$(sLine), 3) = "( _" Then
831:        coBrackets.Remove coBrackets.Count
832:        coBrackets.Remove coBrackets.Count
833:    End If
834:    iParamStart = 0
835:    'Get the position of the unmatched bracket and align to that
836:    For iAlign = 1 To coBrackets.Count
837:        vAlign = coBrackets(iAlign)
838:        If vAlign(1) = "," Then
839:            iParamStart = vAlign(2)
840:        ElseIf vAlign(1) = "(" Then
841:            iParamStart = vAlign(2) + 1
842:        Else
843:            iCheck = vAlign(2)
844:        End If
845:    Next
846:    If iCheck = 1 Or iCheck >= Len(sLine) + iLPad - 1 Then
847:        If coBrackets.Count = 0 And bFirstLine Then
848:            iCheck = miIndentSpaces * 2 + iLPad
849:        Else
850:            iCheck = iLPad
851:        End If
852:    End If
853:    If iParamStart = 0 Then iParamStart = iCheck + 1
854: fnAlignFunction = iCheck + 1
855: End Function

