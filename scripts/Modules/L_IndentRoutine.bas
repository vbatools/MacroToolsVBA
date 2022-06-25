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
70:    Dim moCM   As CodeModule
71:    Dim cmb_txt As String
72:    Dim vbComp As VBIDE.VBComponent
73:    On Error GoTo ErrorHandler
74:    cmb_txt = B_CreateMenus.WhatIsTextInComboBoxHave
75:    Select Case cmb_txt
        Case C_Const.ALLVBAPROJECT:
77:            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
78:                Set moCM = vbComp.CodeModule
79:                Call RebuildModule(moCM, moCM.Parent.Name, 1, moCM.CountOfLines, 0)
80:            Next vbComp
81:        Case C_Const.SELECTEDMODULE:
82:            Set moCM = Application.VBE.ActiveCodePane.CodeModule
83:            Call RebuildModule(moCM, moCM.Parent.Name, 1, moCM.CountOfLines, 0)
84:    End Select
85:    Exit Sub
ErrorHandler:
87:    Select Case Err.Number
        Case 91:
89:            Exit Sub
90:        Case Else:
91:            Debug.Print "Mistake! in ReBild" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
92:            Call WriteErrorLog("ReBild")
93:    End Select
94:    Err.Clear
95: End Sub
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
115:    Dim asCode() As String, asOriginal() As String, i As Long
        If iEndline = 0 Then Exit Sub    'On Error Resume Next
117:    ReDim asCode(0 To iEndline - iStartLine)
118:    ReDim asOriginal(0 To iEndline - iStartLine)
119:    'Это позволило отменить? Если это так, настройте наше хранилище
120:    If mbEnableUndo Then
121:        piUndoCount = piUndoCount + 1
122:        'Make some space in our undo array
123:        If piUndoCount = 1 Then
124:            ReDim pauUndo(1 To 1)
125:        Else
126:            ReDim Preserve pauUndo(1 To piUndoCount)
127:        End If
128:        'Store the undo information
129:        With pauUndo(piUndoCount)
130:            Set .oMod = modCode
131:            .sName = sName
132:            .lStartLine = iStartLine
133:            .lEndLine = iEndline
134:            ReDim .asIndented(0 To iEndline - iStartLine)
135:            ReDim .asOriginal(0 To iEndline - iStartLine)
136:        End With
137:    End If
138:    'Read code module into an array and store the original code in our undo array
139:    For i = 0 To iEndline - iStartLine
140:        asCode(i) = modCode.Lines(iStartLine + i, 1)
141:        asOriginal(i) = asCode(i)
142:        If mbEnableUndo Then pauUndo(piUndoCount).asOriginal(i) = asCode(i)
143:    Next
144:    'Indent the array, showing the progress
145:    RebuildCodeArray asCode, sName, iProgDone
146:    'Copy the changed code back into the module and store in our undo array
147:    For i = 0 To iEndline - iStartLine
148:        If asOriginal(i) <> asCode(i) Then
                On Error Resume Next
149:            modCode.ReplaceLine iStartLine + i, asCode(i)
                On Error GoTo 0
150:        End If
151:        If mbEnableUndo Then pauUndo(piUndoCount).asIndented(i) = asCode(i)
152:    Next
153: End Sub
     Public Sub RebuildCodeArray( _
             ByRef asCodeLines() As String, _
             ByRef sName As String, _
             ByRef iProgDone As Long)
158:    'Переменные, используемые для кода отступа
159:    'Variables used for the indenting code
160:    Dim X As Integer, i As Integer, j As Integer, k As Integer, iGap As Integer, iLineAdjust As Integer
161:    Dim lLineCount As Long, iCommentStart As Long, iStart As Long, iScan As Long, iDebugAdjust As Integer
162:    Dim iIndents As Integer, iIndentNext As Integer, iIn As Integer, iOut As Integer
163:    Dim iFunctionStart As Long, iParamStart As Long
164:    Dim bInCmt As Boolean, bProcStart As Boolean, bAlign As Boolean, bFirstCont As Boolean
165:    Dim bAlreadyPadded As Boolean, bFirstDim As Boolean
166:    Dim sLine As String, sLeft As String, sRight As String, sMatch As String, sItem As String
167:    Dim vaScope As Variant, vaStatic As Variant, vaType As Variant, vaInProc As Variant
168:    Dim iCodeLineNum As Long, sCodeLineNum As String, sOrigLine As String
169:    Dim OptionsTb As ListObject
170:    Set OptionsTb = SHSNIPPETS.ListObjects(C_Const.TB_OPTIONSIDEDENT)
171:    On Error Resume Next
172:    With OptionsTb.ListColumns(2)
173:        mbNoIndent = False
174:        mbInIf = False
175:        'Read the indenting options from the registry
176:        miIndentSpaces = .Range(2, 1)    'Read VB's own setting for tab width
177:        mbIndentProc = .Range(3, 1)
178:        mbIndentFirst = .Range(4, 1)
179:        mbIndentDim = .Range(5, 1)
180:        mbIndentCmt = .Range(6, 1)
181:        mbIndentCase = .Range(7, 1)
182:        mbAlignCont = .Range(8, 1)
183:        mbAlignIgnoreOps = .Range(9, 1)
184:        mbDebugCol1 = .Range(10, 1)
185:        mbAlignDim = .Range(11, 1)
186:        miAlignDimCol = .Range(12, 1)
187:
188:        mbCompilerStuffCol1 = .Range(13, 1)
189:        mbIndentCompilerStuff = .Range(14, 1)
190:
191:        msEOLComment = .Range(15, 1)
192:        miEOLAlignCol = .Range(16, 1)
193:    End With
194:
195:    If mbCompilerStuffCol1 = True Or mbIndentCompilerStuff = True Then
196:        mbInitialised = False
197:    End If
198:
199:    ' Create the list of items to match for the indenting at procedure level
200:    If Not mbInitialised Then
201:        vaScope = Array(vbNullString, "Public ", "Private ", "Friend ")
202:        vaStatic = Array(vbNullString, "Static ")
203:        vaType = Array("Sub", "Function", "Property Let", "Property Get", "Property Set", "Type", "Enum")
204:        X = 1
205:        ReDim vaInProc(1)
206:        For i = 1 To UBound(vaScope)
207:            For j = 1 To UBound(vaStatic)
208:                For k = 1 To UBound(vaType)
209:                    ReDim Preserve vaInProc(X)
210:                    vaInProc(X) = vaScope(i) & vaStatic(j) & vaType(k)
211:                    X = X + 1
212:                Next
213:            Next
214:        Next
215:        ArrayFromVariant masInProc, vaInProc
216:        'Items to match when outdenting at procedure level
217:        ArrayFromVariant masOutProc, Array("End Sub", "End Function", "End Property", "End Type", "End Enum")
218:        If mbIndentCompilerStuff Then
219:            'Items to match when indenting within a procedure
220:            ArrayFromVariant masInCode, Array("If", "ElseIf", "Else", "#If", "#ElseIf", "#Else", "Select Case", "Case", "With", "For", "Do", "While")
221:            'Items to match when outdenting within a procedure
222:            ArrayFromVariant masOutCode, Array("ElseIf", "Else", "End If", "#ElseIf", "#Else", "#End If", "Case", "End Select", "End With", "Next", "Loop", "Wend")
223:        Else
224:            'Items to match when indenting within a procedure
225:            ArrayFromVariant masInCode, Array("If", "ElseIf", "Else", "Select Case", "Case", "With", "For", "Do", "While")
226:            'Items to match when outdenting within a procedure
227:            ArrayFromVariant masOutCode, Array("ElseIf", "Else", "End If", "Case", "End Select", "End With", "Next", "Loop", "Wend")
228:        End If
229:        'Items to match for declarations
230:        ArrayFromVariant masDeclares, Array("Dim", "Const", "Static", "Public", "Private", "#Const")
231:        'Things to look for within a line of code for special handling
232:        ArrayFromVariant masLookFor, Array("""", ": ", " As ", "'", "Rem ", "Stop ", "#If ", "#ElseIf ", "#Else ", "#End If ", "#Const ", "Debug.Print ", "Debug.Assert ")
233:        mbInitialised = True
234:    End If
235:    'Things to skip when finding the function start of a line
236:    ArrayFromVariant masFnAlign, Array("Set ", "Let ", "LSet ", "RSet ", "Declare Function", "Declare Sub", "Private Declare Function", "Private Declare Sub", "Public Declare Function", "Public Declare Sub")
237:    If masInCode(UBound(masInCode)) <> "Select Case" And mbIndentCase Then
238:        'If extra-indenting within Select Case, ensure that we have two items in the arrays
239:        ReDim Preserve masInCode(UBound(masInCode) + 1)
240:        masInCode(UBound(masInCode)) = "Select Case"
241:        ReDim Preserve masOutCode(UBound(masOutCode) + 1)
242:        masOutCode(UBound(masOutCode)) = "End Select"
243:    ElseIf masInCode(UBound(masInCode)) = "Select Case" And Not mbIndentCase Then
244:        'If not extra-indenting within Select Case, ensure that we have one item in the arrays
245:        ReDim Preserve masInCode(UBound(masInCode) - 1)
246:        ReDim Preserve masOutCode(UBound(masOutCode) - 1)
247:    End If
248:    'Flag if the lines are at the top of a procedure
249:    bProcStart = False
250:    bFirstDim = False
251:    bFirstCont = True
252:    'Loop through all the lines to indent
253:    For lLineCount = LBound(asCodeLines) To UBound(asCodeLines)
254:        iLineAdjust = 0
255:        bAlreadyPadded = False
256:        iCodeLineNum = -1
257:        sOrigLine = asCodeLines(lLineCount)
258:        'Read the line of code to indent
259:        sLine = Trim$(asCodeLines(lLineCount))
260:        'If we're not in a continued line, initialise some variables
261:        If Not (mbContinued Or bInCmt) Then
262:            mbFirstProcLine = False
263:            iIndentNext = 0
264:            iCommentStart = 0
265:            iIndents = iIndents + iDebugAdjust
266:            iDebugAdjust = 0
267:            iFunctionStart = 0
268:            iParamStart = 0
269:            i = InStr(1, sLine, " ")
270:            If i > 0 Then
271:                If IsNumeric(Left$(sLine, i - 1)) Then
272:                    iCodeLineNum = Val(Left$(sLine, i - 1))
273:                    sLine = Trim$(Mid$(sLine, i + 1))
274:                    sOrigLine = Space(i) & Mid$(sOrigLine, i + 1)
275:                End If
276:            End If
277:        End If
278:        'Is there anything on the line?
279:        If Len(sLine) > 0 Then
280:            ' Remove leading Tabs
281:            Do Until Left$(sLine, 1) <> Chr$(miTAB)
282:                sLine = Mid$(sLine, 2)
283:            Loop
284:            ' Add an extra space on the end
285:            sLine = sLine & " "
286:            If bInCmt Then
287:                'Within a multi-line comment - indent to line up the comment text
288:                sLine = Space$(iCommentStart) & sLine
289:                'Remember if we're in a continued comment line
290:                bInCmt = Right$(Trim$(sLine), 2) = " _"
291:                GoTo PTR_REPLACE_LINE
292:            End If
293:            'Remember the position of the line segment
294:            iStart = 1
295:            iScan = 0
296:            If mbContinued And mbAlignCont Then
297:                If mbAlignIgnoreOps And Left$(sLine, 2) = ", " Then iParamStart = iFunctionStart - 2
298:                If mbAlignIgnoreOps And (Mid$(sLine, 2, 1) = " " Or Left$(sLine, 2) = ":=") And Left$(sLine, 2) <> ", " Then
299:                    sLine = Space$(iParamStart - 3) & sLine
300:                    iLineAdjust = iLineAdjust + iParamStart - 3
301:                    iScan = iScan + iParamStart - 3
302:                Else
303:                    sLine = Space$(iParamStart - 1) & sLine
304:                    iLineAdjust = iLineAdjust + iParamStart - 1
305:                    iScan = iScan + iParamStart - 1
306:                End If
307:                bAlreadyPadded = True
308:            End If
309:            'Scan through the line, character by character, checking for
310:            'strings, multi-statement lines and comments
311:            Do
312:                iScan = iScan + 1
313:                sItem = fnFindFirstItem(sLine, iScan)
314:                Select Case sItem
                    Case vbNullString
316:                        iScan = iScan + 1
317:                        'Nothing found => Skip the rest of the line
318:                        GoTo PTR_NEXT_PART
319:                    Case """"
320:                        'Start of a string => Jump to the end of it
321:                        iScan = InStr(iScan + 1, sLine, """")
322:                        If iScan = 0 Then iScan = Len(sLine) + 1
323:                    Case ": "
324:                        'A multi-statement line separator => Tidy up and continue
325:                        If Right$(Left$(sLine, iScan), 6) <> " Then:" Then
326:                            sLine = Left$(sLine, iScan + 1) & Trim$(Mid$(sLine, iScan + 2))
327:                            'And check the indenting for the line segment
328:                            CheckLine Mid$(sLine, iStart, iScan - 1), iIn, iOut, bProcStart
329:                            If bProcStart Then bFirstDim = True
330:                            If iStart = 1 Then
331:                                iIndents = iIndents - iOut
332:                                If iIndents < 0 Then iIndents = 0
333:                                iIndentNext = iIndentNext + iIn
334:                            Else
335:                                iIndentNext = iIndentNext + iIn - iOut
336:                            End If
337:                        End If
338:                        'Update the pointer and continue
339:                        iStart = iScan + 2
340:                    Case " As "
341:                        'An " As " in a declaration => Line up to required column
342:                        If mbAlignDim Then
343:                            bAlign = mbNoIndent    'Don't need to check within Type
344:                            If Not bAlign Then
345:                                ' Check if we start with a declaration item
346:                                For i = LBound(masDeclares) To UBound(masDeclares)
347:                                    sMatch = masDeclares(i) & " "
348:                                    If Left$(sLine, Len(sMatch)) = sMatch Then
349:                                        bAlign = True
350:                                        Exit For
351:                                    End If
352:                                Next
353:                            End If
354:                            If bAlign Then
355:                                i = InStr(iScan + 3, sLine, " As ")
356:                                If i = 0 Then
357:                                    'OK to indent
358:                                    If mbIndentProc And bFirstDim And Not mbIndentDim And Not mbNoIndent Then
359:                                        iGap = miAlignDimCol - Len(RTrim$(Left$(sLine, iScan)))
360:                                        'Adjust for a line number at the start of the line
361:                                        If iCodeLineNum > -1 Then iGap = iGap - Len(CStr(iCodeLineNum)) - 1
362:                                    Else
363:                                        iGap = miAlignDimCol - Len(RTrim$(Left$(sLine, iScan))) - iIndents * miIndentSpaces
364:                                        'Adjust for a line number at the start of the line
365:                                        If iCodeLineNum > -1 Then
366:                                            If Len(CStr(iCodeLineNum)) >= iIndents * miIndentSpaces Then
367:                                                iGap = iGap - (Len(CStr(iCodeLineNum)) - iIndents * miIndentSpaces) - 1
368:                                            End If
369:                                        End If
370:                                    End If
371:                                    If iGap < 1 Then iGap = 1
372:                                Else
373:                                    'Multiple declarations on the line, so don't space out
374:                                    iGap = 1
375:                                End If
376:                                'Work out the new spacing
377:                                sLeft = RTrim$(Left$(sLine, iScan))
378:                                sLine = sLeft & Space$(iGap) & Mid$(sLine, iScan + 1)
379:                                'Update the counters
380:                                iLineAdjust = iLineAdjust + iGap + Len(sLeft) - iScan
381:                                iScan = Len(sLeft) + iGap + 3
382:                            End If
383:                        Else
384:                            'Not aligning Dims, so remove any existing spacing
385:                            iScan = Len(RTrim$(Left$(sLine, iScan)))
386:                            sLine = RTrim$(Left$(sLine, iScan)) & " " & Trim$(Mid$(sLine, iScan + 1))
387:                            iScan = iScan + 3
388:                        End If
389:                    Case "'", "Rem "
390:                        'The start of a comment => Handle end-of-line comments properly
391:                        If iScan = 1 Then
392:                            'New comment at start of line
393:                            If bProcStart And Not mbIndentFirst And Not mbNoIndent Then
394:                                'No indenting
395:                            ElseIf mbIndentCmt Or bProcStart Or mbNoIndent Then
396:                                'Inside the procedure, so indent to align with code
397:                                sLine = Space$(iIndents * miIndentSpaces) & sLine
398:                                iCommentStart = iScan + iIndents * miIndentSpaces
399:                            ElseIf iIndents > 0 And mbIndentProc And Not bProcStart Then
400:                                'At the top of the procedure, so indent once if required
401:                                sLine = Space$(miIndentSpaces) & sLine
402:                                iCommentStart = iScan + miIndentSpaces
403:                            End If
404:                        Else
405:                            'New comment at the end of a line
406:                            'Make sure it's a proper 'Rem'
407:                            If sItem = "Rem " And Mid$(sLine, iScan - 1, 1) <> " " And Mid$(sLine, iScan - 1, 1) <> ":" Then GoTo PTR_NEXT_PART
408:                            'Check the indenting of the previous code segment
409:                            CheckLine Mid$(sLine, iStart, iScan - 1), iIn, iOut, bProcStart
410:                            If bProcStart Then bFirstDim = True
411:                            If iStart = 1 Then
412:                                iIndents = iIndents - iOut
413:                                If iIndents < 0 Then iIndents = 0
414:                                iIndentNext = iIndentNext + iIn
415:                            Else
416:                                iIndentNext = iIndentNext + iIn - iOut
417:                            End If
418:                            'Get the text before the comment, and the comment text
419:                            sLeft = Trim$(Left$(sLine, iScan - 1))
420:                            sRight = Trim$(Mid$(sLine, iScan))
421:                            'Indent the code part of the line
422:                            If bAlreadyPadded Then
423:                                sLine = RTrim$(Left$(sLine, iScan - 1))
424:                            Else
425:                                If mbContinued Then
426:                                    sLine = Space$((iIndents + 2) * miIndentSpaces) & sLeft
427:                                Else
428:                                    If mbIndentProc And bFirstDim And Not mbIndentDim Then
429:                                        sLine = sLeft
430:                                    Else
431:                                        sLine = Space$(iIndents * miIndentSpaces) & sLeft
432:                                    End If
433:                                End If
434:                            End If
435:                            mbContinued = (Right$(Trim$(sLine), 2) = " _")
436:                            'How do we handle end-of-line comments?
437:                            Select Case msEOLComment
                                Case "Absolute"
439:                                    iScan = iScan - iLineAdjust + Len(sOrigLine) - Len(LTrim$(sOrigLine))
440:                                    iGap = iScan - Len(sLine) - 1
441:                                Case "SameGap"
442:                                    iScan = iScan - iLineAdjust + Len(sOrigLine) - Len(LTrim$(sOrigLine))
443:                                    iGap = iScan - Len(RTrim$(Left$(sOrigLine, iScan - 1))) - 1
444:                                Case "StandardGap"
445:                                    iGap = miIndentSpaces * 2
446:                                Case "AlignInCol"
447:                                    iGap = miEOLAlignCol - Len(sLine) - 1
448:                            End Select
449:                            'Adjust for a line number at the start of the line
450:                            If iCodeLineNum > -1 Then
451:                                Select Case msEOLComment
                                    Case "Absolute", "AlignInCol"
453:                                        If Len(CStr(iCodeLineNum)) >= iIndents * miIndentSpaces Then
454:                                            iGap = iGap - (Len(CStr(iCodeLineNum)) - iIndents * miIndentSpaces) - 1
455:                                        End If
456:                                End Select
457:                            End If
458:                            If iGap < 2 Then iGap = miIndentSpaces
459:                            iCommentStart = Len(sLine) + iGap
460:                            'Put the comment in the required column
461:                            sLine = sLine & Space$(iGap) & sRight
462:                        End If
463:                        'Work out where the text of the comment starts, to align the next line
464:                        If Mid$(sLine, iCommentStart, 4) = "Rem " Then iCommentStart = iCommentStart + 3
465:                        If Mid$(sLine, iCommentStart, 1) = "'" Then iCommentStart = iCommentStart + 1
466:                        Do Until Mid$(sLine, iCommentStart, 1) <> " "
467:                            iCommentStart = iCommentStart + 1
468:                        Loop
469:                        iCommentStart = iCommentStart - 1
470:                        'Adjust for a line number at the start of the line
471:                        If iCodeLineNum > -1 Then
472:                            If Len(CStr(iCodeLineNum)) >= iIndents * miIndentSpaces Then
473:                                iCommentStart = iCommentStart + (Len(CStr(iCodeLineNum)) - iIndents * miIndentSpaces) + 1
474:                            End If
475:                        End If
476:                        'Remember if we're in a continued comment line
477:                        bInCmt = Right$(Trim$(sLine), 2) = " _"
478:                        'Rest of line is comment, so no need to check any more
479:                        GoTo PTR_REPLACE_LINE
480:                    Case "Stop ", "Debug.Print ", "Debug.Assert "
481:                        'A debugging statement - do we want to force to column 1?
482:                        If mbDebugCol1 And iStart = 1 And iScan = 1 Then
483:                            iLineAdjust = iLineAdjust - (Len(sOrigLine) - LTrim$(Len(sOrigLine)))
484:                            iDebugAdjust = iIndents
485:                            iIndents = 0
486:                        End If
487:                    Case "#If ", "#ElseIf ", "#Else ", "#End If ", "#Const "
488:                        'Do we want to force compiler directives to column 1?
489:                        If mbCompilerStuffCol1 And iStart = 1 And iScan = 1 Then
490:                            iLineAdjust = iLineAdjust - (Len(sOrigLine) - LTrim$(Len(sOrigLine)))
491:                            iDebugAdjust = iIndents
492:                            iIndents = 0
493:                        End If
494:                End Select
PTR_NEXT_PART:
496:            Loop Until iScan > Len(sLine)    'Part of the line
497:            'Do we have some code left to check?
498:            '(i.e. a line without a comment or the last segment of a multi-statement line)
499:            If iStart < Len(sLine) Then
500:                If Not mbContinued Then bProcStart = False
501:                'Check the indenting of the remaining code segment
502:                CheckLine Mid$(sLine, iStart), iIn, iOut, bProcStart
503:                If bProcStart Then bFirstDim = True
504:                If iStart = 1 Then
505:                    iIndents = iIndents - iOut
506:                    If iIndents < 0 Then iIndents = 0
507:                    iIndentNext = iIndentNext + iIn
508:                Else
509:                    iIndentNext = iIndentNext + iIn - iOut
510:                End If
511:            End If
512:            'Start from the left at each procedure start
513:            If mbFirstProcLine Then iIndents = 0
514:            ' What about line continuations?  Here, I indent the continued line by
515:            ' two indents, and check for the end of the continuations.  Note
516:            ' that Excel won't allow comments in the middle of line continuations
517:            ' and that comments are treated differently above.
518:            If mbContinued Then
519:                If Not mbAlignCont Then
520:                    sLine = Space$((iIndents + 2) * miIndentSpaces) & sLine
521:                End If
522:            Else
523:                ' Check if we start with a declaration item
524:                bAlign = False
525:                If mbIndentProc And bFirstDim And Not mbIndentDim And Not bProcStart Then
526:                    For i = LBound(masDeclares) To UBound(masDeclares)
527:                        sMatch = masDeclares(i) & " "
528:                        If Left$(sLine, Len(sMatch)) = sMatch Then
529:                            bAlign = True
530:                            Exit For
531:                        End If
532:                    Next
533:                End If
534:                'Not a declaration item to left-align, so pad it out
535:                If Not bAlign Then
536:                    If Not bProcStart Then bFirstDim = False
537:                    sLine = Space$(iIndents * miIndentSpaces) & sLine
538:                End If
539:            End If
540:            mbContinued = (Right$(Trim$(sLine), 2) = " _")
541:        End If    'Anything there?
PTR_REPLACE_LINE:
543:        'Add the code line number back in
544:        If iCodeLineNum > -1 Then
545:            sCodeLineNum = CStr(iCodeLineNum)
546:            If Len(Trim$(Left$(sLine, Len(sCodeLineNum) + 1))) = 0 Then
547:                sLine = sCodeLineNum & Mid$(sLine, Len(sCodeLineNum) + 1)
548:            Else
549:                sLine = sCodeLineNum & " " & Trim$(sLine)
550:            End If
551:        End If
552:        asCodeLines(lLineCount) = RTrim$(sLine)
553:        'If it's not a continued line, update the indenting for the following lines
554:        If Not mbContinued Then
555:            iIndents = iIndents + iIndentNext
556:            iIndentNext = 0
557:            If iIndents < 0 Then iIndents = 0
558:        Else
559:            'A continued line, so if we're not in a comment and we want smart continuing,
560:            'work out which to continue from
561:            If mbAlignCont And Not bInCmt Then
562:                If Left$(Trim$(sLine), 2) = "& " Or Left$(Trim$(sLine), 2) = "+ " Then sLine = "  " & sLine
563:                iFunctionStart = fnAlignFunction(sLine, bFirstCont, iParamStart)
564:                If iFunctionStart = 0 Then
565:                    iFunctionStart = (iIndents + 2) * miIndentSpaces
566:                    iParamStart = iFunctionStart
567:                End If
568:            End If
569:        End If
570:        bFirstCont = Not mbContinued
571:    Next
572: End Sub
'
'  Find the first occurrence of one of our key items in the list
'
'    Returns the text of the item found
'    Updates the iFrom parameter to point to the location of the found item
'
     Private Function fnFindFirstItem(ByRef sLine As String, ByRef iFrom As Long) As String
580:    Dim sItem As String, iFirst As Long, iFound As Long, iItem As Integer
581:    On Error Resume Next
582:    'Assume we don't find anything
583:    iFirst = Len(sLine)
584:    'Loop through the items to find within the line
585:    For iItem = LBound(masLookFor) To UBound(masLookFor)
586:        'What to find?
587:        sItem = masLookFor(iItem)
588:        'Is it there?
589:        iFound = InStr(iFrom, sLine, sItem)
590:        'Is it before any other items?
591:        If iFound > 0 And iFound < iFirst Then
592:            iFirst = iFound
593:            fnFindFirstItem = sItem
594:        End If
595:    Next
596:    'Update the location of the found item
597:    iFrom = iFirst
598: End Function
'
'  Check the line (segment) to see if it needs in- or out-denting
'
     Private Function CheckLine( _
             ByVal sLine As String, _
             ByRef iIndentNext As Integer, _
             ByRef iOutdentThis As Integer, _
             ByRef bProcStart As Boolean)
607:    Dim i As Integer, j As Integer, sMatch As String
608:    On Error Resume Next
609:    'Assume we don't indent or outdent the code
610:    iIndentNext = 0
611:    iOutdentThis = 0
612:    'Tidy up the line
613:    sLine = Trim$(sLine) & " "
614:    'We don't check within Type and Enums
615:    If Not mbNoIndent Then
616:        ' Check for indenting within the code
617:        For i = LBound(masInCode) To UBound(masInCode)
618:            sMatch = masInCode(i)
619:            If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":")) Then
620:                iIndentNext = iIndentNext + 1
621:            End If
622:        Next
623:        ' Check for out-denting within the code
624:        For i = LBound(masOutCode) To UBound(masOutCode)
625:            sMatch = masOutCode(i)
626:            'Check at start of line for 'real' outdenting
627:            If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":" And Mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
628:                iOutdentThis = iOutdentThis + 1
629:            End If
630:        Next
631:    End If
632:    'Check procedure-level indenting
633:    For i = LBound(masInProc) To UBound(masInProc)
634:        sMatch = masInProc(i)
635:        If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":" And Mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
636:            bProcStart = True
637:            mbFirstProcLine = True
638:            'Don't indent within Type or Enum constructs
639:            If Right$(sMatch, 4) = "Type" Or Right$(sMatch, 4) = "Enum" Then
640:                iIndentNext = iIndentNext + 1
641:                mbNoIndent = True
642:            ElseIf mbIndentProc And Not mbNoIndent Then
643:                iIndentNext = iIndentNext + 1
644:            End If
645:            Exit For
646:        End If
647:    Next
648:    'Check procedure-level outdenting
649:    For i = LBound(masOutProc) To UBound(masOutProc)
650:        sMatch = masOutProc(i)
651:        If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":" And Mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
652:            'Don't indent within Type or Enum constructs
653:            If Right$(sMatch, 4) = "Type" Or Right$(sMatch, 4) = "Enum" Or mbIndentProc Then
654:                iOutdentThis = iOutdentThis + 1
655:                mbNoIndent = False
656:            End If
657:            Exit For
658:        End If
659:    Next
660:    'If we're not indenting, no need to consider the special cases
661:    If mbNoIndent Then Exit Function
662:    ' Treat If as a special case.  If anything other than a comment follows
663:    ' the Then, we don't indent
664:    If Left$(sLine, 3) = "If " Or Left$(sLine, 4) = "#If " Or mbInIf Then
665:        If mbInIf Then iIndentNext = 1
666:        'Strip any strings from the line
667:        i = InStr(1, sLine, """")
668:        Do Until i = 0
669:            j = InStr(i + 1, sLine, """")
670:            If j = 0 Then j = Len(sLine)
671:            sLine = Left$(sLine, i - 1) & Mid$(sLine, j + 1)
672:            i = InStr(1, sLine, """")
673:        Loop
674:        'And strip comments
675:        i = InStr(1, sLine, "'")
676:        If i > 0 Then sLine = Left$(sLine, i - 1)
677:        ' Do we have a Then statement in the line.  Adding a space on the
678:        ' end of the test means we can test for Then being both within or
679:        ' at the end of the line
680:        sLine = " " & sLine & " "
681:        i = InStr(1, sLine, " Then ")
682:        ' Allow for line continuations within the If statement
683:        mbInIf = (Right$(Trim$(sLine), 2) = " _")
684:        If i > 0 Then
685:            ' If there's something after the Then, we don't indent the If
686:            If Trim$(Mid$(sLine, i + 5)) <> vbNullString Then iIndentNext = 0
687:            ' No need to check next time around
688:            mbInIf = False
689:        End If
690:        If mbInIf Then iIndentNext = 0
691:    End If
692: End Function
'
' Convert a Variant array to a string array for faster comparisons
'
     Private Sub ArrayFromVariant(ByRef asString() As String, ByRef vaVariant As Variant)
697:    Dim iLow As Integer, iHigh As Integer, i As Integer
698:    On Error Resume Next
699:    iLow = LBound(vaVariant)
700:    iHigh = UBound(vaVariant)
701:    ReDim asString(iLow To iHigh)
702:    For i = iLow To iHigh
703:        asString(i) = vaVariant(i)
704:    Next
705: End Sub
'
' Locate the start of the first parameter on the line
'
Private Function fnAlignFunction(ByVal sLine As String, ByRef bFirstLine As Boolean, ByRef iParamStart As Long) As Long
710:    Dim iLPad As Integer, iCheck As Long, iBrackets As Long, iChar As Long, sMatch As String, iSpace As Integer
711:    Dim vAlign As Variant, bFound As Boolean, iAlign As Integer
712:    Dim iFirstThisLine As Integer
713:    Static coBrackets As Collection
714:    On Error Resume Next
715:    ReDim vAlign(1 To 2)
716:    If bFirstLine Then Set coBrackets = New Collection
717:    'Convert and numbers at the start of the line to spaces
718:    iChar = InStr(1, sLine, " ")
719:    If iChar > 1 Then
720:        If IsNumeric(Left$(sLine, iChar - 1)) Then
721:            sLine = Mid$(sLine, iChar + 1)
722:            iLPad = iChar
723:        End If
724:    End If
725:    iLPad = iLPad + Len(sLine) - Len(LTrim$(sLine))
726:    iFirstThisLine = coBrackets.Count
727:    sLine = Trim$(sLine)
728:    iCheck = 1
729:    'Skip over stuff that we don't want to locate the start off
730:    For iChar = LBound(masFnAlign) To UBound(masFnAlign)
731:        sMatch = masFnAlign(iChar)
732:        If Left$(sLine, Len(sMatch)) = sMatch Then
733:            iCheck = iCheck + Len(sMatch) + 1
734:            Exit For
735:        End If
736:    Next
737:    iBrackets = 0
738:    iSpace = 999
739:    For iChar = iCheck To Len(sLine)
740:        Select Case Mid$(sLine, iChar, 1)
            Case """"
742:                'A String => jump to the end of it
743:                iChar = InStr(iChar + 1, sLine, """")
744:            Case "("
745:                'Start of another function => remember this position
746:                vAlign(1) = "("
747:                vAlign(2) = iChar + iLPad
748:                coBrackets.Add vAlign
749:                vAlign(1) = ","
750:                vAlign(2) = iChar + iLPad + 1
751:                coBrackets.Add vAlign
752:            Case ")"
                    'Function finished => Remove back to the previous open bracket
754:                vAlign = coBrackets(coBrackets.Count)
755:                Do Until vAlign(1) = "(" Or coBrackets.Count = iFirstThisLine
756:                    coBrackets.Remove coBrackets.Count
757:                    vAlign = coBrackets(coBrackets.Count)
758:                Loop
759:                If coBrackets.Count > iFirstThisLine Then coBrackets.Remove coBrackets.Count
760:            Case " "
761:                If Mid$(sLine, iChar, 3) = " = " Then
762:                    'Space before an = sign => remember it to align to later
763:                    bFound = False
764:                    For iAlign = 1 To coBrackets.Count
765:                        vAlign = coBrackets(iAlign)
766:                        If vAlign(1) = "=" Or vAlign(1) = " " Then
767:                            bFound = True
768:                            Exit For
769:                        End If
770:                    Next
771:                    If Not bFound Then
772:                        vAlign(1) = "="
773:                        vAlign(2) = iChar + iLPad + 2
774:                        coBrackets.Add vAlign
775:                    End If
776:                ElseIf coBrackets.Count = 0 And iChar < Len(sLine) - 2 Then
777:                    'Space after a name before the end of the line => remember it for later
778:                    vAlign(1) = " "
779:                    vAlign(2) = iChar + iLPad
780:                    coBrackets.Add vAlign
781:                ElseIf iChar > 5 Then
782:                    'Clear the collection if we find a Then in an If...Then and set the
783:                    'indenting to align with the bit after the "If "
784:                    If Mid$(sLine, iChar - 5, 6) = " Then " Then
785:                        Do Until coBrackets.Count <= 1
786:                            coBrackets.Remove coBrackets.Count
787:                        Loop
788:                    End If
789:                End If
790:            Case ","
791:                'Start of a new parameter => remember it to align to
792:                vAlign(1) = ","
793:                vAlign(2) = iChar + iLPad + 2
794:                coBrackets.Add vAlign
795:            Case ":"
796:                If Mid$(sLine, iChar, 2) = ":=" Then
797:                    'A named paremeter => remember to align to after the name
798:                    vAlign(1) = ","
799:                    vAlign(2) = iChar + iLPad + 2
800:                    coBrackets.Add vAlign
801:                ElseIf Mid$(sLine, iChar, 2) = ": " Then
802:                    'A new line section, so clear the brackets
803:                    Set coBrackets = New Collection
804:                    iChar = iChar + 1
805:                End If
806:        End Select
807:    Next
808:    'If we end with a comma or a named parameter, get rid of all other comma alignments
809:    If Right$(Trim$(sLine), 3) = ", _" Or Right$(Trim$(sLine), 4) = ":= _" Then
810:        For iAlign = coBrackets.Count To 1 Step -1
811:            vAlign = coBrackets(iAlign)
812:            If vAlign(1) = "," Then
813:                coBrackets.Remove iAlign
814:            Else
815:                Exit For
816:            End If
817:        Next
818:    End If
819:    'If we end with a "( _", remove it and the space alignment after it
820:    If Right$(Trim$(sLine), 3) = "( _" Then
821:        coBrackets.Remove coBrackets.Count
822:        coBrackets.Remove coBrackets.Count
823:    End If
824:    iParamStart = 0
825:    'Get the position of the unmatched bracket and align to that
826:    For iAlign = 1 To coBrackets.Count
827:        vAlign = coBrackets(iAlign)
828:        If vAlign(1) = "," Then
829:            iParamStart = vAlign(2)
830:        ElseIf vAlign(1) = "(" Then
831:            iParamStart = vAlign(2) + 1
832:        Else
833:            iCheck = vAlign(2)
834:        End If
835:    Next
836:    If iCheck = 1 Or iCheck >= Len(sLine) + iLPad - 1 Then
837:        If coBrackets.Count = 0 And bFirstLine Then
838:            iCheck = miIndentSpaces * 2 + iLPad
839:        Else
840:            iCheck = iLPad
841:        End If
842:    End If
843:    If iParamStart = 0 Then iParamStart = iCheck + 1
        fnAlignFunction = iCheck + 1
End Function

