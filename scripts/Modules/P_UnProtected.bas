Attribute VB_Name = "P_UnProtected"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : P_UnProtected - удаление пароля с Excel
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Modified   : Date and Time       Author              Description
'* Updated    : 01-10-2019 15:51    VBATools   add module delete Sheets Password
'* Updated    : 30-10-2019 13:32    VBATools   add new function delete and set unviewable Word
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module

Private Const PAGE_EXECUTE_READWRITE = &H40
#If VBA7 Then
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flNewProtect As LongPtr, lpflOldProtect As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function DialogBoxParam Lib "USER32" Alias "DialogBoxParamA" (ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
#Else
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Any)
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Any, ByVal flNewProtect As Any, lpflOldProtect As Any) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Any, ByVal lpProcName As String) As Long
Private Declare Function DialogBoxParam Lib "USER32" Alias "DialogBoxParamA" (ByVal hInstance As Any, ByVal pTemplateName As Any, ByVal hWndParent As Any, ByVal lpDialogFunc As Any, ByVal dwInitParam As Any) As Integer
#End If

Dim HookBytes(0 To 11) As Byte
Dim OriginBytes(0 To 11) As Byte
Dim pFunc           As LongPtr
Dim Flag            As Boolean

    Private Function GetPtr(ByVal lpValue As LongPtr) As LongPtr
36:    GetPtr = lpValue
37: End Function

    Public Sub RecoverBytes()
40:    If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 12
41: End Sub

    Public Function fHook() As Boolean
44:    Dim TmpBytes(0 To 11) As Byte
45:    Dim p As LongPtr, osi As Byte
46:    Dim OriginProtect As LongPtr
47:
48:    fHook = False
49:
#If Win64 Then
51:    osi = 1
#Else
53:    osi = 0
#End If
55:
56:    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
57:
58:    If VirtualProtect(ByVal pFunc, 12, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then
59:
60:        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, osi + 1
61:        If TmpBytes(osi) <> &HB8 Then
62:
63:            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 12
64:
65:            p = GetPtr(AddressOf MyDialogBoxParam)
66:
67:            If osi Then HookBytes(0) = &H48
68:            HookBytes(osi) = &HB8
69:            osi = osi + 1
70:            MoveMemory ByVal VarPtr(HookBytes(osi)), ByVal VarPtr(p), 4 * osi
71:            HookBytes(osi + 4 * osi) = &HFF
72:            HookBytes(osi + 4 * osi + 1) = &HE0
73:
74:            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 12
75:            Flag = True
76:            fHook = True
77:        End If
78:    End If
79: End Function

    Private Function MyDialogBoxParam(ByVal hInstance As LongPtr, _
                ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, _
                ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
84:
85:    If pTemplateName = 4070 Then
86:        MyDialogBoxParam = 1
87:    Else
88:        RecoverBytes
89:        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
                      hWndParent, lpDialogFunc, dwInitParam)
91:        fHook
92:    End If
93: End Function

     Public Sub unprotected()
96:    If MsgBox("Remove passwords from VBA projects ?", vbInformation + vbYesNo, "Removing passwords:") = vbYes Then
97:        If fHook Then
98:            Call MsgBox("Passwords from VBA projects are disabled!", vbInformation, "******")
99:        Else
100:            Call MsgBox("Passwords from VBA projects could not be disabled!", vbInformation, "******")
101:        End If
102:    End If
103: End Sub
'******************************************************************************************************************************************
'Delete Paswort Sheets
     Public Sub DeletePaswortSheets()
107:    Dim sFileName As String, sMsgErr As String, sMsg As String, Msg As String
108:    Dim sFileNameFull As Variant
109:    Dim bFlag       As Boolean
110:    Dim i As Long, LastRow As Long
111:
112:    On Error GoTo errmsg
113:
114:    sFileNameFull = SelectedFile(vbNullString, True, "*.xls;*.xlsm;*.xlsx")
115:    If TypeName(sFileNameFull) = "Empty" Then Exit Sub
116:
117:    If MsgBox("Create backup files ?", vbYesNo + vbQuestion, "Removing passwords:") = vbYes Then
118:        bFlag = True
119:    End If
120:
121:    Application.ScreenUpdating = False
122:    Application.Calculation = xlCalculationManual
123:
124:    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
125:    With ActiveSheet
126:        .Name = "Passwords_" & Replace(Now(), ":", ".")
127:        .Cells(1, 1).Value = "Book Title"
128:        .Cells(1, 2).Value = "Sheet Name"
129:        .Cells(1, 3).Value = "Description"
130:        For i = 1 To UBound(sFileNameFull)
131:            sFileName = sGetFileName(sFileNameFull(i))
132:            If IsFileOpen(CStr(sFileNameFull(i))) = True Then
133:                LastRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
134:                .Cells(LastRow, 1).Value = sFileName
135:                .Cells(LastRow, 2).Value = "Error"
136:                .Cells(LastRow, 3).Value = "The book is not closed, close the book!"
137:                ActiveSheet.Cells(LastRow, 3).Interior.Color = 255
138:            Else
139:                Call XMLFileDelNodes(sFileNameFull(i), bFlag)
140:                DoEvents
141:            End If
142:        Next i
143:        .Range("A:C").EntireColumn.AutoFit
144:        .Range("A1:C1").AutoFilter
145:    End With
146:    Application.Calculation = xlCalculationAutomatic
147:    Application.ScreenUpdating = True
148:    Call MsgBox("Password deletion is over!", vbInformation, "Deleting passwords:")
149:    Exit Sub
errmsg:
151:    Select Case Err.Number
        Case Else
153:            Call MsgBox("Error in DeletePaswortSheets" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbCritical, "Error:")
154:            Call WriteErrorLog("DeletePaswortSheets")
155:    End Select
156:    Application.ScreenUpdating = True
157:    Application.Calculation = xlCalculationAutomatic
158: End Sub

     Private Sub XMLFileDelNodes(ByVal sFileName As String, Optional bBackUp As Boolean = False)
161:    Const PartWorksheets As String = "worksheets\"
162:    Dim i           As Integer
163:    Dim cEditOpenXML As clsEditOpenXML
164:    Dim sWBName     As String
165:    Dim LastRow     As Long
166:
167:    On Error GoTo errmsg
168:
169:    Set cEditOpenXML = New clsEditOpenXML
170:    With cEditOpenXML
171:        .CreateBackupXML = bBackUp
172:        .SourceFile = sFileName
173:        .UnzipFile
174:        i = 1
175:        sWBName = sGetFileName(sFileName)
176:        Do While FileHave(.XLFolder & PartWorksheets & "sheet" & i & ".xml")
177:            LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
178:            ActiveSheet.Cells(LastRow, 1).Value = sWBName
179:            If .DelPartXMLFromFile(PartWorksheets & "sheet" & i & ".xml", sheetProtection) Then
180:                ActiveSheet.Cells(LastRow, 2).Value = .GetSheetNameFromId(CStr(i - 1))
181:                ActiveSheet.Cells(LastRow, 3).Value = "the password from the sheet was removed"
182:                ActiveSheet.Cells(LastRow, 3).Interior.Color = 13434828
183:            Else
184:                ActiveSheet.Cells(LastRow, 2).Value = .GetSheetNameFromId(CStr(i - 1))
185:                ActiveSheet.Cells(LastRow, 3).Value = "and there was no password"
186:                ActiveSheet.Cells(LastRow, 3).Interior.Color = 13421823
187:            End If
188:            i = i + 1
189:        Loop
190:        LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
191:        ActiveSheet.Cells(LastRow, 1).Value = sWBName
192:        If .DelPartXMLFromFile("workbook.xml", workbookProtection) Then
193:            ActiveSheet.Cells(LastRow, 2).Value = sGetFileName(sFileName)
194:            ActiveSheet.Cells(LastRow, 3).Value = "the password was removed from the structure"
195:            ActiveSheet.Cells(LastRow, 3).Interior.Color = 3381555
196:        Else
197:            ActiveSheet.Cells(LastRow, 2).Value = sGetFileName(sFileName)
198:            ActiveSheet.Cells(LastRow, 3).Value = "of the book there was no password on the structure of the book"
199:            ActiveSheet.Cells(LastRow, 3).Interior.Color = 26367
200:        End If
201:        .ZipAllFilesInFolder
202:    End With
203:    Set cEditOpenXML = Nothing
204:    Exit Sub
errmsg:
206:    Select Case Err.Number
        Case Else
208:            Call MsgBox("Error in XMLFileDelNodes" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbCritical, "Error:")
209:            Call WriteErrorLog("XMLFileDelNodes")
210:    End Select
211: End Sub
'******************************************************************************************************************************************
'Delete Paswort In VBAProject with "Unviewable"
     Public Sub DelPasswordVBAProjectUnivable()
215:    Dim varFileNameFull As Variant
216:    Dim sFileNameFull As String
217:
218:    On Error GoTo errmsg
219:
220:    varFileNameFull = SelectedFile(vbNullString, False, "*.xlsm;*.xlsb;*.xlam;*.docm;*.dotm")
221:    If TypeName(varFileNameFull) = "Empty" Then Exit Sub
222:    sFileNameFull = CStr(varFileNameFull(1))
223:    If IsFileOpen(sFileNameFull) Then
224:        Call MsgBox("Close the file for processing!", vbCritical, "Error:")
225:        Exit Sub
226:    End If
227:
228:    Call WriteBinFileVBAProject(sFileNameFull)
229:
230:    Call MsgBox("Deleting the Unicable password is over" & vbNewLine & "Open the file", vbInformation, "Deleting the Unviewable password:")
231:    Exit Sub
errmsg:
233:    Select Case Err.Number
        Case Else
235:            Call MsgBox("Error in DelPasswordVBAProjectUnisible" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbCritical, "Error:")
236:            Call WriteErrorLog("DelPasswordVBAProjectUnivable")
237:    End Select
238: End Sub
     Private Sub WriteBinFileVBAProject(ByVal sFileNameFull As String)
240:    Dim cEditOpenXML As clsEditOpenXML
241:
242:    On Error GoTo errmsg
243:
244:    Set cEditOpenXML = New clsEditOpenXML
245:    With cEditOpenXML
246:        .CreateBackupXML = True
247:        .SourceFile = sFileNameFull
248:        .UnzipFile
249:        .Sheet2Change = "1"
250:        Call WriteBinFile(.XMLFolder(XMLFolder_xl) & .GetSheetFileNameFromId("*/vbaProject", "Type"))
251:        .ZipAllFilesInFolder
252:    End With
253:    Set cEditOpenXML = Nothing
254:    Exit Sub
errmsg:
256:    Select Case Err.Number
        Case Else
258:            Call MsgBox("Error in WriteBinFileVBAProject" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbCritical, "Error:")
259:            Call WriteErrorLog("WriteBinFileVBAProject")
260:    End Select
261:    Set cEditOpenXML = Nothing
262: End Sub

     Private Sub WriteBinFile(ByVal sPath As String)
265:    Dim iFile As Integer, i As Integer
266:    Dim buf         As String
267:    Dim sPattern    As String
268:
269:    Dim sPATTERN1   As String
270:    Dim sPATTERN2   As String
271:    Dim lf          As Long
272:    Dim varBuf As Variant, varBuf1 As Variant
273:
274:    sPattern = "([^" & VBA.Chr$(34) & "])CMG=(" & VBA.Chr$(34) & "|[\n\r])"
275:    sPATTERN1 = "([^" & VBA.Chr$(34) & "])DPB=(" & VBA.Chr$(34) & "|[\n\r])"
276:    sPATTERN2 = "([^" & VBA.Chr$(34) & "])GC=(" & VBA.Chr$(34) & "|[\n\r])"
277:
278:    iFile = FreeFile()
279:    Open sPath For Binary As #iFile
280:
281:    lf& = LOF(iFile)
282:    buf$ = Space$(lf&)
283:
284:    Get #iFile, , buf
285:
286:    buf = W_RegExp.RegExpFindReplace(buf, sPattern, "$1CMC=$2", True, True, True)
287:    buf = W_RegExp.RegExpFindReplace(buf, sPATTERN1, "$1DPC=$2", True, True, True)
288:    buf = W_RegExp.RegExpFindReplace(buf, sPATTERN2, "$1CC=$2", True, True, True)
289:
290:    Seek #iFile, 1
291:    Put #iFile, , buf
292:    Close iFile
293: End Sub

'******************************************************************************************************************************************
'Set Paswort In VBAProject with "Unviewable"
     Public Sub SetPasswordVBAProjectUnviewable()
298:    Dim sFileName   As String
299:    Dim varFileNameFull As Variant
300:    varFileNameFull = SelectedFile(vbNullString, False, "*.xlsm;*.xlsb;*.xlam;*.docm;*.dotm")
301:    If TypeName(varFileNameFull) = "Empty" Then Exit Sub
302:    sFileName = CStr(varFileNameFull(1))
303:
304:    Select Case sGetExtensionName(sFileName)
            'только для Excel
        Case "xlsm", "xlsb", "xlam":
307:            If AddModule(sFileName) = False Then
308:                Call MsgBox("Remove the password from the VBA project!", vbCritical, "File protected:")
309:                Exit Sub
310:            End If
311:    End Select
312:
313:    Call BinFile(sFileName)
314:    Call MsgBox("The file is encrypted!", vbInformation, "Encryption:")
315: End Sub

     Private Sub BinFile(ByVal sFileName As String)
318:    Const printerSettings As String = "printerSettings"
319:    Const printerSettingsBin As String = printerSettings & ".bin"
320:    Const vbaProjectBin As String = "vbaProject.bin"
321:    Dim cEditOpenXML As clsEditOpenXML
322:    Dim BinPath As String, XML As String
323:    Dim sFolder     As String
324:
325:    Set cEditOpenXML = New clsEditOpenXML
326:
327:    With cEditOpenXML
328:        .CreateBackupXML = True
329:        .SourceFile = sFileName
330:        .UnzipFile
331:        sFolder = .XLFolder
332:
333:        'если есть файл printerSettings
334:        If sFolderHave(sFolder & printerSettings) Then
335:            BinPath = sFolder & vbaProjectBin
336:        Else
337:            XML = Replace(.GetXMLFromFile(.TipeFileRels), vbaProjectBin, printerSettingsBin)
338:            .WriteXML2File XML, .TipeFileRels, XMLFolder_xl
339:            BinPath = sFolder & printerSettingsBin
340:            .CopyFiles2 vbaProjectBin, "", BinPath
341:            Call WriteBinError(sFolder & vbaProjectBin)
342:        End If
343:        'редактирую
344:        Call WriteBinFileUnviewable(BinPath)
345:
346:        .ZipAllFilesInFolder
347:    End With
348:    Set cEditOpenXML = Nothing
349: End Sub

     Private Sub WriteBinFileUnviewable(ByVal sPath As String)
352:    Dim buf As String, NewStr As String, BufPart1 As String, BufPart2 As String
353:    Dim i           As Integer
354:
355:
356:    For i = 1 To 100
357:        NewStr = NewStr & "CMG=" & vbNewLine & "DPB=" & vbNewLine & "GC=" & vbNewLine
358:    Next i
359:
360:    buf = GetBinFile(sPath)
361:    BufPart1 = Left(buf, InStr(buf, "CMG=") - 1)
362:    BufPart2 = Right(buf, Len(buf) - InStrRev(buf, "[Host Extender Info]") + 1)
363:
364:    Call PutBinFile(sPath, BufPart1 & vbNewLine & NewStr & BufPart2)
365:
366: End Sub

     Public Function GetBinFile(ByVal sPath As String) As String
369:    Dim iFile       As Integer
370:    Dim lf          As Long
371:    Dim buf         As String
372:
373:    iFile = FreeFile()
374:    Open sPath For Binary As #iFile
375:    lf& = LOF(iFile)
376:    buf = Space$(lf&)
377:    Get #iFile, , buf
378:    GetBinFile = buf
379:    Close iFile
380: End Function

     Public Sub PutBinFile(ByVal sPath As String, ByVal buf As String)
383:    Dim iFile       As Integer
384:    iFile = FreeFile()
385:    Open sPath For Binary As #iFile
386:
387:    Seek #iFile, 1
388:    Put #iFile, , buf
389:    Close iFile
390: End Sub
     Private Sub WriteBinError(ByVal sPath As String)
392:    Dim buf         As String
393:    buf = GetBinFile(sPath)
394:    buf = Replace(buf, "\", "/")
395:    buf = Replace(buf, "'", "/")
396:    buf = Right(buf, Len(buf) - Len(buf) * 0.15)
397:    Call PutBinFile(sPath, buf)
398: End Sub
     Private Function AddModule(ByVal sPath As String) As Boolean
400:    Dim MyExcel     As Application
401:    Dim MyBook As Workbook, EBook As Workbook
402:    Set MyExcel = CreateObject("Excel.Application")
403:
404:    On Error GoTo errmsg
405:
406:    MyExcel.visible = False
407:    MyExcel.DisplayAlerts = False
408:    Set EBook = MyExcel.Workbooks.Open(sPath, False, False)
409:    If EBook.VBProject.Protection = vbext_pp_locked Then
410:        AddModule = False
411:        MyExcel.Quit
412:        Exit Function
413:    End If
414:    Call AddModuleToProject("DPB", vbext_ct_MSForm, MyExcel, "DPB=19078FB4CA017F0FA5AC1A3")
415:    MyExcel.EnableEvents = False
416:    EBook.Close savechanges:=True
417:    MyExcel.EnableEvents = True
418:    MyExcel.Quit
419:    AddModule = True
420:    Exit Function
errmsg:
422:    Select Case Err.Number
        Case Else
424:            Call MsgBox("Error in AddModule" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbCritical, "Error:")
425:            Call WriteErrorLog("AddModule")
426:    End Select
427:    MyExcel.Quit
428:    AddModule = False
429: End Function
     Private Sub AddModuleToProject(ByVal VBName As String, ByVal TypeModule As String, ByRef MyExcel As Application, sCapiton As String)
431:    Dim vbComp      As VBIDE.VBComponent
432:
433:    On Error GoTo errmsg
434:
435:    Set vbComp = MyExcel.VBE.ActiveVBProject.VBComponents.Add(TypeModule)
436:    vbComp.Name = VBName
437:    If TypeModule = vbext_ct_MSForm Then
438:        vbComp.Properties.Item("Caption").Value = sCapiton
439:    End If
440:    Set vbComp = Nothing
441:    Exit Sub
errmsg:
443:    Select Case Err.Number
        Case 50135
445:            'ничего не делаем
446:            Err.Clear
447:        Case Else
448:            Call MsgBox("Error in Add Module To Project" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbCritical, "Error:")
449:            Call WriteErrorLog("AddModuleToProject")
450:    End Select
451:    Set vbComp = Nothing
452: End Sub


