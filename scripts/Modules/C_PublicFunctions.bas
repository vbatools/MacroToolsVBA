Attribute VB_Name = "C_PublicFunctions"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : C_PublicFunctions - глобальные функции надстройки
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module

#If Win64 Then
Private Declare PtrSafe Function GetKeyboardState Lib "USER32" (pbKeyState As Byte) As Long
#Else
Private Declare Function GetKeyboardState Lib "USER32" (pbKeyState As Byte) As Long
#End If

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SetTextIntoClipboard - поместить текст в буфер обмена
'* Created    : 08-10-2020 13:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):         Description
'*
'* ByVal txt As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub SetTextIntoClipboard(ByVal Txt As String)
30:    Dim MyDataObj   As New DataObject
31:    MyDataObj.SetText Txt
32:    MyDataObj.PutInClipboard
33: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : SelectedLineColumnProcedure - получить номера строк и столбцов выделенных строк в модуле VBA
'* Created    : 08-10-2020 13:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function SelectedLineColumnProcedure() As String
43:    Dim lStartLine   As Long
44:    Dim lStartColumn As Long
45:    Dim lEndLine     As Long
46:    Dim lEndColumn   As Long
47:
48:    On Error GoTo ErrorHandler
49:
50:    With Application.VBE.ActiveCodePane
51:        .GetSelection lStartLine, lStartColumn, lEndLine, lEndColumn
52:        SelectedLineColumnProcedure = lStartLine & "|" & lStartColumn & "|" & lEndLine & "|" & lEndColumn
53:    End With
54:    Exit Function
ErrorHandler:
56:    Select Case Err
        Case 91:
58:            Debug.Print "Error!, the module for inserting code is not activated!" & vbNewLine & Err.Number & vbNewLine & Err.Description
59:        Case Else:
60:            Debug.Print "An error occurred in SelectedLineColumnProcedure" & vbNewLine & Err.Number & vbNewLine & Err.Description
61:            Call WriteErrorLog("SelectedLineColumnProcedure")
62:    End Select
63: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DirLoadFiles - Диалоговое окно выбора директории
'* Created    : 08-10-2020 13:49
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sPath As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function DirLoadFiles(ByVal sPath As String) As String
77:    With Application.FileDialog(msoFileDialogFolderPicker)     ' вывод диалогового окна
78:        .ButtonName = "Choose": .Title = "VBATools": .InitialFileName = sPath
79: If .Show <> -1 Then Exit Function   ' если пользователь отказался от выбора папки
80:        DirLoadFiles = .SelectedItems(1) & Application.PathSeparator
81:    End With
82: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : Num_Not_Stable определение состояния NumLock
'* Created    : 08-10-2020 13:50
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function Num_Not_Stable() As Boolean
92:    ' Определяет, изменчивое ли состояние у NumLock или нет
93:    ' Возвращает false - стабильный, true - изменчевый
94:    Dim keystat(0 To 255) As Byte
95:    Dim state       As String
96:
97:    GetKeyboardState keystat(0)
98:    state = keystat(vbKeyNumlock)
99:
100:    If (state = 0) Then
101:        Num_Not_Stable = False
102:    Else
103:        Num_Not_Stable = True
104:    End If
105: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : WriteErrorLog - процедура ведения Log файла
'* Created    : 08-10-2020 13:51
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal sNameFunc As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub WriteErrorLog(ByVal sNameFunc As String)
119:    Dim LR          As LogRecorder
120:    Set LR = New LogRecorder
121:    LR.WriteErrorLog (sNameFunc)
122: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : URLLinks открытие в брайзере URL ссылки
'* Created    : 08-10-2020 13:51
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal url_str As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub URLLinks(ByVal url_str As String)
136:    On Error GoTo ErrorHandler
137:
138:    Dim appEX       As Object
139:    Set appEX = CreateObject("Wscript.Shell")
140:    appEX.Run url_str
141:    Set appEX = Nothing
142:    Exit Sub
ErrorHandler:
144:    Select Case Err
        Case Else:
146:            Call MsgBox("An error occurred in the URL Links" & vbNewLine & Err.Number & vbNewLine & Err.Description, vbOKOnly + vbCritical, "Error in URL Links")
147:            Call WriteErrorLog("URLLinks")
148:    End Select
149:    Set appEX = Nothing
150:    Err.Clear
151: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : FileSize - определить размер файла
'* Created    : 08-10-2020 13:51
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):     Description
'*
'* sPath As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function FileSize(sPath As String) As Long
165:    'sPathFile - строка, полный путь к файлу.
166:    'возвращает размер файла в байтах.
167:    Dim sz          As Long
168:    Dim FSO As Object, objFile As Object
169:    Set FSO = CreateObject("Scripting.FileSystemObject")
170:    Set objFile = FSO.GetFile(sPath)
171:    FileSize = objFile.Size
172:    Set FSO = Nothing: Set objFile = Nothing
173: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : sGetExtensionName возвращает расширение последнего компонента в заданном пути
'* Created    : 08-10-2020 13:52
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal sPathFile As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function sGetExtensionName(ByVal sPathFile As String) As String
187:    Dim FSO         As Object
188:    Set FSO = CreateObject("Scripting.FileSystemObject")
189:    sGetExtensionName = FSO.GetExtensionName(sPathFile)
190:    Set FSO = Nothing
191: End Function
     Public Function sGetFileName(ByVal sPathFile As String) As String
193:    'sPathFile - строка, путь.
194:    'возвращает имя (с расширением) последнего компонента в заданном пути.
195:    Dim FSO         As Object
196:    Set FSO = CreateObject("Scripting.FileSystemObject")
197:    sGetFileName = FSO.GetFileName(sPathFile)
198:    Set FSO = Nothing
199: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : sGetBaseName -  возвращает имя (без расширения) последнего компонента в заданном пути.
'* Created    : 04-03-2020 13:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal sPathFile As String : строка, путь
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function sGetBaseName(ByVal sPathFile As String) As String
212:    Dim objFSO      As Object
213:    Set objFSO = CreateObject("Scripting.FileSystemObject")
214:    sGetBaseName = objFSO.GetBaseName(sPathFile)
215:    Set objFSO = Nothing
216: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : SelectedFile - диалоговое окно выбора файлов из директории
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                                 Description
'*
'* ByVal sPath As String                    :
'* Optional bMultiSelect As Boolean = True  :
'* Optional ExcelExtens As String = "*.xl*" :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function SelectedFile(ByVal sPath As String, Optional bMultiSelect As Boolean = True, Optional ExcelExtens As String = "*.xl*") As Variant
232:    Dim oFd         As FileDialog
233:    Dim s()         As String
234:    Dim lf          As Long
235:    Set oFd = Application.FileDialog(msoFileDialogFilePicker)
236:
237:    With oFd     'используем короткое обращение к объекту
238:        .AllowMultiSelect = bMultiSelect
239:        .Title = "VBA Tools: Select an Excel file"     'заголовок окна диалога
240:        .Filters.Clear     'очищаем установленные ранее типы файлов
241:        .Filters.Add "Microsoft Excel Files", ExcelExtens, 1     'устанавливаем возможность выбора только файлов Excel
242:        .InitialFileName = sPath     'назначаем папку отображения и имя файла по умолчанию
243:        .InitialView = msoFileDialogViewDetails     'вид диалогового окна(доступно 9 вариантов)
244:        If .Show = 0 Then
245:            SelectedFile = Empty
246: Exit Function   'показывает диалог
247:        End If
248:        ReDim Preserve s(1 To .SelectedItems.Count)
249:        For lf = 1 To .SelectedItems.Count
250:            s(lf) = CStr(.SelectedItems.Item(lf))     'считываем полный путь к файлу
251:        Next
252:    End With
253:    SelectedFile = s
254:    Set oFd = Nothing
255: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : FileHave - проверка существования файла
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                                     Description
'*
'* sPath As String                                :
'* Optional Atributes As FileAttribute = vbNormal :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function FileHave(sPath As String, Optional Atributes As FileAttribute = vbNormal) As Boolean
270:    FileHave = (Dir(sPath, Atributes) <> "")
271:    If sPath = vbNullString Then FileHave = False
272: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WorkbookIsOpen - Возвращает ИСТИНА если открыта книга под названием wname
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef wname As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function WorkbookIsOpen(ByRef wname As String) As Boolean
286:    Dim WB          As Workbook
287:    On Error Resume Next
288:    Set WB = Workbooks(wname)
289:    If Err.Number = 0 Then WorkbookIsOpen = True
290: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : IsFileOpen - если файл открыт то закрывает его
'* Created    : 08-10-2020 13:54
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):         Description
'*
'* sFileName As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function IsFileOpen(sFileName As String) As Boolean
304:    Dim filenum As Integer, errnum As Integer
305:
306:    On Error Resume Next
307:    filenum = FreeFile()
308:    ' Attempt to open the file and lock it.
309:    Open sFileName For Input Lock Read As #filenum
310:    Close filenum
311:    errnum = Err
312:    On Error GoTo 0
313:
314:    Select Case errnum
        Case 0
316:            IsFileOpen = False
317:            ' Error number for "Permission Denied."
318:            ' File is already opened by another user.
319:        Case 70
320:            IsFileOpen = True
321:            ' Another error occurred.
322:        Case Else
323:            Error errnum
324:    End Select
325: End Function
     Public Function sFolderHave(ByVal sPathFile As String) As Boolean
327:    'sPathFile - строка, путь.
328:    'возвращает True, если указанный каталог сущесвтвует, и False в противном случае.
329:    Dim FSO         As Object
330:    Set FSO = CreateObject("Scripting.FileSystemObject")
331:    sFolderHave = FSO.FolderExists(sPathFile)
332:    Set FSO = Nothing
333: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CopyFile - копирование файла
'* Created    : 04-03-2020 13:37
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                     Description
'*
'* ByVal sFileName As String    : от куда
'* ByVal sNewFileName As String : куда
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub CopyFileFSO(ByVal sFileName As String, ByVal sNewFileName As String)
347:    Dim objFSO As Object, objFile As Object
348:
349:    Set objFSO = CreateObject("Scripting.FileSystemObject")
350:    Set objFile = objFSO.GetFile(sFileName)
351:    objFile.Copy sNewFileName
352:    Set objFSO = Nothing
353: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : TXTReadALLFile строковая переменая, возращающает текст файла
'* Created    : 06-03-2020 10:07
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                         Description
'*
'* ByVal FileName As String           : строковая переменая, полный путь файла
'* Optional AddFile As Boolean = True : логическа переменая, по умолчанию True, если нет файла то создаст его
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function TXTReadALLFile(ByVal sFileName As String, Optional AddFile As Boolean = True) As String
367:
368:    Dim FSO         As Object
369:    Dim ts          As Object
370:    On Error Resume Next: Err.Clear
371:    Set FSO = CreateObject("scripting.filesystemobject")
372:    Set ts = FSO.OpenTextFile(sFileName, 1, AddFile): TXTReadALLFile = ts.ReadAll: ts.Close
373:    Set ts = Nothing: Set FSO = Nothing
374: End Function
     Public Function TXTAddIntoTXTFile(ByVal sFileName As String, ByVal Txt As String, Optional AddFile As Boolean = True) As Boolean
376:    'TXTAddIntoTXTFile - логическа переменая, True - добавление удалось, False - нет
377:    'FileName - строковая переменая, полный путь файла
378:    'txt - текст добавляемый в фаил
379:    'AddFile - логическа переменая, по умолчанию True, если нет файла то создаст его
380:
381:    Dim FSO         As Object
382:    Dim ts          As Object
383:    On Error Resume Next: Err.Clear
384:    Set FSO = CreateObject("scripting.filesystemobject")
385:    Set ts = FSO.OpenTextFile(sFileName, 8, AddFile): ts.Write Txt: ts.Close
386:    TXTAddIntoTXTFile = Err = 0
387:    Set ts = Nothing: Set FSO = Nothing
388: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ReplceCode - функция поиска в коде названий и замена их на новое
'* Created    : 26-03-2020 13:11
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sInCode As String : код модуля
'* sOldName As String      : старое имя
'* sNewName As String      : новое имя
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function ReplceCode(ByVal sInCode As String, sOldName As String, sNewName As String) As String
404:    Dim sCode       As String
405:    sCode = sInCode
406:    sCode = VBA.Replace(sCode, " " & sOldName & ".", " " & sNewName & ".", 1, -1, vbTextCompare)
407:    sCode = VBA.Replace(sCode, " " & sOldName & ",", " " & sNewName & ",", 1, -1, vbTextCompare)
408:    sCode = VBA.Replace(sCode, " " & sOldName & ")", " " & sNewName & ")", 1, -1, vbTextCompare)
409:    sCode = VBA.Replace(sCode, "(" & sOldName & ".", "(" & sNewName & ".", 1, -1, vbTextCompare)
410:    sCode = VBA.Replace(sCode, "(" & sOldName & ",", "(" & sNewName & ",", 1, -1, vbTextCompare)
411:    sCode = VBA.Replace(sCode, "=" & sOldName & ".", "=" & sNewName & ".", 1, -1, vbTextCompare)
412:    sCode = VBA.Replace(sCode, "=" & sOldName & vbNewLine, "=" & sNewName & vbNewLine, , , vbTextCompare)
413:    sCode = VBA.Replace(sCode, "(" & sOldName & " ", "(" & sNewName & " ", 1, -1, vbTextCompare)
414:    sCode = VBA.Replace(sCode, "(" & sOldName & ")", "(" & sNewName & ")", 1, -1, vbTextCompare)
415:    sCode = VBA.Replace(sCode, "." & sOldName & ".", "." & sNewName & ".", 1, -1, vbTextCompare)
416:    sCode = VBA.Replace(sCode, "." & sOldName & vbNewLine, "." & sNewName & vbNewLine, , , vbTextCompare)
417:    sCode = VBA.Replace(sCode, " " & sOldName & "_", " " & sNewName & "_", 1, -1, vbTextCompare)
418:    sCode = VBA.Replace(sCode, vbNewLine & sOldName & "_", vbNewLine & sNewName & "_", 1, -1, vbTextCompare)
419:    sCode = VBA.Replace(sCode, """ & sOldName & """, """ & sNewName & """, 1, -1, vbTextCompare)
420:    sCode = VBA.Replace(sCode, " " & sOldName & " ", " " & sNewName & " ", 1, -1, vbTextCompare)
421:    sCode = VBA.Replace(sCode, " " & sOldName & vbNewLine, " " & sNewName & vbNewLine, 1, -1, vbTextCompare)
422:    ReplceCode = sCode
423: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : TrimSpace - удаление всех не одиночных пробелов
'* Created    : 08-10-2020 13:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sTxt As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function TrimSpace(ByVal sTxt As String) As String
437:    Dim sTemp As String
438:    Const LENGHT_CELLS As Long = 32767
439:    sTemp = sTxt
440:    If VBA.Len(sTemp) <= LENGHT_CELLS Then
441:        sTemp = Application.WorksheetFunction.Trim(sTemp)
442:    Else
443:        Dim i As Long
444:        For i = 1 To VBA.Len(sTxt) Step LENGHT_CELLS
445:            sTemp = sTemp & Application.WorksheetFunction.Trim(VBA.Mid$(sTxt, i, LENGHT_CELLS))
446:        Next i
447:    End If
448:    TrimSpace = sTemp
End Function
