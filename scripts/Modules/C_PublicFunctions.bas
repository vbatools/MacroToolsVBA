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
19:    Dim MyDataObj   As New DataObject
20:    MyDataObj.SetText Txt
21:    MyDataObj.PutInClipboard
22: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : SelectedLineColumnProcedure - получить номера строк и столбцов выделенных строк в модуле VBA
'* Created    : 08-10-2020 13:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function SelectedLineColumnProcedure() As String
25:    Dim lStartLine   As Long
26:    Dim lStartColumn As Long
27:    Dim lEndLine     As Long
28:    Dim lEndColumn   As Long
29:
30:    On Error GoTo ErrorHandler
31:
32:    With Application.VBE.ActiveCodePane
33:        .GetSelection lStartLine, lStartColumn, lEndLine, lEndColumn
34:        SelectedLineColumnProcedure = lStartLine & "|" & lStartColumn & "|" & lEndLine & "|" & lEndColumn
35:    End With
36:    Exit Function
ErrorHandler:
38:    Select Case Err
        Case 91:
40:            Debug.Print "Error!, the module for inserting code is not activated!" & vbNewLine & Err.Number & vbNewLine & Err.Description
41:        Case Else:
42:            Debug.Print "An error occurred in SelectedLineColumnProcedure" & vbNewLine & Err.Number & vbNewLine & Err.Description
43:            Call WriteErrorLog("SelectedLineColumnProcedure")
44:    End Select
45: End Function

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
47:    With Application.FileDialog(msoFileDialogFolderPicker)     ' вывод диалогового окна
48:        .ButtonName = "Choose": .Title = "VBATools": .InitialFileName = sPath
49:         If .Show <> -1 Then Exit Function    ' если пользователь отказался от выбора папки
50:        DirLoadFiles = .SelectedItems(1) & Application.PathSeparator
51:    End With
52: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : Num_Not_Stable определение состояния NumLock
'* Created    : 08-10-2020 13:50
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function Num_Not_Stable() As Boolean
55:    ' Определяет, изменчивое ли состояние у NumLock или нет
56:    ' Возвращает false - стабильный, true - изменчевый
57:    Dim keystat(0 To 255) As Byte
58:    Dim state       As String
59:
60:    GetKeyboardState keystat(0)
61:    state = keystat(vbKeyNumlock)
62:
63:    If (state = 0) Then
64:        Num_Not_Stable = False
65:    Else
66:        Num_Not_Stable = True
67:    End If
68: End Function

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
71:    Dim LR          As LogRecorder
72:    Set LR = New LogRecorder
73:    LR.WriteErrorLog (sNameFunc)
74: End Sub

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
77:    On Error GoTo ErrorHandler
78:
79:    Dim appEX       As Object
80:    Set appEX = CreateObject("Wscript.Shell")
81:    appEX.Run url_str
82:    Set appEX = Nothing
83:    Exit Sub
ErrorHandler:
85:    Select Case Err
        Case Else:
87:            Call MsgBox("An error occurred in URLLinks" & vbNewLine & Err.Number & vbNewLine & Err.Description, vbOKOnly + vbCritical, "Error in URLLinks")
88:            Call WriteErrorLog("URLLinks")
89:    End Select
90:    Set appEX = Nothing
91:    Err.Clear
92: End Sub

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
95:    'sPathFile - строка, полный путь к файлу.
96:    'возвращает размер файла в байтах.
97:    Dim sz          As Long
98:    Dim FSO As Object, objFile As Object
99:    Set FSO = CreateObject("Scripting.FileSystemObject")
100:    Set objFile = FSO.GetFile(sPath)
101:    FileSize = objFile.Size
102:    Set FSO = Nothing: Set objFile = Nothing
103: End Function

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
108:    Dim FSO         As Object
109:    Set FSO = CreateObject("Scripting.FileSystemObject")
110:    sGetExtensionName = FSO.GetExtensionName(sPathFile)
111:    Set FSO = Nothing
112: End Function
     Public Function sGetFileName(ByVal sPathFile As String) As String
114:    'sPathFile - строка, путь.
115:    'возвращает имя (с расширением) последнего компонента в заданном пути.
116:    Dim FSO         As Object
117:    Set FSO = CreateObject("Scripting.FileSystemObject")
118:    sGetFileName = FSO.GetFileName(sPathFile)
119:    Set FSO = Nothing
120: End Function
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
133:    Dim objFso      As Object
134:    Set objFso = CreateObject("Scripting.FileSystemObject")
135:    sGetBaseName = objFso.GetBaseName(sPathFile)
136:    Set objFso = Nothing
137: End Function

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
139:    Dim oFd         As FileDialog
140:    Dim s()         As String
141:    Dim lf          As Long
142:    Set oFd = Application.FileDialog(msoFileDialogFilePicker)
143:
144:    With oFd     'используем короткое обращение к объекту
145:        .AllowMultiSelect = bMultiSelect
146:        .Title = "VBATools: Select an Excel file"     'заголовок окна диалога
147:        .Filters.Clear     'очищаем установленные ранее типы файлов
148:        .Filters.Add "Microsoft Excel Files", ExcelExtens, 1     'устанавливаем возможность выбора только файлов Excel
149:        .InitialFileName = sPath     'назначаем папку отображения и имя файла по умолчанию
150:        .InitialView = msoFileDialogViewDetails     'вид диалогового окна(доступно 9 вариантов)
151:        If .Show = 0 Then
152:            SelectedFile = Empty
153:             Exit Function    'показывает диалог
154:        End If
155:        ReDim Preserve s(1 To .SelectedItems.Count)
156:        For lf = 1 To .SelectedItems.Count
157:            s(lf) = CStr(.SelectedItems.Item(lf))     'считываем полный путь к файлу
158:        Next
159:    End With
160:    SelectedFile = s
161:    Set oFd = Nothing
162: End Function

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
164:    FileHave = (Dir(sPath, Atributes) <> "")
165:    If sPath = vbNullString Then FileHave = False
166: End Function

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
169:    Dim wb          As Workbook
170:    On Error Resume Next
171:    Set wb = Workbooks(wname)
172:    If Err.Number = 0 Then WorkbookIsOpen = True
173: End Function

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
     Public Function IsFileOpen(sfileName As String) As Boolean
176:    Dim filenum As Integer, errnum As Integer
177:
178:    On Error Resume Next
179:    filenum = FreeFile()
180:    ' Attempt to open the file and lock it.
181:    Open sfileName For Input Lock Read As #filenum
182:    Close filenum
183:    errnum = Err
184:    On Error GoTo 0
185:
186:    Select Case errnum
        Case 0
188:            IsFileOpen = False
189:            ' Error number for "Permission Denied."
190:            ' File is already opened by another user.
191:        Case 70
192:            IsFileOpen = True
193:            ' Another error occurred.
194:        Case Else
195:            Error errnum
196:    End Select
197: End Function
     Public Function sFolderHave(ByVal sPathFile As String) As Boolean
199:    'sPathFile - строка, путь.
200:    'возвращает True, если указанный каталог сущесвтвует, и False в противном случае.
201:    Dim FSO         As Object
202:    Set FSO = CreateObject("Scripting.FileSystemObject")
203:    sFolderHave = FSO.FolderExists(sPathFile)
204:    Set FSO = Nothing
205: End Function
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
     Public Sub CopyFileFSO(ByVal sfileName As String, ByVal sNewFileName As String)
219:    Dim objFso As Object, objFile As Object
220:
221:    Set objFso = CreateObject("Scripting.FileSystemObject")
222:    Set objFile = objFso.GetFile(sfileName)
223:    objFile.Copy sNewFileName
224:    Set objFso = Nothing
225: End Sub
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
     Public Function TXTReadALLFile(ByVal sfileName As String, Optional AddFile As Boolean = True) As String
239:
240:    Dim FSO         As Object
241:    Dim ts          As Object
242:    On Error Resume Next: Err.Clear
243:    Set FSO = CreateObject("scripting.filesystemobject")
244:    Set ts = FSO.OpenTextFile(sfileName, 1, AddFile): TXTReadALLFile = ts.ReadAll: ts.Close
245:    Set ts = Nothing: Set FSO = Nothing
246: End Function
     Public Function TXTAddIntoTXTFile(ByVal sfileName As String, ByVal Txt As String, Optional AddFile As Boolean = True) As Boolean
248:    'TXTAddIntoTXTFile - логическа переменая, True - добавление удалось, False - нет
249:    'FileName - строковая переменая, полный путь файла
250:    'txt - текст добавляемый в фаил
251:    'AddFile - логическа переменая, по умолчанию True, если нет файла то создаст его
252:
253:    Dim FSO         As Object
254:    Dim ts          As Object
255:    On Error Resume Next: Err.Clear
256:    Set FSO = CreateObject("scripting.filesystemobject")
257:    Set ts = FSO.OpenTextFile(sfileName, 8, AddFile): ts.Write Txt: ts.Close
258:    TXTAddIntoTXTFile = Err = 0
259:    Set ts = Nothing: Set FSO = Nothing
260: End Function

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
276:    Dim sCode       As String
277:    sCode = sInCode
278:    sCode = VBA.Replace(sCode, " " & sOldName & ".", " " & sNewName & ".", 1, -1, vbTextCompare)
279:    sCode = VBA.Replace(sCode, " " & sOldName & ",", " " & sNewName & ",", 1, -1, vbTextCompare)
280:    sCode = VBA.Replace(sCode, " " & sOldName & ")", " " & sNewName & ")", 1, -1, vbTextCompare)
281:    sCode = VBA.Replace(sCode, "(" & sOldName & ".", "(" & sNewName & ".", 1, -1, vbTextCompare)
282:    sCode = VBA.Replace(sCode, "(" & sOldName & ",", "(" & sNewName & ",", 1, -1, vbTextCompare)
283:    sCode = VBA.Replace(sCode, "=" & sOldName & ".", "=" & sNewName & ".", 1, -1, vbTextCompare)
284:    sCode = VBA.Replace(sCode, "=" & sOldName & vbNewLine, "=" & sNewName & vbNewLine, , , vbTextCompare)
285:    sCode = VBA.Replace(sCode, "(" & sOldName & " ", "(" & sNewName & " ", 1, -1, vbTextCompare)
286:    sCode = VBA.Replace(sCode, "(" & sOldName & ")", "(" & sNewName & ")", 1, -1, vbTextCompare)
287:    sCode = VBA.Replace(sCode, "." & sOldName & ".", "." & sNewName & ".", 1, -1, vbTextCompare)
288:    sCode = VBA.Replace(sCode, "." & sOldName & vbNewLine, "." & sNewName & vbNewLine, , , vbTextCompare)
289:    sCode = VBA.Replace(sCode, " " & sOldName & "_", " " & sNewName & "_", 1, -1, vbTextCompare)
290:    sCode = VBA.Replace(sCode, vbNewLine & sOldName & "_", vbNewLine & sNewName & "_", 1, -1, vbTextCompare)
291:    sCode = VBA.Replace(sCode, """ & sOldName & """, """ & sNewName & """, 1, -1, vbTextCompare)
292:    sCode = VBA.Replace(sCode, " " & sOldName & " ", " " & sNewName & " ", 1, -1, vbTextCompare)
293:    sCode = VBA.Replace(sCode, " " & sOldName & vbNewLine, " " & sNewName & vbNewLine, 1, -1, vbTextCompare)
294:    ReplceCode = sCode
295: End Function

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
297:    Dim sTemp As String
298:    Const LENGHT_CELLS As Long = 32767
299:    sTemp = sTxt
300:    If VBA.Len(sTemp) <= LENGHT_CELLS Then
301:        sTemp = Application.WorksheetFunction.Trim(sTemp)
302:    Else
303:        Dim i As Long
304:        For i = 1 To VBA.Len(sTxt) Step LENGHT_CELLS
305:            sTemp = sTemp & Application.WorksheetFunction.Trim(VBA.Mid$(sTxt, i, LENGHT_CELLS))
306:        Next i
307:    End If
308:    TrimSpace = sTemp
End Function
