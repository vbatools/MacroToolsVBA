Attribute VB_Name = "C_PublicFunctions"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : C_PublicFunctions - глобальные функции надстройки
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 11:17    CalDymos
'* Updated    : 12-09-2023 13:28    CalDymos

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
          Dim MyDataObj   As New DataObject
1         MyDataObj.SetText Txt
2         MyDataObj.PutInClipboard
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : SelectedLineColumnProcedure - получить номера строк и столбцов выделенных строк в модуле VBA
'* Created    : 08-10-2020 13:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function SelectedLineColumnProcedure() As String
          Dim lStartLine   As Long
          Dim lStartColumn As Long
          Dim lEndLine     As Long
          Dim lEndColumn   As Long

3         On Error GoTo ErrorHandler

4         With Application.VBE.ActiveCodePane
5             .GetSelection lStartLine, lStartColumn, lEndLine, lEndColumn
6             SelectedLineColumnProcedure = lStartLine & "|" & lStartColumn & "|" & lEndLine & "|" & lEndColumn
7         End With
8         Exit Function
ErrorHandler:
9         Select Case Err
              Case 91:
10                Debug.Print "Error!, the module for inserting code is not activated!" & vbNewLine & Err.Number & vbNewLine & Err.Description
11            Case Else:
12                Debug.Print "An error occurred in SelectedLineColumnProcedure" & vbNewLine & Err.Number & vbNewLine & Err.Description
13                Call WriteErrorLog("SelectedLineColumnProcedure")
14        End Select
End Function

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
15        With Application.FileDialog(msoFileDialogFolderPicker)     ' вывод диалогового окна
16            .ButtonName = "Choose": .Title = "VBATools": .InitialFileName = sPath
17            If .Show <> -1 Then Exit Function    ' если пользователь отказался от выбора папки
18            DirLoadFiles = .SelectedItems(1) & Application.PathSeparator
19        End With
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : Num_Not_Stable определение состояния NumLock
'* Created    : 08-10-2020 13:50
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function Num_Not_Stable() As Boolean
          ' Определяет, изменчивое ли состояние у NumLock или нет
          ' Возвращает false - стабильный, true - изменчевый
          Dim keystat(0 To 255) As Byte
          Dim state       As String

20        GetKeyboardState keystat(0)
21        state = keystat(vbKeyNumlock)

22        If (state = 0) Then
23            Num_Not_Stable = False
24        Else
25            Num_Not_Stable = True
26        End If
End Function

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
          Dim LR          As LogRecorder
27        Set LR = New LogRecorder
28        LR.WriteErrorLog (sNameFunc)
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : URLLinks - URL im Browser цffnen
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
29        On Error GoTo ErrorHandler

          Dim appEX       As Object
30        Set appEX = CreateObject("Wscript.Shell")
31        appEX.Run url_str
32        Set appEX = Nothing
33        Exit Sub
ErrorHandler:
34        Select Case Err
              Case Else:
35                Call MsgBox("An error occurred in URLLinks" & vbNewLine & Err.Number & vbNewLine & Err.Description, vbOKOnly + vbCritical, "Error in URLLinks")
36                Call WriteErrorLog("URLLinks")
37        End Select
38        Set appEX = Nothing
39        Err.Clear
End Sub

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
          'sPathFile - строка, полный путь к файлу.
          'возвращает размер файла в байтах.
          Dim sz          As Long
          Dim FSO As Object, objFile As Object
40        Set FSO = CreateObject("Scripting.FileSystemObject")
41        Set objFile = FSO.GetFile(sPath)
42        FileSize = objFile.Size
43        Set FSO = Nothing: Set objFile = Nothing
End Function

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
          Dim FSO         As Object
44        Set FSO = CreateObject("Scripting.FileSystemObject")
45        sGetExtensionName = FSO.GetExtensionName(sPathFile)
46        Set FSO = Nothing
End Function
Public Function sGetFileName(ByVal sPathFile As String) As String
          'sPathFile - строка, путь.
          'возвращает имя (с расширением) последнего компонента в заданном пути.
          Dim FSO         As Object
47        Set FSO = CreateObject("Scripting.FileSystemObject")
48        sGetFileName = FSO.GetFileName(sPathFile)
49        Set FSO = Nothing
End Function
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
          Dim objFso      As Object
50        Set objFso = CreateObject("Scripting.FileSystemObject")
51        sGetBaseName = objFso.GetBaseName(sPathFile)
52        Set objFso = Nothing
End Function

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
          Dim oFd         As FileDialog
          Dim s()         As String
          Dim lf          As Long
53        Set oFd = Application.FileDialog(msoFileDialogFilePicker)

54        With oFd     'используем короткое обращение к объекту
55            .AllowMultiSelect = bMultiSelect
56            .Title = "VBATools: Select an Excel file"     'заголовок окна диалога
57            .Filters.Clear     'очищаем установленные ранее типы файлов
58            .Filters.Add "Microsoft Excel Files", ExcelExtens, 1     'устанавливаем возможность выбора только файлов Excel
59            .InitialFileName = sPath     'назначаем папку отображения и имя файла по умолчанию
60            .InitialView = msoFileDialogViewDetails     'вид диалогового окна(доступно 9 вариантов)
61            If .Show = 0 Then
62                SelectedFile = Empty
63                Exit Function    'показывает диалог
64            End If
65            ReDim Preserve s(1 To .SelectedItems.Count)
66            For lf = 1 To .SelectedItems.Count
67                s(lf) = CStr(.SelectedItems.Item(lf))     'считываем полный путь к файлу
68            Next
69        End With
70        SelectedFile = s
71        Set oFd = Nothing
End Function

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
72        FileHave = (Dir(sPath, Atributes) <> "")
73        If sPath = vbNullString Then FileHave = False
End Function

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
          Dim wb          As Workbook
74        On Error Resume Next
75        Set wb = Workbooks(wname)
76        If Err.Number = 0 Then WorkbookIsOpen = True
End Function

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
          Dim filenum As Integer, errnum As Integer

77        On Error Resume Next
78        filenum = FreeFile()
          ' Attempt to open the file and lock it.
79        Open sfileName For Input Lock Read As #filenum
80        Close filenum
81        errnum = Err
82        On Error GoTo 0

83        Select Case errnum
              Case 0
84                IsFileOpen = False
                  ' Error number for "Permission Denied."
                  ' File is already opened by another user.
85            Case 70
86                IsFileOpen = True
                  ' Another error occurred.
87            Case Else
88                Error errnum
89        End Select
End Function
Public Function sFolderHave(ByVal sPathFile As String) As Boolean
          'sPathFile - строка, путь.
          'Gibt True zurьck, wenn der angegebene Ordner vorhanden ist, und andernfalls False.
          Dim FSO         As Object
90        Set FSO = CreateObject("Scripting.FileSystemObject")
91        sFolderHave = FSO.FolderExists(sPathFile)
92        Set FSO = Nothing
End Function
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
          Dim objFso As Object, objFile As Object

93        Set objFso = CreateObject("Scripting.FileSystemObject")
94        Set objFile = objFso.GetFile(sfileName)
95        objFile.Copy sNewFileName
96        Set objFso = Nothing
End Sub
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

          Dim FSO         As Object
          Dim ts          As Object
97        On Error Resume Next: Err.Clear
98        Set FSO = CreateObject("scripting.filesystemobject")
99        Set ts = FSO.OpenTextFile(sfileName, 1, AddFile): TXTReadALLFile = ts.ReadAll: ts.Close
100       Set ts = Nothing: Set FSO = Nothing
End Function
Public Function TXTAddIntoTXTFile(ByVal sfileName As String, ByVal Txt As String, Optional AddFile As Boolean = True) As Boolean
          'TXTAddIntoTXTFile - логическа переменая, True - добавление удалось, False - нет
          'FileName - строковая переменая, полный путь файла
          'txt - текст добавляемый в фаил
          'AddFile - логическа переменая, по умолчанию True, если нет файла то создаст его

          Dim FSO         As Object
          Dim ts          As Object
101       On Error Resume Next: Err.Clear
102       Set FSO = CreateObject("scripting.filesystemobject")
103       Set ts = FSO.OpenTextFile(sfileName, 8, AddFile): ts.Write Txt: ts.Close
104       TXTAddIntoTXTFile = Err = 0
105       Set ts = Nothing: Set FSO = Nothing
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ReplceCode - function of searching in the code for names and replacing them with new ones
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
          Dim sCode       As String
106       sCode = sInCode
107       sCode = VBA.Replace(sCode, " " & sOldName & ".", " " & sNewName & ".", 1, -1, vbTextCompare)
108       sCode = VBA.Replace(sCode, " " & sOldName & ",", " " & sNewName & ",", 1, -1, vbTextCompare)
109       sCode = VBA.Replace(sCode, " " & sOldName & ")", " " & sNewName & ")", 1, -1, vbTextCompare)
110       sCode = VBA.Replace(sCode, "(" & sOldName & ".", "(" & sNewName & ".", 1, -1, vbTextCompare)
111       sCode = VBA.Replace(sCode, "(" & sOldName & ",", "(" & sNewName & ",", 1, -1, vbTextCompare)
112       sCode = VBA.Replace(sCode, "=" & sOldName & ".", "=" & sNewName & ".", 1, -1, vbTextCompare)
113       sCode = VBA.Replace(sCode, "=" & sOldName & vbNewLine, "=" & sNewName & vbNewLine, , , vbTextCompare)
114       sCode = VBA.Replace(sCode, "(" & sOldName & " ", "(" & sNewName & " ", 1, -1, vbTextCompare)
115       sCode = VBA.Replace(sCode, "(" & sOldName & ")", "(" & sNewName & ")", 1, -1, vbTextCompare)
116       sCode = VBA.Replace(sCode, "." & sOldName & ".", "." & sNewName & ".", 1, -1, vbTextCompare)
117       sCode = VBA.Replace(sCode, "." & sOldName & vbNewLine, "." & sNewName & vbNewLine, , , vbTextCompare)
118       sCode = VBA.Replace(sCode, " " & sOldName & "_", " " & sNewName & "_", 1, -1, vbTextCompare)
119       sCode = VBA.Replace(sCode, vbNewLine & sOldName & "_", vbNewLine & sNewName & "_", 1, -1, vbTextCompare)
120       sCode = VBA.Replace(sCode, """ & sOldName & """, """ & sNewName & """, 1, -1, vbTextCompare)
121       sCode = VBA.Replace(sCode, " " & sOldName & " ", " " & sNewName & " ", 1, -1, vbTextCompare)
122       sCode = VBA.Replace(sCode, " " & sOldName & vbNewLine, " " & sNewName & vbNewLine, 1, -1, vbTextCompare)
123       ReplceCode = sCode
End Function

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
          Dim sTemp As String
          Const LENGHT_CELLS As Long = 32767
124       sTemp = sTxt
125       If VBA.Len(sTemp) <= LENGHT_CELLS Then
126           sTemp = Application.WorksheetFunction.Trim(sTemp)
127       Else
              Dim i As Long
128           For i = 1 To VBA.Len(sTxt) Step LENGHT_CELLS
129               sTemp = sTemp & Application.WorksheetFunction.Trim(VBA.Mid$(sTxt, i, LENGHT_CELLS))
130           Next i
131       End If
132       TrimSpace = sTemp
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WorksheetExist
'* Created    : 07-09-2023 07:01
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):         Description
'*
'* strName As String :
'* wb As Workbook    :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function WorksheetExist(strName As String, wb As Workbook) As Boolean
133       On Error Resume Next
134       WorksheetExist = wb.Worksheets(strName).index > 0
135       On Error GoTo 0
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DelWorksheet
'* Created    : 07-09-2023 07:01
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):         Description
'*
'* strName As String :
'* wb As Workbook    :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function DelWorksheet(strName As String, wb As Workbook) As Boolean
136       Application.DisplayAlerts = False
137       If WorksheetExist(strName, wb) Then
138           wb.Worksheets(strName).Delete
139           DelWorksheet = Not WorksheetExist(strName, wb)
140       End If
141       Application.DisplayAlerts = True
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : TrimL - Removes all specified characters to the left of the string
'* Created    : 07-09-2023 11:13
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):                         Description
'*
'* ByVal str As String                 :
'* Optional ByVal Char As String = " " :
'* Optional ByVal lCount As Long = 0   :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function TrimL(ByVal str As String, Optional ByVal Char As String = " ", Optional ByVal lCount As Long = 0) As String
       
142       If Char = " " Then
143           TrimL = LTrim$(str)
144       Else
              Dim lLen As Long
       
145           lLen = Len(Char)
146           If lCount > 0 Then
                  Dim i As Long
147               While Len(str) > 0 And Left$(str, lLen) = Char And i < lCount
148                   str = Mid$(str, lLen + 1)
149                   i = i + 1
150               Wend
151           Else
152               While Len(str) > 0 And Left$(str, lLen) = Char
153                   str = Mid$(str, lLen + 1)
154               Wend
155           End If
156       End If
       
157       TrimL = str
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : TrimR - Removes all specified characters to the right of the string
'* Created    : 07-09-2023 11:14
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):                         Description
'*
'* ByVal str As String                 :
'* Optional ByVal Char As String = " " :
'* Optional ByVal lCount As Long = 0   :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function TrimR(ByVal str As String, Optional ByVal Char As String = " ", Optional ByVal lCount As Long = 0) As String
       
158       If Char = " " Then
159           TrimR = RTrim$(str)
160       Else
              Dim lLen As Long
       
161           lLen = Len(Char)
162           If lCount > 0 Then
                  Dim i As Long
163               While Len(str) > 0 And Right$(str, lLen) = Char And i < lCount
164                   str = Left$(str, Len(str) - lLen)
165                   i = i + 1
166               Wend
167           Else
168               While Len(str) > 0 And Right$(str, lLen) = Char
169                   str = Left$(str, Len(str) - lLen)
170               Wend
171           End If
172       End If
       
173       TrimR = str
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : TrimA - Removes all specified characters to the left and right of the string
'* Created    : 07-09-2023 11:14
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):                         Description
'*
'* ByVal str As String                 :
'* Optional ByVal Char As String = " " :
'* Optional ByVal lCount As Long = 0   :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function TrimA(ByVal str As String, Optional ByVal Char As String = " ", Optional ByVal lCount As Long = 0) As String
       
174       TrimA = TrimR(TrimL(str, Char, lCount), Char, lCount)
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : IsArrayEmpty
'* Created    : 13-09-2023 13:28
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function IsArrayEmpty(arr As Variant) As Boolean
       
          Dim b As Integer
       
175       On Error Resume Next
176       b = UBound(arr, 1)
177       IsArrayEmpty = Not (Err.Number = 0)
178       On Error GoTo 0
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ProcInCodeModuleExist
'* Created    : 12-09-2023 12:58
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):             Description
'*
'* objCodeModule As CodeModule :
'* ProcName As String    :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ProcInCodeModuleExist(objCodeModule As CodeModule, ProcName As String) As Boolean
       
179       If Len(ProcName) Then ' Important, otherwise Excel crashes if ProcName is empty and ProcStartLine is called
180           On Error Resume Next
181           ProcInCodeModuleExist = objCodeModule.ProcStartLine(ProcName, vbext_pk_Proc) <> 0
              
182           On Error GoTo 0
183       End If

End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : CodeModuleExist
'* Created    : 12-09-2023 12:58
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):                 Description
'*
'* objWB As Workbook        :
'* CodeModuleName As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function CodeModuleExist(objWB As Workbook, CodeModuleName As String) As Boolean
          Dim objVBCitem  As VBIDE.VBComponent
                
184       For Each objVBCitem In objWB.VBProject.VBComponents
185           If objVBCitem.Name = CodeModuleName Then
186               CodeModuleExist = True
187               Exit Function
188           End If
189       Next

End Function
