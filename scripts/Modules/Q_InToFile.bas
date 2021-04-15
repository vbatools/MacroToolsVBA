Attribute VB_Name = "Q_InToFile"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : Q_InToFile - содержимое файла
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module

    Public Sub InToFile()
5:    Dim strPath As String
6:    On Error GoTo errmsg
7:
8:    strPath = O_XML.OpenAndCloseExcelFileInFolder(bOpenFile:=True, bBackUp:=False)
9:    If strPath = vbNullString Then Exit Sub
10:   Q_InToFile.FilenamesCollectionToPath (strPath)
11:
12:    If MsgBox("Delete the folder of the unpacked Excel file" & vbNewLine & "The Excel file is not deleted!", vbYesNo + vbCritical, "Deleting a folder:") = vbYes Then
13:        Call Q_InToFile.RemoveFolderWithContent(strPath)
14:    End If
15:    Exit Sub
errmsg:
17:    Debug.Print "Error in InTo File!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
18:    Call WriteErrorLog("InToFile")
19: End Sub
    Private Sub FilenamesCollectionToPath(ByVal StrPathToFile As String)
21:    ' Ищем на рабочем столе все файлы TXT, и выводим на лист список их имён.
22:    ' Просматриваются папки с глубиной вложения не более трёх.
23:    Dim i      As Long
24:    Dim coll   As Collection
25:    On Error GoTo errmsg
26:    ' считываем в колекцию coll нужные имена файлов
27:    Set coll = FilenamesCollection(StrPathToFile, "*.*", 3)
28:
29:    Application.ScreenUpdating = False    ' отключаем обновление экрана
30:    ' создаём новую книгу
31:    Dim SH     As Worksheet: Set SH = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
32:    ' формируем заголовки таблицы
33:    With SH.Range("a1").Resize(, 5)
34:        .Value = Array("№", "File name", "Full path", "File size", "File extension")
35:        .Font.Bold = True: .Interior.ColorIndex = 17
36:    End With
37:
38:    ' выводим результаты на лист
39:    For i = 1 To coll.Count    ' перебираем все элементы коллекции, содержащей пути к файлам
40:        SH.Range("a" & SH.Rows.Count).End(xlUp).Offset(1).Resize(, 5).Value = _
                      Array(i, C_PublicFunctions.sGetFileName(coll(i)), coll(i), C_PublicFunctions.FileSize(coll(i)), C_PublicFunctions.sGetExtensionName(coll(i)))    ' выводим на лист очередную строку
42:        DoEvents    ' временно передаём управление ОС
43:    Next
44:    SH.Range("a:e").EntireColumn.AutoFit    ' автоподбор ширины столбцов
45:    [a2].Activate: ActiveWindow.FreezePanes = True    ' закрепляем первую строку листа
46:    Application.ScreenUpdating = True    ' отключаем обновление экрана
47:    Exit Sub
errmsg:
49:    Debug.Print "Error in FilenamesCollectionToPath!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
50:    Call WriteErrorLog("FilenamesCollectionToPath")
51: End Sub

    Private Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal Mask As String = "", _
                Optional ByVal SearchDeep As Long = 999) As Collection
56:    ' Получает в качестве параметра путь к папке FolderPath,
57:    ' маску имени искомых файлов Mask (будут отобраны только файлы с такой маской/расширением)
58:    ' и глубину поиска SearchDeep в подпапках (если SearchDeep=1, то подпапки не просматриваются).
59:    ' Возвращает коллекцию, содержащую полные пути найденных файлов
60:    ' (применяется рекурсивный вызов процедуры GetAllFileNamesUsingFSO)
61:    Dim FSO    As Object
62:    On Error GoTo errmsg
63:    Set FilenamesCollection = New Collection    ' создаём пустую коллекцию
64:    Set FSO = CreateObject("Scripting.FileSystemObject")    ' создаём экземпляр FileSystemObject
65:    Call GetAllFileNamesUsingFSO(FolderPath, Mask, FSO, FilenamesCollection, SearchDeep)  ' поиск
66:    Set FSO = Nothing: Application.StatusBar = False    ' очистка строки состояния Excel
67:    Exit Function
errmsg:
69:    Debug.Print "Error in FilenamesCollection!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
70:    Call WriteErrorLog("FilenamesCollection")
71: End Function

    Private Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal Mask As String, ByRef FSO, _
                ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
75:    ' перебирает все файлы и подпапки в папке FolderPath, используя объект FSO
76:    ' перебор папок осуществляется в том случае, если SearchDeep > 1
77:    ' добавляет пути найденных файлов в коллекцию FileNamesColl
78:    Dim curfold As Object, fil As Object, sfol As Object
79:    On Error Resume Next: Set curfold = FSO.GetFolder(FolderPath)
80:    If Not curfold Is Nothing Then    ' если удалось получить доступ к папке
81:
82:        ' раскомментируйте эту строку для вывода пути к просматриваемой
83:        ' в текущий момент папке в строку состояния Excel
84:        ' Application.StatusBar = "Поиск в папке: " & FolderPath
85:
86:        For Each fil In curfold.Files    ' перебираем все файлы в папке FolderPath
87:            If fil.Name Like "*" & Mask Then FileNamesColl.Add fil.Path
88:        Next
89:        SearchDeep = SearchDeep - 1    ' уменьшаем глубину поиска в подпапках
90:        If SearchDeep Then    ' если надо искать глубже
91:            For Each sfol In curfold.SubFolders    ' перебираем все подпапки в папке FolderPath
92:                GetAllFileNamesUsingFSO sfol.Path, Mask, FSO, FileNamesColl, SearchDeep
93:            Next
94:        End If
95:        Set fil = Nothing: Set curfold = Nothing: Set sfol = Nothing   ' очищаем переменные
96:    End If
97: End Function

     Private Sub RemoveFolderWithContent(ByVal sFolder As String)
102: '    'путь к папке можно задать статично, если он заранее известен и не изменяется
104:   Shell "cmd /c rd /S/Q """ & sFolder & """"
105: End Sub

