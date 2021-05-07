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
13:    Dim strPath As String
14:    On Error GoTo errmsg
15:
16:    strPath = O_XML.OpenAndCloseExcelFileInFolder(bOpenFile:=True, bBackUp:=False)
17:    If strPath = vbNullString Then Exit Sub
18:   Q_InToFile.FilenamesCollectionToPath (strPath)
19:
20:    If MsgBox("Delete the folder of the unpacked Excel file" & vbNewLine & "The Excel file is not deleted!", vbYesNo + vbCritical, "Deleting a folder:") = vbYes Then
21:        Call Q_InToFile.RemoveFolderWithContent(strPath)
22:    End If
23:    Exit Sub
errmsg:
25:    Debug.Print "Error in InTo File!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
26:    Call WriteErrorLog("InToFile")
27: End Sub
    Private Sub FilenamesCollectionToPath(ByVal StrPathToFile As String)
29:    ' Ищем на рабочем столе все файлы TXT, и выводим на лист список их имён.
30:    ' Просматриваются папки с глубиной вложения не более трёх.
31:    Dim i      As Long
32:    Dim coll   As Collection
33:    On Error GoTo errmsg
34:    ' считываем в колекцию coll нужные имена файлов
35:    Set coll = FilenamesCollection(StrPathToFile, "*.*", 3)
36:
37:    Application.ScreenUpdating = False    ' отключаем обновление экрана
38:    ' создаём новую книгу
39:    Dim SH     As Worksheet: Set SH = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
40:    ' формируем заголовки таблицы
41:    With SH.Range("a1").Resize(, 5)
42:        .Value = Array("№", "File name", "Full path", "File size", "File extension")
43:        .Font.Bold = True: .Interior.ColorIndex = 17
44:    End With
45:
46:    ' выводим результаты на лист
47:    For i = 1 To coll.Count    ' перебираем все элементы коллекции, содержащей пути к файлам
48:        SH.Range("a" & SH.Rows.Count).End(xlUp).Offset(1).Resize(, 5).Value = _
                         Array(i, C_PublicFunctions.sGetFileName(coll(i)), coll(i), C_PublicFunctions.FileSize(coll(i)), C_PublicFunctions.sGetExtensionName(coll(i)))    ' выводим на лист очередную строку
50:        DoEvents    ' временно передаём управление ОС
51:    Next
52:    SH.Range("a:e").EntireColumn.AutoFit    ' автоподбор ширины столбцов
53:    [a2].Activate: ActiveWindow.FreezePanes = True    ' закрепляем первую строку листа
54:    Application.ScreenUpdating = True    ' отключаем обновление экрана
55:    Exit Sub
errmsg:
57:    Debug.Print "Error in FilenamesCollectionToPath!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
58:    Call WriteErrorLog("FilenamesCollectionToPath")
59: End Sub

    Private Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal Mask As String = "", _
                    Optional ByVal SearchDeep As Long = 999) As Collection
63:    ' Получает в качестве параметра путь к папке FolderPath,
64:    ' маску имени искомых файлов Mask (будут отобраны только файлы с такой маской/расширением)
65:    ' и глубину поиска SearchDeep в подпапках (если SearchDeep=1, то подпапки не просматриваются).
66:    ' Возвращает коллекцию, содержащую полные пути найденных файлов
67:    ' (применяется рекурсивный вызов процедуры GetAllFileNamesUsingFSO)
68:    Dim FSO    As Object
69:    On Error GoTo errmsg
70:    Set FilenamesCollection = New Collection    ' создаём пустую коллекцию
71:    Set FSO = CreateObject("Scripting.FileSystemObject")    ' создаём экземпляр FileSystemObject
72:    Call GetAllFileNamesUsingFSO(FolderPath, Mask, FSO, FilenamesCollection, SearchDeep)  ' поиск
73:    Set FSO = Nothing: Application.StatusBar = False    ' очистка строки состояния Excel
74:    Exit Function
errmsg:
76:    Debug.Print "Error in FilenamesCollection!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
77:    Call WriteErrorLog("FilenamesCollection")
78: End Function

     Private Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal Mask As String, ByRef FSO, _
                     ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
82:    ' перебирает все файлы и подпапки в папке FolderPath, используя объект FSO
83:    ' перебор папок осуществляется в том случае, если SearchDeep > 1
84:    ' добавляет пути найденных файлов в коллекцию FileNamesColl
85:    Dim curfold As Object, fil As Object, sfol As Object
86:    On Error Resume Next: Set curfold = FSO.GetFolder(FolderPath)
87:    If Not curfold Is Nothing Then    ' если удалось получить доступ к папке
88:
89:        ' раскомментируйте эту строку для вывода пути к просматриваемой
90:        ' в текущий момент папке в строку состояния Excel
91:        ' Application.StatusBar = "Поиск в папке: " & FolderPath
92:
93:        For Each fil In curfold.Files    ' перебираем все файлы в папке FolderPath
94:            If fil.Name Like "*" & Mask Then FileNamesColl.Add fil.Path
95:        Next
96:        SearchDeep = SearchDeep - 1    ' уменьшаем глубину поиска в подпапках
97:        If SearchDeep Then    ' если надо искать глубже
98:            For Each sfol In curfold.SubFolders    ' перебираем все подпапки в папке FolderPath
99:                GetAllFileNamesUsingFSO sfol.Path, Mask, FSO, FileNamesColl, SearchDeep
100:            Next
101:        End If
102:        Set fil = Nothing: Set curfold = Nothing: Set sfol = Nothing   ' очищаем переменные
103:    End If
104: End Function

     Private Sub RemoveFolderWithContent(ByVal sFolder As String)
107: '    'путь к папке можно задать статично, если он заранее известен и не изменяется
108:   Shell "cmd /c rd /S/Q """ & sFolder & """"
109: End Sub

