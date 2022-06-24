Attribute VB_Name = "O_XML"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : O_XML - чтение XML, распоковка, запоковка файла
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module

Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName  As String * MAX_PATH
    cAlternate As String * 14
End Type

#If Win64 Then
Private Declare PtrSafe Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare PtrSafe Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
#Else
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
#End If

'Распоковать или Запоковать файл Ecxel
    Public Function OpenAndCloseExcelFileInFolder(ByVal bOpenFile As Boolean, Optional bBackUp As Boolean = True, Optional sFilePath As String = vbNullString) As String
46:    Dim sfileName As Variant
47:    Dim strFileName As String
48:
49:    On Error GoTo errMsg
50:    sfileName = SelectedFile(sFilePath, False, "*.xlsm;*.xlsb;*.xlam;*.xlsx;*.docm;*.dotm;*.dotx;*.docx;*.pptx;*.pptm;*.potx;*.potm")
51:    If TypeName(sfileName) = "Empty" Then Exit Function
52:
53:    strFileName = sfileName(1)
54:    OpenAndCloseExcelFileInFolder = OpenAndCloseExcelFile(bOpenFile, bBackUp, True, strFileName)
55:
56:    Exit Function
errMsg:
58:    Select Case Err.Number
        Case 70:
60:            Call MsgBox("Ошибка! Нет доступа к файлу!" & vbLf & "Возможно файл открыт, для продолжения закройте его и повторите попытку.", vbCritical, "Нет доступа к файлу:")
61:            Exit Function
62:        Case Else:
63:            Call MsgBox("Ошибка! в OpenAndCloseExcelFileInFolder" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
64:            Call WriteErrorLog("OpenAndCloseExcelFileInFolder")
65:    End Select
66:    Err.Clear
67:    OpenAndCloseExcelFileInFolder = vbNullString
68: End Function
     Public Function OpenAndCloseExcelFile(ByVal bOpenFile As Boolean, Optional bBackUp As Boolean = True, Optional bShowMsg As Boolean = True, Optional sFilePath As String = vbNullString) As String
70:    Dim cEditOpenXML As clsEditOpenXML
71:    Dim sMsg As String, sTitleMsg As String
72:
73:    On Error GoTo errMsg
74:
75:    Set cEditOpenXML = New clsEditOpenXML
76:    With cEditOpenXML
77:
78:        If bOpenFile Then
79:            'не создавать Backup
80:            .CreateBackupXML = bBackUp
81:            'выбор файла для извлечения XML
82:            .SourceFile = sFilePath
83:            'Распоковка файла
84:            .UnzipFile
85:            OpenAndCloseExcelFile = .XMLFolder(XMLFolder_root)
86:            sMsg = "Распоковка файла выполнена!"
87:            sTitleMsg = "Распоковка файла Excel:"
88:        Else
89:            .CreateBackupXML = bBackUp
90:            .SourceFile = sFilePath
91:            'Запоковка файла
92:            .ZipAllFilesInFolder
93:            sMsg = "Запоковка файла выполнена!" & vbLf & "Так же создан Backup файла"
94:            sTitleMsg = "Запоковка файла Excel:"
95:        End If
96:    End With
97:    Set cEditOpenXML = Nothing
98:    If bShowMsg Then Call MsgBox(sMsg, vbInformation, sTitleMsg)
99:
100:    Exit Function
errMsg:
102:    Select Case Err.Number
        Case 70:
104:            Call MsgBox("Ошибка! Нет доступа к файлу!" & vbLf & "Возможно файл открыт, для продолжения закройте его и повторите попытку.", vbCritical, "Нет доступа к файлу:")
105:            Exit Function
106:        Case Else:
107:            Call MsgBox("Ошибка! в OpenAndCloseExcelFileInFolder" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
108:            Call WriteErrorLog("OpenAndCloseExcelFileInFolder")
109:    End Select
110:    Err.Clear
111:    Set cEditOpenXML = Nothing
112:    OpenAndCloseExcelFile = vbNullString
113: End Function
     Public Function FolderExists(ByVal sFolder As String) As Boolean
115:
116:    Dim hFile  As Long
117:    Dim WFD    As WIN32_FIND_DATA
118:
119:    'remove training slash before verifying
120:    sFolder = UnQualifyPath(sFolder)
121:
122:    'call the API pasing the folder
123:    hFile = FindFirstFile(sFolder, WFD)
124:
125:    'if a valid file handle was returned,
126:    'and the directory attribute is set
127:    'the folder exists
128:    FolderExists = (hFile <> INVALID_HANDLE_VALUE) And _
                (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY)
130:
131:    'clean up
132:    Call FindClose(hFile)
133: End Function

Private Function UnQualifyPath(ByVal sFolder As String) As String
136:
137:    'trim and remove any trailing slash
138:    sFolder = Trim$(sFolder)
139:
140:    If Right$(sFolder, 1) = Application.PathSeparator Then
141:        UnQualifyPath = Left$(sFolder, Len(sFolder) - 1)
142:    Else
143:        UnQualifyPath = sFolder
144:    End If
End Function
