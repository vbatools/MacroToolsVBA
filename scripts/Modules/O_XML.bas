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
46:    Dim sFileName As Variant
47:    Dim strFileName As String
48:
49:    On Error GoTo errmsg
50:    sFileName = SelectedFile(sFilePath, False, "*.xlsm;*.xlsb;*.xlam;*.xlsx;*.docm;*.dotm;*.dotx;*.docx;*.pptx;*.pptm;*.potx;*.potm")
51:    If TypeName(sFileName) = "Empty" Then Exit Function
52:
53:    strFileName = sFileName(1)
54:    OpenAndCloseExcelFileInFolder = OpenAndCloseExcelFile(bOpenFile, bBackUp, True, strFileName)
55:
56:    Exit Function
errmsg:
58:    Select Case Err.Number
        Case 70:
60:            Call MsgBox("Error No access to the file!" & vbLf & "Perhaps the file is open, to continue, close it and try again.", vbCritical, "No file access:")
61:            Exit Function
62:        Case Else:
63:            Call MsgBox("Error! in Open And Close Excel File In Folder" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the row" & Erl, vbCritical, "Error:")
64:            Call WriteErrorLog("OpenAndCloseExcelFileInFolder")
65:    End Select
66:    Err.Clear
67:    OpenAndCloseExcelFileInFolder = vbNullString
68: End Function
     Public Function OpenAndCloseExcelFile(ByVal bOpenFile As Boolean, Optional bBackUp As Boolean = True, Optional bShowMsg As Boolean = True, Optional sFilePath As String = vbNullString) As String
70:    Dim cEditOpenXML As clsEditOpenXML
71:    Dim sMsg As String, sTitleMsg As String
72:
73:    On Error GoTo errmsg
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
86:            sMsg = "Unpacking the file has been completed!"
87:            sTitleMsg = "Unpacking an Excel file:"
88:        Else
89:            .CreateBackupXML = bBackUp
90:            .SourceFile = sFilePath
91:            'Запоковка файла
92:            .ZipAllFilesInFolder
93:            sMsg = "The file has been packed!" & vbLf & "Also created a Backup file"
94:            sTitleMsg = "Packing an Excel file:"
95:        End If
96:    End With
97:    Set cEditOpenXML = Nothing
98:    If bShowMsg Then Call MsgBox(sMsg, vbInformation, sTitleMsg)
99:
100:    Exit Function
errmsg:
102:    Select Case Err.Number
        Case 70:
104:            Call MsgBox("Error No access to the file!" & vbLf & "Perhaps the file is open, to continue, close it and try again.", vbCritical, "No file access:")
105:            Exit Function
106:        Case Else:
107:            Call MsgBox("Error! in Open And Close Excel File In Folder" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the row" & Erl, vbCritical, "Error:")
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
145: End Function

