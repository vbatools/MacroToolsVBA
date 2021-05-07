VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : LogRecorder - модуль логирования ошибок
'* Created    : 01-20-2020 15:55
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
'добавление разделителей в лог
Public Enum LOG_SEPARATOR_TYPE
    LOG_SEPARATOR_NONE = 0
    LOG_SEPARATOR_BEFORE = 1
    LOG_SEPARATOR_AFTER = 2
    LOG_SEPARATOR_BEFORE_AND_AFTER = 3
End Enum

Private Const DefaultTypeExtenFile As String = ".log"
Private Const DefaultNameLogFile As String = "common" & DefaultTypeExtenFile

Private PathWorkbook As String
Private LogFileFullName As String

Public Buffer  As String
Public maxLogSizeBeforeSave As Long
Public LogSeparator As String

    Public Property Let LogFileName(ByVal LogFileName As String)
29:    If LogFileName = vbNullString Then
30:        LogFileFullName = PathWorkbook & DefaultNameLogFile
31:    Else
32:        LogFileFullName = PathWorkbook & LogFileName & DefaultTypeExtenFile
33:    End If
34: End Property

    Public Sub WriteErrorLog(Optional ByVal sNameFunction As String = vbNullString)
37:    Dim Txt    As String
38: If sNameFunction = vbNullString Then sNameFunction = "Procedure name not specified"
39: Txt = String(2, vbNewLine) & "Function/procedure name:" & sNameFunction & vbNewLine
40:    Txt = Txt & "Error number:" & Err.Number & vbNewLine
41:    Txt = Txt & "Line number:" & Erl
42:    Call AddRecord(Txt, vbTab & vbTab & "Description:" & Err.Description, 0, LOG_SEPARATOR_BEFORE_AND_AFTER, True)
43: End Sub

    Public Sub SaveLog()
46:    If AddIntoTXTfile(LogFileFullName, Buffer) Then
47:        Buffer = Empty
48:    Else
49:        Debug.Print "Error writing the log to the file" & LogFileFullName
50:    End If
51: End Sub

    Public Sub ClearLog()
54:    Buffer = vbNullString                                        ' очистка буфера
55:    SaveTXTfile LogFileFullName, vbNullString                    ' сохранение пустого файла
56: End Sub

    Public Sub OpenLog()
59:    Buffer = vbNullString                                             ' очистка буфера
60:    SaveTXTfile LogFileFullName, vbNullString                         ' сохранение пустого файла
61:    AddRecord "Starting Log Record...", , , LOG_SEPARATOR_BEFORE, True    ' стартовая запись
62: End Sub

    Sub CloseLog()
65:    AddRecord "Finishing Log Record...", , , LOG_SEPARATOR_AFTER, True    ' конечная запись
66:    Buffer = vbNullString                                             ' очистка буфера
67: End Sub

    Public Sub ShowLog(Optional ShowAsDebugWindow As Boolean = False)
70:    On Error Resume Next
71:    Dim sPath$
72:    If ShowAsDebugWindow Then
73:        Debug.Print ReadTXTfile(LogFileFullName)
74:    Else
75:        sPath$ = Chr(34) & LogFileFullName & Chr(34)
76:        CreateObject("wscript.shell").Run sPath$
77:    End If
78: End Sub

    Public Sub AddRecord(ByVal EventX As String, Optional ByVal Info As String, _
                Optional ByVal Level As Integer = 0, _
                Optional ByVal LogSeparatorType As LOG_SEPARATOR_TYPE = LOG_SEPARATOR_NONE, _
                Optional ByVal ForceSavingLog As Boolean = False)
84:    Dim Txt    As String
85:    Txt = String(160, " "): Mid(Txt, 1) = Left(EventX, 200): Mid(Txt, 100) = Info
86:    Txt = String(Level, vbTab) & Format(Now, "YYYY-MM-DD HH:NN:SS") & String(2, vbTab) & Trim(Txt) & vbNewLine
87:
88:    Txt = IIf(LogSeparatorType = LOG_SEPARATOR_BEFORE Or LogSeparatorType = LOG_SEPARATOR_BEFORE_AND_AFTER, LogSeparator, "") & _
                  Txt & IIf(LogSeparatorType = LOG_SEPARATOR_AFTER Or LogSeparatorType = LOG_SEPARATOR_BEFORE_AND_AFTER, LogSeparator, "")
90:
91:    Buffer = Buffer & Txt
92:    If ForceSavingLog Or (Len(Buffer) > maxLogSizeBeforeSave) Then SaveLog
93: End Sub


     Public Sub SaveLogAs(ByVal sFileName As String, Optional ByVal Overwrite As Boolean = False)
97:    On Error Resume Next
98:    Dim Txt    As String
99:    sFileName = PathWorkbook & sFileName
100:    If Not sFileName Like "*" & DefaultTypeExtenFile Then sFileName = sFileName & DefaultTypeExtenFile
101:    Txt = ReadTXTfile(LogFileFullName)
102:    If Overwrite Then
103:        SaveTXTfile sFileName, Txt
104:    Else
105:        AddIntoTXTfile sFileName, Txt
106:    End If
107: End Sub

' ================================================================================================================================
     Private Sub Class_Terminate()
111:    Buffer = vbNullString    'CloseLog
112: End Sub

     Private Sub Class_Initialize()
115:    PathWorkbook = ThisWorkbook.Path & Application.PathSeparator
116:    LogSeparator = String(100, "=") & vbNewLine
117:    If maxLogSizeBeforeSave = 0 Then maxLogSizeBeforeSave = 5000
118:    LogFileFullName = PathWorkbook & DefaultNameLogFile
119: End Sub

     Private Function ReadTXTfile(ByVal sFileName As String) As String
122:    Dim FSO    As Object
123:    Dim ts     As Object
124:    Set FSO = CreateObject("scripting.filesystemobject")
125:    Set ts = FSO.OpenTextFile(sFileName, 1, True): ReadTXTfile = ts.ReadAll: ts.Close
126:    Set ts = Nothing: Set FSO = Nothing
127: End Function

     Private Function SaveTXTfile(ByVal sFileName As String, ByVal Txt As String) As Boolean
130:    Dim FSO    As Object
131:    Dim ts     As Object
132:    On Error Resume Next: Err.Clear
133:    Set FSO = CreateObject("scripting.filesystemobject")
134:    Set ts = FSO.CreateTextFile(sFileName, True)
135:    ts.Write Txt: ts.Close
136:    SaveTXTfile = Err = 0
137:    Set ts = Nothing: Set FSO = Nothing
138: End Function

     Private Function AddIntoTXTfile(ByVal sFileName As String, ByVal Txt As String) As Boolean
141:    Dim FSO    As Object
142:    Dim ts     As Object
143:    On Error Resume Next: Err.Clear
144:    Set FSO = CreateObject("scripting.filesystemobject")
145:    Set ts = FSO.OpenTextFile(sFileName, 8, True)
146:    ts.Write Txt: ts.Close
147:    Set ts = Nothing: Set FSO = Nothing
148:    AddIntoTXTfile = Err = 0
149: End Function

