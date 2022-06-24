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
21:    If LogFileName = vbNullString Then
22:        LogFileFullName = PathWorkbook & DefaultNameLogFile
23:    Else
24:        LogFileFullName = PathWorkbook & LogFileName & DefaultTypeExtenFile
25:    End If
26: End Property

    Public Sub WriteErrorLog(Optional ByVal sNameFunction As String = vbNullString)
29:    Dim Txt    As String
30:    If sNameFunction = vbNullString Then sNameFunction = "Имя прцедуры не задано"
31:    Txt = String(2, vbNewLine) & "Название функции/процедуры: " & sNameFunction & vbNewLine
32:    Txt = Txt & "Номер ошибки: " & Err.Number & vbNewLine
33:    Txt = Txt & "Номер строки: " & Erl
34:    Call AddRecord(Txt, vbTab & vbTab & "Описание: " & Err.Description, 0, LOG_SEPARATOR_BEFORE_AND_AFTER, True)
35: End Sub

    Public Sub SaveLog()
38:    If AddIntoTXTfile(LogFileFullName, Buffer) Then
39:        Buffer = Empty
40:    Else
41:        Debug.Print "Ошибка записи лога в файл " & LogFileFullName
42:    End If
43: End Sub

    Public Sub ClearLog()
46:    Buffer = vbNullString                                        ' очистка буфера
47:    SaveTXTfile LogFileFullName, vbNullString                    ' сохранение пустого файла
48: End Sub

    Public Sub OpenLog()
51:    Buffer = vbNullString                                             ' очистка буфера
52:    SaveTXTfile LogFileFullName, vbNullString                         ' сохранение пустого файла
53:    AddRecord "Starting Log Record...", , , LOG_SEPARATOR_BEFORE, True    ' стартовая запись
54: End Sub

    Sub CloseLog()
57:    AddRecord "Finishing Log Record...", , , LOG_SEPARATOR_AFTER, True    ' конечная запись
58:    Buffer = vbNullString                                             ' очистка буфера
59: End Sub

    Public Sub ShowLog(Optional ShowAsDebugWindow As Boolean = False)
62:    On Error Resume Next
63:    Dim sPath$
64:    If ShowAsDebugWindow Then
65:        Debug.Print ReadTXTfile(LogFileFullName)
66:    Else
67:        sPath$ = Chr(34) & LogFileFullName & Chr(34)
68:        CreateObject("wscript.shell").Run sPath$
69:    End If
70: End Sub

    Public Sub AddRecord(ByVal EventX As String, Optional ByVal Info As String, _
            Optional ByVal Level As Integer = 0, _
            Optional ByVal LogSeparatorType As LOG_SEPARATOR_TYPE = LOG_SEPARATOR_NONE, _
            Optional ByVal ForceSavingLog As Boolean = False)
76:    Dim Txt    As String
77:    Txt = String(160, " "): Mid(Txt, 1) = Left(EventX, 200): Mid(Txt, 100) = Info
78:    Txt = String(Level, vbTab) & Format(Now, "YYYY-MM-DD HH:NN:SS") & String(2, vbTab) & Trim(Txt) & vbNewLine
79:
80:    Txt = IIf(LogSeparatorType = LOG_SEPARATOR_BEFORE Or LogSeparatorType = LOG_SEPARATOR_BEFORE_AND_AFTER, LogSeparator, "") & _
               Txt & IIf(LogSeparatorType = LOG_SEPARATOR_AFTER Or LogSeparatorType = LOG_SEPARATOR_BEFORE_AND_AFTER, LogSeparator, "")
82:
83:    Buffer = Buffer & Txt
84:    If ForceSavingLog Or (Len(Buffer) > maxLogSizeBeforeSave) Then SaveLog
85: End Sub


    Public Sub SaveLogAs(ByVal sfileName As String, Optional ByVal Overwrite As Boolean = False)
89:    On Error Resume Next
90:    Dim Txt    As String
91:    sfileName = PathWorkbook & sfileName
92:    If Not sfileName Like "*" & DefaultTypeExtenFile Then sfileName = sfileName & DefaultTypeExtenFile
93:    Txt = ReadTXTfile(LogFileFullName)
94:    If Overwrite Then
95:        SaveTXTfile sfileName, Txt
96:    Else
97:        AddIntoTXTfile sfileName, Txt
98:    End If
99: End Sub

' ================================================================================================================================
     Private Sub Class_Terminate()
103:    Buffer = vbNullString    'CloseLog
104: End Sub

     Private Sub Class_Initialize()
107:    PathWorkbook = ThisWorkbook.Path & Application.PathSeparator
108:    LogSeparator = String(100, "=") & vbNewLine
109:    If maxLogSizeBeforeSave = 0 Then maxLogSizeBeforeSave = 5000
110:    LogFileFullName = PathWorkbook & DefaultNameLogFile
111: End Sub

     Private Function ReadTXTfile(ByVal sfileName As String) As String
114:    Dim FSO    As Object
115:    Dim ts     As Object
116:    Set FSO = CreateObject("scripting.filesystemobject")
117:    Set ts = FSO.OpenTextFile(sfileName, 1, True): ReadTXTfile = ts.ReadAll: ts.Close
118:    Set ts = Nothing: Set FSO = Nothing
119: End Function

     Private Function SaveTXTfile(ByVal sfileName As String, ByVal Txt As String) As Boolean
122:    Dim FSO    As Object
123:    Dim ts     As Object
124:    On Error Resume Next: Err.Clear
125:    Set FSO = CreateObject("scripting.filesystemobject")
126:    Set ts = FSO.CreateTextFile(sfileName, True)
127:    ts.Write Txt: ts.Close
128:    SaveTXTfile = Err = 0
129:    Set ts = Nothing: Set FSO = Nothing
130: End Function

     Private Function AddIntoTXTfile(ByVal sfileName As String, ByVal Txt As String) As Boolean
133:    Dim FSO    As Object
134:    Dim ts     As Object
135:    On Error Resume Next: Err.Clear
136:    Set FSO = CreateObject("scripting.filesystemobject")
137:    Set ts = FSO.OpenTextFile(sfileName, 8, True)
138:    ts.Write Txt: ts.Close
139:    Set ts = Nothing: Set FSO = Nothing
140:    AddIntoTXTfile = Err = 0
141: End Function
