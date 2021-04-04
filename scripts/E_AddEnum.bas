Attribute VB_Name = "E_AddEnum"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : E_AddEnum - вставка модуля со снипетами перечесления, анализ кода на неиспользуеммые переменные
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Private Module
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddSnippetEnumModule - вставка модуля сниппетов
'* Created    : 08-10-2020 14:01
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub AddSnippetEnumModule()
12:    Call AddModuleToProject(C_Const.MOD_ENUM_NAME, vbext_ct_StdModule, AddEnumCode, Application.VBE.ActiveVBProject)
13: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : DeleteSnippetEnumModule - удаление модуля сниппетов
'* Created    : 08-10-2020 14:01
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub DeleteSnippetEnumModule()
15:    Call DeleteModuleToProject(C_Const.MOD_ENUM_NAME)
16: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddLogRecorderClass - создание класса логирования
'* Created    : 08-10-2020 14:01
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub AddLogRecorderClass()
18:    Dim LogCode     As String
19:    LogCode = SHSNIPPETS.ListObjects(C_Const.TB_LOG_CODE).DataBodyRange.Cells(1, 1).Value2
20:    Call AddModuleToProject(C_Const.CLS_LOG_NAME, vbext_ct_ClassModule, LogCode, Application.VBE.ActiveVBProject)
21: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddModuleToProject - создание модуля VBA
'* Created    : 08-10-2020 14:04
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                     Description
'*
'* ByVal VBName As String          :
'* ByVal TypeModule As String      :
'* ByVal code As String            :
'* ByRef vbProj As VBIDE.VBProject :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub AddModuleToProject(ByVal VBName As String, ByVal TypeModule As String, ByVal code As String, ByRef vbProj As VBIDE.VBProject)
23:    Dim vbComp      As VBIDE.VBComponent
24:    Dim CodeMod     As VBIDE.CodeModule
25:    Dim LineNum     As Long
26:    Dim i           As Integer
27:    On Error GoTo errmsg
28:    If TypeModule = vbext_ct_Document Then TypeModule = vbext_ct_StdModule
29:    Set vbComp = vbProj.VBComponents.Add(TypeModule)
30:
31:    If VBName = C_Const.CLS_LOG_NAME Or VBName = C_Const.MOD_ENUM_NAME Then
32:        vbComp.Name = VBName
33:    Else
34:        VBName = AddModuleName(vbProj, VBName)
35:        vbComp.Name = VBName
36:    End If
37:
38:    Set CodeMod = vbComp.CodeModule
39:    Application.DisplayAlerts = False
40:    With CodeMod
41:        LineNum = .CountOfLines + 1
42:        .InsertLines LineNum, code
43:    End With
44:    Application.DisplayAlerts = True
45:    Debug.Print "Модуль: [" & VBName & "] добавлен в книгу [" & AddFileName & "]"
46:    Exit Sub
errmsg:
48:    Select Case Err.Number
        Case 32813:
50:            Debug.Print "Модуль: [" & VBName & "] уже был добавлен в книгу [" & AddFileName & "]"
51:            vbProj.VBComponents.Remove vbComp
52:        Case 76:
53:            Debug.Print "Модуль: [" & VBName & "] добавлен в книгу: " & ActiveWorkbook.Name & vbLf & "Файл не сохранен!"
54:        Case Else:
55:            Debug.Print "Ошибка в AddModuleToProject" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
56:            Call WriteErrorLog("AddModuleToProject")
57:    End Select
58: End Sub
'создание имении модуля
    Private Function AddModuleName(ByRef vbProj As VBIDE.VBProject, ByVal NameModule As String) As String
61:    Dim objCol      As Collection
62:    Dim vbCompObj   As VBIDE.VBComponent
63:    Dim i           As Integer
64:    Dim bFlag       As Boolean
65:    Set objCol = New Collection
66:
67:    For Each vbCompObj In vbProj.VBComponents
68:        objCol.Add vbCompObj.Name, vbCompObj.Name
69:    Next vbCompObj
70:    bFlag = True
71:    Do While bFlag
72:        On Error Resume Next
73:        Err.Clear
74:        i = i + 1
75:        objCol.Add NameModule & "_" & i, NameModule & "_" & i
76:        If Err.Number = 0 Then
77:            bFlag = False
78:            AddModuleName = NameModule & "_" & i
79:        End If
80:    Loop
81: End Function

     Private Function AddEnumCode() As String
84:    Dim snippets    As ListObject
85:    Dim i           As Long
86:    Dim str1 As String, str2 As String, str0 As String, code As String
87:    Dim str3()      As String
88:    Dim Flag        As Boolean
89:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
90:    i = 1
91:    code = vbNullString
92:    Flag = True
93:    Do While snippets.DataBodyRange.Cells(i, 2) <> vbNullString
94:        str1 = snippets.DataBodyRange.Cells(i, 1).Value
95:        str2 = snippets.DataBodyRange.Cells(i + 1, 1).Value
96:        If str1 = str2 Or str2 = vbNullString Then
97:            If Flag Then
98:                str3 = Split(snippets.DataBodyRange.Cells(i, 3).Value, ".")
99:                code = code & "Public Enum " & str3(0) & vbLf
100:                Flag = False
101:            End If
102:            code = code & Space(4) & snippets.DataBodyRange.Cells(i, 2).Value2 & vbLf
103:        Else
104:            If i <> 1 Then str0 = snippets.DataBodyRange.Cells(i - 1, 1).Value
105:            If str2 <> vbNullString And str1 <> str0 Then
106:                str3 = Split(snippets.DataBodyRange.Cells(i, 3).Value, ".")
107:                code = code & "Public Enum " & str3(0) & vbLf
108:            End If
109:            code = code & Space(4) & snippets.DataBodyRange.Cells(i, 2).Value2 & vbLf
110:            code = code & "End Enum" & vbLf & vbLf
111:            Flag = True
112:        End If
113:        If str2 = vbNullString Then code = code & "End Enum" & vbLf
114:        i = i + 1
115:    Loop
116:    AddEnumCode = code
117: End Function
     Private Sub DeleteModuleToProject(ByVal VBName As String)
119:    Dim vbProj      As VBIDE.VBProject
120:    Dim vbComp      As VBIDE.VBComponent
121:    On Error GoTo ErrorHandler
122:    Set vbProj = Application.VBE.ActiveVBProject
123:    Set vbComp = vbProj.VBComponents(VBName)
124:    vbProj.VBComponents.Remove vbComp
125:    Debug.Print "Модуль: [" & VBName & "] был удален, из книги [" & AddFileName & "]"
126:    Exit Sub
ErrorHandler:
128:    Select Case Err.Number
        Case 9:
130:            Debug.Print "Модуля: [" & VBName & "] нет, в книге [" & AddFileName & "]"
131:        Case 76:
132:            Debug.Print "Модуль: [" & VBName & "] был удален, из книги: " & ActiveWorkbook.Name & vbLf & "Файл не сохранен!"
133:        Case Else:
134:            Debug.Print "Ошибка в DeleteModuleToProject" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
135:            Call WriteErrorLog("DeleteModuleToProject")
136:    End Select
137:    Err.Clear
138: End Sub
     Private Function AddFileName() As String
140:    Dim strVar      As String
141:    Dim arr_str()   As String
142:    On Error GoTo ErrorHandler
143:    strVar = Application.VBE.ActiveVBProject.Filename
144:    arr_str = Split(strVar, "\")
145:    AddFileName = arr_str(UBound(arr_str))
146:    Exit Function
ErrorHandler:
148:    Select Case Err.Number
        Case 76:
150:            AddFileName = ActiveWorkbook.Name
151:        Case Else:
152:            Debug.Print "Ошибка в AddFileName" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
153:            Call WriteErrorLog("AddFileName")
154:    End Select
155:    Err.Clear
156: End Function
'копирование модуля для меню
     Public Sub CopyModyleVBE()
159:    Dim vbProj      As VBIDE.VBProject
160:    Dim vbCompObj   As VBIDE.VBComponent
161:    Dim txtCode     As String
162:    Set vbProj = Application.VBE.ActiveVBProject
163:    Set vbCompObj = vbProj.VBE.SelectedVBComponent
164:    If vbCompObj Is Nothing Then Exit Sub
165:    With vbCompObj
166:        If vbCompObj.Type = vbext_ct_MSForm Then
167:            Call CopyModuleForm(vbProj, vbCompObj)
168:        Else
169:            If .CodeModule.CountOfLines = 0 Then
170:                txtCode = vbNullString
171:            Else
172:                txtCode = .CodeModule.Lines(1, .CodeModule.CountOfLines)
173:            End If
174:            txtCode = Replace(txtCode, "Option Explicit", vbNullString)
175:            Call E_AddEnum.AddModuleToProject(.Name, vbCompObj.Type, txtCode, vbProj)
176:        End If
177:    End With
178: End Sub
     Private Sub CopyModuleForm(ByRef vbProj As VBIDE.VBProject, ByRef vbCompObj As VBIDE.VBComponent)
180:    Dim sFullFileName As String
181:    Dim sNameFile   As String
182:    Dim sNameMod    As String
183:    Dim sNameModNew As String
184:    Dim sNamePath   As String
185:    Dim sFullFileNameNew As String
186:    Const sEXT      As String = ".bas"
187:    Const sEXTFRX   As String = ".frx"
188:
189:    On Error GoTo ErrorHandler
190:
191:    sFullFileName = vbProj.Filename
192:    sNameFile = C_PublicFunctions.sGetFileName(sFullFileName)
193:    sNamePath = VBA.Replace(sFullFileName, sNameFile, vbNullString)
194:
195:    sNameMod = vbCompObj.Name
196:    sNameModNew = AddModuleName(vbProj, vbCompObj.Name)
197:    sFullFileNameNew = sNamePath & sNameModNew & sEXT
198:
199:    vbCompObj.Name = sNameModNew
200:    Call vbCompObj.Export(Filename:=sFullFileNameNew)
201:    vbCompObj.Name = sNameMod
202:    Call vbProj.VBComponents.Import(Filename:=sFullFileNameNew)
203:    Call Kill(sFullFileNameNew)
204:    Call Kill(sNamePath & sNameModNew & sEXTFRX)
205:
206:    Exit Sub
ErrorHandler:
208:    Select Case Err.Number
        Case 76:
210:            Call MsgBox("Файл не сохранен, сохраните файл!", vbCritical, "Ошибка:")
211:        Case Else:
212:            Debug.Print "Ошибка в CopyModuleForm" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
213:            Call WriteErrorLog("CopyModuleForm")
214:    End Select
215:    Err.Clear
216: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SerchVariableUnUsedInSelectedWorkBook - анализ кода на не используеммые переменные
'* Created    : 08-10-2020 14:00
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub SerchVariableUnUsedInSelectedWorkBook()
218:    Call VariableUnUsed.Show
End Sub
