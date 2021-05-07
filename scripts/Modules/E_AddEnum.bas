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
20:    Call AddModuleToProject(C_Const.MOD_ENUM_NAME, vbext_ct_StdModule, AddEnumCode, Application.VBE.ActiveVBProject)
21: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : DeleteSnippetEnumModule - удаление модуля сниппетов
'* Created    : 08-10-2020 14:01
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub DeleteSnippetEnumModule()
31:    Call DeleteModuleToProject(C_Const.MOD_ENUM_NAME)
32: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddLogRecorderClass - создание класса логирования
'* Created    : 08-10-2020 14:01
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub AddLogRecorderClass()
42:    Dim LogCode     As String
43:    LogCode = SHSNIPPETS.ListObjects(C_Const.TB_LOG_CODE).DataBodyRange.Cells(1, 1).Value2
44:    Call AddModuleToProject(C_Const.CLS_LOG_NAME, vbext_ct_ClassModule, LogCode, Application.VBE.ActiveVBProject)
45: End Sub

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
62:    Dim vbComp      As VBIDE.VBComponent
63:    Dim CodeMod     As VBIDE.CodeModule
64:    Dim LineNum     As Long
65:    Dim i           As Integer
66:    On Error GoTo errmsg
67:    If TypeModule = vbext_ct_Document Then TypeModule = vbext_ct_StdModule
68:    Set vbComp = vbProj.VBComponents.Add(TypeModule)
69:
70:    If VBName = C_Const.CLS_LOG_NAME Or VBName = C_Const.MOD_ENUM_NAME Then
71:        vbComp.Name = VBName
72:    Else
73:        VBName = AddModuleName(vbProj, VBName)
74:        vbComp.Name = VBName
75:    End If
76:
77:    Set CodeMod = vbComp.CodeModule
78:    Application.DisplayAlerts = False
79:    With CodeMod
80:        LineNum = .CountOfLines + 1
81:        .InsertLines LineNum, code
82:    End With
83:    Application.DisplayAlerts = True
84:    Debug.Print "Module: [" & VBName & "] added to the book [" & AddFileName & "]"
85:    Exit Sub
errmsg:
87:    Select Case Err.Number
        Case 32813:
89:            Debug.Print "Module: [" & VBName & "] has already been added to the book [" & AddFileName & "]"
90:            vbProj.VBComponents.Remove vbComp
91:        Case 76:
92:            Debug.Print "Module: [" & VBName & "] added to the book:" & ActiveWorkbook.Name & vbLf & "The file is not saved!"
93:        Case Else:
94:            Debug.Print "Error in Add Module To Project" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
95:            Call WriteErrorLog("AddModuleToProject")
96:    End Select
97: End Sub
'создание имении модуля
     Private Function AddModuleName(ByRef vbProj As VBIDE.VBProject, ByVal NameModule As String) As String
100:    Dim objCol      As Collection
101:    Dim vbCompObj   As VBIDE.VBComponent
102:    Dim i           As Integer
103:    Dim bFlag       As Boolean
104:    Set objCol = New Collection
105:
106:    For Each vbCompObj In vbProj.VBComponents
107:        objCol.Add vbCompObj.Name, vbCompObj.Name
108:    Next vbCompObj
109:    bFlag = True
110:    Do While bFlag
111:        On Error Resume Next
112:        Err.Clear
113:        i = i + 1
114:        objCol.Add NameModule & "_" & i, NameModule & "_" & i
115:        If Err.Number = 0 Then
116:            bFlag = False
117:            AddModuleName = NameModule & "_" & i
118:        End If
119:    Loop
120: End Function

     Private Function AddEnumCode() As String
123:    Dim snippets    As ListObject
124:    Dim i           As Long
125:    Dim str1 As String, str2 As String, str0 As String, code As String
126:    Dim str3()      As String
127:    Dim Flag        As Boolean
128:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
129:    i = 1
130:    code = vbNullString
131:    Flag = True
132:    Do While snippets.DataBodyRange.Cells(i, 2) <> vbNullString
133:        str1 = snippets.DataBodyRange.Cells(i, 1).Value
134:        str2 = snippets.DataBodyRange.Cells(i + 1, 1).Value
135:        If str1 = str2 Or str2 = vbNullString Then
136:            If Flag Then
137:                str3 = Split(snippets.DataBodyRange.Cells(i, 3).Value, ".")
138:                code = code & "Public Enum " & str3(0) & vbLf
139:                Flag = False
140:            End If
141:            code = code & Space(4) & snippets.DataBodyRange.Cells(i, 2).Value2 & vbLf
142:        Else
143:            If i <> 1 Then str0 = snippets.DataBodyRange.Cells(i - 1, 1).Value
144:            If str2 <> vbNullString And str1 <> str0 Then
145:                str3 = Split(snippets.DataBodyRange.Cells(i, 3).Value, ".")
146:                code = code & "Public Enum " & str3(0) & vbLf
147:            End If
148:            code = code & Space(4) & snippets.DataBodyRange.Cells(i, 2).Value2 & vbLf
149:            code = code & "End Enum" & vbLf & vbLf
150:            Flag = True
151:        End If
152:        If str2 = vbNullString Then code = code & "End Enum" & vbLf
153:        i = i + 1
154:    Loop
155:    AddEnumCode = code
156: End Function
     Private Sub DeleteModuleToProject(ByVal VBName As String)
158:    Dim vbProj      As VBIDE.VBProject
159:    Dim vbComp      As VBIDE.VBComponent
160:    On Error GoTo ErrorHandler
161:    Set vbProj = Application.VBE.ActiveVBProject
162:    Set vbComp = vbProj.VBComponents(VBName)
163:    vbProj.VBComponents.Remove vbComp
164:    Debug.Print "Module: [" & VBName & "] has been removed, from the workbook [" & AddFileName & "]"
165:    Exit Sub
ErrorHandler:
167:    Select Case Err.Number
        Case 9:
169:            Debug.Print "Module: [" & VBName & "] no, in the book [" & AddFileName & "]"
170:        Case 76:
171:            Debug.Print "Module: [" & VBName & "] has been removed, from the workbook:" & ActiveWorkbook.Name & vbLf & "The file is not saved!"
172:        Case Else:
173:            Debug.Print "Error in Delete Module To Project" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
174:            Call WriteErrorLog("DeleteModuleToProject")
175:    End Select
176:    Err.Clear
177: End Sub
     Private Function AddFileName() As String
179:    Dim strVar      As String
180:    Dim arr_str()   As String
181:    On Error GoTo ErrorHandler
182:    strVar = Application.VBE.ActiveVBProject.Filename
183:    arr_str = Split(strVar, "\")
184:    AddFileName = arr_str(UBound(arr_str))
185:    Exit Function
ErrorHandler:
187:    Select Case Err.Number
        Case 76:
189:            AddFileName = ActiveWorkbook.Name
190:        Case Else:
191:            Debug.Print "Error in Add FileName" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
192:            Call WriteErrorLog("AddFileName")
193:    End Select
194:    Err.Clear
195: End Function
'копирование модуля для меню
     Public Sub CopyModyleVBE()
198:    Dim vbProj      As VBIDE.VBProject
199:    Dim vbCompObj   As VBIDE.VBComponent
200:    Dim txtCode     As String
201:    Set vbProj = Application.VBE.ActiveVBProject
202:    Set vbCompObj = vbProj.VBE.SelectedVBComponent
203:    If vbCompObj Is Nothing Then Exit Sub
204:    With vbCompObj
205:        If vbCompObj.Type = vbext_ct_MSForm Then
206:            Call CopyModuleForm(vbProj, vbCompObj)
207:        Else
208:            If .CodeModule.CountOfLines = 0 Then
209:                txtCode = vbNullString
210:            Else
211:                txtCode = .CodeModule.Lines(1, .CodeModule.CountOfLines)
212:            End If
213:            txtCode = Replace(txtCode, "Option Explicit", vbNullString)
214:            Call E_AddEnum.AddModuleToProject(.Name, vbCompObj.Type, txtCode, vbProj)
215:        End If
216:    End With
217: End Sub
     Private Sub CopyModuleForm(ByRef vbProj As VBIDE.VBProject, ByRef vbCompObj As VBIDE.VBComponent)
219:    Dim sFullFileName As String
220:    Dim sNameFile   As String
221:    Dim sNameMod    As String
222:    Dim sNameModNew As String
223:    Dim sNamePath   As String
224:    Dim sFullFileNameNew As String
225:    Const sEXT      As String = ".bas"
226:    Const sEXTFRX   As String = ".frx"
227:
228:    On Error GoTo ErrorHandler
229:
230:    sFullFileName = vbProj.Filename
231:    sNameFile = C_PublicFunctions.sGetFileName(sFullFileName)
232:    sNamePath = VBA.Replace(sFullFileName, sNameFile, vbNullString)
233:
234:    sNameMod = vbCompObj.Name
235:    sNameModNew = AddModuleName(vbProj, vbCompObj.Name)
236:    sFullFileNameNew = sNamePath & sNameModNew & sEXT
237:
238:    vbCompObj.Name = sNameModNew
239:    Call vbCompObj.Export(Filename:=sFullFileNameNew)
240:    vbCompObj.Name = sNameMod
241:    Call vbProj.VBComponents.Import(Filename:=sFullFileNameNew)
242:    Call Kill(sFullFileNameNew)
243:    Call Kill(sNamePath & sNameModNew & sEXTFRX)
244:
245:    Exit Sub
ErrorHandler:
247:    Select Case Err.Number
        Case 76:
249:            Call MsgBox("The file is not saved, save the file!", vbCritical, "Error:")
250:        Case Else:
251:            Debug.Print "Error in Copy Module Form" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
252:            Call WriteErrorLog("CopyModuleForm")
253:    End Select
254:    Err.Clear
255: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SerchVariableUnUsedInSelectedWorkBook - анализ кода на не используеммые переменные
'* Created    : 08-10-2020 14:00
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub SerchVariableUnUsedInSelectedWorkBook()
265:    Call VariableUnUsed.Show
End Sub
