Attribute VB_Name = "S_ModuleCommander"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : S_ModuleCommander - копирование вставка и удаление модулей VBA
'* Created    : 25-12-2019 14:22
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module
    Public Sub ImportAllModules(ByRef wb As Workbook)
11:    Dim sfileName As Variant
12:    Dim sFileNameTxt As String
13:    Dim i As Long, lBas As Long, lCls As Long, lFrm As Long
14:    On Error GoTo Error_Handler_
15:    sfileName = SelectedFile(wb.FullName, True, "*.bas;*.cls;*.frm")
16:    If TypeName(sfileName) = "Empty" Then Exit Sub
17:    For i = 1 To UBound(sfileName)
18:        sFileNameTxt = CStr(sfileName(i))
19:        On Error Resume Next
20:        wb.VBProject.VBComponents.Import Filename:=sFileNameTxt
21:        On Error GoTo 0
22:        Select Case C_PublicFunctions.sGetExtensionName(sFileNameTxt)
            Case "bas": lBas = lBas + 1
24:            Case "cls": lCls = lCls + 1
25:            Case "frm": lFrm = lFrm + 1
26:        End Select
27:    Next i
28:    Call MsgBox("Импорт проекта VBA в книгу" & wb.Name & " выполнен!" & vbCrLf & vbCrLf & "Импортировано: " & _
                  vbCrLf & vbCrLf & "Модулей: " & lBas & _
                  vbCrLf & "Классов: " & lCls & _
                  vbCrLf & "Форм: " & lFrm & _
                  vbCrLf & VBA.String(14, "-") & _
                  vbCrLf & "Всего: " & lFrm + lCls + lBas, vbInformation, "Импорт проекта VBA:")
34:    Exit Sub
Error_Handler_:
36:    Call MsgBox("Ошибка! в ImportAllModules" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
37:    Call WriteErrorLog("ImportAllModules")
38: End Sub
    Public Sub ExportAllModules(ByRef wb As Workbook, ByRef arrVBComp() As String)
40:    Dim oVBComp As VBIDE.VBComponent
41:    Dim sDestinationFolder As String
42:    Dim sFullPathFile As String
43:    Dim sFileExt As String
44:    Dim C      As Long
45:    Dim i      As Long
46:    On Error GoTo Error_Handler_
47:    If DoesActiveVBAprojectExist(wb) Then
48:        sDestinationFolder = DirLoadFiles(wb.Path)
49:        If sDestinationFolder = vbNullString Then Exit Sub
50:    Else
51:        Exit Sub
52:    End If
53:    If MsgBox("Вы хотите экспортировать все модули, из книги - " & wb.Name & " ?" & vbCrLf & vbCrLf & "Экспорт в папку: " & sDestinationFolder, vbYesNo, "Выгрузка проекта:") = vbNo Then Exit Sub
54:    If Dir(sDestinationFolder, vbDirectory) = vbNullString Then MkDir sDestinationFolder
55:    On Error Resume Next
56:    For i = 0 To 2
57:        If Dir(sDestinationFolder, vbDirectory) <> vbNullString Then
58:            Kill sDestinationFolder & "\*." & Array("bas", "cls", "fr?")(i)
59:        End If
60:    Next
61:    On Error GoTo Error_Handler_
62:    For i = 0 To UBound(arrVBComp)
63:        Set oVBComp = wb.VBProject.VBComponents(arrVBComp(i))
64:        If ModuleLineCount(oVBComp) > 0 Then
65:            Select Case oVBComp.Type
                Case vbext_ct_ClassModule
67:                    sFileExt = ".cls"
68:                Case vbext_ct_Document
69:                    sFileExt = ".cls"
70:                Case vbext_ct_StdModule
71:                    sFileExt = ".bas"
72:                Case vbext_ct_MSForm
73:                    sFileExt = ".frm"
74:                Case Else
75:                    sFileExt = ".txt"
76:            End Select
77:            If sFileExt <> vbNullString Then
78:                sFullPathFile = sDestinationFolder & "\" & oVBComp.Name & sFileExt
79:                oVBComp.Export sFullPathFile
80:                C = C + 1
81:            End If
82:        End If
83:    Next i
84:    If C > 0 Then Shell "C:\WINDOWS\explorer.exe """ & sDestinationFolder & "", vbNormalFocus
85:    Exit Sub
Error_Handler_:
87:    Call MsgBox("Ошибка! в ExportAllModules" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
88:    Call WriteErrorLog("ExportAllModules")
89: End Sub
     Public Sub DeleteAllModulesInActiveProject(ByRef wb As Workbook, ByRef arrVBComp() As String)
91:    Dim oVbProject As VBIDE.VBProject
92:    Dim oVBComp As VBIDE.VBComponent
93:    Dim m      As Long
94:    Dim r      As Long
95:    Dim i As Integer
96:    On Error GoTo Error_Handler_
97:    If Not DoesActiveVBAprojectExist(wb:=wb, bFlagBlankLines:=False) Then Exit Sub
98:    If MsgBox("Вы хотите удалить все модули кода в активном проекте VBA?, из книги - " & wb.Name, vbYesNo, "Удаление проекта:") = vbNo Then Exit Sub
99:    If MsgBox("ПОЖАЛУЙСТА, ПОДТВЕРДИТЕ-УДАЛЕНИЕ КОДА ЯВЛЯЕТСЯ НЕ ОБРАТИМЫМ ДЕЙСТВИЕМ" & vbCrLf & vbCrLf & "Вы хотите удалить весь код из активного проекта VBA?, из книги - " & wb.Name, vbYesNo, "Удаление проекта:") = vbNo Then Exit Sub
100:    Set oVbProject = wb.VBProject
101:    For i = 0 To UBound(arrVBComp)
102:        Set oVBComp = wb.VBProject.VBComponents(arrVBComp(i))
103:        If oVBComp.Type = vbext_ct_Document Then
104:            With oVBComp.CodeModule
105:                If .CountOfLines > 1 Then
106:                    .DeleteLines 1, .CountOfLines
107:                    .InsertLines 1, "Option Explicit"
108:                    m = m + 1
109:                End If
110:            End With
111:        Else
112:            oVbProject.VBComponents.Remove oVBComp
113:            m = m + 1
114:        End If
115:    Next i
116:    'R = RemoveAllReferences(WB)
117:    If r + m > 0 Then MsgBox "Modules удалено: " & m & vbCrLf & "References удалено: " & r, vbInformation, ""
118:    Exit Sub
Error_Handler_:
120:    Call MsgBox("Ошибка! в DeleteAllModulesInActiveProject" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
121:    Call WriteErrorLog("DeleteAllModulesInActiveProject")
122: End Sub
'**********************************************************************************************************************************************************************
     Public Function DoesActiveVBAprojectExist(ByRef wb As Workbook, Optional bFlagBlankLines As Boolean = True) As Boolean
125:    Dim oVBComp As VBIDE.VBComponent
126:    Dim C      As Long
127:    On Error GoTo Error_Handler_
128:
129:    If Not IsWorkbookOpen Then Exit Function
130:
131:    If wb.VBProject.Protection = vbext_pp_locked Then
132:        Call MsgBox("VBA проект в книге - " & wb.Name & " защищен, паролем!" & vbCrLf & "Снимите пароль!", vbCritical, "Ошибка:")
133:        DoesActiveVBAprojectExist = False: Exit Function
134:    End If
135:
136:    If bFlagBlankLines Then
137:        For Each oVBComp In wb.VBProject.VBComponents
138:            C = C + ModuleLineCount(oVBComp)
139:        Next oVBComp
140:    Else
141:        DoesActiveVBAprojectExist = True
142:        Exit Function
143:    End If
144:
145:    If C = 0 Then
146:        Call MsgBox("Экспорт проекта VBA - в книге: " & wb.Name & " нет проекта VBA!", vbCritical, "Отсутствие проекта VBA:")
147:        DoesActiveVBAprojectExist = False
148:        Exit Function
149:    Else
150:        DoesActiveVBAprojectExist = True
151:    End If
152:
153:    Exit Function
Error_Handler_:
155:    Call MsgBox("Ошибка! в DoesActiveVBAprojectExist" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
156:    Call WriteErrorLog("DoesActiveVBAprojectExist")
157:    DoesActiveVBAprojectExist = False
158: End Function
     Public Function ModuleLineCount(oVBComp As VBIDE.VBComponent) As Long
160:    Dim sLine  As String
161:    Dim C      As Long
162:    Dim i      As Long
163:    With oVBComp.CodeModule
164:        If .CountOfLines > 0 Then
165:            For i = 1 To .CountOfLines
166:                sLine = Trim(.Lines(i, 1))
167:                If Left(sLine, 11) = "Option Base" Or Left(sLine, 14) = "Option Compare" Or Left(sLine, 15) = "Option Explicit" Or Left(sLine, 21) = "Option Private Module" Then sLine = vbNullString
168:                If sLine <> vbNullString Then C = C + 1
169:            Next i
170:        End If
171:    End With
172:    ModuleLineCount = C
173: End Function
     Private Function RemoveAllReferences(ByRef wb As Workbook) As Long
175:    Dim oVbProject As VBProject
176:    Dim oRef   As Reference
177:    Dim i      As Long
178:    On Error GoTo ErrorHandler
179:    Set oVbProject = wb.VBProject
180:    For Each oRef In oVbProject.References
181:        If Not oRef.BuiltIn Then
182:            oVbProject.References.Remove oRef
183:            i = i + 1
184:        End If
185:    Next oRef
186:    RemoveAllReferences = i
187:    Exit Function
ErrorHandler:
189:    RemoveAllReferences = i
190:    Call MsgBox("Ошибка! в RemoveAllReferences" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
191: End Function
Private Function IsWorkbookOpen() As Boolean
193:    If Workbooks.Count > 0 Then
194:        IsWorkbookOpen = True
195:    Else
196:        Call MsgBox("Активная рабочая книга не найдена. Пожалуйста, сначала откройте книгу!", vbCritical, "Экспорт проекта VBA:")
197:    End If
End Function
