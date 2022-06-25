Attribute VB_Name = "I_StatisticVBAProj"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : I_StatisticVBAProj - формирование статистики кодовой базы на листе Excel
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Private Module
Option Explicit
Public Enum LineSplits
    LineSplitRemove = 0
    LineSplitKeep = 1
    LineSplitConvert = 2
End Enum
    Public Sub AddSheetStatistica()
9:    Dim i      As Integer
10:    Dim Rng    As Range
11:    Dim Form   As AddStatistic
12:    Dim strVar As String
13:    On Error GoTo errMsg
14:    If Not VBAIsTrusted Then
15:        Call MsgBox("Warning!" & vbLf & "Disabled: [Trust access to the VBE object model]" _
                   & vbLf & "To enable it, go to: File->Settings->Security Management Center->Macro Settings" _
                   & vbLf & "And restart Excel", vbCritical, "Access to the object model is denied:")
18:        Exit Sub
19:    End If
20:    Application.DisplayAlerts = False
21:    Set Form = New AddStatistic
22:    Form.Show
23:    strVar = Form.cmbMain.Value
24:    If strVar = vbNullString Then Exit Sub
25:    i = 1
26:    ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
27:    With ActiveSheet
28:        .Name = C_Const.SH_STATISTICA
29:        Set Rng = .Range("A1")
30:        Rng.Value = "Module name"
31:        i = i + 1
32:        Rng(1, i).Value = "Module type"
33:        i = i + 1
34:        Rng(1, i).Value = "Type of modifier"
35:        i = i + 1
36:        Rng(1, i).Value = "Type of procedure"
37:        i = i + 1
38:        Rng(1, i).Value = "Name of the procedure"
39:        i = i + 1
40:        Rng(1, i).Value = "Initial line"
41:        i = i + 1
42:        Rng(1, i).Value = "Number of rows"
43:        i = i + 1
44:        Rng(1, i).Value = "Declaration of the procedure"
45:    End With
46:    Call AddInfoProject(Workbooks(strVar))
47:    With ActiveSheet.UsedRange
48:        .EntireColumn.AutoFit
49:        .EntireRow.AutoFit
50:    End With
51:    Application.DisplayAlerts = True
52:    Exit Sub
errMsg:
54:    If Err.Number = 1004 Then
55:        ActiveWorkbook.Sheets(C_Const.SH_STATISTICA).Delete
56:        ActiveSheet.Name = C_Const.SH_STATISTICA
57:        Err.Clear
58:        Resume Next
59:    Else
60:        Call MsgBox("Error in I_StatisticVBAProj.AddSheetStatistica" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
61:        Call WriteErrorLog("I_StatisticVBAProj.AddSheetStatistica")
62:    End If
63:    Application.DisplayAlerts = True
64: End Sub
    Private Sub AddInfoProject(Optional ByRef wb As Workbook = Nothing)
66:    Dim VBP    As VBIDE.VBProject
67:    Dim vbComp As VBIDE.VBComponent
68:    If wb Is Nothing Then
69:        Set VBP = ActiveWorkbook.VBProject
70:    Else
71:        Set VBP = wb.VBProject
72:    End If
73:    If VBP.Protection = vbext_pp_locked Then
74:        Call MsgBox("On a VBA project, [" & VBP.Name & "] password set!", vbCritical, "Error, the project is password protected:")
75:        Exit Sub
76:    End If
77:    For Each vbComp In VBP.VBComponents
78:        Call ListProcedures(vbComp)
79:    Next vbComp
80: End Sub
     Public Sub ListProcedures(ByRef vbComp As VBIDE.VBComponent)
82:    Dim CodeMod As VBIDE.CodeModule
83:    Dim LineNum As Long
84:    Dim NumLines As Long
85:    Dim WS     As Worksheet
86:    Dim Rng    As Range
87:    Dim procName As String
88:    Dim ProcKind As VBIDE.vbext_ProcKind
89:    Dim i      As Integer
90:    Dim StrDeclarationProcedure As String
91:    Set CodeMod = vbComp.CodeModule
92:    Set WS = ActiveSheet
93:    Set Rng = WS.Range("A" & LastRowOrColumn(1) + 1)
94:    With CodeMod
95:        LineNum = .CountOfDeclarationLines + 1
96:        Do Until LineNum >= .CountOfLines
97:            procName = .ProcOfLine(LineNum, ProcKind)
98:            StrDeclarationProcedure = GetProcedureDeclaration(CodeMod, procName, ProcKind, LineSplitKeep)
99:            Rng.Value = vbComp.Name
100:            i = 1
101:            i = i + 1
102:            Rng(1, i).Value = ComponentTypeToString(vbComp.Type)
103:            i = i + 1
104:            Rng(1, i).Value = TypeOfAccessModifier(StrDeclarationProcedure)
105:            i = i + 1
106:            Rng(1, i).Value = TypeProcedyre(StrDeclarationProcedure)
107:            i = i + 1
108:            Rng(1, i).Value = procName
109:            i = i + 1
110:            Rng(1, i).Value = .ProcStartLine(procName, ProcKind)
111:            i = i + 1
112:            Rng(1, i).Value = .ProcCountLines(procName, ProcKind)
113:            i = i + 1
114:            Rng(1, i).Value = StrDeclarationProcedure
115:            LineNum = .ProcStartLine(procName, ProcKind) + _
                        .ProcCountLines(procName, ProcKind) + 1
117:            Set Rng = Rng(2, 1)
118:        Loop
119:    End With
120: End Sub
     Private Function TypeOfAccessModifier(ByRef StrDeclarationProcedure As String) As String
122:    If StrDeclarationProcedure Like "Private*" Then
123:        TypeOfAccessModifier = "Private"
124:    Else
125:        TypeOfAccessModifier = "Public"
126:    End If
127: End Function
     Public Function TypeProcedyre(ByRef StrDeclarationProcedure As String) As String
129:    If StrDeclarationProcedure Like "*Sub*" Then
130:        TypeProcedyre = "Sub"
131:    ElseIf StrDeclarationProcedure Like "*Function*" Then
132:        TypeProcedyre = "Function"
133:    ElseIf StrDeclarationProcedure Like "*Property Set*" Then
134:        TypeProcedyre = "Property Set"
135:    ElseIf StrDeclarationProcedure Like "*Property Get*" Then
136:        TypeProcedyre = "Property Get"
137:    ElseIf StrDeclarationProcedure Like "*Property Let*" Then
138:        TypeProcedyre = "Property Let"
139:    Else
140:        TypeProcedyre = "Unknown Type"
141:    End If
142: End Function
     Public Function ComponentTypeToString(ByRef ComponentType As VBIDE.vbext_ComponentType) As String
144:    Select Case ComponentType
        Case vbext_ct_ActiveXDesigner
146:            ComponentTypeToString = "ActiveX Designer"
147:        Case vbext_ct_ClassModule
148:            ComponentTypeToString = "Class Module"
149:        Case vbext_ct_Document
150:            ComponentTypeToString = "Document Module"
151:        Case vbext_ct_MSForm
152:            ComponentTypeToString = "UserForm"
153:        Case vbext_ct_StdModule
154:            ComponentTypeToString = "Code Module"
155:        Case Else
156:            ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
157:    End Select
158: End Function
     Public Function GetProcedureDeclaration( _
             ByRef CodeMod As VBIDE.CodeModule, _
             ByRef procName As String, ByRef ProcKind As VBIDE.vbext_ProcKind, _
             Optional ByRef LineSplitBehavior As LineSplits = LineSplitRemove) As Variant
163:    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
164:    ' GetProcedureDeclaration
165:    ' This return the procedure declaration of ProcName in CodeMod. The LineSplitBehavior
166:    ' determines what to do with procedure declaration that span more than one line using
167:    ' the "_" line continuation character. If LineSplitBehavior is LineSplitRemove, the
168:    ' entire procedure declaration is converted to a single line of text. If
169:    ' LineSplitBehavior is LineSplitKeep the "_" characters are retained and the
170:    ' declaration is split with vbNewLine into multiple lines. If LineSplitBehavior is
171:    ' LineSplitConvert, the "_" characters are removed and replaced with vbNewLine.
172:    ' The function returns vbNullString if the procedure could not be found.
173:    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
174:    Dim LineNum As Long
175:    Dim s      As String
176:    Dim Declaration As String
177:    On Error Resume Next
178:    LineNum = CodeMod.ProcBodyLine(procName, ProcKind)
179:    If Err.Number <> 0 Then
180:        Exit Function
181:    End If
182:    s = CodeMod.Lines(LineNum, 1)
183:    Do While Right$(s, 1) = "_"
184:        Select Case True
            Case LineSplitBehavior = LineSplitConvert
186:                s = Left$(s, Len(s) - 1) & vbNewLine
187:            Case LineSplitBehavior = LineSplitKeep
188:                s = s & vbNewLine
189:            Case LineSplitBehavior = LineSplitRemove
190:                s = Left$(s, Len(s) - 1) & " "
191:        End Select
192:        Declaration = Declaration & s
193:        LineNum = LineNum + 1
194:        s = CodeMod.Lines(LineNum, 1)
195:    Loop
196:    Declaration = SingleSpace(Declaration & s)
197:    GetProcedureDeclaration = Declaration
198: End Function
     Private Function SingleSpace(ByVal sText As String) As String
200:    Dim pos    As String
201:    pos = VBA.InStr(1, sText, Space(2), vbBinaryCompare)
202:    Do Until pos = 0
203:        sText = VBA.Replace(sText, Space(2), Space(1))
204:        pos = VBA.InStr(1, sText, Space(2), vbBinaryCompare)
205:    Loop
206:    SingleSpace = sText
207: End Function
     Private Function LastRowOrColumn(ByVal NomerRowOrColumn As Variant, _
             Optional ByRef WorkSheetName As Variant = vbNullString, _
             Optional ByRef RowOrColumn As Boolean = True) As Long
211:    'NomerRowOrColumn - номер или буква искомого столбца для строк, строки для столбцов, обезательный параметр
212:    'WorkSheetName - на каком листе искать, по умолчанию используется активный, не обезательный параметр
213:    'RowOrColumn - поиск строки или столбца, по умолчанию по ищем строку, не обезательный параметр
214:    On Error GoTo Err_msg_WSN
215:    Dim WH     As Worksheet
216:    If WorkSheetName = vbNullString Then
217:        Set WH = ActiveSheet
218:    ElseIf IsNumeric(WorkSheetName) Then
219:        Set WH = ActiveWorkbook.Worksheets(CInt(WorkSheetName))
220:    Else
221:        Set WH = ActiveWorkbook.Worksheets(WorkSheetName)
222:    End If
223:    If RowOrColumn Then
224:        If Not IsNumeric(NomerRowOrColumn) Then
225:            LastRowOrColumn = WH.Cells(Rows.Count, NomerRowOrColumn).End(xlUp).Row
226:        Else
227:            LastRowOrColumn = WH.Cells(Rows.Count, CInt(NomerRowOrColumn)).End(xlUp).Row
228:        End If
229:    Else
230:        If Not IsNumeric(NomerRowOrColumn) Then
231:            LastRowOrColumn = WH.Cells(NomerRowOrColumn, Columns.Count).End(xlToLeft).Column
232:        Else
233:            LastRowOrColumn = WH.Cells(CInt(NomerRowOrColumn), Columns.Count).End(xlToLeft).Column
234:        End If
235:    End If
236:    Exit Function
Err_msg_WSN:
238:    Select Case Err.Number
        Case 13, 1004:
240:            Call MsgBox("The following is not a valid value for a column or row number: [" & NomerRowOrColumn & "] ", vbCritical, "Input error:")
241:        Case 9:
242:            Call MsgBox("There is an invalid value in the file name: [" & WorkSheetName & "] ", vbCritical, "Input error:")
243:        Case Else:
244:            Call MsgBox("Error in LastRowOrColumn:" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
245:            Call WriteErrorLog("LastRowOrColumn")
246:    End Select
247: End Function

