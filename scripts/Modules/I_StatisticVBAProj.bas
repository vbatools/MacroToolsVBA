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
17:    Dim i      As Integer
18:    Dim Rng    As Range
19:    Dim Form   As AddStatistic
20:    Dim strVar As String
21:    On Error GoTo errmsg
22:    If Not VBAIsTrusted Then
23:        Call MsgBox("Warning!" & vbLf & "Disabled: [Trust access to the VBA object model]" _
                      & vbLf & "To enable it, go to: File->Settings->Security Management Center->Macro Settings" _
                      & vbLf & "And restart Excel", vbCritical, "No access to the object model:")
26:        Exit Sub
27:    End If
28:    Application.DisplayAlerts = False
29:    Set Form = New AddStatistic
30:    Form.Show
31:    strVar = Form.cmbMain.Value
32:    If strVar = vbNullString Then Exit Sub
33:    i = 1
34:    ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
35:    With ActiveSheet
36:        .Name = C_Const.SH_STATISTICA
37:        Set Rng = .Range("A1")
38:        Rng.Value = "Module name"
39:        i = i + 1
40:        Rng(1, i).Value = "Module type"
41:        i = i + 1
42:        Rng(1, i).Value = "Modifier type"
43:        i = i + 1
44:        Rng(1, i).Value = "Type of procedure"
45:        i = i + 1
46:        Rng(1, i).Value = "Name of the procedure"
47:        i = i + 1
48:        Rng(1, i).Value = "Start line"
49:        i = i + 1
50:        Rng(1, i).Value = "Number of rows"
51:        i = i + 1
52:        Rng(1, i).Value = "Declaring the procedure"
53:    End With
54:    Call AddInfoProject(Workbooks(strVar))
55:    With ActiveSheet.UsedRange
56:        .EntireColumn.AutoFit
57:        .EntireRow.AutoFit
58:    End With
59:    Application.DisplayAlerts = True
60:    Exit Sub
errmsg:
62:    If Err.Number = 1004 Then
63:        ActiveWorkbook.Sheets(C_Const.SH_STATISTICA).Delete
64:        ActiveSheet.Name = C_Const.SH_STATISTICA
65:        Err.Clear
66:        Resume Next
67:    Else
68:        Call MsgBox("Error in I_StatisticVBAProj.AddSheetStatistica" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbCritical, "Error:")
69:        Call WriteErrorLog("I_StatisticVBAProj.AddSheetStatistica")
70:    End If
71:    Application.DisplayAlerts = True
72: End Sub
    Private Sub AddInfoProject(Optional ByRef WB As Workbook = Nothing)
74:    Dim VBP    As VBIDE.VBProject
75:    Dim vbComp As VBIDE.VBComponent
76:    If WB Is Nothing Then
77:        Set VBP = ActiveWorkbook.VBProject
78:    Else
79:        Set VBP = WB.VBProject
80:    End If
81:    If VBP.Protection = vbext_pp_locked Then
82:        Call MsgBox("On a VBA project, [" & VBP.Name & "] password set!", vbCritical, "Error, the project is password protected:")
83:        Exit Sub
84:    End If
85:    For Each vbComp In VBP.VBComponents
86:        Call ListProcedures(vbComp)
87:    Next vbComp
88: End Sub
     Public Sub ListProcedures(ByRef vbComp As VBIDE.VBComponent)
90:    Dim CodeMod As VBIDE.CodeModule
91:    Dim LineNum As Long
92:    Dim NumLines As Long
93:    Dim WS     As Worksheet
94:    Dim Rng    As Range
95:    Dim procName As String
96:    Dim ProcKind As VBIDE.vbext_ProcKind
97:    Dim i      As Integer
98:    Dim StrDeclarationProcedure As String
99:    Set CodeMod = vbComp.CodeModule
100:    Set WS = ActiveSheet
101:    Set Rng = WS.Range("A" & LastRowOrColumn(1) + 1)
102:    With CodeMod
103:        LineNum = .CountOfDeclarationLines + 1
104:        Do Until LineNum >= .CountOfLines
105:            procName = .ProcOfLine(LineNum, ProcKind)
106:            StrDeclarationProcedure = GetProcedureDeclaration(CodeMod, procName, ProcKind, LineSplitKeep)
107:            Rng.Value = vbComp.Name
108:            i = 1
109:            i = i + 1
110:            Rng(1, i).Value = ComponentTypeToString(vbComp.Type)
111:            i = i + 1
112:            Rng(1, i).Value = TypeOfAccessModifier(StrDeclarationProcedure)
113:            i = i + 1
114:            Rng(1, i).Value = TypeProcedyre(StrDeclarationProcedure)
115:            i = i + 1
116:            Rng(1, i).Value = procName
117:            i = i + 1
118:            Rng(1, i).Value = .ProcStartLine(procName, ProcKind)
119:            i = i + 1
120:            Rng(1, i).Value = .ProcCountLines(procName, ProcKind)
121:            i = i + 1
122:            Rng(1, i).Value = StrDeclarationProcedure
123:            LineNum = .ProcStartLine(procName, ProcKind) + _
                            .ProcCountLines(procName, ProcKind) + 1
125:            Set Rng = Rng(2, 1)
126:        Loop
127:    End With
128: End Sub
     Private Function TypeOfAccessModifier(ByRef StrDeclarationProcedure As String) As String
130:    If StrDeclarationProcedure Like "Private*" Then
131:        TypeOfAccessModifier = "Private"
132:    Else
133:        TypeOfAccessModifier = "Public"
134:    End If
135: End Function
     Public Function TypeProcedyre(ByRef StrDeclarationProcedure As String) As String
137:    If StrDeclarationProcedure Like "*Sub*" Then
138:        TypeProcedyre = "Sub"
139:    ElseIf StrDeclarationProcedure Like "*Function*" Then
140:        TypeProcedyre = "Function"
141:    ElseIf StrDeclarationProcedure Like "*Property Set*" Then
142:        TypeProcedyre = "Property Set"
143:    ElseIf StrDeclarationProcedure Like "*Property Get*" Then
144:        TypeProcedyre = "Property Get"
145:    ElseIf StrDeclarationProcedure Like "*Property Let*" Then
146:        TypeProcedyre = "Property Let"
147:    Else
148:        TypeProcedyre = "Unknown Type"
149:    End If
150: End Function
     Public Function ComponentTypeToString(ByRef ComponentType As VBIDE.vbext_ComponentType) As String
152:    Select Case ComponentType
        Case vbext_ct_ActiveXDesigner
154:            ComponentTypeToString = "ActiveX Designer"
155:        Case vbext_ct_ClassModule
156:            ComponentTypeToString = "Class Module"
157:        Case vbext_ct_Document
158:            ComponentTypeToString = "Document Module"
159:        Case vbext_ct_MSForm
160:            ComponentTypeToString = "UserForm"
161:        Case vbext_ct_StdModule
162:            ComponentTypeToString = "Code Module"
163:        Case Else
164:            ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
165:    End Select
166: End Function
     Public Function GetProcedureDeclaration( _
                  ByRef CodeMod As VBIDE.CodeModule, _
                  ByRef procName As String, ByRef ProcKind As VBIDE.vbext_ProcKind, _
                  Optional ByRef LineSplitBehavior As LineSplits = LineSplitRemove) As Variant
171:    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
172:    ' GetProcedureDeclaration
173:    ' This return the procedure declaration of ProcName in CodeMod. The LineSplitBehavior
174:    ' determines what to do with procedure declaration that span more than one line using
175:    ' the "_" line continuation character. If LineSplitBehavior is LineSplitRemove, the
176:    ' entire procedure declaration is converted to a single line of text. If
177:    ' LineSplitBehavior is LineSplitKeep the "_" characters are retained and the
178:    ' declaration is split with vbNewLine into multiple lines. If LineSplitBehavior is
179:    ' LineSplitConvert, the "_" characters are removed and replaced with vbNewLine.
180:    ' The function returns vbNullString if the procedure could not be found.
181:    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
182:    Dim LineNum As Long
183:    Dim s      As String
184:    Dim Declaration As String
185:    On Error Resume Next
186:    LineNum = CodeMod.ProcBodyLine(procName, ProcKind)
187:    If Err.Number <> 0 Then
188:        Exit Function
189:    End If
190:    s = CodeMod.Lines(LineNum, 1)
191:    Do While Right$(s, 1) = "_"
192:        Select Case True
            Case LineSplitBehavior = LineSplitConvert
194:                s = Left$(s, Len(s) - 1) & vbNewLine
195:            Case LineSplitBehavior = LineSplitKeep
196:                s = s & vbNewLine
197:            Case LineSplitBehavior = LineSplitRemove
198:                s = Left$(s, Len(s) - 1) & " "
199:        End Select
200:        Declaration = Declaration & s
201:        LineNum = LineNum + 1
202:        s = CodeMod.Lines(LineNum, 1)
203:    Loop
204:    Declaration = SingleSpace(Declaration & s)
205:    GetProcedureDeclaration = Declaration
206: End Function
     Private Function SingleSpace(ByVal sText As String) As String
208:    Dim pos    As String
209:    pos = VBA.InStr(1, sText, Space(2), vbBinaryCompare)
210:    Do Until pos = 0
211:        sText = VBA.Replace(sText, Space(2), Space(1))
212:        pos = VBA.InStr(1, sText, Space(2), vbBinaryCompare)
213:    Loop
214:    SingleSpace = sText
215: End Function
     Private Function LastRowOrColumn(ByVal NomerRowOrColumn As Variant, _
                  Optional ByRef WorkSheetName As Variant = vbNullString, _
                  Optional ByRef RowOrColumn As Boolean = True) As Long
219:    'NomerRowOrColumn - номер или буква искомого столбца для строк, строки для столбцов, обезательный параметр
220:    'WorkSheetName - на каком листе искать, по умолчанию используется активный, не обезательный параметр
221:    'RowOrColumn - поиск строки или столбца, по умолчанию по ищем строку, не обезательный параметр
222:    On Error GoTo Err_msg_WSN
223:    Dim WH     As Worksheet
224:    If WorkSheetName = vbNullString Then
225:        Set WH = ActiveSheet
226:    ElseIf IsNumeric(WorkSheetName) Then
227:        Set WH = ActiveWorkbook.Worksheets(CInt(WorkSheetName))
228:    Else
229:        Set WH = ActiveWorkbook.Worksheets(WorkSheetName)
230:    End If
231:    If RowOrColumn Then
232:        If Not IsNumeric(NomerRowOrColumn) Then
233:            LastRowOrColumn = WH.Cells(Rows.Count, NomerRowOrColumn).End(xlUp).Row
234:        Else
235:            LastRowOrColumn = WH.Cells(Rows.Count, CInt(NomerRowOrColumn)).End(xlUp).Row
236:        End If
237:    Else
238:        If Not IsNumeric(NomerRowOrColumn) Then
239:            LastRowOrColumn = WH.Cells(NomerRowOrColumn, Columns.Count).End(xlToLeft).Column
240:        Else
241:            LastRowOrColumn = WH.Cells(CInt(NomerRowOrColumn), Columns.Count).End(xlToLeft).Column
242:        End If
243:    End If
244:    Exit Function
Err_msg_WSN:
246:    Select Case Err.Number
        Case 13, 1004:
248:            Call MsgBox("Invalid column or row number value entered: [" & NomerRowOrColumn & "] ", vbCritical, "Input error:")
249:        Case 9:
250:            Call MsgBox("Invalid value entered in the file name: [" & WorkSheetName & "] ", vbCritical, "Input error:")
251:        Case Else:
252:            Call MsgBox("Error in LastRowOrColumn:" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbCritical, "Error:")
253:            Call WriteErrorLog("LastRowOrColumn")
254:    End Select
255: End Function

