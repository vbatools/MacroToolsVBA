Attribute VB_Name = "K_AddNumbersLine"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : K_AddNumbersLine - создание номерации строк кода VBA
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Private Module
Option Explicit
Public Enum vbLineNumbers_LabelTypes
    vbLabelColon    ' 0
    vbLabelTab    ' 1
End Enum
Private Enum vbLineNumbers_ScopeToAddLineNumbersTo
    vbScopeAllProc    ' 1
    vbScopeThisProc    ' 2
End Enum
    Public Sub AddLineNumbers_()
20:    On Error GoTo ErrorHandler
21:    Dim cmb_txt As String
22:    Dim vbComp As VBIDE.VBComponent
23:    cmb_txt = B_CreateMenus.WhatIsTextInComboBoxHave
24:    Select Case cmb_txt
        Case C_Const.ALLVBAPROJECT:
26:            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
27:                AddLineNumbers vbCompObj:=vbComp, LabelType:=vbLabelColon, AddLineNumbersToEmptyLines:=True, AddLineNumbersToEndOfProc:=True, Scope:=vbScopeAllProc
28:            Next vbComp
29:        Case C_Const.SELECTEDMODULE:
30:            AddLineNumbers vbCompObj:=Application.VBE.ActiveCodePane.CodeModule.Parent, LabelType:=vbLabelColon, AddLineNumbersToEmptyLines:=True, AddLineNumbersToEndOfProc:=True, Scope:=vbScopeAllProc
31:    End Select
32:    Exit Sub
ErrorHandler:
34:    Select Case Err.Number
        Case 91:
36:            Exit Sub
37:        Case Else:
38:            Debug.Print "Ошибка! в AddLineNumbers_" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
39:            Call WriteErrorLog("AddLineNumbers_")
40:    End Select
41:    Err.Clear
42: End Sub
    Public Sub RemoveLineNumbers_()
44:    On Error GoTo ErrorHandler
45:    Dim cmb_txt As String
46:    Dim vbComp As VBIDE.VBComponent
47:    cmb_txt = B_CreateMenus.WhatIsTextInComboBoxHave
48:    Select Case cmb_txt
        Case C_Const.ALLVBAPROJECT:
50:            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
51:                RemoveLineNumbers vbCompObj:=vbComp, LabelType:=vbLabelColon
52:                RemoveLineNumbers vbCompObj:=vbComp, LabelType:=vbLabelTab
53:            Next vbComp
54:        Case C_Const.SELECTEDMODULE:
55:            RemoveLineNumbers vbCompObj:=Application.VBE.ActiveCodePane.CodeModule.Parent, LabelType:=vbLabelColon
56:            RemoveLineNumbers vbCompObj:=Application.VBE.ActiveCodePane.CodeModule.Parent, LabelType:=vbLabelTab
57:    End Select
58:    Exit Sub
ErrorHandler:
60:    Select Case Err.Number
        Case 91:
62:            Exit Sub
63:        Case Else:
64:            Debug.Print "Ошибка! в RemoveLineNumbers_" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
65:            Call WriteErrorLog("RemoveLineNumbers_")
66:    End Select
67:    Err.Clear
68: End Sub
     Private Sub AddLineNumbers( _
                  ByVal vbCompObj As VBIDE.VBComponent, _
                  ByVal LabelType As vbLineNumbers_LabelTypes, _
                  ByVal AddLineNumbersToEmptyLines As Boolean, _
                  ByVal AddLineNumbersToEndOfProc As Boolean, _
                  ByVal Scope As vbLineNumbers_ScopeToAddLineNumbersTo)
75:    ' USAGE RULES
76:    ' DO NOT MIX LABEL TYPES FOR LINE NUMBERS! IF ADDING LINE NUMBERS AS COLON TYPE, ANY LINE NUMBERS AS VBTAB TYPE MUST BE REMOVE BEFORE, AND RECIPROCALLY ADDING LINE NUMBERS AS VBTAB TYPE
77:    Dim i      As Long
78:    Dim procName As String
79:    Dim startOfProcedure As Long
80:    Dim lengthOfProcedure As Long
81:    Dim endOfProcedure As Long
82:    Dim bodyOfProcedure As Long
83:    Dim countOfProcedure As Long
84:    Dim prelinesOfProcedure As Long
85:    Dim PreviousIndentAdded As Long
86:    Dim strLine As String
87:    Dim temp_strLine As String
88:    Dim new_strLine As String
89:    Dim tupe_procedure As vbext_ProcKind
90:    Dim InProcBodyLines As Boolean
91:    Dim FlagSelect As Boolean
92:    With vbCompObj.CodeModule
93:
94:        If Scope = vbScopeAllProc Then
95:            For i = 1 To .CountOfLines - 1
96:                strLine = .Lines(i, 1)
97:                If FlagSelect Then
98:                    FlagSelect = False
99:                    GoTo NextLine
100:                End If
101:                If strLine Like "*Select Case *" Then FlagSelect = True
                procName = .ProcOfLine(i, tupe_procedure)    ' Type d'argument ByRef incompatible ~~> Requires VBIDE library as a Reference for the VBA Project
103:                If procName <> vbNullString Then
104:                    startOfProcedure = .ProcStartLine(procName, tupe_procedure)
105:                    bodyOfProcedure = .ProcBodyLine(procName, tupe_procedure)
106:                    countOfProcedure = .ProcCountLines(procName, tupe_procedure)
107:                    prelinesOfProcedure = bodyOfProcedure - startOfProcedure
108:                    'postlineOfProcedure = ??? not directly available since endOfProcedure is itself not directly available.
109:                    lengthOfProcedure = countOfProcedure - prelinesOfProcedure    ' includes postlinesOfProcedure !
110:                    'endOfProcedure = ??? not directly available, each line of the proc must be tested until the End statement is reached. See below.
111:                    If endOfProcedure <> 0 And startOfProcedure < endOfProcedure And i > endOfProcedure Then
112:                        GoTo NextLine
113:                    End If
114:                    If i = bodyOfProcedure Then InProcBodyLines = True
115:                    If bodyOfProcedure < i And i < startOfProcedure + countOfProcedure Then
116:                        If Not (.Lines(i - 1, 1) Like "* _") Then
117:                            InProcBodyLines = False
118:                            PreviousIndentAdded = 0
119:                            If Trim$(strLine) = vbNullString And Not AddLineNumbersToEmptyLines Then GoTo NextLine
120:                            If IsProcEndLine(vbCompObj, i) Then
121:                                endOfProcedure = i
122:                                If AddLineNumbersToEndOfProc Then
123:                                    Call IndentProcBodyLinesAsProcEndLine(vbCompObj, LabelType, endOfProcedure, tupe_procedure)
124:                                Else
125:                                    GoTo NextLine
126:                                End If
127:                            End If
128:                            If LabelType = vbLabelColon Then
129:                                If HasLabel(strLine, vbLabelColon) Then strLine = RemoveOneLineNumber(.Lines(i, 1), vbLabelColon)
130:                                If Not HasLabel(strLine, vbLabelColon) Then
131:                                    temp_strLine = strLine
132:                                    On Error Resume Next
133:                                    .ReplaceLine i, CStr(i) & ":" & strLine
134:                                    On Error GoTo 0
135:                                    new_strLine = .Lines(i, 1)
136:                                    If Len(new_strLine) = Len(CStr(i) & ":" & temp_strLine) Then
137:                                        PreviousIndentAdded = Len(CStr(i) & ":")
138:                                    Else
139:                                        PreviousIndentAdded = Len(CStr(i) & ": ")
140:                                    End If
141:                                End If
142:                            ElseIf LabelType = vbLabelTab Then
143:                                If Not HasLabel(strLine, vbLabelTab) Then strLine = RemoveOneLineNumber(.Lines(i, 1), vbLabelTab)
144:                                If Not HasLabel(strLine, vbLabelColon) Then
145:                                    temp_strLine = strLine
146:                                    On Error Resume Next
147:                                    .ReplaceLine i, CStr(i) & vbTab & strLine
148:                                    On Error GoTo 0
149:                                    PreviousIndentAdded = Len(strLine) - Len(temp_strLine)
150:                                End If
151:                            End If
152:                        Else
153:                            If Not InProcBodyLines Then
154:                                If LabelType = vbLabelColon Then
155:                                    On Error Resume Next
156:                                    .ReplaceLine i, Space(PreviousIndentAdded) & strLine
157:                                    On Error GoTo 0
158:                                ElseIf LabelType = vbLabelTab Then
159:                                    On Error Resume Next
160:                                    .ReplaceLine i, Space(4) & strLine
161:                                    On Error GoTo 0
162:                                End If
163:                            Else
164:                            End If
165:                        End If
166:                    End If
167:                End If
NextLine:
169:            Next i
170:        ElseIf AddLineNumbersToEmptyLines And Scope = vbScopeThisProc Then
171:            'TODO selected prosedure
172:        End If
173:
174:    End With
175: End Sub
     Private Function IsProcEndLine( _
                  ByVal vbCompObj As VBIDE.VBComponent, _
                  ByVal lLine As Long) As Boolean
179:    With vbCompObj.CodeModule
180:        If Trim$(.Lines(lLine, 1)) Like "End Sub*" _
                        Or Trim$(.Lines(lLine, 1)) Like "End Function*" _
                        Or Trim$(.Lines(lLine, 1)) Like "End Property*" _
                        Then IsProcEndLine = True
184:    End With
185: End Function
     Private Sub IndentProcBodyLinesAsProcEndLine( _
                  ByVal vbCompObj As VBIDE.VBComponent, _
                  ByVal LabelType As vbLineNumbers_LabelTypes, _
                  ByVal ProcEndLine As Long, _
                  ByVal VBEXT As vbext_ProcKind)
191:    Dim procName As String
192:    Dim bodyOfProcedure As Long
193:    Dim j      As Long
194:    Dim endOfProcedure As Long
195:    Dim strEnd As String
196:    Dim strLine As String
197:    With vbCompObj.CodeModule
198:        procName = .ProcOfLine(ProcEndLine, VBEXT)
199:        bodyOfProcedure = .ProcBodyLine(procName, VBEXT)
200:        endOfProcedure = ProcEndLine
201:        strEnd = .Lines(endOfProcedure, 1)
202:        j = bodyOfProcedure
203:        If j = 1 Then j = 2
204:        Do Until Not .Lines(j - 1, 1) Like "* _" And j <> bodyOfProcedure
205:            strLine = .Lines(j, 1)
206:            If LabelType = vbLabelColon Then
207:                If Mid$(strEnd, Len(CStr(endOfProcedure)) + 1 + 1 + 1, 1) = " " Then
208:                    On Error Resume Next
209:                    .ReplaceLine j, Space(Len(CStr(endOfProcedure)) + 1) & strLine
210:                    On Error GoTo 0
211:                Else
212:                    On Error Resume Next
213:                    .ReplaceLine j, Space(Len(CStr(endOfProcedure)) + 2) & strLine
214:                    On Error GoTo 0
215:                End If
216:            ElseIf LabelType = vbLabelTab Then
217:                If endOfProcedure < 1000 Then
218:                    On Error Resume Next
219:                    .ReplaceLine j, Space(4) & strLine
220:                    On Error GoTo 0
221:                Else
222:                    Debug.Print "Этот инструмент ограничен 999 строками кода для правильной работы."
223:                End If
224:            End If
225:            j = j + 1
226:        Loop
227:    End With
228: End Sub
     Public Sub RemoveLineNumbers(ByVal vbCompObj As VBIDE.VBComponent, ByVal LabelType As vbLineNumbers_LabelTypes)
230:    Dim i      As Long
231:    Dim RemovedChars_previous_i As Long
232:    Dim procName As String
233:    Dim InProcBodyLines As Boolean
234:    Dim tupe_procedure As vbext_ProcKind
235:    With vbCompObj.CodeModule
236:        'Debug.Print ("nr of lines = " & .CountOfLines & vbNewLine & "Procname = " & procName)
237:        'Debug.Print ("nr of lines REMEMBER MUST BE LARGER THAN 7! = " & .CountOfLines)
238:        For i = 1 To .CountOfLines
239:            procName = .ProcOfLine(i, tupe_procedure)
240:            If procName <> vbNullString Then
241:                If i > 1 Then
242:                    'Debug.Print ("Line " & i & " is a body line " & .ProcBodyLine(procName, tupe_procedure))
243:                    If i = .ProcBodyLine(procName, tupe_procedure) Then InProcBodyLines = True
244:                    If Not .Lines(i - 1, 1) Like "* _" Then
245:                        'Debug.Print (InProcBodyLines)
246:                        InProcBodyLines = False
247:                        'Debug.Print ("recoginized a line that should be substituted: " & i)
248:                        'Debug.Print ("about to replace " & .Lines(i, 1) & vbNewLine & " with: " & RemoveOneLineNumber(.Lines(i, 1), LabelType) & vbNewLine & " with label type: " & LabelType)
249:                        On Error Resume Next
250:                        .ReplaceLine i, RemoveOneLineNumber(.Lines(i, 1), LabelType)
251:                        On Error GoTo 0
252:                    Else
253:                        If InProcBodyLines Then
254:                            ' do nothing
255:                            'Debug.Print i
256:                        Else
257:                            On Error Resume Next
258:                            .ReplaceLine i, Mid$(.Lines(i, 1), RemovedChars_previous_i + 1)
259:                            On Error GoTo 0
260:                        End If
261:                    End If
262:                End If
263:            Else
264:            End If
265:        Next i
266:    End With
267: End Sub
     Private Function RemoveOneLineNumber(ByVal aString As String, ByVal LabelType As vbLineNumbers_LabelTypes) As Variant
269:    RemoveOneLineNumber = aString
270:    If LabelType = vbLabelColon Then
271:        If aString Like "#:*" Or aString Like "##:*" Or aString Like "###:*" Or aString Like "####:*" Then
272:            RemoveOneLineNumber = Mid$(aString, 1 + InStr(1, aString, ":", vbTextCompare))
273:            If Left$(RemoveOneLineNumber, 2) Like " [! ]*" Then RemoveOneLineNumber = Mid$(RemoveOneLineNumber, 2)
274:        End If
275:    ElseIf LabelType = vbLabelTab Then
276:        If aString Like "#   *" Or aString Like "##  *" Or aString Like "### *" Or aString Like "#### *" Then RemoveOneLineNumber = Mid$(aString, 5)
277:        If aString Like "#" Or aString Like "##" Or aString Like "###" Or aString Like "####" Then RemoveOneLineNumber = vbNullString
278:    End If
211:     If RemoveOneLineNumber Like "*Function *" Or RemoveOneLineNumber Like "*Sub *" _
            Or RemoveOneLineNumber Like "*Property Set *" Or RemoveOneLineNumber Like "*Property Get *" Or RemoveOneLineNumber Like "*Property Let *" Then
281:        RemoveOneLineNumber = RemoveLeadingSpaces(RemoveOneLineNumber)
282:    End If
283: End Function
     Private Function HasLabel(ByVal aString As String, ByVal LabelType As vbLineNumbers_LabelTypes) As Boolean
285:    If LabelType = vbLabelColon Then HasLabel = InStr(1, aString & ":", ":") < InStr(1, aString & " ", " ")
286:    If LabelType = vbLabelTab Then
287:        HasLabel = Mid$(aString, 1, 4) Like "#   " Or Mid$(aString, 1, 4) Like "##  " Or Mid$(aString, 1, 4) Like "### " Or Mid$(aString, 1, 5) Like "#### "
288:    End If
289: End Function
'удаляет все пробелы вначале строки
Private Function RemoveLeadingSpaces(ByVal aString As String) As String
292:    Do Until Left$(aString, 1) <> " "
293:        aString = Mid$(aString, 2)
294:    Loop
295:    RemoveLeadingSpaces = aString
End Function
