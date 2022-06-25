Attribute VB_Name = "N_ObfMainNew"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : N_ObfMainNew - модуль обфускации кода
'* Created    : 08-10-2020 14:11
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

    Public Sub StartObfuscation()
12:    Dim Form        As AddStatistic
13:    Dim sNameWB     As String
14:    Dim objWB       As Object
15:
16:    'On Error GoTo ErrStartParser
17:    Set Form = New AddStatistic
18:    With Form
19:        .Caption = "Code obfuscation:"
20:        .lbOK.Caption = "OBFUSCATE"
21:        .chQuestion.visible = True
22:        .chQuestion.Value = True
23:        .lbWord.Caption = 1
24:        .Show
25:        sNameWB = .cmbMain.Value
26:    End With
27:    If sNameWB = vbNullString Then Exit Sub
28:
29:    If sNameWB Like "*.docm" Or sNameWB Like "*.DOCM" Then
30:        Dim objWrdApp As Object
31:        Set objWrdApp = GetObject(, "Word.Application")
32:        Set objWB = objWrdApp.Documents(sNameWB)
33:    Else
34:        Set objWB = Workbooks(sNameWB)
35:    End If
36:
37:    Call MainObfuscation(objWB, Form.chQuestion.Value)
38:    Set Form = Nothing
39:    Exit Sub
ErrStartParser:
41:    Application.Calculation = xlCalculationAutomatic
42:    Application.ScreenUpdating = True
43:    Call MsgBox("Error in N_ObfParserVBA.StartParser" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
44:    Call WriteErrorLog("AddShapeStatistic")
45: End Sub

    Private Sub MainObfuscation(ByRef objWB As Object, Optional bEncodeStr As Boolean = False)
48:    'On Error GoTo ErrStartParser
49:    If objWB.VBProject.Protection = vbext_pp_locked Then
50:        Call MsgBox("The project is protected, remove the password!", vbCritical, "Project:")
51:    Else
52:        If ActiveSheet.Name = NAME_SH Then
53:            Application.ScreenUpdating = False
54:            Application.Calculation = xlCalculationManual
55:            Application.EnableEvents = False
56:
57:            Call Obfuscation(objWB, bEncodeStr)
58:
59:            With ActiveWorkbook.Worksheets(NAME_SH)
60:                Call SortTabel(ActiveWorkbook.Worksheets(NAME_SH), .Range(.Cells(1, 1), .Cells(1, 13)).Address, "M1", 1)
61:            End With
62:
63:            Application.EnableEvents = True
64:            Application.Calculation = xlCalculationAutomatic
65:            Application.ScreenUpdating = True
66:            Call MsgBox("Book code [" & objWB.Name & "] encrypted!", vbInformation, "Code encryption:")
67:        Else
68:            Call MsgBox("Create or navigate to the sheet: [" & NAME_SH & "]", vbCritical, "Activating the sheet:")
69:        End If
70:    End If
71:    Exit Sub
ErrStartParser:
73:    Application.EnableEvents = True
74:    Application.Calculation = xlCalculationAutomatic
75:    Application.ScreenUpdating = True
76:    Call MsgBox("Error in MainObfuscation" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
77: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : FelterAdd - фильтрация в нужном порядке перед шифрованием
'* Created    : 29-07-2020 09:58
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub FelterAdd()
87:    Dim LastRow     As Long
88:    With ActiveWorkbook.Worksheets(NAME_SH)
89:        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
90:        If LastRow > 1 Then
91:            .Range(.Cells(2, 12), .Cells(LastRow, 12)).FormulaR1C1 = "=LEN(RC[-4])"
92:            .Range(.Cells(2, 13), .Cells(LastRow, 13)).FormulaR1C1 = "=R[-1]C+1"
93:            .Range(.Cells(2, 13), .Cells(LastRow, 13)).Value = .Range(.Cells(2, 13), .Cells(LastRow, 13)).Value
94:            Call SortTabel(ActiveWorkbook.Worksheets(NAME_SH), .Range(.Cells(1, 1), .Cells(1, 13)).Address, "L1", 2)
95:        End If
96:    End With
97: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : Obfuscation - главная процедура шифрования
'* Created    : 20-04-2020 18:26
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objWB As Workbook               : книга
'* Optional bEncodeStr As Boolean = True : шифровать строковые значения
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub Obfuscation(ByRef objWB As Object, Optional bEncodeStr As Boolean = True)
112:    Dim arrData     As Variant
113:    Dim i           As Long
114:    Dim j           As Long
115:    Dim sKey        As String
116:    Dim sCode       As String
117:    Dim sFinde      As String
118:    Dim sReplace    As String
119:    Dim sPattern    As String
120:
121:    Dim objDictName As Scripting.Dictionary
122:    Dim objDictModule As Scripting.Dictionary
123:    Dim objDictModuleOld As Scripting.Dictionary
124:    Dim objVBCitem  As VBIDE.VBComponent
125:    Dim dTime       As Date
126:
127:    dTime = Now()
128:    Debug.Print "Start:" & VBA.Format$(Now() - dTime, "Long Time")
129:
130:    Set objDictName = New Scripting.Dictionary
131:    Set objDictModule = New Scripting.Dictionary
132:    Set objDictModuleOld = New Scripting.Dictionary
133:
134:    'сохранение и загрузка
135:    objWB.SaveAs Filename:=objWB.Path & Application.PathSeparator & C_PublicFunctions.sGetBaseName(objWB.FullName) & "_obf_" & Replace(Now(), ":", ".") & "." & C_PublicFunctions.sGetExtensionName(objWB.FullName)    ', FileFormat:=objWB.FileFormat
136:
137:    Debug.Print "File saving - completed:" & VBA.Format$(Now() - dTime, "Long Time")
138:    'фильтрация
139:    Call FelterAdd
140:
141:    'считывание данных
142:    With ActiveWorkbook.Worksheets(NAME_SH)
143:        .Activate
144:        i = .Cells(Rows.Count, 1).End(xlUp).Row
145:        arrData = .Range(Cells(2, 1), Cells(i, 10)).Value2
146:    End With
147:
148:
149:    'сбор имен с шифрами
150:    For i = LBound(arrData) To UBound(arrData)
151:        If arrData(i, 9) = "yes" Then
152:            'сбор имен с шифрами
153:            If objDictName.Exists(arrData(i, 8)) = False Then objDictName.Add arrData(i, 8), arrData(i, 10)
154:        End If
155:    Next i
156:
157:    'сбор кода из модулей
158:    For Each objVBCitem In objWB.VBProject.VBComponents
159:        If objDictModule.Exists(objVBCitem.Name) = False Then
160:            sCode = GetCodeFromModule(objVBCitem)
161:            'убираю перенос строк
162:            sCode = VBA.Replace(sCode, " _" & vbNewLine, "B")
163:            objDictModule.Add objVBCitem.Name, sCode
164:            objDictModuleOld.Add objVBCitem.Name, sCode
165:            sCode = vbNullString
166:        End If
167:    Next objVBCitem
168:    'конец сбора
169:
170:    Debug.Print "Data collection - completed:" & VBA.Format$(Now() - dTime, "Long Time")
171:
172:    'шифрование
173:    sCode = vbNullString
174:    With objDictName
175:        For i = 0 To .Count - 1
176:            For j = 0 To objDictModule.Count - 1
177:                sFinde = .Keys(i)
178:                sReplace = .Items(i)
179:                sKey = objDictModule.Keys(j)
180:                sCode = objDictModule.Item(sKey)
181:                If sCode Like "*" & sFinde & "*" And VBA.Len(sFinde) > 1 Then
182:                    sPattern = "([\*\.\^\*\+\#\(\)\-\=\/\,\:\;\s\" & VBA.Chr$(34) & "])" & sFinde & "([\*\.\^\*\+\!\@\#\$\%\&\(\)\-\=\/\,\:\;\s\" & VBA.Chr$(34) & "]|$)"
183:                    sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
184:                    If sCode <> vbNullString Then objDictModule.Item(sKey) = sCode
185:                End If
186:                'регулярка для событий в основном для форм
187:                If sCode Like "* " & Chr$(83) & "ub *" & sFinde & "_*(*)*" Then
188:                    sPattern = "([\s])(Sub)([\s])" & sFinde & "(\_{1}[A-Za-zА-Яа-яЁё]{4,40}\([A-Za-zА-Яа-яЁё\s\.\,]{0,100}\))"
189:                    sCode = RegExpFindReplace(sCode, sPattern, "$1$2$3" & sReplace & "$4", True, False, False)
190:                    If sCode <> vbNullString Then objDictModule.Item(sKey) = sCode
191:                    sPattern = "([\s])" & sFinde & "(\_{1}[A-Za-zА-Яа-яЁё]{4,40}(?:\:\s|\n|\r))"
192:                    sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
193:                    If sCode <> vbNullString Then objDictModule.Item(sKey) = sCode
194:                End If
195:                sCode = vbNullString
196:            Next j
197:            DoEvents
198:            Application.StatusBar = "Data encryption - completed:" & Format(i / (.Count - 1), "Percent") & ", " & i & "from" & .Count - 1
199:        Next i
200:    End With
201:    Application.StatusBar = False
202:    'конец
203:
204:    'чередование
205:    sCode = vbNullString
206:
207:    For j = 0 To objDictModule.Count - 1
208:        Dim arrNew  As Variant
209:        Dim arrOld  As Variant
210:        Dim sTemp   As String
211:        arrNew = VBA.Split(objDictModule.Items(j), vbNewLine)
212:        arrOld = VBA.Split(objDictModuleOld.Items(j), vbNewLine)
213:        For i = LBound(arrNew) To UBound(arrNew)
214:            If arrNew(i) = vbNullString Or VBA.Left$(VBA.Trim$(arrNew(i)), 1) = "'" Then
215:                sTemp = vbNullString
216:            Else
217:                sTemp = "'" & arrOld(i) & vbNewLine
218:            End If
219:            sCode = sCode & sTemp & arrNew(i) & vbNewLine
220:            sTemp = vbNullString
221:        Next i
222:        sKey = objDictModule.Keys(j)
223:        objDictModule.Item(sKey) = sCode
224:        sCode = vbNullString
225:    Next j
226:    Debug.Print "String alternation - completed:" & VBA.Format$(Now() - dTime, "Long Time")
227:
228:
229:    Debug.Print "Data encryption - completed:" & VBA.Format$(Now() - dTime, "Long Time")
230:    'загрузка кода
231:    For j = 0 To objDictModule.Count - 1
232:        Set objVBCitem = objWB.VBProject.VBComponents(objDictModule.Keys(j))
233:        sCode = objDictModule.Items(j)
234:        'возврат перенос строк
235:        sCode = VBA.Replace(sCode, "B", " _" & vbNewLine)
236:        Call SetCodeInModule(objVBCitem, sCode)
237:    Next j
238:
239:    Debug.Print "Code loading- completed:" & VBA.Format$(Now() - dTime, "Long Time")
240:
241:    'переименование контролов
242:    For i = LBound(arrData) To UBound(arrData)
243:        If arrData(i, 9) = "yes" And objDictName.Exists(arrData(i, 8)) Then
244:            If arrData(i, 1) = "Control" Then
245:                Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
246:                objVBCitem.Designer.Controls(arrData(i, 8)).Name = arrData(i, 10)
247:            End If
248:        End If
249:    Next i
250:
251:    Debug.Print "Renaming of controls - completed:" & VBA.Format$(Now() - dTime, "Long Time")
252:
253:    'переименование модулей
254:    For i = LBound(arrData) To UBound(arrData)
255:        If arrData(i, 9) = "yes" And objDictName.Exists(arrData(i, 8)) Then
256:            If arrData(i, 1) = "Module" And VBA.CByte(arrData(i, 2)) <> 100 Then
257:                Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
258:                objVBCitem.Name = arrData(i, 10)
259:            End If
260:        End If
261:    Next i
262:
263:    'шифрование строк
264:    If bEncodeStr Then Call EncodedStringCode(objWB)
265:
266:    Debug.Print "Renaming modules- completed:" & VBA.Format$(Now() - dTime, "Long Time")
267:    objWB.Save
268:
269: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : EncodedStringCode - шифрование строковый значений кода
'* Created    : 29-07-2020 10:00
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef objWB As Workbook :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub EncodedStringCode(ByRef objWB As Object)
283:    Dim arrData     As Variant
284:    Dim i           As Long
285:    Dim sCodeString As String
286:    Dim objVBCitem  As VBIDE.VBComponent
287:    Dim sCode       As String
288:
289:    'считывание данных
290:    With ActiveWorkbook.Worksheets(NAME_SH_STR)
291:        .Activate
292:        arrData = .Range(Cells(2, 1), Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 9)).Value2
293:    End With
294:    'сбор строк
295:    sCodeString = "Option Explicit" & VBA.Chr$(13)
296:    For i = LBound(arrData) To UBound(arrData)
297:        If arrData(i, 7) = "yes" Then
298:            sCodeString = sCodeString & "Public Const " & arrData(i, 8) & " as string=" & arrData(i, 5) & VBA.Chr$(13)
299:        End If
300:    Next i
301:    Dim NameOldMOdule As String
302:    For i = LBound(arrData) To UBound(arrData)
303:        If arrData(i, 7) = "yes" Then
304:            If NameOldMOdule <> arrData(i, 9) Then
305:                sCode = vbNullString
306:                Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 9))
307:                sCode = GetCodeFromModule(objVBCitem)
308:                If VBA.InStr(1, sCode, arrData(i, 5)) <> 0 Then
309:                    sCode = VBA.Trim$(VBA.Replace(sCode, arrData(i, 5), arrData(i, 8)))
310:                End If
311:                NameOldMOdule = arrData(i, 9)
312:            Else
313:                If VBA.InStr(1, sCode, arrData(i, 5)) <> 0 Then
314:                    sCode = VBA.Trim$(VBA.Replace(sCode, arrData(i, 5), arrData(i, 8)))
315:                End If
316:            End If
317:            If i = UBound(arrData) Then
318:                Call SetCodeInModule(objVBCitem, sCode)
319:                Set objVBCitem = Nothing
320:            Else
321:                If arrData(i + 1, 9) <> arrData(i, 9) Then
322:                    Call SetCodeInModule(objVBCitem, sCode)
323:                    Set objVBCitem = Nothing
324:                End If
325:            End If
326:
327:        End If
328:        DoEvents
329:        If i Mod 100 = 0 Then Application.StatusBar = "String encryption - completed:" & Format(i / UBound(arrData), "Percent") & ", " & i & "from" & UBound(arrData)
330:    Next i
331:    Application.StatusBar = False
332:    Dim sName       As String
333:    sName = ActiveWorkbook.Worksheets(NAME_SH_STR).Cells(2, 11).Value
334:    If sName <> vbNullString Then
335:        Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
336:        objVBCitem.Name = sName
337:        Call SetCodeInModule(objVBCitem, sCodeString)
338:    End If
339: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetCodeFromModule - получить код из модуля в строковую переменную
'* Created    : 20-04-2020 18:20
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objVBComp As VBIDE.VBComponent : модуль VBA
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function GetCodeFromModule(ByRef objVBComp As VBIDE.VBComponent) As String
352:    GetCodeFromModule = vbNullString
353:    With objVBComp.CodeModule
354:        If .CountOfLines > 0 Then
355:            GetCodeFromModule = .Lines(1, .CountOfLines)
356:        End If
357:    End With
358: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SetCodeInModule загрузить код из строковой переменой в модуль
'* Created    : 20-04-2020 18:21
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objVBComp As VBIDE.VBComponent : модуль VBA
'* ByVal SCode As String                : строковая переменная
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub SetCodeInModule(ByRef objVBComp As VBIDE.VBComponent, ByVal sCode As String)
373:    With objVBComp.CodeModule
374:        If .CountOfLines > 0 Then
375:            'Debug.Print .CountOfLines
376:            Call .DeleteLines(1, .CountOfLines)
377:            'Debug.Print sCode
378:            Call .InsertLines(1, VBA.Trim$(sCode))
379:        End If
380:    End With
381: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SortTabel - сортировка диапазона данных
'* Created    : 29-07-2020 10:03
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByRef WS As Worksheet       :
'* ByVal sRng As String        :
'* sKey1 As String             :
'* Optional bOrder As Byte = 2 :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub SortTabel(ByRef WS As Worksheet, ByVal sRng As String, sKey1 As String, Optional bOrder As Byte = 2)
398:    With WS
399:        On Error GoTo errMsg
400:        .Activate
401:        .Range(sRng).AutoFilter
Repeatnext:
403:        .AutoFilter.Sort.SortFields.Clear
404:        .AutoFilter.Sort.SortFields.Add Key:=Range(sKey1), SortOn:=xlSortOnValues, Order:=bOrder, DataOption:=xlSortNormal
405:        With .AutoFilter.Sort
406:            .Header = xlYes
407:            .MatchCase = False
408:            .Orientation = xlTopToBottom
409:            .SortMethod = xlPinYin
410:            .Apply
411:        End With
412:    End With
413:    Exit Sub
errMsg:
415:    If Err.Number = 91 Then
416:        WS.Range(sRng).AutoFilter
417:        Err.Clear
418:        GoTo Repeatnext
419:    Else
420:        Call MsgBox(Err.Description, vbCritical, "Mistake:")
421:    End If
End Sub
