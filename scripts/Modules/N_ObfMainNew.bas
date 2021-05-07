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
14:    Dim objWB       As Workbook
15:
16:    On Error GoTo ErrStartParser
17:    Set Form = New AddStatistic
18:    With Form
19:        .Caption = "Code Obfuscation:"
20:        .lbOK.Caption = "OBFUSCATE"
21:        .chQuestion.visible = True
22:        .chQuestion.Value = True
23:        .Show
24:        sNameWB = .cmbMain.Value
25:    End With
26:    If sNameWB = vbNullString Then Exit Sub
27:    Set objWB = Workbooks(sNameWB)
28:    Call MainObfuscation(objWB, Form.chQuestion.Value)
29:    Set Form = Nothing
30:    Exit Sub
ErrStartParser:
32:    Application.Calculation = xlCalculationAutomatic
33:    Application.ScreenUpdating = True
34:    Call MsgBox("Error in N_ObfParserVBA. Start Parser" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbCritical, "Error:")
35:    Call WriteErrorLog("AddShapeStatistic")
36: End Sub

    Private Sub MainObfuscation(ByRef objWB As Workbook, Optional bEncodeStr As Boolean = False)
39:    On Error GoTo ErrStartParser
40:    If objWB.VBProject.Protection = vbext_pp_locked Then
41:        Call MsgBox("The project is protected, remove the password!", vbCritical, "Project:")
42:    Else
43:        If ActiveSheet.Name = NAME_SH Then
44:            Application.ScreenUpdating = False
45:            Application.Calculation = xlCalculationManual
46:            Application.EnableEvents = False
47:
48:            Call Obfuscation(objWB, bEncodeStr)
49:
50:            With ActiveWorkbook.Worksheets(NAME_SH)
51:                Call SortTabel(ActiveWorkbook.Worksheets(NAME_SH), .Range(.Cells(1, 1), .Cells(1, 13)).Address, "M1", 1)
52:            End With
53:
54:            Application.EnableEvents = True
55:            Application.Calculation = xlCalculationAutomatic
56:            Application.ScreenUpdating = True
57:            Call MsgBox("The book code [" & objWB.Name & "] is encrypted!", vbInformation, "Code Encryption:")
58:        Else
59:            Call MsgBox("Create or navigate to a sheet: [" & NAME_SH & "]", vbCritical, "Activating a sheet:")
60:        End If
61:    End If
62:    Exit Sub
ErrStartParser:
64:    Application.EnableEvents = True
65:    Application.Calculation = xlCalculationAutomatic
66:    Application.ScreenUpdating = True
67:    Call MsgBox("Error in Main Obfuscation" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbCritical, "Error:")
68: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : FelterAdd - фильтрация в нужном порядке перед шифрованием
'* Created    : 29-07-2020 09:58
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub FelterAdd()
78:    Dim LastRow     As Long
79:    With ActiveWorkbook.Worksheets(NAME_SH)
80:        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
81:        If LastRow > 1 Then
82:            .Range(.Cells(2, 12), .Cells(LastRow, 12)).FormulaR1C1 = "=LEN(RC[-4])"
83:            .Range(.Cells(2, 13), .Cells(LastRow, 13)).FormulaR1C1 = "=R[-1]C+1"
84:            .Range(.Cells(2, 13), .Cells(LastRow, 13)).Value = .Range(.Cells(2, 13), .Cells(LastRow, 13)).Value
85:            Call SortTabel(ActiveWorkbook.Worksheets(NAME_SH), .Range(.Cells(1, 1), .Cells(1, 13)).Address, "L1", 2)
86:        End If
87:    End With
88: End Sub

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
     Private Sub Obfuscation(ByRef objWB As Workbook, Optional bEncodeStr As Boolean = True)
103:    Dim arrData     As Variant
104:    Dim i           As Long
105:    Dim j           As Long
106:    Dim sKey        As String
107:    Dim sCode       As String
108:    Dim sFinde      As String
109:    Dim sReplace    As String
110:    Dim sPattern    As String
111:
112:    Dim objDictName As Scripting.Dictionary
113:    Dim objDictModule As Scripting.Dictionary
114:    Dim objDictModuleOld As Scripting.Dictionary
115:    Dim objVBCitem  As VBIDE.VBComponent
116:    Dim dTime       As Date
117:
118:    dTime = Now()
119:    Debug.Print "Start:" & VBA.Format$(Now() - dTime, "Long Time")
120:
121:    Set objDictName = New Scripting.Dictionary
122:    Set objDictModule = New Scripting.Dictionary
123:    Set objDictModuleOld = New Scripting.Dictionary
124:
125:    'сохранение и загрузка
126:    objWB.SaveAs Filename:=objWB.Path & Application.PathSeparator & C_PublicFunctions.sGetBaseName(objWB.FullName) & "_obf_" & Replace(Now(), ":", ".") & "." & C_PublicFunctions.sGetExtensionName(objWB.FullName), FileFormat:=objWB.FileFormat
127:
128:    Debug.Print "Saving the file-Completed:" & VBA.Format$(Now() - dTime, "Long Time")
129:    'фильтрация
130:    Call FelterAdd
131:
132:    'считывание данных
133:    With ActiveWorkbook.Worksheets(NAME_SH)
134:        .Activate
135:        i = .Cells(Rows.Count, 1).End(xlUp).Row
136:        arrData = .Range(Cells(2, 1), Cells(i, 10)).Value2
137:    End With
138:
139:
140:    'сбор имен с шифрами
141:    For i = LBound(arrData) To UBound(arrData)
142:        If arrData(i, 9) = "YES" Then
143:            'сбор имен с шифрами
144:            If objDictName.Exists(arrData(i, 8)) = False Then objDictName.Add arrData(i, 8), arrData(i, 10)
145:        End If
146:    Next i
147:
148:    'сбор кода из модулей
149:    For Each objVBCitem In objWB.VBProject.VBComponents
150:        If objDictModule.Exists(objVBCitem.Name) = False Then
151:            sCode = GetCodeFromModule(objVBCitem)
152:            'убираю перенос строк
153:            sCode = VBA.Replace(sCode, " _" & vbNewLine, "ЪЪЪЪЪ")
154:            objDictModule.Add objVBCitem.Name, sCode
155:            objDictModuleOld.Add objVBCitem.Name, sCode
156:            sCode = vbNullString
157:        End If
158:    Next objVBCitem
159:    'конец сбора
160:
161:    Debug.Print "Data collection-completed:" & VBA.Format$(Now() - dTime, "Long Time")
162:
163:    'шифрование
164:    sCode = vbNullString
165:    With objDictName
166:        For i = 0 To .Count - 1
167:            For j = 0 To objDictModule.Count - 1
168:                sFinde = .Keys(i)
169:                sReplace = .Items(i)
170:                sKey = objDictModule.Keys(j)
171:                sCode = objDictModule.Item(sKey)
172:                If sCode Like "*" & sFinde & "*" And VBA.Len(sFinde) > 1 Then
173:                    sPattern = "([\*\.\^\*\+\#\(\)\-\=\/\,\:\;\s\" & VBA.Chr$(34) & "])" & sFinde & "([\*\.\^\*\+\!\@\#\$\%\&\(\)\-\=\/\,\:\;\s\" & VBA.Chr$(34) & "]|$)"
174:                    sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
175:                    If sCode <> vbNullString Then objDictModule.Item(sKey) = sCode
176:                End If
177:                'регулярка для событий в основном для форм
178:                If sCode Like "* " & Chr$(83) & "ub *" & sFinde & "_*(*)*" Then
179:                    sPattern = "([\s])(Sub)([\s])" & sFinde & "(\_{1}[A-Za-zА-Яа-яЁё]{4,40}\([A-Za-zА-Яа-яЁё\s\.\,]{0,100}\))"
180:                    sCode = RegExpFindReplace(sCode, sPattern, "$1$2$3" & sReplace & "$4", True, False, False)
181:                    If sCode <> vbNullString Then objDictModule.Item(sKey) = sCode
182:                    sPattern = "([\s])" & sFinde & "(\_{1}[A-Za-zА-Яа-яЁё]{4,40}(?:\:\s|\n|\r))"
183:                    sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
184:                    If sCode <> vbNullString Then objDictModule.Item(sKey) = sCode
185:                End If
186:                sCode = vbNullString
187:            Next j
188:            DoEvents
189:            Application.StatusBar = "Data encryption-performed by:" & Format(i / (.Count - 1), "Percent") & ", " & i & "from" & .Count - 1
190:        Next i
191:    End With
192:    Application.StatusBar = False
193:    'конец
194:
195:    'чередование
196:    sCode = vbNullString
197:
198:    For j = 0 To objDictModule.Count - 1
199:        Dim arrNew  As Variant
200:        Dim arrOld  As Variant
201:        Dim sTemp   As String
202:        arrNew = VBA.Split(objDictModule.Items(j), vbNewLine)
203:        arrOld = VBA.Split(objDictModuleOld.Items(j), vbNewLine)
204:        For i = LBound(arrNew) To UBound(arrNew)
205:            If arrNew(i) = vbNullString Or VBA.Left$(VBA.Trim$(arrNew(i)), 1) = "'" Then
206:                sTemp = vbNullString
207:            Else
208:                sTemp = "'" & arrOld(i) & vbNewLine
209:            End If
210:            sCode = sCode & sTemp & arrNew(i) & vbNewLine
211:            sTemp = vbNullString
212:        Next i
213:        sKey = objDictModule.Keys(j)
214:        objDictModule.Item(sKey) = sCode
215:        sCode = vbNullString
216:    Next j
217:    Debug.Print "String alternation-completed:" & VBA.Format$(Now() - dTime, "Long Time")
218:
219:
220:    Debug.Print "Data encryption-performed by:" & VBA.Format$(Now() - dTime, "Long Time")
221:    'загрузка кода
222:    For j = 0 To objDictModule.Count - 1
223:        Set objVBCitem = objWB.VBProject.VBComponents(objDictModule.Keys(j))
224:        sCode = objDictModule.Items(j)
225:        'возврат перенос строк
226:        sCode = VBA.Replace(sCode, "ЪЪЪЪЪ", " _" & vbNewLine)
227:        Call SetCodeInModule(objVBCitem, sCode)
228:    Next j
229:
230:    Debug.Print "Code download-Completed:" & VBA.Format$(Now() - dTime, "Long Time")
231:
232:    'переименование контролов
233:    For i = LBound(arrData) To UBound(arrData)
234:        If arrData(i, 9) = "YES" And objDictName.Exists(arrData(i, 8)) Then
235:            If arrData(i, 1) = "Control" Then
236:                Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
237:                objVBCitem.Designer.Controls(arrData(i, 8)).Name = arrData(i, 10)
238:            End If
239:        End If
240:    Next i
241:
242:    Debug.Print "Renaming controls-completed:" & VBA.Format$(Now() - dTime, "Long Time")
243:
244:    'переименование модулей
245:    For i = LBound(arrData) To UBound(arrData)
246:        If arrData(i, 9) = "YES" And objDictName.Exists(arrData(i, 8)) Then
247:            If arrData(i, 1) = "Module" And VBA.CByte(arrData(i, 2)) <> 100 Then
248:                Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
249:                objVBCitem.Name = arrData(i, 10)
250:            End If
251:        End If
252:    Next i
253:
254:    'шифрование строк
255:    If bEncodeStr Then Call EncodedStringCode(objWB)
256:
257:    Debug.Print "Renaming of modules-completed:" & VBA.Format$(Now() - dTime, "Long Time")
258:    objWB.Save
259:
260: End Sub

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
     Private Sub EncodedStringCode(ByRef objWB As Workbook)
274:    Dim arrData     As Variant
275:    Dim i           As Long
276:    Dim sCodeString As String
277:    Dim objVBCitem  As VBIDE.VBComponent
278:    Dim sCode       As String
279:
280:    'считывание данных
281:    With ActiveWorkbook.Worksheets(NAME_SH_STR)
282:        .Activate
283:        arrData = .Range(Cells(2, 1), Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 9)).Value2
284:    End With
285:    'сбор строк
286:    sCodeString = "Option Explicit" & VBA.Chr$(13)
287:    For i = LBound(arrData) To UBound(arrData)
288:        If arrData(i, 7) = "YES" Then
289:            sCodeString = sCodeString & "Public Const " & arrData(i, 8) & " as string=" & arrData(i, 5) & VBA.Chr$(13)
290:        End If
291:    Next i
292:    Dim NameOldMOdule As String
293:    For i = LBound(arrData) To UBound(arrData)
294:        If arrData(i, 7) = "YES" Then
295:            If NameOldMOdule <> arrData(i, 9) Then
296:                sCode = vbNullString
297:                Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 9))
298:                sCode = GetCodeFromModule(objVBCitem)
299:                If VBA.InStr(1, sCode, arrData(i, 5)) <> 0 Then
300:                    sCode = VBA.Trim$(VBA.Replace(sCode, arrData(i, 5), arrData(i, 8)))
301:                End If
302:                NameOldMOdule = arrData(i, 9)
303:            Else
304:                If VBA.InStr(1, sCode, arrData(i, 5)) <> 0 Then
305:                    sCode = VBA.Trim$(VBA.Replace(sCode, arrData(i, 5), arrData(i, 8)))
306:                End If
307:            End If
308:            If i = UBound(arrData) Then
309:                Call SetCodeInModule(objVBCitem, sCode)
310:                Set objVBCitem = Nothing
311:            Else
312:                If arrData(i + 1, 9) <> arrData(i, 9) Then
313:                    Call SetCodeInModule(objVBCitem, sCode)
314:                    Set objVBCitem = Nothing
315:                End If
316:            End If
317:
318:        End If
319:        DoEvents
320:        If i Mod 100 = 0 Then Application.StatusBar = "String encryption-performed by:" & Format(i / UBound(arrData), "Percent") & ", " & i & "from" & UBound(arrData)
321:    Next i
322:    Application.StatusBar = False
323:    Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
324:    objVBCitem.Name = ActiveWorkbook.Worksheets(NAME_SH_STR).Cells(2, 11).Value
325:    Call SetCodeInModule(objVBCitem, sCodeString)
326: End Sub
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
339:    GetCodeFromModule = vbNullString
340:    With objVBComp.CodeModule
341:        If .CountOfLines > 0 Then
342:            GetCodeFromModule = .Lines(1, .CountOfLines)
343:        End If
344:    End With
345: End Function

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
360:    With objVBComp.CodeModule
361:        If .CountOfLines > 0 Then
362:            'Debug.Print .CountOfLines
363:            Call .DeleteLines(1, .CountOfLines)
364:            'Debug.Print sCode
365:            Call .InsertLines(1, VBA.Trim$(sCode))
366:        End If
367:    End With
368: End Sub

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
385:    With WS
386:        On Error GoTo errmsg
387:        .Activate
388:        .Range(sRng).AutoFilter
Repeatnext:
390:        .AutoFilter.Sort.SortFields.Clear
391:        .AutoFilter.Sort.SortFields.Add Key:=Range(sKey1), SortOn:=xlSortOnValues, Order:=bOrder, DataOption:=xlSortNormal
392:        With .AutoFilter.Sort
393:            .Header = xlYes
394:            .MatchCase = False
395:            .Orientation = xlTopToBottom
396:            .SortMethod = xlPinYin
397:            .Apply
398:        End With
399:    End With
400:    Exit Sub
errmsg:
402:    If Err.Number = 91 Then
403:        WS.Range(sRng).AutoFilter
404:        Err.Clear
405:        GoTo Repeatnext
406:    Else
407:        Call MsgBox(Err.Description, vbCritical, "Error:")
408:    End If
409: End Sub

