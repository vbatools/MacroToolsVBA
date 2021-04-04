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
5:    Dim Form        As AddStatistic
6:    Dim sNameWB     As String
7:    Dim objWB       As Workbook
8:
9:    On Error GoTo ErrStartParser
10:    Set Form = New AddStatistic
11:    With Form
12:        .Caption = "Обфусцирование кода:"
13:        .lbOK.Caption = "ОБФУЦИРОВАТЬ"
14:        .chQuestion.visible = True
15:        .chQuestion.Value = True
16:        .Show
17:        sNameWB = .cmbMain.Value
18:    End With
19:    If sNameWB = vbNullString Then Exit Sub
20:    Set objWB = Workbooks(sNameWB)
21:    Call MainObfuscation(objWB, Form.chQuestion.Value)
22:    Set Form = Nothing
23:    Exit Sub
ErrStartParser:
25:    Application.Calculation = xlCalculationAutomatic
26:    Application.ScreenUpdating = True
27:    Call MsgBox("Ошибка в N_ObfParserVBA.StartParser" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
28:    Call WriteErrorLog("AddShapeStatistic")
29: End Sub

    Private Sub MainObfuscation(ByRef objWB As Workbook, Optional bEncodeStr As Boolean = False)
32:    On Error GoTo ErrStartParser
33:    If objWB.VBProject.Protection = vbext_pp_locked Then
34:        Call MsgBox("Проект защищен, снимите пароль!", vbCritical, "Проект:")
35:    Else
36:        If ActiveSheet.Name = NAME_SH Then
37:            Application.ScreenUpdating = False
38:            Application.Calculation = xlCalculationManual
39:            Application.EnableEvents = False
40:
41:            Call Obfuscation(objWB, bEncodeStr)
42:
43:            With ActiveWorkbook.Worksheets(NAME_SH)
44:                Call SortTabel(ActiveWorkbook.Worksheets(NAME_SH), .Range(.Cells(1, 1), .Cells(1, 13)).Address, "M1", 1)
45:            End With
46:
47:            Application.EnableEvents = True
48:            Application.Calculation = xlCalculationAutomatic
49:            Application.ScreenUpdating = True
50:            Call MsgBox("Код книги [" & objWB.Name & "] зашифрован!", vbInformation, "Шифрование кода:")
51:        Else
52:            Call MsgBox("Создайте или перейдите на лист: [" & NAME_SH & "]", vbCritical, "Активация листа:")
53:        End If
54:    End If
55:    Exit Sub
ErrStartParser:
57:    Application.EnableEvents = True
58:    Application.Calculation = xlCalculationAutomatic
59:    Application.ScreenUpdating = True
60:    Call MsgBox("Ошибка в MainObfuscation" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
61: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : FelterAdd - фильтрация в нужном порядке перед шифрованием
'* Created    : 29-07-2020 09:58
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub FelterAdd()
71:    Dim LastRow     As Long
72:    With ActiveWorkbook.Worksheets(NAME_SH)
73:        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
74:        If LastRow > 1 Then
75:            .Range(.Cells(2, 12), .Cells(LastRow, 12)).FormulaR1C1 = "=LEN(RC[-4])"
76:            .Range(.Cells(2, 13), .Cells(LastRow, 13)).FormulaR1C1 = "=R[-1]C+1"
77:            .Range(.Cells(2, 13), .Cells(LastRow, 13)).Value = .Range(.Cells(2, 13), .Cells(LastRow, 13)).Value
78:            Call SortTabel(ActiveWorkbook.Worksheets(NAME_SH), .Range(.Cells(1, 1), .Cells(1, 13)).Address, "L1", 2)
79:        End If
80:    End With
81: End Sub

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
96:    Dim arrData     As Variant
97:    Dim i           As Long
98:    Dim j           As Long
99:    Dim sKey        As String
100:    Dim sCode       As String
101:    Dim sFinde      As String
102:    Dim sReplace    As String
103:    Dim sPattern    As String
104:
105:    Dim objDictName As Scripting.Dictionary
106:    Dim objDictModule As Scripting.Dictionary
107:    Dim objDictModuleOld As Scripting.Dictionary
108:    Dim objVBCitem  As VBIDE.VBComponent
109:    Dim dTime       As Date
110:
111:    dTime = Now()
112:    Debug.Print "Старт: " & VBA.Format$(Now() - dTime, "Long Time")
113:
114:    Set objDictName = New Scripting.Dictionary
115:    Set objDictModule = New Scripting.Dictionary
116:    Set objDictModuleOld = New Scripting.Dictionary
117:
118:    'сохранение и загрузка
119:    objWB.SaveAs Filename:=objWB.Path & Application.PathSeparator & C_PublicFunctions.sGetBaseName(objWB.FullName) & "_obf_" & Replace(Now(), ":", ".") & "." & C_PublicFunctions.sGetExtensionName(objWB.FullName), FileFormat:=objWB.FileFormat
120:
121:    Debug.Print "Сохранение файла - выполнено: " & VBA.Format$(Now() - dTime, "Long Time")
122:    'фильтрация
123:    Call FelterAdd
124:
125:    'считывание данных
126:    With ActiveWorkbook.Worksheets(NAME_SH)
127:        .Activate
128:        i = .Cells(Rows.Count, 1).End(xlUp).Row
129:        arrData = .Range(Cells(2, 1), Cells(i, 10)).Value2
130:    End With
131:
132:
133:    'сбор имен с шифрами
134:    For i = LBound(arrData) To UBound(arrData)
135:        If arrData(i, 9) = "ДА" Then
136:            'сбор имен с шифрами
137:            If objDictName.Exists(arrData(i, 8)) = False Then objDictName.Add arrData(i, 8), arrData(i, 10)
138:        End If
139:    Next i
140:
141:    'сбор кода из модулей
142:    For Each objVBCitem In objWB.VBProject.VBComponents
143:        If objDictModule.Exists(objVBCitem.Name) = False Then
144:            sCode = GetCodeFromModule(objVBCitem)
145:            'убираю перенос строк
146:            sCode = VBA.Replace(sCode, " _" & vbNewLine, vbNullString)
147:            objDictModule.Add objVBCitem.Name, sCode
148:            objDictModuleOld.Add objVBCitem.Name, sCode
149:            sCode = vbNullString
150:        End If
151:    Next objVBCitem
152:    'конец сбора
153:
154:    Debug.Print "Сбор данных - выполнено: " & VBA.Format$(Now() - dTime, "Long Time")
155:
156:    'шифрование
157:    sCode = vbNullString
158:    With objDictName
159:        For i = 0 To .Count - 1
160:            For j = 0 To objDictModule.Count - 1
161:                sFinde = .Keys(i)
162:                sReplace = .Items(i)
163:                sKey = objDictModule.Keys(j)
164:                sCode = objDictModule.Item(sKey)
165:                If sCode Like "*" & sFinde & "*" And VBA.Len(sFinde) > 1 Then
166:                    sPattern = "([\*\.\^\*\+\#\(\)\-\=\/\,\:\;\s\" & VBA.Chr$(34) & "])" & sFinde & "([\*\.\^\*\+\!\@\#\$\%\&\(\)\-\=\/\,\:\;\s\" & VBA.Chr$(34) & "]|$)"
167:                    sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
168:                    If sCode <> vbNullString Then objDictModule.Item(sKey) = sCode
169:                End If
170:                'регулярка для событий в основном для форм
171:                If sCode Like "* " & Chr$(83) & "ub *" & sFinde & "_*(*)*" Then
172:                    sPattern = "([\s])(Sub)([\s])" & sFinde & "(\_{1}[A-Za-zА-Яа-яЁё]{4,40}\([A-Za-zА-Яа-яЁё\s\.\,]{0,100}\))"
173:                    sCode = RegExpFindReplace(sCode, sPattern, "$1$2$3" & sReplace & "$4", True, False, False)
174:                    If sCode <> vbNullString Then objDictModule.Item(sKey) = sCode
175:                    sPattern = "([\s])" & sFinde & "(\_{1}[A-Za-zА-Яа-яЁё]{4,40}(?:\:\s|\n|\r))"
176:                    sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
177:                    If sCode <> vbNullString Then objDictModule.Item(sKey) = sCode
178:                End If
179:                sCode = vbNullString
180:            Next j
181:            DoEvents
182:            Application.StatusBar = "Шифрование данных - выполнено: " & Format(i / (.Count - 1), "Percent") & ", " & i & " из " & .Count - 1
183:        Next i
184:    End With
185:    Application.StatusBar = False
186:    'конец
187:
188:    'чередование
189:    sCode = vbNullString
190:
191:    For j = 0 To objDictModule.Count - 1
192:        Dim arrNew  As Variant
193:        Dim arrOld  As Variant
194:        Dim sTemp   As String
195:        arrNew = VBA.Split(objDictModule.Items(j), vbNewLine)
196:        arrOld = VBA.Split(objDictModuleOld.Items(j), vbNewLine)
197:        For i = LBound(arrNew) To UBound(arrNew)
198:            If arrNew(i) = vbNullString Or VBA.Left$(VBA.Trim$(arrNew(i)), 1) = "'" Then
199:                sTemp = vbNullString
200:            Else
201:                sTemp = "'" & arrOld(i) & vbNewLine
202:            End If
203:            sCode = sCode & sTemp & arrNew(i) & vbNewLine
204:            sTemp = vbNullString
205:        Next i
206:        sKey = objDictModule.Keys(j)
207:        objDictModule.Item(sKey) = sCode
208:        sCode = vbNullString
209:    Next j
210:    Debug.Print "Чередование строк - выполнено: " & VBA.Format$(Now() - dTime, "Long Time")
211:
212:
213:    Debug.Print "Шифрование данных - выполнено: " & VBA.Format$(Now() - dTime, "Long Time")
214:    'загрузка кода
215:    For j = 0 To objDictModule.Count - 1
216:        Set objVBCitem = objWB.VBProject.VBComponents(objDictModule.Keys(j))
217:        Call SetCodeInModule(objVBCitem, objDictModule.Items(j))
218:    Next j
219:
220:    Debug.Print "Загрузка кода- выполнено: " & VBA.Format$(Now() - dTime, "Long Time")
221:
222:    'переименование контролов
223:    For i = LBound(arrData) To UBound(arrData)
224:        If arrData(i, 9) = "ДА" And objDictName.Exists(arrData(i, 8)) Then
225:            If arrData(i, 1) = "Контрол" Then
226:                Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
227:                objVBCitem.Designer.Controls(arrData(i, 8)).Name = arrData(i, 10)
228:            End If
229:        End If
230:    Next i
231:
232:    Debug.Print "Переименование контролов - выполнено: " & VBA.Format$(Now() - dTime, "Long Time")
233:
234:    'переименование модулей
235:    For i = LBound(arrData) To UBound(arrData)
236:        If arrData(i, 9) = "ДА" And objDictName.Exists(arrData(i, 8)) Then
237:            If arrData(i, 1) = "Модуль" And VBA.CByte(arrData(i, 2)) <> 100 Then
238:                Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
239:                objVBCitem.Name = arrData(i, 10)
240:            End If
241:        End If
242:    Next i
243:
244:    'шифрование строк
245:    If bEncodeStr Then Call EncodedStringCode(objWB)
246:
247:    Debug.Print "Переименование модулей- выполнено: " & VBA.Format$(Now() - dTime, "Long Time")
248:    objWB.Save
249:
250: End Sub

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
264:    Dim arrData     As Variant
265:    Dim i           As Long
266:    Dim sCodeString As String
267:    Dim objVBCitem  As VBIDE.VBComponent
268:    Dim sCode       As String
269:
270:    'считывание данных
271:    With ActiveWorkbook.Worksheets(NAME_SH_STR)
272:        .Activate
273:        arrData = .Range(Cells(2, 1), Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 9)).Value2
274:    End With
275:    'сбор строк
276:    sCodeString = "Option Explicit" & VBA.Chr$(13)
277:    For i = LBound(arrData) To UBound(arrData)
278:        If arrData(i, 7) = "ДА" Then
279:            sCodeString = sCodeString & "Public Const " & arrData(i, 8) & " as string=" & arrData(i, 5) & VBA.Chr$(13)
280:        End If
281:    Next i
282:    Dim NameOldMOdule As String
283:    For i = LBound(arrData) To UBound(arrData)
284:        If arrData(i, 7) = "ДА" Then
285:            If NameOldMOdule <> arrData(i, 9) Then
286:                sCode = vbNullString
287:                Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 9))
288:                sCode = GetCodeFromModule(objVBCitem)
289:                If VBA.InStr(1, sCode, arrData(i, 5)) <> 0 Then
290:                    sCode = VBA.Trim$(VBA.Replace(sCode, arrData(i, 5), arrData(i, 8)))
291:                End If
292:                NameOldMOdule = arrData(i, 9)
293:            Else
294:                If VBA.InStr(1, sCode, arrData(i, 5)) <> 0 Then
295:                    sCode = VBA.Trim$(VBA.Replace(sCode, arrData(i, 5), arrData(i, 8)))
296:                End If
297:            End If
298:            sCode = CheckStringForLength(sCode)
299:            If i <> UBound(arrData) Then
300:                If arrData(i + 1, 9) <> arrData(i, 9) Or i = UBound(arrData) Then
301:                    Call SetCodeInModule(objVBCitem, sCode)
302:                    Set objVBCitem = Nothing
303:                End If
304:            End If
305:
306:        End If
307:        DoEvents
308:        If i Mod 100 = 0 Then Application.StatusBar = "Шифрование строк - выполнено: " & Format(i / UBound(arrData), "Percent") & ", " & i & " из " & UBound(arrData)
309:    Next i
310:    Application.StatusBar = False
311:    Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
312:    objVBCitem.Name = ActiveWorkbook.Worksheets(NAME_SH_STR).Cells(2, 11).Value
313:    Call SetCodeInModule(objVBCitem, sCodeString)
314: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : CheckStringForLength - разбиение дилинных строк кода на короткие по 200 символов
'* Created    : 29-07-2020 09:59
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sSTR As String : строка кода
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function CheckStringForLength(ByVal sSTR As String) As String
328:    Dim arrStr      As Variant
329:    Dim i           As Long
330:    Dim sCode       As String
331:    arrStr = VBA.Split(sSTR, VBA.Chr$(13))
332:    For i = LBound(arrStr) To UBound(arrStr)
333:        If VBA.Len(arrStr(i)) > 500 Then
334:            Dim j   As Integer
335:            j = VBA.InStr(200, arrStr(i), " ")
336:            arrStr(i) = VBA.Left$(arrStr(i), j) & " _" & vbNewLine & VBA.Right$(arrStr(i), VBA.Len(arrStr(i)) - j + 1)
337:        End If
338:        If arrStr(i) <> vbNullString Then sCode = sCode & VBA.Chr$(13) & arrStr(i)
339:    Next i
340:    CheckStringForLength = VBA.Trim$(sCode)
341: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : RegExpFindReplace - поиск и замена через регулярное выражение
'* Created    : 20-04-2020 18:24
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* str As String                          : исходный текст
'* Pattern As String                      : паттерн
'* Replace As String                      : на то что меняем
'* Optional Globa1 As Boolean = True      : Все совпадения или только первое?
'* Optional IgnoreCase As Boolean = False : Регистр неважен?
'* Optional Multiline As Boolean = False  : Игнорировать переносы строк?
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'Public Function RegExpFindReplace(str As String, _
 '        Pattern As String, _
 '        Replace As String, _
 '        Optional Globa1 As Boolean = True, _
 '        Optional IgnoreCase As Boolean = False, _
 '        Optional Multiline As Boolean = False) As String
'    RegExpFindReplace = vbNullString
'    'Пока ничего не меняли
'    If str <> vbNullString And Pattern <> vbNullString Then
'        Dim RegExp  As New RegExp
'        'Set RegExp = CreateObject("VBScript.RegExp")
'
'        With RegExp
'            'Все совпадения или только первое?
'            .Global = Globa1
'            'Регистр неважен?
'            .IgnoreCase = IgnoreCase
'            'Игнорировать переносы строк?
'            .Multiline = Multiline
'            .Pattern = Pattern
'
'            'Найти/заменить
'            'On Error Resume Next
'            RegExpFindReplace = str
'            If .Test(str) Then
'                RegExpFindReplace = RegExp.Replace(str, Replace)
'            End If
'            Set RegExp = Nothing    'Очистка памяти
'        End With
'    End If
'End Function

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
403:    GetCodeFromModule = vbNullString
404:    With objVBComp.CodeModule
405:        If .CountOfLines > 0 Then
406:            GetCodeFromModule = .Lines(1, .CountOfLines)
407:        End If
408:    End With
409: End Function

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
424:    With objVBComp.CodeModule
425:        If .CountOfLines > 0 Then
426:            'Debug.Print .CountOfLines
427:            Call .DeleteLines(1, .CountOfLines)
428:            'Debug.Print sCode
429:            Call .InsertLines(1, VBA.Trim$(sCode))
430:        End If
431:    End With
432: End Sub

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
449:    With WS
450:        On Error GoTo errmsg
451:        .Activate
452:        .Range(sRng).AutoFilter
Repeatnext:
454:        .AutoFilter.Sort.SortFields.Clear
455:        .AutoFilter.Sort.SortFields.Add Key:=Range(sKey1), SortOn:=xlSortOnValues, Order:=bOrder, DataOption:=xlSortNormal
456:        With .AutoFilter.Sort
457:            .Header = xlYes
458:            .MatchCase = False
459:            .Orientation = xlTopToBottom
460:            .SortMethod = xlPinYin
461:            .Apply
462:        End With
463:    End With
464:    Exit Sub
errmsg:
466:    If Err.Number = 91 Then
467:        WS.Range(sRng).AutoFilter
468:        Err.Clear
469:        GoTo Repeatnext
470:    Else
471:        Call MsgBox(Err.Description, vbCritical, "Ошибка:")
472:    End If
End Sub
