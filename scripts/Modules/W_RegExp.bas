Attribute VB_Name = "W_RegExp"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : modRegExp - тестирование регулярных выражений
'* Created    : 22-04-2020 23:02
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RegExpStart запуск регулярного выражения
'* Created    : 23-04-2020 00:03
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub RegExpStart()
18:    Dim sSTR        As String
19:    Dim sPattern    As String
20:    Dim sReplace    As String
21:    Dim sMsgErr     As String
22:    Dim bGloba1     As Boolean
23:    Dim bIgnoreCase As Boolean
24:    Dim bMultiLine  As Boolean
25:
26:    Application.ScreenUpdating = False
27:    With ActiveSheet
28:        sSTR = VBA.Trim$(.Cells(11, 3).Value)
29:        sPattern = VBA.Trim$(.Cells(2, 3).Value)
30:        sReplace = VBA.Trim$(.Cells(24, 3).Value)
31:        bGloba1 = VBA.CBool(.Cells(7, 3).Value)
32:        bIgnoreCase = VBA.CBool(.Cells(8, 3).Value)
33:        bMultiLine = VBA.CBool(.Cells(9, 3).Value)
34:    End With
35:
36:    If sPattern = vbNullString Then sMsgErr = "Не указано регулярное выражение!" & vbNewLine
37:    If sSTR = vbNullString Then sMsgErr = sMsgErr & "Не указан исходный текст!"
38:
39:    Call RegExpClearCells
40:    If sMsgErr <> vbNullString Then
41:        Call MsgBox(sMsgErr, vbCritical, "Поиск совпадений:")
42:        Exit Sub
43:    End If
44:    'сброс форматирования
45:    With ActiveSheet.Cells(11, 3).Font
46:        .ColorIndex = xlAutomatic
47:        .Underline = xlUnderlineStyleNone
48:    End With
49:    With ActiveSheet.Cells(26, 3).Font
50:        .ColorIndex = xlAutomatic
51:        .Underline = xlUnderlineStyleNone
52:    End With
53:
54:    Call RegExpEnjoyReplace(sSTR, sPattern, sReplace, bGloba1, bIgnoreCase, bMultiLine)
55:    Call RegExpGetMatches(sSTR, sPattern, sReplace, bGloba1, bIgnoreCase, bMultiLine)
56:    Application.ScreenUpdating = True
57: End Sub

     Private Sub RegExpGetMatches(ByVal sSTR As String, ByVal sPattern As String, ByVal sReplace As String, Optional bGloba1 As Boolean = True, Optional bIgnoreCase As Boolean = False, Optional bMultiLine As Boolean = False)
60:
61:    Dim objMatches  As Object
62:    Dim itemMatch   As Object
63:    Dim lRow        As Long
64:    Dim iFerstChr   As Integer
65:    Dim i           As Integer
66:
67:    lRow = 2
68:    i = 1
69:    iFerstChr = 0
70:
71:    With ActiveSheet
72:        Set objMatches = RegExpExecuteCollection(sSTR, sPattern, bGloba1, bIgnoreCase, bMultiLine)
73:        If objMatches Is Nothing Then
74:            Call MsgBox("Совпадений не найдено!", vbInformation + vbOKOnly, "Поиск совпадений:")
75:            .Range("M:P").EntireColumn.AutoFit
76:        ElseIf objMatches.Count = 0 Then
77:            Call MsgBox("Совпадений не найдено!", vbInformation + vbOKOnly, "Поиск совпадений:")
78:            .Range("M:P").EntireColumn.AutoFit
79:        Else
80:            For Each itemMatch In objMatches
81:                With itemMatch
82:                    ActiveSheet.Cells(lRow, 13).Value = lRow - 1
83:                    ActiveSheet.Cells(lRow, 14).Value = .FirstIndex
84:                    ActiveSheet.Cells(lRow, 15).Value = .Length
85:                    ActiveSheet.Cells(lRow, 16).Value = .Value
86:                End With
87:
88:                With ActiveSheet.Cells(11, 3).Characters(Start:=itemMatch.FirstIndex + 1, Length:=itemMatch.Length).Font
89:                    .Color = -16776961
90:                    .Underline = xlUnderlineStyleSingle
91:                End With
92:
93:                sReplace = RegExpFindReplace(sReplace, "\$[1-9]{1}", vbNullString, True, False, True)
94:                iFerstChr = VBA.InStr(iFerstChr + 1, ActiveSheet.Cells(26, 3).Value, sReplace)
95:                If iFerstChr > 0 And sReplace <> vbNullString Then
96:                    With ActiveSheet.Cells(26, 3).Characters(Start:=iFerstChr, Length:=VBA.Len(sReplace)).Font
97:                        .Color = -16776961
98:                        .Underline = xlUnderlineStyleSingle
99:                    End With
100:                End If
101:                lRow = lRow + 1
102:            Next itemMatch
103:            .Range("M:P").EntireColumn.AutoFit
104:        End If
105:    End With
106:    'Очистка объектного мусора
107:    Set itemMatch = Nothing
108:    Set objMatches = Nothing
109: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RegExpEnjoyReplace - запуск замены
'* Created    : 22-04-2020 23:24
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub RegExpEnjoyReplace(ByVal sSTR As String, ByVal sPattern As String, ByVal sReplace As String, Optional bGloba1 As Boolean = True, Optional bIgnoreCase As Boolean = False, Optional bMultiLine As Boolean = False)
119:    With ActiveSheet
120:        .Cells(26, 3).Value = RegExpFindReplace(sSTR, sPattern, sReplace, bGloba1, bIgnoreCase, bMultiLine)
121:    End With
122: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : RegExpFindReplace
'* Created    : 22-04-2020 23:07
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* sStr As String                          : исходная строка
'* sPattern As String                      : паттерн поиска
'* sReplace As String                      : строка для замены
'* Optional bGloba1 As Boolean = True      : ЛОЖЬ - проверять до первого соответствия, ИСТИНА- проверять по всему тексту
'* Optional bIgnoreCase As Boolean = False : ЛОЖЬ - учитывать регистр символов, ИСТИНА - игнорировать регистр символов
'* Optional bMultiline As Boolean = False  : ЛОЖЬ - однострочный объект, ИСТИНА - многострочный
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function RegExpFindReplace(ByVal sSTR As String, ByVal sPattern As String, ByVal sReplace As String, Optional bGloba1 As Boolean = True, Optional bIgnoreCase As Boolean = False, Optional bMultiLine As Boolean = False) As String
140:    RegExpFindReplace = sSTR
141:    If Not sPattern Like vbNullString Then
142:        Dim RegExp  As New RegExp
143:        With RegExp
144:            .Global = bGloba1
145:            .IgnoreCase = bIgnoreCase
146:            .Multiline = bMultiLine
147:            .Pattern = sPattern
148:        End With
149:
150:        On Error Resume Next
151:        RegExpFindReplace = RegExp.Replace(sSTR, sReplace)
152:        Set RegExp = Nothing
153:    End If
154: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : RegExpExecuteCollection   - колекция найденого
'* Created    : 22-04-2020 22:59
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* sStr As String                          : исходная строка
'* Pattern As String                       : паттерн поиска
'* Optional bGloba1 As Boolean = True      : ЛОЖЬ - проверять до первого соответствия, ИСТИНА- проверять по всему тексту
'* Optional bIgnoreCase As Boolean = False : ЛОЖЬ - учитывать регистр символов, ИСТИНА - игнорировать регистр символов
'* Optional bMultiline As Boolean = False  : ЛОЖЬ - однострочный объект, ИСТИНА - многострочный
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function RegExpExecuteCollection(ByVal sSTR As String, ByVal sPattern As String, Optional bGloba1 As Boolean = True, Optional bIgnoreCase As Boolean = False, Optional bMultiLine As Boolean = False) As Object
171:    Set RegExpExecuteCollection = Nothing
172:    If Not sPattern Like vbNullString Then
173:        Dim RegExp  As New RegExp
174:        With RegExp
175:            .Global = bGloba1
176:            .IgnoreCase = bIgnoreCase
177:            .Multiline = bMultiLine
178:            .Pattern = sPattern
179:        End With
180:
181:        On Error Resume Next
182:        'Получаем коллекцию совпадений
183:        Set RegExpExecuteCollection = RegExp.Execute(sSTR)
184:        Set RegExp = Nothing
185:    End If
186: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RegExpClearCells - процедура по очистке формы перед запуском
'* Created    : 22-04-2020 23:11
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub RegExpClearCellsAll()
196:    With ActiveSheet
197:        .Range("C24:K24").ClearContents
198:    End With
199:    Call RegExpClearCells
200:    Call RegExpClearCellsPattern
201:    Call RegExpClearCellsStr
202: End Sub
     Private Sub RegExpClearCells()
204:    With ActiveSheet
205:        .Range("C26:K37").ClearContents
206:        .Range("M2:P" & .Cells(Rows.Count, 13).End(xlUp).Row + 1).ClearContents
207:    End With
208: End Sub
     Public Sub RegExpClearCellsPattern()
210:    With ActiveSheet
211:        .Range("C2:K3").ClearContents
212:    End With
213: End Sub
     Public Sub RegExpClearCellsStr()
215:    With ActiveSheet
216:        .Range("C11:K22").ClearContents
217:    End With
218: End Sub
     Public Sub ShowTempleteManedger()
220:    Call RegExpTemplateManager.Show
221: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddSheetTestRegExp - создание листа тестирования шаблона
'* Created    : 25-04-2020 21:27
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub AddSheetTestRegExp()
231:    Const sSHNAMETEST As String = "TestRegExpVBATools"
232:    'создание листа в активной книге
233:    Application.DisplayAlerts = False
234:    On Error Resume Next
235:    ActiveWorkbook.Worksheets(sSHNAMETEST).Delete
236:    On Error GoTo 0
237:    Application.DisplayAlerts = True
238:    ThisWorkbook.Sheets(sSHNAMETEST).Copy After:=ActiveWorkbook.ActiveSheet
239:    With ActiveWorkbook.Sheets(sSHNAMETEST)
240:        .visible = True
241:        .Activate
242:    End With
243: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : РЕГВЫР_ПОЛУЧЗНАЧПОНОМЕРУ - получение значения по номеру
'* Created    : 25-04-2020 18:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByVal Текст As String                  : исходный текст
'* ByVal Паттерн As String                : паттерн рег. выражения
'* Optional НомерИскомого As Integer = 0  : номер искомого значения если 0 то ищем все через разделитель
'* Optional Разделитель As String = " "   : разделитель если ищем все
'* Optional ИскатьВсе As Boolean = True   : ЛОЖЬ - проверять до первого соответствия, ИСТИНА- проверять по всему тексту
'* Optional Регист As Boolean = False     : ЛОЖЬ - учитывать регистр символов, ИСТИНА - игнорировать регистр символов
'* Optional МногоСтроч As Boolean = False : ЛОЖЬ - однострочный объект, ИСТИНА - многострочный
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function РЕГВЫР_ПОЛУЧЗНАЧПОНОМЕРУ(ByVal Текст As String, ByVal Паттерн As String, Optional НомерИскомого As Integer = 0, Optional Разделитель As String = " ", Optional ИскатьВсе As Boolean = True, Optional Регист As Boolean = False, Optional МногоСтроч As Boolean = False) As Variant
Attribute РЕГВЫР_ПОЛУЧЗНАЧПОНОМЕРУ.VB_Description = "Получает значение соответствующее рег. выражению по его порядковому номеру в тексте"
Attribute РЕГВЫР_ПОЛУЧЗНАЧПОНОМЕРУ.VB_ProcData.VB_Invoke_Func = " \n7"
263:    Dim ObjColl     As MatchCollection
264:    Dim sTxt        As String
265:    Dim i           As Integer
266:    Set ObjColl = RegExpExecuteCollection(Текст, Паттерн, ИскатьВсе, Регист, МногоСтроч)
267:    With ObjColl
268:        If .Count > 0 Then
269:            If НомерИскомого > 0 Then
270:                sTxt = .Item(НомерИскомого - 1)
271:            Else
272:                For i = 0 To .Count - 1
273:                    sTxt = sTxt & Разделитель & .Item(i)
274:                Next i
275:                sTxt = VBA.Right$(sTxt, VBA.Len(sTxt) - VBA.Len(Разделитель))
276:            End If
277:        End If
278:    End With
279:    РЕГВЫР_ПОЛУЧЗНАЧПОНОМЕРУ = sTxt
280: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : РЕГВЫР_СЧЁТ - счет количества значений удовлетворяющих паттерну
'* Created    : 25-04-2020 19:00
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByVal Текст As String                  : исходный текст
'* ByVal Паттерн As String                : паттерн рег. выражения
'* Optional ИскатьВсе As Boolean = True   : ЛОЖЬ - проверять до первого соответствия, ИСТИНА- проверять по всему тексту
'* Optional Регист As Boolean = False     : ЛОЖЬ - учитывать регистр символов, ИСТИНА - игнорировать регистр символов
'* Optional МногоСтроч As Boolean = False : ЛОЖЬ - однострочный объект, ИСТИНА - многострочный
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function РЕГВЫР_СЧЁТ(ByVal Текст As String, ByVal Паттерн As String, Optional ИскатьВсе As Boolean = True, Optional Регист As Boolean = False, Optional МногоСтроч As Boolean = False) As LongPtr
Attribute РЕГВЫР_СЧЁТ.VB_Description = "Подсчитывает количества значений удовлетворяющих паттерну"
Attribute РЕГВЫР_СЧЁТ.VB_ProcData.VB_Invoke_Func = " \n7"
298:    Dim ObjColl     As MatchCollection
299:    Dim lCount      As Long
300:    Dim i           As Integer
301:    Set ObjColl = RegExpExecuteCollection(Текст, Паттерн, ИскатьВсе, Регист, МногоСтроч)
302:    With ObjColl
303:        If .Count > 0 Then
304:            For i = 0 To .Count - 1
305:                lCount = lCount + 1
306:            Next i
307:        End If
308:    End With
309:    РЕГВЫР_СЧЁТ = lCount
310: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : РЕГВЫР_ТЕСТ - проверяет существует ли значения удовлетворяющие паттерну
'* Created    : 25-04-2020 19:01
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByVal Текст As String                  : исходный текст
'* ByVal Паттерн As String                : паттерн рег. выражения
'* Optional ИскатьВсе As Boolean = True   : ЛОЖЬ - проверять до первого соответствия, ИСТИНА- проверять по всему тексту
'* Optional Регист As Boolean = False     : ЛОЖЬ - учитывать регистр символов, ИСТИНА - игнорировать регистр символов
'* Optional МногоСтроч As Boolean = False : ЛОЖЬ - однострочный объект, ИСТИНА - многострочный
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function РЕГВЫР_ТЕСТ(ByVal Текст As String, ByVal Паттерн As String, Optional ИскатьВсе As Boolean = True, Optional Регист As Boolean = False, Optional МногоСтроч As Boolean = False) As Boolean
328:    If Not Текст Like vbNullString And Not Паттерн Like vbNullString Then
329:        Dim RegExp  As New RegExp
330:        With RegExp
331:            .Global = ИскатьВсе
332:            .IgnoreCase = Регист
333:            .Multiline = МногоСтроч
334:            .Pattern = Паттерн
335:        End With
336:
337:        On Error Resume Next
338:        РЕГВЫР_ТЕСТ = RegExp.Test(Текст)
339:        Set RegExp = Nothing
340:    End If
341: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : РЕГВЫР_ЗАМЕНИТЬ - заменить значения удовлетворяющие рег. выражению
'* Created    : 25-04-2020 19:02
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByVal Текст As String                  : исходный текст
'* ByVal Паттерн As String                : паттерн рег. выражения
'* ByVal Заменить_на As String            : текст на который заменяется найденые рег. выражения
'* Optional ИскатьВсе As Boolean = True   : ЛОЖЬ - проверять до первого соответствия, ИСТИНА- проверять по всему тексту
'* Optional Регист As Boolean = False     : ЛОЖЬ - учитывать регистр символов, ИСТИНА - игнорировать регистр символов
'* Optional МногоСтроч As Boolean = False : ЛОЖЬ - однострочный объект, ИСТИНА - многострочный
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function РЕГВЫР_ЗАМЕНИТЬ(ByVal Текст As String, ByVal Паттерн As String, ByVal Заменить_на As String, Optional ИскатьВсе As Boolean = True, Optional Регист As Boolean = False, Optional МногоСтроч As Boolean = False) As String
Attribute РЕГВЫР_ЗАМЕНИТЬ.VB_Description = "Заменяет значения удовлетворяющие рег. выражению в тексте на текст замены"
Attribute РЕГВЫР_ЗАМЕНИТЬ.VB_ProcData.VB_Invoke_Func = " \n7"
360:    РЕГВЫР_ЗАМЕНИТЬ = RegExpFindReplace(Текст, Паттерн, Заменить_на, ИскатьВсе, Регист, МногоСтроч)
End Function
