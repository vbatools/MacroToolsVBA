Attribute VB_Name = "N_ObfParserVBA"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : N_ObfParserVBA - парсер кода VBA
'* Created    : 08-10-2020 14:12
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module

Private objCollUnical As New Collection
Private Const CHR_TO As String = "|ЪЪ|"

Private Type obfModule
    objName         As Scripting.Dictionary
    objNameGlobVar  As Scripting.Dictionary
    objContr        As Scripting.Dictionary
    objSubFun       As Scripting.Dictionary
    objDimVar       As Scripting.Dictionary
    objTypeEnum     As Scripting.Dictionary
    objAPI          As Scripting.Dictionary
    objStringCode   As Scripting.Dictionary
End Type

    Public Sub StartParser()
27:    Dim Form        As AddStatistic
28:    Dim sNameWB     As String
29:    Dim objWB       As Object
30:
31:    On Error GoTo ErrStartParser
32:    Application.Calculation = xlCalculationManual
33:    Set Form = New AddStatistic
34:    With Form
35:        .Caption = "Code base data collection:"
36:        .lbOK.Caption = "Parse code"
37:        .chQuestion.visible = True
38:        .chQuestion.Value = True
39:        .chQuestion.Caption = "Collect string values?"
40:        .lbWord.Caption = 1
41:        .Show
42:        sNameWB = .cmbMain.Value
43:    End With
44:    If sNameWB = vbNullString Then Exit Sub
45:    If sNameWB Like "*.docm" Or sNameWB Like "*.DOCM" Then
46:        Dim objWrdApp As Object
47:        Set objWrdApp = GetObject(, "Word.Application")
48:        Set objWB = objWrdApp.Documents(sNameWB)
49:    Else
50:        Set objWB = Workbooks(sNameWB)
51:    End If
52:
53:    Call MainObfParser(objWB, Form.chQuestion.Value)
54:    Set Form = Nothing
55:    Application.Calculation = xlCalculationAutomatic
56:    Exit Sub
ErrStartParser:
58:    Application.Calculation = xlCalculationAutomatic
59:    Application.ScreenUpdating = True
60:    Call MsgBox("Error in N_ObfParserVBA.StartParser" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
61:    Call WriteErrorLog("AddShapeStatistic")
62: End Sub

    Private Sub MainObfParser(ByRef objWB As Object, Optional bEncodeStr As Boolean = False)
65:    If objWB.VBProject.Protection = vbext_pp_locked Then
66:        Call MsgBox("The project is protected, remove the password!", vbCritical, "The project is protected:")
67:    Else
68:        Call ParserProjectVBA(objWB, bEncodeStr)
69:        Call MsgBox("Book code [" & objWB.Name & "] assembled!", vbInformation, "Code parsing:")
70:    End If
71: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserProjectVBA - основной парсер кода, собирает названия модулей и присваивает им шифры
'* Created    : 27-03-2020 13:21
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef objWB As Workbook : выбранная книга
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ParserProjectVBA(ByRef objWB As Object, Optional bEncodeStr As Boolean = False)
85:    Dim objVBComp   As VBIDE.VBComponent
86:    Dim varModule   As obfModule
87:    Dim i           As Long
88:    Dim k           As Long
89:    Dim objDict     As Scripting.Dictionary
90:
91:
92:    With varModule
93:        'главный парсер
94:        Set .objName = AddNewDictionary(.objName)
95:        Set .objDimVar = AddNewDictionary(.objDimVar)
96:        Set .objSubFun = AddNewDictionary(.objSubFun)
97:        Set .objContr = AddNewDictionary(.objContr)
98:        Set .objTypeEnum = AddNewDictionary(.objTypeEnum)
99:        Set .objNameGlobVar = AddNewDictionary(.objNameGlobVar)
100:        Set .objStringCode = AddNewDictionary(.objStringCode)
101:        Set .objAPI = AddNewDictionary(.objAPI)
102:
103:        For Each objVBComp In objWB.VBProject.VBComponents
104:            'собираю названия модулей
105:            Dim sKey As String
106:            sKey = objVBComp.Type & CHR_TO & objVBComp.Name
107:            If Not .objName.Exists(sKey) Then .objName.Add sKey, 0
108:            'собираю названия контролов форм
109:            Call ParserNameControlsForm(objVBComp.Name, objVBComp, .objContr)
110:            'собираю названия процедур и функций
111:            Call ParserNameSubFunc(objVBComp.Name, objVBComp, .objSubFun)
112:            'собираю названия глобальные переменые
113:            Call ParserNameGlobalVariable(objVBComp.Name, objVBComp, .objNameGlobVar, .objTypeEnum, .objAPI)
114:            'собираю переменные процедур и функций, строковые переменые
115:            Call ParserVariebleSubFunc(objVBComp, .objDimVar, .objStringCode)
116:        Next objVBComp
117:        'конец парсера
118:    End With
119:
120:    'создание листа в активной книге
121:    Call AddShhetInWBook(NAME_SH, ActiveWorkbook)
122:
123:    ReDim arrRange(1 To varModule.objName.Count + varModule.objNameGlobVar.Count + varModule.objSubFun.Count + varModule.objContr.Count + varModule.objDimVar.Count + varModule.objTypeEnum.Count + varModule.objAPI.Count, 1 To 10) As String
124:
125:    Set objDict = New Scripting.Dictionary
126:
127:
128:    For i = 1 To varModule.objName.Count
129:        arrRange(i, 1) = "Module"
130:        arrRange(i, 2) = VBA.Split(varModule.objName.Keys(i - 1), CHR_TO)(0)
131:        arrRange(i, 3) = VBA.Split(varModule.objName.Keys(i - 1), CHR_TO)(1)
132:        arrRange(i, 4) = "Public"
133:        arrRange(i, 8) = arrRange(i, 3)
134:        arrRange(i, 9) = "yes"
135:
136:        If objDict.Exists(arrRange(i, 8)) = False Then
137:            objDict.Add arrRange(i, 8), AddEncodeName()
138:        End If
139:        arrRange(i, 10) = objDict.Item(arrRange(i, 8))
140:    Next i
141:    k = i
142:    Application.StatusBar = "Data collection: Module names, completed:" & VBA.Format(1 / 7, "Percent")
143:    For i = 1 To varModule.objNameGlobVar.Count
144:        arrRange(k, 1) = "Global variable"
145:        arrRange(k, 2) = varModule.objNameGlobVar.Items(i - 1)
146:        arrRange(k, 3) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(0)
147:        arrRange(k, 4) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(1)
148:        arrRange(k, 6) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(2)
149:        arrRange(k, 7) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(3)
150:        arrRange(k, 8) = arrRange(k, 7)
151:        arrRange(k, 9) = "yes"
152:
153:        If objDict.Exists(arrRange(k, 8)) = False Then
154:            objDict.Add arrRange(k, 8), AddEncodeName()
155:        End If
156:        arrRange(k, 10) = objDict.Item(arrRange(k, 8))
157:        k = k + 1
158:    Next i
159:
160:    Application.StatusBar = "Data collection: Global variables, completed:" & VBA.Format(2 / 7, "Percent")
161:    For i = 1 To varModule.objSubFun.Count
162:        arrRange(k, 1) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(1)
163:        arrRange(k, 2) = varModule.objSubFun.Items(i - 1)
164:        arrRange(k, 3) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(0)
165:        arrRange(k, 4) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(2)
166:        arrRange(k, 5) = arrRange(k, 1)
167:        arrRange(k, 6) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(3)
168:        arrRange(k, 8) = arrRange(k, 6)
169:        arrRange(k, 9) = "yes"
170:
171:        If objDict.Exists(arrRange(k, 8)) = False Then
172:            objDict.Add arrRange(k, 8), AddEncodeName()
173:        End If
174:        arrRange(k, 10) = objDict.Item(arrRange(k, 8))
175:        k = k + 1
176:    Next i
177:
178:    Application.StatusBar = "Data collection: Procedure names, completed:" & VBA.Format(3 / 7, "Percent")
179:    For i = 1 To varModule.objContr.Count
180:        arrRange(k, 1) = "Control"
181:        arrRange(k, 2) = varModule.objContr.Items(i - 1)
182:        arrRange(k, 3) = VBA.Split(varModule.objContr.Keys(i - 1), CHR_TO)(0)
183:        arrRange(k, 4) = "Private"
184:        arrRange(k, 6) = VBA.Split(varModule.objContr.Keys(i - 1), CHR_TO)(1)
185:        arrRange(k, 8) = arrRange(k, 6)
186:        arrRange(k, 9) = "yes"
187:
188:        If objDict.Exists(arrRange(k, 8)) = False Then
189:            objDict.Add arrRange(k, 8), AddEncodeName()
190:        End If
191:        arrRange(k, 10) = objDict.Item(arrRange(k, 8))
192:        k = k + 1
193:    Next i
194:
195:    Application.StatusBar = "Data collection: Names of controls, completed:" & VBA.Format(4 / 7, "Percent")
196:    For i = 1 To varModule.objDimVar.Count
197:        arrRange(k, 1) = "Variable"
198:        arrRange(k, 2) = varModule.objDimVar.Items(i - 1)
199:        arrRange(k, 3) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(0)
200:        arrRange(k, 4) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(3)
201:        arrRange(k, 5) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(1)
202:        arrRange(k, 6) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(2)
203:        arrRange(k, 7) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(4)
204:        arrRange(k, 8) = arrRange(k, 7)
205:        arrRange(k, 9) = "yes"
206:
207:        If objDict.Exists(arrRange(k, 8)) = False Then
208:            objDict.Add arrRange(k, 8), AddEncodeName()
209:        End If
210:        arrRange(k, 10) = objDict.Item(arrRange(k, 8))
211:        k = k + 1
212:        If i Mod 50 = 0 Then
213:            Application.StatusBar = "Data collection: Names of controls, completed:" & VBA.Format(i / varModule.objDimVar.Count, "Percent")
214:            DoEvents
215:        End If
216:    Next i
217:
218:    Application.StatusBar = "Data collection: Variable names, completed:" & VBA.Format(5 / 7, "Percent")
219:    For i = 1 To varModule.objTypeEnum.Count
220:        arrRange(k, 1) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(2)
221:        arrRange(k, 2) = varModule.objTypeEnum.Items(i - 1)
222:        arrRange(k, 3) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(0)
223:        arrRange(k, 4) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(1)
224:        arrRange(k, 6) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(3)
225:        arrRange(k, 8) = arrRange(k, 6)
226:        arrRange(k, 9) = "yes"
227:
228:        If objDict.Exists(arrRange(k, 8)) = False Then
229:            objDict.Add arrRange(k, 8), AddEncodeName()
230:        End If
231:        arrRange(k, 10) = objDict.Item(arrRange(k, 8))
232:        k = k + 1
233:    Next i
234:
235:    Application.StatusBar = "Data collection: Names of enumerations and types, completed:" & VBA.Format(6 / 7, "Percent")
236:    For i = 1 To varModule.objAPI.Count
237:        arrRange(k, 1) = "API"
238:        arrRange(k, 2) = varModule.objAPI.Items(i - 1)
239:        arrRange(k, 3) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(0)
240:        arrRange(k, 4) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(1)
241:        arrRange(k, 5) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(2)
242:        arrRange(k, 6) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(3)
243:        arrRange(k, 8) = arrRange(k, 6)
244:        arrRange(k, 9) = "yes"
245:
246:        If objDict.Exists(arrRange(k, 8)) = False Then
247:            objDict.Add arrRange(k, 8), AddEncodeName()
248:        End If
249:        arrRange(k, 10) = objDict.Item(arrRange(k, 8))
250:        k = k + 1
251:    Next i
252:    Application.StatusBar = "Data collection: API names, completed:" & VBA.Format(7 / 7, "Percent")
253:
254:    With ActiveSheet
255:        Application.StatusBar = "Application of formats"
256:        .Cells.ClearContents
257:        .Cells(1, 1).Value = "Type"
258:        .Cells(1, 2).Value = "Module type"
259:        .Cells(1, 3).Value = "Module name"
260:        .Cells(1, 4).Value = "Access Modifiers"
261:        .Cells(1, 5).Value = "Percentage type. and funk."
262:        .Cells(1, 6).Value = "The name of the percentage. and funk."
263:        .Cells(1, 7).Value = "Name of variables"
264:        .Cells(1, 8).Value = "Encryption Object"
265:        .Cells(1, 9).Value = "Encrypt yes/No"
266:        .Cells(1, 10).Value = "Code"
267:        .Cells(1, 11).Value = "Mistakes"
268:
269:        .Cells(2, 1).Resize(UBound(arrRange), 10) = arrRange
270:
271:        .Range(.Cells(2, 11), .Cells(k, 11)).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3]," & SHSNIPPETS.ListObjects(C_Const.TB_SERVICEWORDS).DataBodyRange.Address(ReferenceStyle:=xlR1C1, External:=True) & ",1,0),"""")"
272:        .Range(.Cells(2, 9), .Cells(k, 9)).FormulaR1C1 = "=IF(RC[2]="""",""yes"",""no"")"
273:        .Columns("A:K").AutoFilter
274:        .Columns("A:K").EntireColumn.AutoFit
275:        .Range(Cells(2, 9), Cells(UBound(arrRange) + 1, 9)).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="YES, NO"
276:        Application.StatusBar = "Application of formats, finished"
277:    End With
278:
279:    'выгрузка строковых переменых
280:    If bEncodeStr Then
281:        Call AddShhetInWBook(NAME_SH_STR, ActiveWorkbook)
282:        Application.StatusBar = "Collecting String variables"
283:        If varModule.objStringCode.Count <> 0 Then
284:            ReDim arrRange(1 To varModule.objStringCode.Count, 1 To 8) As String
285:            For i = 1 To varModule.objStringCode.Count
286:                arrRange(i, 1) = varModule.objStringCode.Items(i - 1)
287:                arrRange(i, 2) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(0)
288:                arrRange(i, 3) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(1)
289:                arrRange(i, 4) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(2)
290:                arrRange(i, 5) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(3)
291:                arrRange(i, 6) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(4)
292:                arrRange(i, 7) = "yes"
293:                arrRange(i, 8) = AddEncodeName()
294:                If i Mod 50 = 0 Then
295:                    Application.StatusBar = "Collecting String variables, completed:" & VBA.Format(i / varModule.objStringCode.Count, "Percent")
296:                    DoEvents
297:                End If
298:            Next i
299:            Application.StatusBar = "Collecting String variables, completed"
300:            With ActiveSheet
301:                .Cells(1, 1).Value = "Module type"
302:                .Cells(1, 2).Value = "Module name"
303:                .Cells(1, 3).Value = "Type Sub or Fun"
304:                .Cells(1, 4).Value = "Name Sub or Fun"
305:                .Cells(1, 5).Value = "Line"
306:                .Cells(1, 6).Value = "Array Strings"
307:                .Cells(1, 7).Value = "Encrypt yes/No"
308:                .Cells(1, 8).Value = "Code"
309:                .Cells(1, 9).Value = "Module cipher"
310:
311:                .Cells(1, 11).Value = "The cipher of the Const module"
312:                .Cells(2, 11).Value = AddEncodeName()
313:
314:                .Cells(2, 1).Resize(UBound(arrRange), 8) = arrRange
315:
316:                .Range(Cells(2, 7), Cells(UBound(arrRange) + 1, 7)).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="YES, NO"
317:                .Range(Cells(2, 9), Cells(UBound(arrRange) + 1, 9)).FormulaR1C1 = "=IF(RC1*1=100,RC2,VLOOKUP(RC2,DATA_OBF_VBATools!R2C3:R" & k & "C10,8,0))"
318:                .Columns("A:I").AutoFilter
319:                .Columns("A:D").EntireColumn.AutoFit
320:                .Columns("E").ColumnWidth = 60
321:                .Columns("F:K").EntireColumn.AutoFit
322:                .Rows("2:" & UBound(arrRange) + 1).RowHeight = 12
323:            End With
324:        End If
325:    End If
326:    ActiveWorkbook.Worksheets(NAME_SH).Activate
327:
328:    Application.StatusBar = False
329: End Sub
     Public Sub AddShhetInWBook(ByVal WSheetName As String, ByRef wb As Workbook)
331:    'создание листа в активной книге
332:    Application.DisplayAlerts = False
333:    On Error Resume Next
334:    wb.Worksheets(WSheetName).Delete
335:    On Error GoTo 0
336:    Application.DisplayAlerts = True
337:    wb.Sheets.Add Before:=ActiveSheet
338:    ActiveSheet.Name = WSheetName
339: End Sub


     Private Sub ParserVariebleSubFunc(ByRef objVBC As VBIDE.VBComponent, ByRef objDic As Scripting.Dictionary, ByRef objDicStr As Scripting.Dictionary)
343:    Dim lLine       As Long
344:    Dim sCode       As String
345:    Dim sVar        As String
346:    Dim sSubName    As String
347:    Dim sNumTypeName As String
348:    Dim sType       As String
349:    Dim arrStrCode  As Variant
350:    Dim arrEnum     As Variant
351:    Dim itemArr     As Variant
352:    Dim itemVar     As Variant
353:    Dim arrVar      As Variant
354:
355:    With objVBC.CodeModule
356:        lLine = .CountOfLines
357:        If lLine > 0 Then
358:            sCode = .Lines(1, lLine)
359:            If sCode <> vbNullString Then
360:                'убираю перенос строк
361:                sCode = VBA.Replace(sCode, " _" & vbNewLine, vbNullString)
362:                arrStrCode = VBA.Split(sCode, vbNewLine)
363:                For Each itemArr In arrStrCode
364:                    itemArr = C_PublicFunctions.TrimSpace(itemArr)
365:                    If itemArr <> vbNullString And VBA.Left$(itemArr, 1) <> "'" Then
366:                        sVar = vbNullString
367:                        'если есть коментарий в строке кода то удаляем его
368:                        itemArr = DeleteCommentString(itemArr)
369:                        'из строки декларирования и определение что вошли в процедуру
370:                        If (itemArr Like "* Sub *(*)*" Or itemArr Like "* Function *(*)*" Or itemArr Like "* Property Let *(*)*" Or itemArr Like "* Property Set *(*)*" Or itemArr Like "* Property Get *(*)*" Or _
                                    itemArr Like "Sub *(*)*" Or itemArr Like "Function *(*)*" Or itemArr Like "Property Let *(*)*" Or itemArr Like "Property Set *(*)*" Or itemArr Like "Property Get *(*)*") _
                                    And (Not itemArr Like "*As IRibbonControl*" And Not itemArr Like "* Declare *(*)*") Then
373:
374:                            sSubName = TypeProcedyre(VBA.CStr(itemArr))
375:                            sSubName = sSubName & CHR_TO & GetNameSubFromString(itemArr)
376:                            sVar = ParserStrDimConst(itemArr, sSubName, .Name)
377:
378:                        End If
379:                        'если в перечислении и типе данных
380:                        If itemArr Like "Private Enum *" Or itemArr Like "Public Enum *" Or itemArr Like "Enum *" Or itemArr Like "Private Type *" Or itemArr Like "Public Type *" Or itemArr Like "Type *" Then
381:                            arrEnum = VBA.Split(itemArr, " ")
382:                            If VBA.CStr(itemArr) Like "Private *" Then
383:                                sNumTypeName = "Private"
384:                            Else
385:                                sNumTypeName = "Public"
386:                            End If
387:                            sNumTypeName = arrEnum(UBound(arrEnum)) & CHR_TO & sNumTypeName
388:                            If itemArr Like "* Enum *" Or itemArr Like "Enum *" Then
389:                                sType = "Enum"
390:                            Else
391:                                sType = "Type"
392:                            End If
393:                        End If
394:                        'вышли из процедуры или перечисления
395:                        If itemArr Like "*End Sub" Or itemArr Like "*End Function" Or itemArr Like "*End Property" Or itemArr Like "*End Enum" Or itemArr Like "*End Type" Then
396:                            sSubName = vbNullString
397:                            sNumTypeName = vbNullString
398:                        End If
399:                        'если внутри типа или перечисления
400:                        If sNumTypeName <> vbNullString And Not itemArr Like "* Enum *" And Not itemArr Like "Enum *" And Not itemArr Like "* Type *" And Not itemArr Like "Type *" Then
401:                            arrEnum = VBA.Split(VBA.Trim$(itemArr), " ")
402:                            sVar = arrEnum(0)
403:                            If sVar Like "*(*" Then sVar = VBA.Left$(sVar, VBA.InStr(1, sVar, "(") - 1)
404:                            sVar = .Name & CHR_TO & sType & CHR_TO & sNumTypeName & CHR_TO & ReplaceType(sVar)
405:                        End If
406:                        'если находимся только внутри процедуры
407:                        If (itemArr Like "* Dim *" Or itemArr Like "* Const *" Or itemArr Like "Dim *" Or itemArr Like "Const *") And sSubName <> vbNullString Then
408:                            sVar = ParserStrDimConst(itemArr, sSubName, .Name)
409:                        End If
410:                        arrVar = VBA.Split(sVar, vbNewLine)
411:                        For Each itemVar In arrVar
412:                            If itemVar <> vbNullString And objDic.Exists(itemVar) = False Then
413:                                objDic.Add itemVar, objVBC.Type
414:                            End If
415:                        Next itemVar
416:                        Call ParserStringInCode(itemArr, sSubName, objVBC, objDicStr)
417:                    End If
418:                Next itemArr
419:            End If
420:        End If
421:    End With
422: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : GetNameSubFromString - получение названия процедуры из строки
'* Created    : 20-04-2020 18:19
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal sStrCode As String : строка
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function GetNameSubFromString(ByVal sStrCode As String) As String
436:    Dim sTemp       As String
437:    sTemp = VBA.Trim$(VBA.Left$(sStrCode, VBA.InStr(1, sStrCode, "(") - 1))
438:    Select Case True
        Case sTemp Like "*Sub *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Sub ") - 3)
440:        Case sTemp Like "*Function *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Function ") - 8)
441:        Case sTemp Like "*Property Let *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Let ") - 12)
442:        Case sTemp Like "*Property Set *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Set ") - 12)
443:        Case sTemp Like "*Property Get *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Get ") - 12)
444:    End Select
445:    GetNameSubFromString = VBA.Trim$(sTemp)
446: End Function

     Private Sub ParserStringInCode(ByVal sSTR As String, ByVal sNameSub As String, ByRef objVBC As VBIDE.VBComponent, ByRef objDicStr As Scripting.Dictionary)
449:    Dim sTxt        As String
450:    Dim arrStr      As Variant
451:    Dim Arr         As Variant
452:    Dim sReplace    As String
453:    Dim i           As Integer
454:    Dim sArray      As String
455:    Const CHAR_REPLACE As String = "B"
456:
457:    sSTR = VBA.Trim$(sSTR)
458:
459:    If sSTR Like "*" & VBA.Chr$(34) & "*" And sSTR <> vbNullString And Not sSTR Like "*Declare * Lib *(*)*" Then
460:
461:        sTxt = VBA.Right$(sSTR, VBA.Len(sSTR) - VBA.InStr(1, sSTR, VBA.Chr$(34)) + 1)
462:        sTxt = VBA.Replace(sTxt, VBA.Chr$(34) & VBA.Chr$(34), CHAR_REPLACE)
463:        arrStr = VBA.Split(sTxt, VBA.Chr$(34))
464:
465:        sArray = VBA.Left$(sSTR, VBA.InStr(1, sSTR, VBA.Chr$(34)) - 1)
466:        If sArray Like "* = Array(" Then
467:            sArray = VBA.Replace(sArray, " = Array(", vbNullString)
468:            Arr = VBA.Split(sArray, " ")
469:            sArray = Arr(UBound(Arr))
470:        Else
471:            sArray = vbNullString
472:        End If
473:        For i = 1 To UBound(arrStr) Step 2
474:            If arrStr(i) <> vbNullString Then
475:                If sNameSub = vbNullString Then sNameSub = "Declaration" & CHR_TO
476:
477:                sReplace = VBA.Replace(arrStr(i), CHAR_REPLACE, VBA.Chr$(34) & VBA.Chr$(34))
478:                sTxt = objVBC.Name & CHR_TO & sNameSub & CHR_TO & VBA.Chr$(34) & sReplace & VBA.Chr$(34) & CHR_TO & sArray    '& CHR_TO & sYesNo
479:                If arrStr(i + 1) Like "*: * = *" Then sArray = vbNullString
480:                If arrStr(i + 1) Like "*: * = Array(*" Then
481:                    sArray = VBA.Replace(arrStr(i + 1), ": ", vbNullString)
482:                    sArray = VBA.Replace(sArray, " = Array(", vbNullString)
483:                    sArray = VBA.Replace(sArray, ")", vbNullString)
484:                End If
485:                If objDicStr.Exists(sTxt) = False Then objDicStr.Add sTxt, objVBC.Type
486:            End If
487:        Next i
488:        sArray = vbNullString
489:    End If
490: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ParserStrDimConst - парсер для строк инициализации переменых и констант
'* Created    : 14-04-2020 22:45
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sTxt As String : - строка кода
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function ParserStrDimConst(ByVal sTxt As String, ByVal sNameSub As String, ByVal sNameMod As String) As String
504:    Dim sTemp       As String
505:    Dim sWord       As String
506:    Dim sWordTemp   As String
507:    Dim arrStr      As Variant
508:    Dim itemArr     As Variant
509:    Dim arrWord     As Variant
510:    Dim sType       As String
511:
512:    sTemp = C_PublicFunctions.TrimSpace(sTxt)
513:    sType = "Dim"
514:    If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
515:        'если есть коментарий в строке кода то удаляем его
516:        sTemp = DeleteCommentString(sTemp)
517:        If sTemp Like "*Sub *(*)*" Or sTemp Like "*Function *(*)*" Or sTemp Like "*Property Let *(*)*" Or sTemp Like "*Property Set *(*)*" Or sTemp Like "*Property Get *(*)*" Then
518:            If VBA.InStr(1, sTemp, ")") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, ")") - 1)
519:            If VBA.InStr(1, sTemp, " = ") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ") - 1)
520:            If VBA.Len(sTemp) - VBA.InStr(1, sTemp, "(") >= 0 Then
521:                sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "("))
522:            End If
523:        ElseIf sTemp Like "* Dim *" Or sTemp Like Chr$(68) & "im *" Then
524:            sType = "Dim"
525:            If VBA.InStr(1, sTemp, "Dim ") >= 3 Then sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Dim ") - 3)
526:        ElseIf sTemp Like "* Const *" Or sTemp Like Chr$(67) & "onst *" Then
527:            sType = "Const"
528:            If VBA.InStr(1, sTemp, "Const ") >= 5 Then sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Const ") - 5)
529:            If VBA.InStr(1, sTemp, " = ") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ") - 1)
530:        Else
531:            sTemp = vbNullString
532:        End If
533:    End If
534:
535:    If sTemp Like "*: *" Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, ": ") - 1)
536:    If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
537:        arrStr = VBA.Split(sTemp, ",")
538:        For Each itemArr In arrStr
539:            If itemArr Like "*(*" Then itemArr = VBA.Left$(itemArr, VBA.InStr(1, itemArr, "(") - 1)
540:            If Not itemArr Like "*)*" And Not itemArr Like "* To *" Then
541:                arrWord = VBA.Split(itemArr, " As ")
542:                arrWord = VBA.Split(VBA.Trim$(arrWord(0)), " ")
543:                If UBound(arrWord) = -1 Then
544:                    sWord = vbNullString
545:                Else
546:                    sWordTemp = VBA.Trim$(arrWord(UBound(arrWord)))
547:                    sWordTemp = ReplaceType(sWordTemp)
548:                    sWord = sWord & vbNewLine & sNameMod & CHR_TO & sNameSub & CHR_TO & sType & CHR_TO & sWordTemp
549:                End If
550:            End If
551:        Next itemArr
552:    End If
553:    sWord = VBA.Trim$(sWord)
554:    If VBA.Len(sWord) = 0 Then
555:        sWord = vbNullString
556:    Else
557:        sWord = VBA.Trim$(VBA.Right$(sWord, VBA.Len(sWord) - 2))
558:    End If
559:    ParserStrDimConst = sWord
560: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserNameSubFunc - сбор названий процедур и функций
'* Created    : 27-03-2020 13:20
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objCodeModule As VBIDE.CodeModule : объект модуль
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ParserNameSubFunc(ByVal sNameVBC As String, ByRef objVBC As VBIDE.VBComponent, ByRef varSubFun As Scripting.Dictionary)
574:    Dim ProcKind    As VBIDE.vbext_ProcKind
575:    Dim lLine       As Long
576:    Dim lineOld     As Long
577:    Dim sNameSub    As String
578:    Dim strFunctionBody As String
579:    With objVBC.CodeModule
580:        If .CountOfLines > 0 Then
581:            lLine = .CountOfDeclarationLines
582:            If lLine = 0 Then lLine = 2
583:            Do Until lLine >= .CountOfLines
584:
585:                'сбор названий процедур и функций
586:                sNameSub = .ProcOfLine(lLine, ProcKind)
587:                If sNameSub <> vbNullString Then
588:                    strFunctionBody = C_PublicFunctions.TrimSpace(.Lines(lLine - 1, .ProcCountLines(sNameSub, ProcKind)))
589:                    If (Not strFunctionBody Like "*As IRibbonControl*") And _
                                (Not WorkBookAndSheetsEvents(strFunctionBody, objVBC.Type)) And _
                                (Not (strFunctionBody Like "* UserForm_*" And objVBC.Type = vbext_ct_MSForm)) And _
                                (Not UserFormsEvents(strFunctionBody, objVBC.Type)) Then
593:                        Dim sKey As String
594:                        sKey = sNameVBC & CHR_TO & TypeProcedyre(strFunctionBody) & CHR_TO & TypeOfAccessModifier(strFunctionBody) & CHR_TO & sNameSub
595:                        If Not varSubFun.Exists(sKey) Then
596:                            varSubFun.Add sKey, objVBC.Type
597:                        End If
598:                    End If
599:                    lLine = .ProcStartLine(sNameSub, ProcKind) + .ProcCountLines(sNameSub, ProcKind) + 1
600:                Else
601:                    lLine = lLine + 1
602:                End If
603:                If lineOld > lLine Then Exit Do
604:                lineOld = lLine
605:            Loop
606:        End If
607:    End With
608: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ParserNameControlsForm - сбор названий контролов юзерформ
'* Created    : 27-03-2020 13:50
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                         Description
'*
'* ByRef objVBC As VBIDE.VBComponent :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ParserNameControlsForm(ByVal sNameVBC As String, ByRef objVBC As VBIDE.VBComponent, ByRef obfNewDict As Scripting.Dictionary)
622:    Dim objCont     As MSForms.control
623:    If Not objVBC.Designer Is Nothing Then
624:        With objVBC.Designer
625:            For Each objCont In .Controls
626:                obfNewDict.Add sNameVBC & CHR_TO & objCont.Name, objVBC.Type
627:            Next objCont
628:        End With
629:    End If
630: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserNameGlobalVariable - сбор глобальных переменных
'* Created    : 27-03-2020 15:38
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                         Description
'*
'* ByVal sDeclarationLines As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ParserNameGlobalVariable(ByVal sNameVBC As String, ByRef objVBC As VBIDE.VBComponent, ByRef dicGloblVar As Scripting.Dictionary, ByRef dicTypeEnum As Scripting.Dictionary, ByRef dicAPI As Scripting.Dictionary)
644:    Dim varArr      As Variant
645:    Dim varArrWord  As Variant
646:    Dim varStr      As Variant
647:    Dim itemVarStr  As Variant
648:    Dim varAPI      As Variant
649:    Dim sTemp       As String
650:    Dim sTempArr    As String
651:    Dim i           As Long
652:    Dim bFlag       As Boolean
653:    Dim j           As Byte
654:    Dim itemArr     As Byte
655:    bFlag = True
656:    If objVBC.CodeModule.CountOfDeclarationLines <> 0 Then
657:        sTemp = objVBC.CodeModule.Lines(1, objVBC.CodeModule.CountOfDeclarationLines)
658:        sTemp = VBA.Replace(sTemp, " _" & vbNewLine, vbNullString)
659:        If sTemp <> vbNullString Then
660:            varArr = VBA.Split(sTemp, vbNewLine)
661:            For i = 0 To UBound(varArr)
662:                sTemp = C_PublicFunctions.TrimSpace(DeleteCommentString(varArr(i)))
663:                If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
664:                    If sTemp Like "* Type *" Or sTemp Like "* Enum *" Or sTemp Like "Type *" Or sTemp Like "Enum *" Then
665:                        varArrWord = VBA.Split(sTemp, " ")
666:                        If UBound(varArrWord) = 2 Then
667:                            sTemp = VBA.Trim$(varArrWord(0)) & CHR_TO & VBA.Trim$(varArrWord(1)) & CHR_TO & VBA.Trim$(varArrWord(2))
668:                        ElseIf UBound(varArrWord) = 1 Then
669:                            sTemp = "Public" & CHR_TO & VBA.Trim$(varArrWord(0)) & CHR_TO & VBA.Trim$(varArrWord(1))
670:                        End If
671:                        sTemp = sNameVBC & CHR_TO & sTemp
672:                        If Not dicTypeEnum.Exists(sTemp) Then dicTypeEnum.Add sTemp, objVBC.Type
673:                        bFlag = False
674:                    End If
675:                    If bFlag And Not (sTemp Like "Implements *" Or sTemp Like "Option *" Or VBA.Left$(sTemp, 1) = "'" Or sTemp = vbNullString Or VBA.Left$(sTemp, 1) = "#" Or sTemp Like "*Declare *(*)*" Or sTemp Like "*Event *(*)") Then
676:
677:                        If sTemp Like "* = *" Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ", vbTextCompare) + 2)
678:                        If sTemp Like "* *(* To *) *" Then
679:                            sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, "(", vbTextCompare) - 1)
680:                        End If
681:                        varStr = VBA.Split(sTemp, ",")
682:                        For Each itemVarStr In varStr
683:                            sTemp = VBA.Trim$(itemVarStr)
684:                            varArrWord = VBA.Split(sTemp, " As ")
685:                            varArrWord = VBA.Split(varArrWord(0), " = ")
686:                            sTemp = varArrWord(0)
687:                            varArrWord = VBA.Split(sTemp, " ")
688:
689:                            j = UBound(varArrWord)
690:                            If j > 1 Then
691:                                If varArrWord(0) = "Dim" Or varArrWord(0) = "Const" Then
692:                                    sTemp = "Private" & CHR_TO & varArrWord(0) & CHR_TO
693:                                    sTempArr = varArrWord(1)
694:                                ElseIf (varArrWord(0) = "Private" Or varArrWord(0) = "Public") And (varArrWord(1) = "Dim" Or varArrWord(1) = "Const" Or varArrWord(1) = "WithEvents") Then
695:                                    sTemp = varArrWord(0) & CHR_TO & varArrWord(1) & CHR_TO
696:                                    sTempArr = varArrWord(2)
697:                                ElseIf (varArrWord(0) = "Private" Or varArrWord(0) = "Public") And Not (varArrWord(1) = "Dim" Or varArrWord(1) = "Const" Or varArrWord(1) = "WithEvents") Then
698:                                    sTemp = varArrWord(0) & CHR_TO & "Dim" & CHR_TO
699:                                    sTempArr = varArrWord(1)
700:                                End If
701:                            ElseIf j = 1 And varArrWord(0) = "Global" Then
702:                                sTemp = "Public" & CHR_TO & varArrWord(0) & CHR_TO
703:                                sTempArr = varArrWord(1)
704:                            ElseIf j = 1 And (varArrWord(0) = "Private" Or varArrWord(0) = "Public") Then
705:                                sTemp = varArrWord(0) & CHR_TO & "Dim" & CHR_TO
706:                                sTempArr = varArrWord(1)
707:                            ElseIf j = 1 And (varArrWord(0) = "Dim" Or varArrWord(0) = "Const") Then
708:                                sTemp = "Private" & CHR_TO & varArrWord(0) & CHR_TO
709:                                sTempArr = varArrWord(1)
710:                            ElseIf j = 0 Then
711:                                sTemp = "Private" & CHR_TO & " Dim" & CHR_TO
712:                                sTempArr = varArrWord(0)
713:                            End If
714:
715:                            sTempArr = ReplaceType(sTempArr)
716:                            If sTempArr Like "*(*" Then sTempArr = VBA.Left$(sTempArr, VBA.InStr(1, sTempArr, "(") - 1)
717:                            sTemp = sNameVBC & CHR_TO & sTemp & sTempArr
718:                            If Not dicGloblVar.Exists(sTemp) Then dicGloblVar.Add sTemp, objVBC.Type
719:
720:                            sTemp = vbNullString
721:                        Next itemVarStr
722:                        sTemp = vbNullString
723:                    End If
724:                    If sTemp Like "*End Type" Or sTemp Like "*End Enum" Then
725:                        bFlag = True
726:                    End If
727:                    If sTemp Like "*Declare * Lib " & VBA.Chr$(34) & "*" & VBA.Chr$(34) & " (*)*" Then
728:                        sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " Lib ", vbTextCompare) - 1)
729:                        varAPI = VBA.Split(sTemp, VBA.Chr$(32))
730:                        itemArr = UBound(varAPI)
731:                        sTemp = CHR_TO & varAPI(itemArr - 1) & CHR_TO & varAPI(itemArr)
732:                        If varAPI(1) = "Declare" Then
733:                            sTemp = sNameVBC & CHR_TO & varAPI(0) & sTemp
734:                        Else
735:                            sTemp = sNameVBC & CHR_TO & "Private" & sTemp
736:                        End If
737:                        If Not dicAPI.Exists(sTemp) Then dicAPI.Add sTemp, objVBC.Type
738:                    End If
739:                    If sTemp Like "*Event *(*)" Then
740:                        sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, "(", vbTextCompare) - 1)
741:                        varAPI = VBA.Split(sTemp, VBA.Chr$(32))
742:                        itemArr = UBound(varAPI)
743:                        sTemp = CHR_TO & varAPI(itemArr - 1) & CHR_TO & varAPI(itemArr)
744:                        If varAPI(1) = "Event" Then
745:                            sTemp = sNameVBC & CHR_TO & varAPI(0) & sTemp
746:                        Else
747:                            sTemp = sNameVBC & CHR_TO & "Private" & sTemp
748:                        End If
749:                        If Not dicAPI.Exists(sTemp) Then dicAPI.Add sTemp, objVBC.Type
750:                    End If
751:                End If
752:            Next i
753:        End If
754:    End If
755: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddNewDictionary -функция инициализации словаря
'* Created    : 27-03-2020 13:21
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objDict As Scripting.Dictionary :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function AddNewDictionary(ByRef objDict As Scripting.Dictionary) As Scripting.Dictionary
768:    Set objDict = Nothing
769:    Set objDict = New Scripting.Dictionary
770:    Set AddNewDictionary = objDict
771: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DeleteCommentString - удаление в строке комментария
'* Created    : 20-04-2020 18:18
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sWord As String : строка
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function DeleteCommentString(ByVal sWord As String) As String
785:    'есть '
786:    Dim sTemp       As String
787:    sTemp = sWord
788:    If VBA.InStr(1, sTemp, "'") <> 0 Then
789:        If VBA.InStr(1, sTemp, VBA.Chr(34)) <> 0 Then
790:            'есть "
791:            If VBA.InStr(1, sTemp, "'") < VBA.InStr(1, sTemp, VBA.Chr(34)) Then
792:                'если так -> '"
793:                sTemp = VBA.Trim$(VBA.Left$(sTemp, VBA.InStr(1, sTemp, "'") - 1))
794:            End If
795:        Else
796:            'нет " -> '
797:            sTemp = VBA.Trim$(VBA.Left$(sTemp, VBA.InStr(1, sTemp, "'") - 1))
798:        End If
799:    End If
800:    DeleteCommentString = sTemp
801: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddEncodeName - функция генерации случайного зашифрованного имени
'* Created    : 27-03-2020 13:22
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function AddEncodeName() As String
810:    Const CharCount As Integer = 20
811:    Dim i           As Integer
812:    Dim sName       As String
813:
814:    Const FIRST_CODE_SIGN As String = "1"
815:    Const SECOND_CODE_SIGN As String = "0"
tryAgain:
817:    Err.Clear
818:    sName = vbNullString
819:    Randomize
820:    sName = "o"
821:    For i = 2 To CharCount
822:        If (VBA.Round(VBA.Rnd() * 1000)) Mod 2 = 1 Then sName = sName & FIRST_CODE_SIGN Else sName = sName & SECOND_CODE_SIGN
823:    Next i
824:    On Error Resume Next
825:    'добовляем новое имя в коллекцию, если имя существует то
826:    'генерируется ошибка, запуск повторной генерации имени
827:    objCollUnical.Add sName, sName
828:    If Err.Number <> 0 Then GoTo tryAgain
829:    AddEncodeName = sName
830: End Function
'взято
     Private Function TypeOfAccessModifier(ByRef StrDeclarationProcedure As String) As String
833:    If StrDeclarationProcedure Like "*Private *(*)*" Then
834:        TypeOfAccessModifier = "Private"
835:    Else
836:        TypeOfAccessModifier = "Public"
837:    End If
838: End Function
     Private Function TypeProcedyre(ByRef StrDeclarationProcedure As String) As String
840:    If StrDeclarationProcedure Like "*Sub *" Then
841:        TypeProcedyre = "Sub"
842:    ElseIf StrDeclarationProcedure Like "*Function *" Then
843:        TypeProcedyre = "Function"
844:    ElseIf StrDeclarationProcedure Like "*Property Set *" Then
845:        TypeProcedyre = "Property Set"
846:    ElseIf StrDeclarationProcedure Like "*Property Get *" Then
847:        TypeProcedyre = "Property Get"
848:    ElseIf StrDeclarationProcedure Like "*Property Let *" Then
849:        TypeProcedyre = "Property Let"
850:    Else
851:        TypeProcedyre = "Unknown Type"
852:    End If
853: End Function

     Private Function ReplaceType(ByVal sVar As String) As String
856:    sVar = Replace(sVar, "%", vbNullString)     'Integer
857:    sVar = Replace(sVar, "&", vbNullString)     'Long
858:    sVar = Replace(sVar, "$", vbNullString)     'String
859:    sVar = Replace(sVar, "!", vbNullString)     'Single
860:    sVar = Replace(sVar, "#", vbNullString)     'Double
861:    sVar = Replace(sVar, "@", vbNullString)     'Currency
862:    ReplaceType = sVar
863: End Function

     Private Function WorkBookAndSheetsEvents(ByVal sTxt As String, ByVal TypeModule As VBIDE.vbext_ComponentType) As Boolean
866:    Dim Flag        As Boolean
867:    Flag = False
868:    'только для модулей листов, книг и класов
869:    If TypeModule = vbext_ct_Document Or TypeModule = vbext_ct_ClassModule Then
870:        Select Case True
            Case sTxt Like "*_Activate(*": Flag = True
872:            Case sTxt Like "*_AddinInstall(*": Flag = True
873:            Case sTxt Like "*_AddinUninstall(*": Flag = True
874:            Case sTxt Like "*_AfterSave(*": Flag = True
875:            Case sTxt Like "*_AfterXmlExport(*": Flag = True
876:            Case sTxt Like "*_AfterXmlImport(*": Flag = True
877:            Case sTxt Like "*_BeforeClose(*": Flag = True
878:            Case sTxt Like "*_BeforeDoubleClick(*": Flag = True
879:            Case sTxt Like "*_BeforePrint(*": Flag = True
880:            Case sTxt Like "*_BeforeRightClick(*": Flag = True
881:            Case sTxt Like "*_BeforeSave(*": Flag = True
882:            Case sTxt Like "*_BeforeXmlExport(*": Flag = True
883:            Case sTxt Like "*_BeforeXmlImport(*": Flag = True
884:            Case sTxt Like "*_Calculate(*": Flag = True
885:            Case sTxt Like "*_Change(*": Flag = True
886:            Case sTxt Like "*_Deactivate(*": Flag = True
887:            Case sTxt Like "*_FollowHyperlink(*": Flag = True
888:            Case sTxt Like "*_MouseDown(*": Flag = True
889:            Case sTxt Like "*_MouseMove(*": Flag = True
890:            Case sTxt Like "*_MouseUp(*": Flag = True
891:            Case sTxt Like "*_NewChart(*": Flag = True
892:            Case sTxt Like "*_NewSheet(*": Flag = True
893:            Case sTxt Like "*_Open(*": Flag = True
894:            Case sTxt Like "*_PivotTableAfterValueChange(*": Flag = True
895:            Case sTxt Like "*_PivotTableBeforeAllocateChanges(*": Flag = True
896:            Case sTxt Like "*_PivotTableBeforeCommitChanges(*": Flag = True
897:            Case sTxt Like "*_PivotTableBeforeDiscardChanges(*": Flag = True
898:            Case sTxt Like "*_PivotTableChangeSync(*": Flag = True
899:            Case sTxt Like "*_PivotTableCloseConnection(*": Flag = True
900:            Case sTxt Like "*_PivotTableOpenConnection(*": Flag = True
901:            Case sTxt Like "*_PivotTableUpdate(*": Flag = True
902:            Case sTxt Like "*_Resize(*": Flag = True
903:            Case sTxt Like "*_RowsetComplete(*": Flag = True
904:            Case sTxt Like "*_SelectionChange(*": Flag = True
905:            Case sTxt Like "*_SeriesChange(*": Flag = True
906:            Case sTxt Like "*_SheetActivate(*": Flag = True
907:            Case sTxt Like "*_SheetBeforeDoubleClick(*": Flag = True
908:            Case sTxt Like "*_SheetBeforeRightClick(*": Flag = True
909:            Case sTxt Like "*_SheetCalculate(*": Flag = True
910:            Case sTxt Like "*_SheetChange(*": Flag = True
911:            Case sTxt Like "*_SheetDeactivate(*": Flag = True
912:            Case sTxt Like "*_SheetFollowHyperlink(*": Flag = True
913:            Case sTxt Like "*_SheetPivotTableAfterValueChange(*": Flag = True
914:            Case sTxt Like "*_SheetPivotTableBeforeAllocateChanges(*": Flag = True
915:            Case sTxt Like "*_SheetPivotTableBeforeCommitChanges(*": Flag = True
916:            Case sTxt Like "*_SheetPivotTableBeforeDiscardChanges(*": Flag = True
917:            Case sTxt Like "*_SheetPivotTableChangeSync(*": Flag = True
918:            Case sTxt Like "*_SheetPivotTableUpdate(*": Flag = True
919:            Case sTxt Like "*_SheetSelectionChange(*": Flag = True
920:            Case sTxt Like "*_Sync(*": Flag = True
921:            Case sTxt Like "*_WindowActivate(*": Flag = True
922:            Case sTxt Like "*_WindowDeactivate(*": Flag = True
923:            Case sTxt Like "*_WindowResize(*": Flag = True
924:            Case sTxt Like "*_NewWorkbook(*": Flag = True
925:            Case sTxt Like "*_WorkbookActivate(*": Flag = True
926:            Case sTxt Like "*_WorkbookAddinInstall(*": Flag = True
927:            Case sTxt Like "*_WorkbookAddinUninstall(*": Flag = True
928:            Case sTxt Like "*_WorkbookAfterSave(*": Flag = True
929:            Case sTxt Like "*_WorkbookAfterXmlExport(*": Flag = True
930:            Case sTxt Like "*_WorkbookAfterXmlImport(*": Flag = True
931:            Case sTxt Like "*_WorkbookBeforeClose(*": Flag = True
932:            Case sTxt Like "*_WorkbookBeforePrint(*": Flag = True
933:            Case sTxt Like "*_WorkbookBeforeSave(*": Flag = True
934:            Case sTxt Like "*_WorkbookBeforeXmlExport(*": Flag = True
935:            Case sTxt Like "*_WorkbookBeforeXmlImport(*": Flag = True
936:            Case sTxt Like "*_WorkbookDeactivate(*": Flag = True
937:            Case sTxt Like "*_WorkbookModelChange(*": Flag = True
938:            Case sTxt Like "*_WorkbookNewChart(*": Flag = True
939:            Case sTxt Like "*_WorkbookNewSheet(*": Flag = True
940:            Case sTxt Like "*_WorkbookOpen(*": Flag = True
941:            Case sTxt Like "*_WorkbookPivotTableCloseConnection(*": Flag = True
942:            Case sTxt Like "*_WorkbookPivotTableOpenConnection(*": Flag = True
943:            Case sTxt Like "*_WorkbookRowsetComplete(*": Flag = True
944:            Case sTxt Like "*_WorkbookSync(*": Flag = True
945:        End Select
946:    End If
947:    WorkBookAndSheetsEvents = Flag
948: End Function

Private Function UserFormsEvents(ByVal sTxt As String, ByVal TypeModule As VBIDE.vbext_ComponentType) As Boolean
951:    Dim Flag        As Boolean
952:    Flag = False
953:    'только для событий юзер форм и класов
954:    If TypeModule = vbext_ct_MSForm Or TypeModule = vbext_ct_ClassModule Then
955:        Select Case True
            Case sTxt Like "*_AfterUpdate(*": Flag = True
957:            Case sTxt Like "*_BeforeDragOver(*": Flag = True
958:            Case sTxt Like "*_BeforeDropOrPaste(*": Flag = True
959:            Case sTxt Like "*_BeforeUpdate(*": Flag = True
960:            Case sTxt Like "*_Change(*": Flag = True
961:            Case sTxt Like "*_Click(*": Flag = True
962:            Case sTxt Like "*_DblClick(*": Flag = True
963:            Case sTxt Like "*_Deactivate(*": Flag = True
964:            Case sTxt Like "*_DropButtonClick(*": Flag = True
965:            Case sTxt Like "*_Enter(*": Flag = True
966:            Case sTxt Like "*_Error(*": Flag = True
967:            Case sTxt Like "*_Exit(*": Flag = True
968:            Case sTxt Like "*_Initialize(*": Flag = True
969:            Case sTxt Like "*_KeyDown(*": Flag = True
970:            Case sTxt Like "*_KeyPress(*": Flag = True
971:            Case sTxt Like "*_KeyUp(*": Flag = True
972:            Case sTxt Like "*_Layout(*": Flag = True
973:            Case sTxt Like "*_MouseDown(*": Flag = True
974:            Case sTxt Like "*_MouseMove(*": Flag = True
975:            Case sTxt Like "*_MouseUp(*": Flag = True
976:            Case sTxt Like "*_QueryClose(*": Flag = True
977:            Case sTxt Like "*_RemoveControl(*": Flag = True
978:            Case sTxt Like "*_Resize(*": Flag = True
979:            Case sTxt Like "*_Scroll(*": Flag = True
980:            Case sTxt Like "*_Terminate(*": Flag = True
981:            Case sTxt Like "*_Zoom(*": Flag = True
982:        End Select
983:    End If
984:    UserFormsEvents = Flag
End Function

