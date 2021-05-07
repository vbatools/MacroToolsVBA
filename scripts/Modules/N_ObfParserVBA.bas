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
29:    Dim objWB       As Workbook
30:
31:    On Error GoTo ErrStartParser
32:    Application.Calculation = xlCalculationManual
33:    Set Form = New AddStatistic
34:    With Form
35:        .Caption = "Code base data collection:"
36:        .lbOK.Caption = "PARSE CODE"
37:        .chQuestion.visible = True
38:        .chQuestion.Value = True
39:        .chQuestion.Caption = "Collect string values?"
40:        .Show
41:        sNameWB = .cmbMain.Value
42:    End With
43:    If sNameWB = vbNullString Then Exit Sub
44:    Set objWB = Workbooks(sNameWB)
45:    Call MainObfParser(objWB, Form.chQuestion.Value)
46:    Set Form = Nothing
47:    Application.Calculation = xlCalculationAutomatic
48:    Exit Sub
ErrStartParser:
50:    Application.Calculation = xlCalculationAutomatic
51:    Application.ScreenUpdating = True
52:    Call MsgBox("Error in N_ObfParserVBA. Start Parser" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbCritical, "Error:")
53:    Call WriteErrorLog("AddShapeStatistic")
54: End Sub

    Private Sub MainObfParser(ByRef objWB As Workbook, Optional bEncodeStr As Boolean = False)
57:    If objWB.VBProject.Protection = vbext_pp_locked Then
58:        Call MsgBox("The project is protected, remove the password!", vbCritical, "The project is protected:")
59:    Else
60:        Call ParserProjectVBA(objWB, bEncodeStr)
61:        Call MsgBox("The code of the book [" & objWB.Name & "] is collected!", vbInformation, "Code Parsing:")
62:    End If
63: End Sub

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
     Private Sub ParserProjectVBA(ByRef objWB As Workbook, Optional bEncodeStr As Boolean = False)
77:    Dim objVBComp   As VBIDE.VBComponent
78:    Dim varModule   As obfModule
79:    Dim i           As Long
80:    Dim k           As Long
81:    Dim objDict     As Scripting.Dictionary
82:
83:
84:    With varModule
85:        'главный парсер
86:        Set .objName = AddNewDictionary(.objName)
87:        Set .objDimVar = AddNewDictionary(.objDimVar)
88:        Set .objSubFun = AddNewDictionary(.objSubFun)
89:        Set .objContr = AddNewDictionary(.objContr)
90:        Set .objTypeEnum = AddNewDictionary(.objTypeEnum)
91:        Set .objNameGlobVar = AddNewDictionary(.objNameGlobVar)
92:        Set .objStringCode = AddNewDictionary(.objStringCode)
93:        Set .objAPI = AddNewDictionary(.objAPI)
94:
95:        For Each objVBComp In objWB.VBProject.VBComponents
96:            'собираю названия модулей
97:            Dim sKey As String
98:            sKey = objVBComp.Type & CHR_TO & objVBComp.Name
99:            If Not .objName.Exists(sKey) Then .objName.Add sKey, 0
100:            'собираю названия контролов форм
101:            Call ParserNameControlsForm(objVBComp.Name, objVBComp, .objContr)
102:            'собираю названия процедур и функций
103:            Call ParserNameSubFunc(objVBComp.Name, objVBComp, .objSubFun)
104:            'собираю названия глобальные переменые
105:            Call ParserNameGlobalVariable(objVBComp.Name, objVBComp, .objNameGlobVar, .objTypeEnum, .objAPI)
106:            'собираю переменные процедур и функций, строковые переменые
107:            Call ParserVariebleSubFunc(objVBComp, .objDimVar, .objStringCode)
108:        Next objVBComp
109:        'конец парсера
110:    End With
111:
112:    'создание листа в активной книге
113:    Call AddShhetInWBook(NAME_SH, ActiveWorkbook)
114:
115:    ReDim arrRange(1 To varModule.objName.Count + varModule.objNameGlobVar.Count + varModule.objSubFun.Count + varModule.objContr.Count + varModule.objDimVar.Count + varModule.objTypeEnum.Count + varModule.objAPI.Count, 1 To 10) As String
116:
117:    Set objDict = New Scripting.Dictionary
118:
119:
120:    For i = 1 To varModule.objName.Count
121:        arrRange(i, 1) = "Module"
122:        arrRange(i, 2) = VBA.Split(varModule.objName.Keys(i - 1), CHR_TO)(0)
123:        arrRange(i, 3) = VBA.Split(varModule.objName.Keys(i - 1), CHR_TO)(1)
124:        arrRange(i, 4) = "Public"
125:        arrRange(i, 8) = arrRange(i, 3)
126:        arrRange(i, 9) = "YES"
127:
128:        If objDict.Exists(arrRange(i, 8)) = False Then
129:            objDict.Add arrRange(i, 8), AddEncodeName()
130:        End If
131:        arrRange(i, 10) = objDict.Item(arrRange(i, 8))
132:    Next i
133:    k = i
134:    Application.StatusBar = "Data collection: Module names, completed:" & VBA.Format(1 / 7, "Percent")
135:    For i = 1 To varModule.objNameGlobVar.Count
136:        arrRange(k, 1) = "Global variable"
137:        arrRange(k, 2) = varModule.objNameGlobVar.Items(i - 1)
138:        arrRange(k, 3) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(0)
139:        arrRange(k, 4) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(1)
140:        arrRange(k, 6) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(2)
141:        arrRange(k, 7) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(3)
142:        arrRange(k, 8) = arrRange(k, 7)
143:        arrRange(k, 9) = "YES"
144:
145:        If objDict.Exists(arrRange(k, 8)) = False Then
146:            objDict.Add arrRange(k, 8), AddEncodeName()
147:        End If
148:        arrRange(k, 10) = objDict.Item(arrRange(k, 8))
149:        k = k + 1
150:    Next i
151:
152:    Application.StatusBar = "Data collection: Global variables, completed:" & VBA.Format(2 / 7, "Percent")
153:    For i = 1 To varModule.objSubFun.Count
154:        arrRange(k, 1) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(1)
155:        arrRange(k, 2) = varModule.objSubFun.Items(i - 1)
156:        arrRange(k, 3) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(0)
157:        arrRange(k, 4) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(2)
158:        arrRange(k, 5) = arrRange(k, 1)
159:        arrRange(k, 6) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(3)
160:        arrRange(k, 8) = arrRange(k, 6)
161:        arrRange(k, 9) = "YES"
162:
163:        If objDict.Exists(arrRange(k, 8)) = False Then
164:            objDict.Add arrRange(k, 8), AddEncodeName()
165:        End If
166:        arrRange(k, 10) = objDict.Item(arrRange(k, 8))
167:        k = k + 1
168:    Next i
169:
170:    Application.StatusBar = "Data collection: Procedure names, completed:" & VBA.Format(3 / 7, "Percent")
171:    For i = 1 To varModule.objContr.Count
172:        arrRange(k, 1) = "Control"
173:        arrRange(k, 2) = varModule.objContr.Items(i - 1)
174:        arrRange(k, 3) = VBA.Split(varModule.objContr.Keys(i - 1), CHR_TO)(0)
175:        arrRange(k, 4) = "Private"
176:        arrRange(k, 6) = VBA.Split(varModule.objContr.Keys(i - 1), CHR_TO)(1)
177:        arrRange(k, 8) = arrRange(k, 6)
178:        arrRange(k, 9) = "YES"
179:
180:        If objDict.Exists(arrRange(k, 8)) = False Then
181:            objDict.Add arrRange(k, 8), AddEncodeName()
182:        End If
183:        arrRange(k, 10) = objDict.Item(arrRange(k, 8))
184:        k = k + 1
185:    Next i
186:
187:    Application.StatusBar = "Data collection: Control names, completed:" & VBA.Format(4 / 7, "Percent")
188:    For i = 1 To varModule.objDimVar.Count
189:        arrRange(k, 1) = "Variable"
190:        arrRange(k, 2) = varModule.objDimVar.Items(i - 1)
191:        arrRange(k, 3) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(0)
192:        arrRange(k, 4) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(3)
193:        arrRange(k, 5) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(1)
194:        arrRange(k, 6) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(2)
195:        arrRange(k, 7) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(4)
196:        arrRange(k, 8) = arrRange(k, 7)
197:        arrRange(k, 9) = "YES"
198:
199:        If objDict.Exists(arrRange(k, 8)) = False Then
200:            objDict.Add arrRange(k, 8), AddEncodeName()
201:        End If
202:        arrRange(k, 10) = objDict.Item(arrRange(k, 8))
203:        k = k + 1
204:        If i Mod 50 = 0 Then
205:            Application.StatusBar = "Data collection: Control names, completed:" & VBA.Format(i / varModule.objDimVar.Count, "Percent")
206:            DoEvents
207:        End If
208:    Next i
209:
210:    Application.StatusBar = "Data collection: Variable names, completed:" & VBA.Format(5 / 7, "Percent")
211:    For i = 1 To varModule.objTypeEnum.Count
212:        arrRange(k, 1) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(2)
213:        arrRange(k, 2) = varModule.objTypeEnum.Items(i - 1)
214:        arrRange(k, 3) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(0)
215:        arrRange(k, 4) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(1)
216:        arrRange(k, 6) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(3)
217:        arrRange(k, 8) = arrRange(k, 6)
218:        arrRange(k, 9) = "YES"
219:
220:        If objDict.Exists(arrRange(k, 8)) = False Then
221:            objDict.Add arrRange(k, 8), AddEncodeName()
222:        End If
223:        arrRange(k, 10) = objDict.Item(arrRange(k, 8))
224:        k = k + 1
225:    Next i
226:
227:    Application.StatusBar = "Data collection: Names of enumerations and types, completed:" & VBA.Format(6 / 7, "Percent")
228:    For i = 1 To varModule.objAPI.Count
229:        arrRange(k, 1) = "API"
230:        arrRange(k, 2) = varModule.objAPI.Items(i - 1)
231:        arrRange(k, 3) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(0)
232:        arrRange(k, 4) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(1)
233:        arrRange(k, 5) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(2)
234:        arrRange(k, 6) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(3)
235:        arrRange(k, 8) = arrRange(k, 6)
236:        arrRange(k, 9) = "YES"
237:
238:        If objDict.Exists(arrRange(k, 8)) = False Then
239:            objDict.Add arrRange(k, 8), AddEncodeName()
240:        End If
241:        arrRange(k, 10) = objDict.Item(arrRange(k, 8))
242:        k = k + 1
243:    Next i
244:    Application.StatusBar = "Data collection: API names, completed:" & VBA.Format(7 / 7, "Percent")
245:
246:    With ActiveSheet
247:        Application.StatusBar = "Applying formats"
248:        .Cells.ClearContents
249:        .Cells(1, 1).Value = "Type"
250:        .Cells(1, 2).Value = "Module type"
251:        .Cells(1, 3).Value = "Module name"
252:        .Cells(1, 4).Value = "Access Modifiers"
253:        .Cells(1, 5).Value = "Percentage type. and funk."
254:        .Cells(1, 6).Value = "The name of the percentage. and funk."
255:        .Cells(1, 7).Value = "Variable names"
256:        .Cells(1, 8).Value = "Encryption object"
257:        .Cells(1, 9).Value = "Encrypt yes/No"
258:        .Cells(1, 10).Value = "Cipher"
259:        .Cells(1, 11).Value = "Errors"
260:
261:        .Cells(2, 1).Resize(UBound(arrRange), 10) = arrRange
262:
263:        .Range(.Cells(2, 11), .Cells(k, 11)).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3]," & SHSNIPPETS.ListObjects(C_Const.TB_SERVICEWORDS).DataBodyRange.Address(ReferenceStyle:=xlR1C1, External:=True) & ",1,0),"""")"
264:        .Range(.Cells(2, 9), .Cells(k, 9)).FormulaR1C1 = "=IF(RC[2]="""",""YES"",""NO"")"
265:        .Columns("A:K").AutoFilter
266:        .Columns("A:K").EntireColumn.AutoFit
267:        .Range(Cells(2, 9), Cells(UBound(arrRange) + 1, 9)).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="YES, NO"
268:        Application.StatusBar = "Application of formats, finished"
269:    End With
270:
271:    'выгрузка строковых переменых
272:    If bEncodeStr Then
273:        Call AddShhetInWBook(NAME_SH_STR, ActiveWorkbook)
274:        Application.StatusBar = "Collecting String variables"
275:        If varModule.objStringCode.Count <> 0 Then
276:            ReDim arrRange(1 To varModule.objStringCode.Count, 1 To 8) As String
277:            For i = 1 To varModule.objStringCode.Count
278:                arrRange(i, 1) = varModule.objStringCode.Items(i - 1)
279:                arrRange(i, 2) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(0)
280:                arrRange(i, 3) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(1)
281:                arrRange(i, 4) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(2)
282:                arrRange(i, 5) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(3)
283:                arrRange(i, 6) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(4)
284:                arrRange(i, 7) = "YES"
285:                arrRange(i, 8) = AddEncodeName()
286:                If i Mod 50 = 0 Then
287:                    Application.StatusBar = "Collection of String variables, completed:" & VBA.Format(i / varModule.objStringCode.Count, "Percent")
288:                    DoEvents
289:                End If
290:            Next i
291:            Application.StatusBar = "Collecting String variables, completed"
292:            With ActiveSheet
293:                .Cells(1, 1).Value = "Module type"
294:                .Cells(1, 2).Value = "Module name"
295:                .Cells(1, 3).Value = "Тип Sub или Fun"
296:                .Cells(1, 4).Value = "Name of Sub or Function"
297:                .Cells(1, 5).Value = "String"
298:                .Cells(1, 6).Value = "Array strings"
299:                .Cells(1, 7).Value = "Encrypt yes/No"
300:                .Cells(1, 8).Value = "Cipher"
301:                .Cells(1, 9).Value = "Module cipher"
302:
303:                .Cells(1, 11).Value = "Code of the Const module"
304:                .Cells(2, 11).Value = AddEncodeName()
305:
306:                .Cells(2, 1).Resize(UBound(arrRange), 8) = arrRange
307:
308:                .Range(Cells(2, 7), Cells(UBound(arrRange) + 1, 7)).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="YES, NO"
309:                .Range(Cells(2, 9), Cells(UBound(arrRange) + 1, 9)).FormulaR1C1 = "=IF(RC1*1=100,RC2,VLOOKUP(RC2,DATA_OBF_VBATools!R2C3:R" & k & "C10,8,0))"
310:                .Columns("A:I").AutoFilter
311:                .Columns("A:D").EntireColumn.AutoFit
312:                .Columns("E").ColumnWidth = 60
313:                .Columns("F:K").EntireColumn.AutoFit
314:                .Rows("2:" & UBound(arrRange) + 1).RowHeight = 12
315:            End With
316:        End If
317:    End If
318:    ActiveWorkbook.Worksheets(NAME_SH).Activate
319:
320:    Application.StatusBar = False
321: End Sub
     Public Sub AddShhetInWBook(ByVal WSheetName As String, ByRef WB As Workbook)
323:    'создание листа в активной книге
324:    Application.DisplayAlerts = False
325:    On Error Resume Next
326:    WB.Worksheets(WSheetName).Delete
327:    On Error GoTo 0
328:    Application.DisplayAlerts = True
329:    WB.Sheets.Add Before:=ActiveSheet
330:    ActiveSheet.Name = WSheetName
331: End Sub

     Private Sub ParserVariebleSubFunc(ByRef objVBC As VBIDE.VBComponent, ByRef objDic As Scripting.Dictionary, ByRef objDicStr As Scripting.Dictionary)
334:    Dim lLine       As Long
335:    Dim sCode       As String
336:    Dim sVar        As String
337:    Dim sSubName    As String
338:    Dim sNumTypeName As String
339:    Dim sType       As String
340:    Dim arrStrCode  As Variant
341:    Dim arrEnum     As Variant
342:    Dim itemArr     As Variant
343:    Dim itemVar     As Variant
344:    Dim arrVar      As Variant
345:
346:    With objVBC.CodeModule
347:        lLine = .CountOfLines
348:        If lLine > 0 Then
349:            sCode = .Lines(1, lLine)
350:            If sCode <> vbNullString Then
351:                'убираю перенос строк
352:                sCode = VBA.Replace(sCode, " _" & vbNewLine, vbNullString)
353:                arrStrCode = VBA.Split(sCode, vbNewLine)
354:                For Each itemArr In arrStrCode
355:                    itemArr = C_PublicFunctions.TrimSpace(itemArr)
356:                    If itemArr <> vbNullString And VBA.Left$(itemArr, 1) <> "'" Then
357:                        sVar = vbNullString
358:                        'если есть коментарий в строке кода то удаляем его
359:                        itemArr = DeleteCommentString(itemArr)
360:                        'из строки декларирования и определение что вошли в процедуру
361:                        If (itemArr Like "* Sub *(*)*" Or itemArr Like "* Function *(*)*" Or itemArr Like "* Property Let *(*)*" Or itemArr Like "* Property Set *(*)*" Or itemArr Like "* Property Get *(*)*" Or _
                                    itemArr Like "Sub *(*)*" Or itemArr Like "Function *(*)*" Or itemArr Like "Property Let *(*)*" Or itemArr Like "Property Set *(*)*" Or itemArr Like "Property Get *(*)*") _
                                    And (Not itemArr Like "*As IRibbonControl*" And Not itemArr Like "* Declare *(*)*") Then
364:
365:                            sSubName = TypeProcedyre(VBA.CStr(itemArr))
366:                            sSubName = sSubName & CHR_TO & GetNameSubFromString(itemArr)
367:                            sVar = ParserStrDimConst(itemArr, sSubName, .Name)
368:
369:                        End If
370:                        'если в перечислении и типе данных
371:                        If itemArr Like "Private Enum *" Or itemArr Like "Public Enum *" Or itemArr Like "Enum *" Or itemArr Like "Private Type *" Or itemArr Like "Public Type *" Or itemArr Like "Type *" Then
372:                            arrEnum = VBA.Split(itemArr, " ")
373:                            If VBA.CStr(itemArr) Like "Private *" Then
374:                                sNumTypeName = "Private"
375:                            Else
376:                                sNumTypeName = "Public"
377:                            End If
378:                            sNumTypeName = arrEnum(UBound(arrEnum)) & CHR_TO & sNumTypeName
379:                            If itemArr Like "* Enum *" Or itemArr Like "Enum *" Then
380:                                sType = "Enum"
381:                            Else
382:                                sType = "Type"
383:                            End If
384:                        End If
385:                        'вышли из процедуры или перечисления
386:                        If itemArr Like "*End Sub" Or itemArr Like "*End Function" Or itemArr Like "*End Property" Or itemArr Like "*End Enum" Or itemArr Like "*End Type" Then
387:                            sSubName = vbNullString
388:                            sNumTypeName = vbNullString
389:                        End If
390:                        'если внутри типа или перечисления
391:                        If sNumTypeName <> vbNullString And Not itemArr Like "* Enum *" And Not itemArr Like "Enum *" And Not itemArr Like "* Type *" And Not itemArr Like "Type *" Then
392:                            arrEnum = VBA.Split(VBA.Trim$(itemArr), " ")
393:                            sVar = arrEnum(0)
394:                            If sVar Like "*(*" Then sVar = VBA.Left$(sVar, VBA.InStr(1, sVar, "(") - 1)
395:                            sVar = .Name & CHR_TO & sType & CHR_TO & sNumTypeName & CHR_TO & ReplaceType(sVar)
396:                        End If
397:                        'если находимся только внутри процедуры
398:                        If (itemArr Like "* Dim *" Or itemArr Like "* Const *" Or itemArr Like "Dim *" Or itemArr Like "Const *") And sSubName <> vbNullString Then
399:                            sVar = ParserStrDimConst(itemArr, sSubName, .Name)
400:                        End If
401:                        arrVar = VBA.Split(sVar, vbNewLine)
402:                        For Each itemVar In arrVar
403:                            If itemVar <> vbNullString And objDic.Exists(itemVar) = False Then
404:                                objDic.Add itemVar, objVBC.Type
405:                            End If
406:                        Next itemVar
407:                        Call ParserStringInCode(itemArr, sSubName, objVBC, objDicStr)
408:                    End If
409:                Next itemArr
410:            End If
411:        End If
412:    End With
413: End Sub

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
427:    Dim sTemp       As String
428:    sTemp = VBA.Trim$(VBA.Left$(sStrCode, VBA.InStr(1, sStrCode, "(") - 1))
429:    Select Case True
        Case sTemp Like "*Sub *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Sub ") - 3)
431:        Case sTemp Like "*Function *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Function ") - 8)
432:        Case sTemp Like "*Property Let *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Let ") - 12)
433:        Case sTemp Like "*Property Set *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Set ") - 12)
434:        Case sTemp Like "*Property Get *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Get ") - 12)
435:    End Select
436:    GetNameSubFromString = VBA.Trim$(sTemp)
437: End Function

     Private Sub ParserStringInCode(ByVal sSTR As String, ByVal sNameSub As String, ByRef objVBC As VBIDE.VBComponent, ByRef objDicStr As Scripting.Dictionary)
440:    Dim sTxt        As String
441:    Dim arrStr      As Variant
442:    Dim Arr         As Variant
443:    Dim sReplace    As String
444:    Dim i           As Integer
445:    Dim sArray      As String
446:    Const CHAR_REPLACE As String = "ЪЪЪЪ"
447:
448:    sSTR = VBA.Trim$(sSTR)
449:
450:    If sSTR Like "*" & VBA.Chr$(34) & "*" And sSTR <> vbNullString And Not sSTR Like "*Declare * Lib *(*)*" Then
451:
452:        sTxt = VBA.Right$(sSTR, VBA.Len(sSTR) - VBA.InStr(1, sSTR, VBA.Chr$(34)) + 1)
453:        sTxt = VBA.Replace(sTxt, VBA.Chr$(34) & VBA.Chr$(34), CHAR_REPLACE)
454:        arrStr = VBA.Split(sTxt, VBA.Chr$(34))
455:
456:        sArray = VBA.Left$(sSTR, VBA.InStr(1, sSTR, VBA.Chr$(34)) - 1)
457:        If sArray Like "* = Array(" Then
458:            sArray = VBA.Replace(sArray, " = Array(", vbNullString)
459:            Arr = VBA.Split(sArray, " ")
460:            sArray = Arr(UBound(Arr))
461:        Else
462:            sArray = vbNullString
463:        End If
464:        For i = 1 To UBound(arrStr) Step 2
465:            If arrStr(i) <> vbNullString Then
466:                If sNameSub = vbNullString Then sNameSub = "Declaration" & CHR_TO
467:
468:                sReplace = VBA.Replace(arrStr(i), CHAR_REPLACE, VBA.Chr$(34) & VBA.Chr$(34))
469:                sTxt = objVBC.Name & CHR_TO & sNameSub & CHR_TO & VBA.Chr$(34) & sReplace & VBA.Chr$(34) & CHR_TO & sArray    '& CHR_TO & sYesNo
470:                If arrStr(i + 1) Like "*: * = *" Then sArray = vbNullString
471:                If arrStr(i + 1) Like "*: * = Array(*" Then
472:                    sArray = VBA.Replace(arrStr(i + 1), ": ", vbNullString)
473:                    sArray = VBA.Replace(sArray, " = Array(", vbNullString)
474:                    sArray = VBA.Replace(sArray, ")", vbNullString)
475:                End If
476:                If objDicStr.Exists(sTxt) = False Then objDicStr.Add sTxt, objVBC.Type
477:            End If
478:        Next i
479:        sArray = vbNullString
480:    End If
481: End Sub

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
495:    Dim sTemp       As String
496:    Dim sWord       As String
497:    Dim sWordTemp   As String
498:    Dim arrStr      As Variant
499:    Dim itemArr     As Variant
500:    Dim arrWord     As Variant
501:    Dim sType       As String
502:
503:    sTemp = C_PublicFunctions.TrimSpace(sTxt)
504:    sType = "Dim"
505:    If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
506:        'если есть коментарий в строке кода то удаляем его
507:        sTemp = DeleteCommentString(sTemp)
508:        If sTemp Like "*Sub *(*)*" Or sTemp Like "*Function *(*)*" Or sTemp Like "*Property Let *(*)*" Or sTemp Like "*Property Set *(*)*" Or sTemp Like "*Property Get *(*)*" Then
509:            If VBA.InStr(1, sTemp, ")") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, ")") - 1)
510:            If VBA.InStr(1, sTemp, " = ") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ") - 1)
511:            If VBA.Len(sTemp) - VBA.InStr(1, sTemp, "(") >= 0 Then
512:                sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "("))
513:            End If
514:        ElseIf sTemp Like "* Dim *" Or sTemp Like Chr$(68) & "im *" Then
515:            sType = "Dim"
516:            If VBA.InStr(1, sTemp, "Dim ") >= 3 Then sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Dim ") - 3)
517:        ElseIf sTemp Like "* Const *" Or sTemp Like Chr$(67) & "onst *" Then
518:            sType = "Const"
519:            If VBA.InStr(1, sTemp, "Const ") >= 5 Then sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Const ") - 5)
520:            If VBA.InStr(1, sTemp, " = ") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ") - 1)
521:        Else
522:            sTemp = vbNullString
523:        End If
524:    End If
525:
526:    If sTemp Like "*: *" Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, ": ") - 1)
527:    If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
528:        arrStr = VBA.Split(sTemp, ",")
529:        For Each itemArr In arrStr
530:            If itemArr Like "*(*" Then itemArr = VBA.Left$(itemArr, VBA.InStr(1, itemArr, "(") - 1)
531:            If Not itemArr Like "*)*" And Not itemArr Like "* To *" Then
532:                arrWord = VBA.Split(itemArr, " As ")
533:                arrWord = VBA.Split(VBA.Trim$(arrWord(0)), " ")
534:                If UBound(arrWord) = -1 Then
535:                    sWord = vbNullString
536:                Else
537:                    sWordTemp = VBA.Trim$(arrWord(UBound(arrWord)))
538:                    sWordTemp = ReplaceType(sWordTemp)
539:                    sWord = sWord & vbNewLine & sNameMod & CHR_TO & sNameSub & CHR_TO & sType & CHR_TO & sWordTemp
540:                End If
541:            End If
542:        Next itemArr
543:    End If
544:    sWord = VBA.Trim$(sWord)
545:    If VBA.Len(sWord) = 0 Then
546:        sWord = vbNullString
547:    Else
548:        sWord = VBA.Trim$(VBA.Right$(sWord, VBA.Len(sWord) - 2))
549:    End If
550:    ParserStrDimConst = sWord
551: End Function

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
565:    Dim ProcKind    As VBIDE.vbext_ProcKind
566:    Dim lLine       As Long
567:    Dim lineOld     As Long
568:    Dim sNameSub    As String
569:    Dim strFunctionBody As String
570:    With objVBC.CodeModule
571:        If .CountOfLines > 0 Then
572:            lLine = .CountOfDeclarationLines
573:            If lLine = 0 Then lLine = 2
574:            Do Until lLine >= .CountOfLines
575:
576:                'сбор названий процедур и функций
577:                sNameSub = .ProcOfLine(lLine, ProcKind)
578:                If sNameSub <> vbNullString Then
579:                    strFunctionBody = C_PublicFunctions.TrimSpace(.Lines(lLine - 1, .ProcCountLines(sNameSub, ProcKind)))
580:                    If (Not strFunctionBody Like "*As IRibbonControl*") And _
                                (Not WorkBookAndSheetsEvents(strFunctionBody, objVBC.Type)) And _
                                (Not (strFunctionBody Like "* UserForm_*" And objVBC.Type = vbext_ct_MSForm)) And _
                                (Not UserFormsEvents(strFunctionBody, objVBC.Type)) Then
584:                        Dim sKey As String
585:                        sKey = sNameVBC & CHR_TO & TypeProcedyre(strFunctionBody) & CHR_TO & TypeOfAccessModifier(strFunctionBody) & CHR_TO & sNameSub
586:                        If Not varSubFun.Exists(sKey) Then
587:                            varSubFun.Add sKey, objVBC.Type
588:                        End If
589:                    End If
590:                    lLine = .ProcStartLine(sNameSub, ProcKind) + .ProcCountLines(sNameSub, ProcKind) + 1
591:                Else
592:                    lLine = lLine + 1
593:                End If
594:                If lineOld > lLine Then Exit Do
595:                lineOld = lLine
596:            Loop
597:        End If
598:    End With
599: End Sub

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
613:    Dim objCont     As MSForms.control
614:    If Not objVBC.Designer Is Nothing Then
615:        With objVBC.Designer
616:            For Each objCont In .Controls
617:                obfNewDict.Add sNameVBC & CHR_TO & objCont.Name, objVBC.Type
618:            Next objCont
619:        End With
620:    End If
621: End Sub

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
635:    Dim varArr      As Variant
636:    Dim varArrWord  As Variant
637:    Dim varStr      As Variant
638:    Dim itemVarStr  As Variant
639:    Dim varAPI      As Variant
640:    Dim sTemp       As String
641:    Dim sTempArr    As String
642:    Dim i           As Long
643:    Dim bFlag       As Boolean
644:    Dim j           As Byte
645:    Dim itemArr     As Byte
646:    bFlag = True
647:    If objVBC.CodeModule.CountOfDeclarationLines <> 0 Then
648:        sTemp = objVBC.CodeModule.Lines(1, objVBC.CodeModule.CountOfDeclarationLines)
649:        sTemp = VBA.Replace(sTemp, " _" & vbNewLine, vbNullString)
650:        If sTemp <> vbNullString Then
651:            varArr = VBA.Split(sTemp, vbNewLine)
652:            For i = 0 To UBound(varArr)
653:                sTemp = C_PublicFunctions.TrimSpace(DeleteCommentString(varArr(i)))
654:                If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
655:                    If sTemp Like "* Type *" Or sTemp Like "* Enum *" Or sTemp Like "Type *" Or sTemp Like "Enum *" Then
656:                        varArrWord = VBA.Split(sTemp, " ")
657:                        If UBound(varArrWord) = 2 Then
658:                            sTemp = VBA.Trim$(varArrWord(0)) & CHR_TO & VBA.Trim$(varArrWord(1)) & CHR_TO & VBA.Trim$(varArrWord(2))
659:                        ElseIf UBound(varArrWord) = 1 Then
660:                            sTemp = "Public" & CHR_TO & VBA.Trim$(varArrWord(0)) & CHR_TO & VBA.Trim$(varArrWord(1))
661:                        End If
662:                        sTemp = sNameVBC & CHR_TO & sTemp
663:                        If Not dicTypeEnum.Exists(sTemp) Then dicTypeEnum.Add sTemp, objVBC.Type
664:                        bFlag = False
665:                    End If
666:                    If bFlag And Not (sTemp Like "Implements *" Or sTemp Like "Option *" Or VBA.Left$(sTemp, 1) = "'" Or sTemp = vbNullString Or VBA.Left$(sTemp, 1) = "#" Or sTemp Like "*Declare *(*)*" Or sTemp Like "*Event *(*)") Then
667:
668:                        If sTemp Like "* = *" Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ", vbTextCompare) + 2)
669:                        If sTemp Like "* *(* To *) *" Then
670:                            sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, "(", vbTextCompare) - 1)
671:                        End If
672:                        varStr = VBA.Split(sTemp, ",")
673:                        For Each itemVarStr In varStr
674:                            sTemp = VBA.Trim$(itemVarStr)
675:                            varArrWord = VBA.Split(sTemp, " As ")
676:                            varArrWord = VBA.Split(varArrWord(0), " = ")
677:                            sTemp = varArrWord(0)
678:                            varArrWord = VBA.Split(sTemp, " ")
679:
680:                            j = UBound(varArrWord)
681:                            If j > 1 Then
682:                                If varArrWord(0) = "Dim" Or varArrWord(0) = "Const" Then
683:                                    sTemp = "Private" & CHR_TO & varArrWord(0) & CHR_TO
684:                                    sTempArr = varArrWord(1)
685:                                ElseIf (varArrWord(0) = "Private" Or varArrWord(0) = "Public") And (varArrWord(1) = "Dim" Or varArrWord(1) = "Const" Or varArrWord(1) = "WithEvents") Then
686:                                    sTemp = varArrWord(0) & CHR_TO & varArrWord(1) & CHR_TO
687:                                    sTempArr = varArrWord(2)
688:                                ElseIf (varArrWord(0) = "Private" Or varArrWord(0) = "Public") And Not (varArrWord(1) = "Dim" Or varArrWord(1) = "Const" Or varArrWord(1) = "WithEvents") Then
689:                                    sTemp = varArrWord(0) & CHR_TO & "Dim" & CHR_TO
690:                                    sTempArr = varArrWord(1)
691:                                End If
692:                            ElseIf j = 1 And varArrWord(0) = "Global" Then
693:                                sTemp = "Public" & CHR_TO & varArrWord(0) & CHR_TO
694:                                sTempArr = varArrWord(1)
695:                            ElseIf j = 1 And (varArrWord(0) = "Private" Or varArrWord(0) = "Public") Then
696:                                sTemp = varArrWord(0) & CHR_TO & "Dim" & CHR_TO
697:                                sTempArr = varArrWord(1)
698:                            ElseIf j = 1 And (varArrWord(0) = "Dim" Or varArrWord(0) = "Const") Then
699:                                sTemp = "Private" & CHR_TO & varArrWord(0) & CHR_TO
700:                                sTempArr = varArrWord(1)
701:                            ElseIf j = 0 Then
702:                                sTemp = "Private" & CHR_TO & " Dim" & CHR_TO
703:                                sTempArr = varArrWord(0)
704:                            End If
705:
706:                            sTempArr = ReplaceType(sTempArr)
707:                            If sTempArr Like "*(*" Then sTempArr = VBA.Left$(sTempArr, VBA.InStr(1, sTempArr, "(") - 1)
708:                            sTemp = sNameVBC & CHR_TO & sTemp & sTempArr
709:                            If Not dicGloblVar.Exists(sTemp) Then dicGloblVar.Add sTemp, objVBC.Type
710:
711:                            sTemp = vbNullString
712:                        Next itemVarStr
713:                        sTemp = vbNullString
714:                    End If
715:                    If sTemp Like "*End Type" Or sTemp Like "*End Enum" Then
716:                        bFlag = True
717:                    End If
718:                    If sTemp Like "*Declare * Lib " & VBA.Chr$(34) & "*" & VBA.Chr$(34) & " (*)*" Then
719:                        sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " Lib ", vbTextCompare) - 1)
720:                        varAPI = VBA.Split(sTemp, VBA.Chr$(32))
721:                        itemArr = UBound(varAPI)
722:                        sTemp = CHR_TO & varAPI(itemArr - 1) & CHR_TO & varAPI(itemArr)
723:                        If varAPI(1) = "Declare" Then
724:                            sTemp = sNameVBC & CHR_TO & varAPI(0) & sTemp
725:                        Else
726:                            sTemp = sNameVBC & CHR_TO & "Private" & sTemp
727:                        End If
728:                        If Not dicAPI.Exists(sTemp) Then dicAPI.Add sTemp, objVBC.Type
729:                    End If
730:                    If sTemp Like "*Event *(*)" Then
731:                        sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, "(", vbTextCompare) - 1)
732:                        varAPI = VBA.Split(sTemp, VBA.Chr$(32))
733:                        itemArr = UBound(varAPI)
734:                        sTemp = CHR_TO & varAPI(itemArr - 1) & CHR_TO & varAPI(itemArr)
735:                        If varAPI(1) = "Event" Then
736:                            sTemp = sNameVBC & CHR_TO & varAPI(0) & sTemp
737:                        Else
738:                            sTemp = sNameVBC & CHR_TO & "Private" & sTemp
739:                        End If
740:                        If Not dicAPI.Exists(sTemp) Then dicAPI.Add sTemp, objVBC.Type
741:                    End If
742:                End If
743:            Next i
744:        End If
745:    End If
746: End Sub
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
759:    Set objDict = Nothing
760:    Set objDict = New Scripting.Dictionary
761:    Set AddNewDictionary = objDict
762: End Function

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
776:    'есть '
777:    Dim sTemp       As String
778:    sTemp = sWord
779:    If VBA.InStr(1, sTemp, "'") <> 0 Then
780:        If VBA.InStr(1, sTemp, VBA.Chr(34)) <> 0 Then
781:            'есть "
782:            If VBA.InStr(1, sTemp, "'") < VBA.InStr(1, sTemp, VBA.Chr(34)) Then
783:                'если так -> '"
784:                sTemp = VBA.Trim$(VBA.Left$(sTemp, VBA.InStr(1, sTemp, "'") - 1))
785:            End If
786:        Else
787:            'нет " -> '
788:            sTemp = VBA.Trim$(VBA.Left$(sTemp, VBA.InStr(1, sTemp, "'") - 1))
789:        End If
790:    End If
791:    DeleteCommentString = sTemp
792: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddEncodeName - функция генерации случайного зашифрованного имени
'* Created    : 27-03-2020 13:22
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function AddEncodeName() As String
801:    Const CharCount As Integer = 20
802:    Dim i           As Integer
803:    Dim sName       As String
804:
805:    Const FIRST_CODE_SIGN As String = "1"
806:    Const SECOND_CODE_SIGN As String = "0"
tryAgain:
808:    Err.Clear
809:    sName = vbNullString
810:    Randomize
811:    sName = "o"
812:    For i = 2 To CharCount
813:        If (VBA.Round(VBA.Rnd() * 1000)) Mod 2 = 1 Then sName = sName & FIRST_CODE_SIGN Else sName = sName & SECOND_CODE_SIGN
814:    Next i
815:    On Error Resume Next
816:    'добовляем новое имя в коллекцию, если имя существует то
817:    'генерируется ошибка, запуск повторной генерации имени
818:    objCollUnical.Add sName, sName
819:    If Err.Number <> 0 Then GoTo tryAgain
820:    AddEncodeName = sName
821: End Function
'взято
     Private Function TypeOfAccessModifier(ByRef StrDeclarationProcedure As String) As String
824:    If StrDeclarationProcedure Like "*Private *(*)*" Then
825:        TypeOfAccessModifier = "Private"
826:    Else
827:        TypeOfAccessModifier = "Public"
828:    End If
829: End Function
     Private Function TypeProcedyre(ByRef StrDeclarationProcedure As String) As String
831:    If StrDeclarationProcedure Like "*Sub *" Then
832:        TypeProcedyre = "Sub"
833:    ElseIf StrDeclarationProcedure Like "*Function *" Then
834:        TypeProcedyre = "Function"
835:    ElseIf StrDeclarationProcedure Like "*Property Set *" Then
836:        TypeProcedyre = "Property Set"
837:    ElseIf StrDeclarationProcedure Like "*Property Get *" Then
838:        TypeProcedyre = "Property Get"
839:    ElseIf StrDeclarationProcedure Like "*Property Let *" Then
840:        TypeProcedyre = "Property Let"
841:    Else
842:        TypeProcedyre = "Unknown Type"
843:    End If
844: End Function

     Private Function ReplaceType(ByVal sVar As String) As String
847:    sVar = Replace(sVar, "%", vbNullString)     'Integer
848:    sVar = Replace(sVar, "&", vbNullString)     'Long
849:    sVar = Replace(sVar, "$", vbNullString)     'String
850:    sVar = Replace(sVar, "!", vbNullString)     'Single
851:    sVar = Replace(sVar, "#", vbNullString)     'Double
852:    sVar = Replace(sVar, "@", vbNullString)     'Currency
853:    ReplaceType = sVar
854: End Function

     Private Function WorkBookAndSheetsEvents(ByVal sTxt As String, ByVal TypeModule As VBIDE.vbext_ComponentType) As Boolean
857:    Dim Flag        As Boolean
858:    Flag = False
859:    'только для модулей листов, книг и класов
860:    If TypeModule = vbext_ct_Document Or TypeModule = vbext_ct_ClassModule Then
861:        Select Case True
            Case sTxt Like "*_Activate(*": Flag = True
863:            Case sTxt Like "*_AddinInstall(*": Flag = True
864:            Case sTxt Like "*_AddinUninstall(*": Flag = True
865:            Case sTxt Like "*_AfterSave(*": Flag = True
866:            Case sTxt Like "*_AfterXmlExport(*": Flag = True
867:            Case sTxt Like "*_AfterXmlImport(*": Flag = True
868:            Case sTxt Like "*_BeforeClose(*": Flag = True
869:            Case sTxt Like "*_BeforeDoubleClick(*": Flag = True
870:            Case sTxt Like "*_BeforePrint(*": Flag = True
871:            Case sTxt Like "*_BeforeRightClick(*": Flag = True
872:            Case sTxt Like "*_BeforeSave(*": Flag = True
873:            Case sTxt Like "*_BeforeXmlExport(*": Flag = True
874:            Case sTxt Like "*_BeforeXmlImport(*": Flag = True
875:            Case sTxt Like "*_Calculate(*": Flag = True
876:            Case sTxt Like "*_Change(*": Flag = True
877:            Case sTxt Like "*_Deactivate(*": Flag = True
878:            Case sTxt Like "*_FollowHyperlink(*": Flag = True
879:            Case sTxt Like "*_MouseDown(*": Flag = True
880:            Case sTxt Like "*_MouseMove(*": Flag = True
881:            Case sTxt Like "*_MouseUp(*": Flag = True
882:            Case sTxt Like "*_NewChart(*": Flag = True
883:            Case sTxt Like "*_NewSheet(*": Flag = True
884:            Case sTxt Like "*_Open(*": Flag = True
885:            Case sTxt Like "*_PivotTableAfterValueChange(*": Flag = True
886:            Case sTxt Like "*_PivotTableBeforeAllocateChanges(*": Flag = True
887:            Case sTxt Like "*_PivotTableBeforeCommitChanges(*": Flag = True
888:            Case sTxt Like "*_PivotTableBeforeDiscardChanges(*": Flag = True
889:            Case sTxt Like "*_PivotTableChangeSync(*": Flag = True
890:            Case sTxt Like "*_PivotTableCloseConnection(*": Flag = True
891:            Case sTxt Like "*_PivotTableOpenConnection(*": Flag = True
892:            Case sTxt Like "*_PivotTableUpdate(*": Flag = True
893:            Case sTxt Like "*_Resize(*": Flag = True
894:            Case sTxt Like "*_RowsetComplete(*": Flag = True
895:            Case sTxt Like "*_SelectionChange(*": Flag = True
896:            Case sTxt Like "*_SeriesChange(*": Flag = True
897:            Case sTxt Like "*_SheetActivate(*": Flag = True
898:            Case sTxt Like "*_SheetBeforeDoubleClick(*": Flag = True
899:            Case sTxt Like "*_SheetBeforeRightClick(*": Flag = True
900:            Case sTxt Like "*_SheetCalculate(*": Flag = True
901:            Case sTxt Like "*_SheetChange(*": Flag = True
902:            Case sTxt Like "*_SheetDeactivate(*": Flag = True
903:            Case sTxt Like "*_SheetFollowHyperlink(*": Flag = True
904:            Case sTxt Like "*_SheetPivotTableAfterValueChange(*": Flag = True
905:            Case sTxt Like "*_SheetPivotTableBeforeAllocateChanges(*": Flag = True
906:            Case sTxt Like "*_SheetPivotTableBeforeCommitChanges(*": Flag = True
907:            Case sTxt Like "*_SheetPivotTableBeforeDiscardChanges(*": Flag = True
908:            Case sTxt Like "*_SheetPivotTableChangeSync(*": Flag = True
909:            Case sTxt Like "*_SheetPivotTableUpdate(*": Flag = True
910:            Case sTxt Like "*_SheetSelectionChange(*": Flag = True
911:            Case sTxt Like "*_Sync(*": Flag = True
912:            Case sTxt Like "*_WindowActivate(*": Flag = True
913:            Case sTxt Like "*_WindowDeactivate(*": Flag = True
914:            Case sTxt Like "*_WindowResize(*": Flag = True
915:            Case sTxt Like "*_NewWorkbook(*": Flag = True
916:            Case sTxt Like "*_WorkbookActivate(*": Flag = True
917:            Case sTxt Like "*_WorkbookAddinInstall(*": Flag = True
918:            Case sTxt Like "*_WorkbookAddinUninstall(*": Flag = True
919:            Case sTxt Like "*_WorkbookAfterSave(*": Flag = True
920:            Case sTxt Like "*_WorkbookAfterXmlExport(*": Flag = True
921:            Case sTxt Like "*_WorkbookAfterXmlImport(*": Flag = True
922:            Case sTxt Like "*_WorkbookBeforeClose(*": Flag = True
923:            Case sTxt Like "*_WorkbookBeforePrint(*": Flag = True
924:            Case sTxt Like "*_WorkbookBeforeSave(*": Flag = True
925:            Case sTxt Like "*_WorkbookBeforeXmlExport(*": Flag = True
926:            Case sTxt Like "*_WorkbookBeforeXmlImport(*": Flag = True
927:            Case sTxt Like "*_WorkbookDeactivate(*": Flag = True
928:            Case sTxt Like "*_WorkbookModelChange(*": Flag = True
929:            Case sTxt Like "*_WorkbookNewChart(*": Flag = True
930:            Case sTxt Like "*_WorkbookNewSheet(*": Flag = True
931:            Case sTxt Like "*_WorkbookOpen(*": Flag = True
932:            Case sTxt Like "*_WorkbookPivotTableCloseConnection(*": Flag = True
933:            Case sTxt Like "*_WorkbookPivotTableOpenConnection(*": Flag = True
934:            Case sTxt Like "*_WorkbookRowsetComplete(*": Flag = True
935:            Case sTxt Like "*_WorkbookSync(*": Flag = True
936:        End Select
937:    End If
938:    WorkBookAndSheetsEvents = Flag
939: End Function

     Private Function UserFormsEvents(ByVal sTxt As String, ByVal TypeModule As VBIDE.vbext_ComponentType) As Boolean
942:    Dim Flag        As Boolean
943:    Flag = False
944:    'только для событий юзер форм и класов
945:    If TypeModule = vbext_ct_MSForm Or TypeModule = vbext_ct_ClassModule Then
946:        Select Case True
            Case sTxt Like "*_AfterUpdate(*": Flag = True
948:            Case sTxt Like "*_BeforeDragOver(*": Flag = True
949:            Case sTxt Like "*_BeforeDropOrPaste(*": Flag = True
950:            Case sTxt Like "*_BeforeUpdate(*": Flag = True
951:            Case sTxt Like "*_Change(*": Flag = True
952:            Case sTxt Like "*_Click(*": Flag = True
953:            Case sTxt Like "*_DblClick(*": Flag = True
954:            Case sTxt Like "*_Deactivate(*": Flag = True
955:            Case sTxt Like "*_DropButtonClick(*": Flag = True
956:            Case sTxt Like "*_Enter(*": Flag = True
957:            Case sTxt Like "*_Error(*": Flag = True
958:            Case sTxt Like "*_Exit(*": Flag = True
959:            Case sTxt Like "*_Initialize(*": Flag = True
960:            Case sTxt Like "*_KeyDown(*": Flag = True
961:            Case sTxt Like "*_KeyPress(*": Flag = True
962:            Case sTxt Like "*_KeyUp(*": Flag = True
963:            Case sTxt Like "*_Layout(*": Flag = True
964:            Case sTxt Like "*_MouseDown(*": Flag = True
965:            Case sTxt Like "*_MouseMove(*": Flag = True
966:            Case sTxt Like "*_MouseUp(*": Flag = True
967:            Case sTxt Like "*_QueryClose(*": Flag = True
968:            Case sTxt Like "*_RemoveControl(*": Flag = True
969:            Case sTxt Like "*_Resize(*": Flag = True
970:            Case sTxt Like "*_Scroll(*": Flag = True
971:            Case sTxt Like "*_Terminate(*": Flag = True
972:            Case sTxt Like "*_Zoom(*": Flag = True
973:        End Select
974:    End If
975:    UserFormsEvents = Flag
976: End Function

