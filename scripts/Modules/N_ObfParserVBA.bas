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
36:        .lbOK.Caption = "Parse code"
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
52:    Call MsgBox("Error in N_ObfParserVBA. Start Parser" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Error:")
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
335:    Dim lLine       As Long
336:    Dim sCode       As String
337:    Dim sVar        As String
338:    Dim sSubName    As String
339:    Dim sNumTypeName As String
340:    Dim sType       As String
341:    Dim arrStrCode  As Variant
342:    Dim arrEnum     As Variant
343:    Dim itemArr     As Variant
344:    Dim itemVar     As Variant
345:    Dim arrVar      As Variant
346:
347:    With objVBC.CodeModule
348:        lLine = .CountOfLines
349:        If lLine > 0 Then
350:            sCode = .Lines(1, lLine)
351:            If sCode <> vbNullString Then
352:                'убираю перенос строк
353:                sCode = VBA.Replace(sCode, " _" & vbNewLine, vbNullString)
354:                arrStrCode = VBA.Split(sCode, vbNewLine)
355:                For Each itemArr In arrStrCode
356:                    itemArr = C_PublicFunctions.TrimSpace(itemArr)
357:                    If itemArr <> vbNullString And VBA.Left$(itemArr, 1) <> "'" Then
358:                        sVar = vbNullString
359:                        'если есть коментарий в строке кода то удаляем его
360:                        itemArr = DeleteCommentString(itemArr)
361:                        'из строки декларирования и определение что вошли в процедуру
362:                        If (itemArr Like "* Sub *(*)*" Or itemArr Like "* Function *(*)*" Or itemArr Like "* Property Let *(*)*" Or itemArr Like "* Property Set *(*)*" Or itemArr Like "* Property Get *(*)*" Or _
                                    itemArr Like "Sub *(*)*" Or itemArr Like "Function *(*)*" Or itemArr Like "Property Let *(*)*" Or itemArr Like "Property Set *(*)*" Or itemArr Like "Property Get *(*)*") _
                                    And (Not itemArr Like "*As IRibbonControl*" And Not itemArr Like "* Declare *(*)*") Then
365:
366:                            sSubName = TypeProcedyre(VBA.CStr(itemArr))
367:                            sSubName = sSubName & CHR_TO & GetNameSubFromString(itemArr)
368:                            sVar = ParserStrDimConst(itemArr, sSubName, .Name)
369:
370:                        End If
371:                        'если в перечислении и типе данных
372:                        If itemArr Like "Private Enum *" Or itemArr Like "Public Enum *" Or itemArr Like "Enum *" Or itemArr Like "Private Type *" Or itemArr Like "Public Type *" Or itemArr Like "Type *" Then
373:                            arrEnum = VBA.Split(itemArr, " ")
374:                            If VBA.CStr(itemArr) Like "Private *" Then
375:                                sNumTypeName = "Private"
376:                            Else
377:                                sNumTypeName = "Public"
378:                            End If
379:                            sNumTypeName = arrEnum(UBound(arrEnum)) & CHR_TO & sNumTypeName
380:                            If itemArr Like "* Enum *" Or itemArr Like "Enum *" Then
381:                                sType = "Enum"
382:                            Else
383:                                sType = "Type"
384:                            End If
385:                        End If
386:                        'вышли из процедуры или перечисления
387:                        If itemArr Like "*End Sub" Or itemArr Like "*End Function" Or itemArr Like "*End Property" Or itemArr Like "*End Enum" Or itemArr Like "*End Type" Then
388:                            sSubName = vbNullString
389:                            sNumTypeName = vbNullString
390:                        End If
391:                        'если внутри типа или перечисления
392:                        If sNumTypeName <> vbNullString And Not itemArr Like "* Enum *" And Not itemArr Like "Enum *" And Not itemArr Like "* Type *" And Not itemArr Like "Type *" Then
393:                            arrEnum = VBA.Split(VBA.Trim$(itemArr), " ")
394:                            sVar = arrEnum(0)
395:                            If sVar Like "*(*" Then sVar = VBA.Left$(sVar, VBA.InStr(1, sVar, "(") - 1)
396:                            sVar = .Name & CHR_TO & sType & CHR_TO & sNumTypeName & CHR_TO & ReplaceType(sVar)
397:                        End If
398:                        'если находимся только внутри процедуры
399:                        If (itemArr Like "* Dim *" Or itemArr Like "* Const *" Or itemArr Like "Dim *" Or itemArr Like "Const *") And sSubName <> vbNullString Then
400:                            sVar = ParserStrDimConst(itemArr, sSubName, .Name)
401:                        End If
402:                        arrVar = VBA.Split(sVar, vbNewLine)
403:                        For Each itemVar In arrVar
404:                            If itemVar <> vbNullString And objDic.Exists(itemVar) = False Then
405:                                objDic.Add itemVar, objVBC.Type
406:                            End If
407:                        Next itemVar
408:                        Call ParserStringInCode(itemArr, sSubName, objVBC, objDicStr)
409:                    End If
410:                Next itemArr
411:            End If
412:        End If
413:    End With
414: End Sub

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
428:    Dim sTemp       As String
429:    sTemp = VBA.Trim$(VBA.Left$(sStrCode, VBA.InStr(1, sStrCode, "(") - 1))
430:    Select Case True
        Case sTemp Like "*Sub *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Sub ") - 3)
432:        Case sTemp Like "*Function *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Function ") - 8)
433:        Case sTemp Like "*Property Let *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Let ") - 12)
434:        Case sTemp Like "*Property Set *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Set ") - 12)
435:        Case sTemp Like "*Property Get *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Get ") - 12)
436:    End Select
437:    GetNameSubFromString = VBA.Trim$(sTemp)
438: End Function

     Private Sub ParserStringInCode(ByVal sSTR As String, ByVal sNameSub As String, ByRef objVBC As VBIDE.VBComponent, ByRef objDicStr As Scripting.Dictionary)
441:    Dim stxt        As String
442:    Dim arrStr      As Variant
443:    Dim Arr         As Variant
444:    Dim sReplace    As String
445:    Dim i           As Integer
446:    Dim sArray      As String
447:    Const CHAR_REPLACE As String = "ЪЪЪЪ"
448:
449:    sSTR = VBA.Trim$(sSTR)
450:
451:    If sSTR Like "*" & VBA.Chr$(34) & "*" And sSTR <> vbNullString Then
452:
453:        stxt = VBA.Right$(sSTR, VBA.Len(sSTR) - VBA.InStr(1, sSTR, VBA.Chr$(34)) + 1)
454:        stxt = VBA.Replace(stxt, VBA.Chr$(34) & VBA.Chr$(34), CHAR_REPLACE)
455:        arrStr = VBA.Split(stxt, VBA.Chr$(34))
456:
457:        sArray = VBA.Left$(sSTR, VBA.InStr(1, sSTR, VBA.Chr$(34)) - 1)
458:        If sArray Like "* = Array(" Then
459:            sArray = VBA.Replace(sArray, " = Array(", vbNullString)
460:            Arr = VBA.Split(sArray, " ")
461:            sArray = Arr(UBound(Arr))
462:        Else
463:            sArray = vbNullString
464:        End If
465:        For i = 1 To UBound(arrStr) Step 2
466:            If arrStr(i) <> vbNullString Then
467:                If sNameSub = vbNullString Then sNameSub = "Declaration" & CHR_TO
468:
469:                sReplace = VBA.Replace(arrStr(i), CHAR_REPLACE, VBA.Chr$(34) & VBA.Chr$(34))
470:                stxt = objVBC.Name & CHR_TO & sNameSub & CHR_TO & VBA.Chr$(34) & sReplace & VBA.Chr$(34) & CHR_TO & sArray    '& CHR_TO & sYesNo
471:                If arrStr(i + 1) Like "*: * = *" Then sArray = vbNullString
472:                If arrStr(i + 1) Like "*: * = Array(*" Then
473:                    sArray = VBA.Replace(arrStr(i + 1), ": ", vbNullString)
474:                    sArray = VBA.Replace(sArray, " = Array(", vbNullString)
475:                    sArray = VBA.Replace(sArray, ")", vbNullString)
476:                End If
477:                If objDicStr.Exists(stxt) = False Then objDicStr.Add stxt, objVBC.Type
478:            End If
479:        Next i
480:        sArray = vbNullString
481:    End If
482: End Sub

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
     Private Function ParserStrDimConst(ByVal stxt As String, ByVal sNameSub As String, ByVal sNameMod As String) As String
496:    Dim sTemp       As String
497:    Dim sWord       As String
498:    Dim sWordTemp   As String
499:    Dim arrStr      As Variant
500:    Dim itemArr     As Variant
501:    Dim arrWord     As Variant
502:    Dim sType       As String
503:
504:    sTemp = C_PublicFunctions.TrimSpace(stxt)
505:    sType = "Dim"
506:    If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
507:        'если есть коментарий в строке кода то удаляем его
508:        sTemp = DeleteCommentString(sTemp)
509:        If sTemp Like "*Sub *(*)*" Or sTemp Like "*Function *(*)*" Or sTemp Like "*Property Let *(*)*" Or sTemp Like "*Property Set *(*)*" Or sTemp Like "*Property Get *(*)*" Then
510:            If VBA.InStr(1, sTemp, ")") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, ")") - 1)
511:            If VBA.InStr(1, sTemp, " = ") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ") - 1)
512:            If VBA.Len(sTemp) - VBA.InStr(1, sTemp, "(") >= 0 Then
513:                sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "("))
514:            End If
515:        ElseIf sTemp Like "* Dim *" Or sTemp Like Chr$(68) & "im *" Then
516:            sType = "Dim"
517:            If VBA.InStr(1, sTemp, "Dim ") >= 3 Then sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Dim ") - 3)
518:        ElseIf sTemp Like "* Const *" Or sTemp Like Chr$(67) & "onst *" Then
519:            sType = "Const"
520:            If VBA.InStr(1, sTemp, "Const ") >= 5 Then sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Const ") - 5)
521:            If VBA.InStr(1, sTemp, " = ") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ") - 1)
522:        Else
523:            sTemp = vbNullString
524:        End If
525:    End If
526:
527:    If sTemp Like "*: *" Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, ": ") - 1)
528:    If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
529:        arrStr = VBA.Split(sTemp, ",")
530:        For Each itemArr In arrStr
531:            If itemArr Like "*(*" Then itemArr = VBA.Left$(itemArr, VBA.InStr(1, itemArr, "(") - 1)
532:            If Not itemArr Like "*)*" And Not itemArr Like "* To *" Then
533:                arrWord = VBA.Split(itemArr, " As ")
534:                arrWord = VBA.Split(VBA.Trim$(arrWord(0)), " ")
535:                If UBound(arrWord) = -1 Then
536:                    sWord = vbNullString
537:                Else
538:                    sWordTemp = VBA.Trim$(arrWord(UBound(arrWord)))
539:                    sWordTemp = ReplaceType(sWordTemp)
540:                    sWord = sWord & vbNewLine & sNameMod & CHR_TO & sNameSub & CHR_TO & sType & CHR_TO & sWordTemp
541:                End If
542:            End If
543:        Next itemArr
544:    End If
545:    sWord = VBA.Trim$(sWord)
546:    If VBA.Len(sWord) = 0 Then
547:        sWord = vbNullString
548:    Else
549:        sWord = VBA.Trim$(VBA.Right$(sWord, VBA.Len(sWord) - 2))
550:    End If
551:    ParserStrDimConst = sWord
552: End Function

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
566:    Dim ProcKind    As VBIDE.vbext_ProcKind
567:    Dim lLine       As Long
568:    Dim lineOld     As Long
569:    Dim sNameSub    As String
570:    Dim strFunctionBody As String
571:    With objVBC.CodeModule
572:        If .CountOfLines > 0 Then
573:            lLine = .CountOfDeclarationLines
574:            If lLine = 0 Then lLine = 2
575:            Do Until lLine >= .CountOfLines
576:
577:                'сбор названий процедур и функций
578:                sNameSub = .ProcOfLine(lLine, ProcKind)
579:                If sNameSub <> vbNullString Then
580:                    strFunctionBody = C_PublicFunctions.TrimSpace(.Lines(lLine - 1, .ProcCountLines(sNameSub, ProcKind)))
581:                    If (Not strFunctionBody Like "*As IRibbonControl*") And _
                                (Not WorkBookAndSheetsEvents(strFunctionBody, objVBC.Type)) And _
                                (Not (strFunctionBody Like "* UserForm_*" And objVBC.Type = vbext_ct_MSForm)) And _
                                (Not UserFormsEvents(strFunctionBody, objVBC.Type)) Then
585:                        Dim sKey As String
586:                        sKey = sNameVBC & CHR_TO & TypeProcedyre(strFunctionBody) & CHR_TO & TypeOfAccessModifier(strFunctionBody) & CHR_TO & sNameSub
587:                        If Not varSubFun.Exists(sKey) Then
588:                            varSubFun.Add sKey, objVBC.Type
589:                        End If
590:                    End If
591:                    lLine = .ProcStartLine(sNameSub, ProcKind) + .ProcCountLines(sNameSub, ProcKind) + 1
592:                Else
593:                    lLine = lLine + 1
594:                End If
595:                If lineOld > lLine Then Exit Do
596:                lineOld = lLine
597:            Loop
598:        End If
599:    End With
600: End Sub

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
614:    Dim objCont     As MSForms.control
615:    If Not objVBC.Designer Is Nothing Then
616:        With objVBC.Designer
617:            For Each objCont In .Controls
618:                obfNewDict.Add sNameVBC & CHR_TO & objCont.Name, objVBC.Type
619:            Next objCont
620:        End With
621:    End If
622: End Sub

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
636:    Dim varArr      As Variant
637:    Dim varArrWord  As Variant
638:    Dim varStr      As Variant
639:    Dim itemVarStr  As Variant
640:    Dim varAPI      As Variant
641:    Dim sTemp       As String
642:    Dim sTempArr    As String
643:    Dim i           As Long
644:    Dim bFlag       As Boolean
645:    Dim j           As Byte
646:    Dim itemArr     As Byte
647:    bFlag = True
648:    If objVBC.CodeModule.CountOfDeclarationLines <> 0 Then
649:        sTemp = objVBC.CodeModule.Lines(1, objVBC.CodeModule.CountOfDeclarationLines)
650:        sTemp = VBA.Replace(sTemp, " _" & vbNewLine, vbNullString)
651:        If sTemp <> vbNullString Then
652:            varArr = VBA.Split(sTemp, vbNewLine)
653:            For i = 0 To UBound(varArr)
654:                sTemp = C_PublicFunctions.TrimSpace(DeleteCommentString(varArr(i)))
655:                If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
656:                    If sTemp Like "* Type *" Or sTemp Like "* Enum *" Or sTemp Like "Type *" Or sTemp Like "Enum *" Then
657:                        varArrWord = VBA.Split(sTemp, " ")
658:                        If UBound(varArrWord) = 2 Then
659:                            sTemp = VBA.Trim$(varArrWord(0)) & CHR_TO & VBA.Trim$(varArrWord(1)) & CHR_TO & VBA.Trim$(varArrWord(2))
660:                        ElseIf UBound(varArrWord) = 1 Then
661:                            sTemp = "Public" & CHR_TO & VBA.Trim$(varArrWord(0)) & CHR_TO & VBA.Trim$(varArrWord(1))
662:                        End If
663:                        sTemp = sNameVBC & CHR_TO & sTemp
664:                        If Not dicTypeEnum.Exists(sTemp) Then dicTypeEnum.Add sTemp, objVBC.Type
665:                        bFlag = False
666:                    End If
667:                    If bFlag And Not (sTemp Like "Implements *" Or sTemp Like "Option *" Or VBA.Left$(sTemp, 1) = "'" Or sTemp = vbNullString Or VBA.Left$(sTemp, 1) = "#" Or sTemp Like "*Declare *(*)*") Then
668:
669:                        If sTemp Like "* = *" Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ", vbTextCompare) + 2)
670:                        If sTemp Like "* *(* To *) *" Then
671:                            sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, "(", vbTextCompare) - 1)
672:                        End If
673:                        varStr = VBA.Split(sTemp, ",")
674:                        For Each itemVarStr In varStr
675:                            sTemp = VBA.Trim$(itemVarStr)
676:                            varArrWord = VBA.Split(sTemp, " As ")
677:                            varArrWord = VBA.Split(varArrWord(0), " = ")
678:                            sTemp = varArrWord(0)
679:                            varArrWord = VBA.Split(sTemp, " ")
680:
681:                            j = UBound(varArrWord)
682:                            If j > 1 Then
683:                                If varArrWord(0) = "Dim" Or varArrWord(0) = "Const" Then
684:                                    sTemp = "Private" & CHR_TO & varArrWord(0) & CHR_TO
685:                                    sTempArr = varArrWord(1)
686:                                ElseIf (varArrWord(0) = "Private" Or varArrWord(0) = "Public") And (varArrWord(1) = "Dim" Or varArrWord(1) = "Const" Or varArrWord(1) = "WithEvents") Then
687:                                    sTemp = varArrWord(0) & CHR_TO & varArrWord(1) & CHR_TO
688:                                    sTempArr = varArrWord(2)
689:                                ElseIf (varArrWord(0) = "Private" Or varArrWord(0) = "Public") And Not (varArrWord(1) = "Dim" Or varArrWord(1) = "Const" Or varArrWord(1) = "WithEvents") Then
690:                                    sTemp = varArrWord(0) & CHR_TO & "Dim" & CHR_TO
691:                                    sTempArr = varArrWord(1)
692:                                End If
693:                            ElseIf j = 1 And varArrWord(0) = "Global" Then
694:                                sTemp = "Public" & CHR_TO & varArrWord(0) & CHR_TO
695:                                sTempArr = varArrWord(1)
696:                            ElseIf j = 1 And (varArrWord(0) = "Private" Or varArrWord(0) = "Public") Then
697:                                sTemp = varArrWord(0) & CHR_TO & "Dim" & CHR_TO
698:                                sTempArr = varArrWord(1)
699:                            ElseIf j = 1 And (varArrWord(0) = "Dim" Or varArrWord(0) = "Const") Then
700:                                sTemp = "Private" & CHR_TO & varArrWord(0) & CHR_TO
701:                                sTempArr = varArrWord(1)
702:                            ElseIf j = 0 Then
703:                                sTemp = "Private" & CHR_TO & " Dim" & CHR_TO
704:                                sTempArr = varArrWord(0)
705:                            End If
706:
707:                            sTempArr = ReplaceType(sTempArr)
708:                            If sTempArr Like "*(*" Then sTempArr = VBA.Left$(sTempArr, VBA.InStr(1, sTempArr, "(") - 1)
709:                            sTemp = sNameVBC & CHR_TO & sTemp & sTempArr
710:                            If Not dicGloblVar.Exists(sTemp) Then dicGloblVar.Add sTemp, objVBC.Type
711:
712:                            sTemp = vbNullString
713:                        Next itemVarStr
714:                        sTemp = vbNullString
715:                    End If
716:                    If sTemp Like "*End Type" Or sTemp Like "*End Enum" Then
717:                        bFlag = True
718:                    End If
719:                    If sTemp Like "*Declare * Lib " & VBA.Chr$(34) & "*" & VBA.Chr$(34) & " (*)*" Then
720:                        sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " Lib ", vbTextCompare) - 1)
721:                        varAPI = VBA.Split(sTemp, VBA.Chr$(32))
722:                        itemArr = UBound(varAPI)
723:                        sTemp = CHR_TO & varAPI(itemArr - 1) & CHR_TO & varAPI(itemArr)
724:                        If varAPI(1) = "Declare" Then
725:                            sTemp = sNameVBC & CHR_TO & varAPI(0) & sTemp
726:                        Else
727:                            sTemp = sNameVBC & CHR_TO & "Private" & sTemp
728:                        End If
729:                        If Not dicAPI.Exists(sTemp) Then dicAPI.Add sTemp, objVBC.Type
730:                    End If
731:                End If
732:            Next i
733:        End If
734:    End If
735: End Sub
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
748:    Set objDict = Nothing
749:    Set objDict = New Scripting.Dictionary
750:    Set AddNewDictionary = objDict
751: End Function

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
765:    'есть '
766:    Dim sTemp       As String
767:    sTemp = sWord
768:    If VBA.InStr(1, sTemp, "'") <> 0 Then
769:        If VBA.InStr(1, sTemp, VBA.Chr(34)) <> 0 Then
770:            'есть "
771:            If VBA.InStr(1, sTemp, "'") < VBA.InStr(1, sTemp, VBA.Chr(34)) Then
772:                'если так -> '"
773:                sTemp = VBA.Trim$(VBA.Left$(sTemp, VBA.InStr(1, sTemp, "'") - 1))
774:            End If
775:        Else
776:            'нет " -> '
777:            sTemp = VBA.Trim$(VBA.Left$(sTemp, VBA.InStr(1, sTemp, "'") - 1))
778:        End If
779:    End If
780:    DeleteCommentString = sTemp
781: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddEncodeName - функция генерации случайного зашифрованного имени
'* Created    : 27-03-2020 13:22
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function AddEncodeName() As String
790:    Const CharCount As Integer = 20
791:    Dim i           As Integer
792:    Dim sName       As String
793:
794:    Const FIRST_CODE_SIGN As String = "1"
795:    Const SECOND_CODE_SIGN As String = "0"
tryAgain:
797:    Err.Clear
798:    sName = vbNullString
799:    Randomize
800:    sName = "o"
801:    For i = 2 To CharCount
802:        If (VBA.Round(VBA.Rnd() * 1000)) Mod 2 = 1 Then sName = sName & FIRST_CODE_SIGN Else sName = sName & SECOND_CODE_SIGN
803:    Next i
804:    On Error Resume Next
805:    'добовляем новое имя в коллекцию, если имя существует то
806:    'генерируется ошибка, запуск повторной генерации имени
807:    objCollUnical.Add sName, sName
808:    If Err.Number <> 0 Then GoTo tryAgain
809:    AddEncodeName = sName
810: End Function
'взято
     Private Function TypeOfAccessModifier(ByRef StrDeclarationProcedure As String) As String
813:    If StrDeclarationProcedure Like "*Private *(*)*" Then
814:        TypeOfAccessModifier = "Private"
815:    Else
816:        TypeOfAccessModifier = "Public"
817:    End If
818: End Function
     Private Function TypeProcedyre(ByRef StrDeclarationProcedure As String) As String
820:    If StrDeclarationProcedure Like "*Sub *" Then
821:        TypeProcedyre = "Sub"
822:    ElseIf StrDeclarationProcedure Like "*Function *" Then
823:        TypeProcedyre = "Function"
824:    ElseIf StrDeclarationProcedure Like "*Property Set *" Then
825:        TypeProcedyre = "Property Set"
826:    ElseIf StrDeclarationProcedure Like "*Property Get *" Then
827:        TypeProcedyre = "Property Get"
828:    ElseIf StrDeclarationProcedure Like "*Property Let *" Then
829:        TypeProcedyre = "Property Let"
830:    Else
831:        TypeProcedyre = "Unknown Type"
832:    End If
833: End Function

     Private Function ReplaceType(ByVal sVar As String) As String
836:    sVar = Replace(sVar, "%", vbNullString)     'Integer
837:    sVar = Replace(sVar, "&", vbNullString)     'Long
838:    sVar = Replace(sVar, "$", vbNullString)     'String
839:    sVar = Replace(sVar, "!", vbNullString)     'Single
840:    sVar = Replace(sVar, "#", vbNullString)     'Double
841:    sVar = Replace(sVar, "@", vbNullString)     'Currency
842:    ReplaceType = sVar
843: End Function

     Private Function WorkBookAndSheetsEvents(ByVal stxt As String, ByVal TypeModule As VBIDE.vbext_ComponentType) As Boolean
846:    Dim Flag        As Boolean
847:    Flag = False
848:    'только для модулей листов, книг и класов
849:    If TypeModule = vbext_ct_Document Or TypeModule = vbext_ct_ClassModule Then
850:        Select Case True
            Case stxt Like "*_Activate(*": Flag = True
852:            Case stxt Like "*_AddinInstall(*": Flag = True
853:            Case stxt Like "*_AddinUninstall(*": Flag = True
854:            Case stxt Like "*_AfterSave(*": Flag = True
855:            Case stxt Like "*_AfterXmlExport(*": Flag = True
856:            Case stxt Like "*_AfterXmlImport(*": Flag = True
857:            Case stxt Like "*_BeforeClose(*": Flag = True
858:            Case stxt Like "*_BeforeDoubleClick(*": Flag = True
859:            Case stxt Like "*_BeforePrint(*": Flag = True
860:            Case stxt Like "*_BeforeRightClick(*": Flag = True
861:            Case stxt Like "*_BeforeSave(*": Flag = True
862:            Case stxt Like "*_BeforeXmlExport(*": Flag = True
863:            Case stxt Like "*_BeforeXmlImport(*": Flag = True
864:            Case stxt Like "*_Calculate(*": Flag = True
865:            Case stxt Like "*_Change(*": Flag = True
866:            Case stxt Like "*_Deactivate(*": Flag = True
867:            Case stxt Like "*_FollowHyperlink(*": Flag = True
868:            Case stxt Like "*_MouseDown(*": Flag = True
869:            Case stxt Like "*_MouseMove(*": Flag = True
870:            Case stxt Like "*_MouseUp(*": Flag = True
871:            Case stxt Like "*_NewChart(*": Flag = True
872:            Case stxt Like "*_NewSheet(*": Flag = True
873:            Case stxt Like "*_Open(*": Flag = True
874:            Case stxt Like "*_PivotTableAfterValueChange(*": Flag = True
875:            Case stxt Like "*_PivotTableBeforeAllocateChanges(*": Flag = True
876:            Case stxt Like "*_PivotTableBeforeCommitChanges(*": Flag = True
877:            Case stxt Like "*_PivotTableBeforeDiscardChanges(*": Flag = True
878:            Case stxt Like "*_PivotTableChangeSync(*": Flag = True
879:            Case stxt Like "*_PivotTableCloseConnection(*": Flag = True
880:            Case stxt Like "*_PivotTableOpenConnection(*": Flag = True
881:            Case stxt Like "*_PivotTableUpdate(*": Flag = True
882:            Case stxt Like "*_Resize(*": Flag = True
883:            Case stxt Like "*_RowsetComplete(*": Flag = True
884:            Case stxt Like "*_SelectionChange(*": Flag = True
885:            Case stxt Like "*_SeriesChange(*": Flag = True
886:            Case stxt Like "*_SheetActivate(*": Flag = True
887:            Case stxt Like "*_SheetBeforeDoubleClick(*": Flag = True
888:            Case stxt Like "*_SheetBeforeRightClick(*": Flag = True
889:            Case stxt Like "*_SheetCalculate(*": Flag = True
890:            Case stxt Like "*_SheetChange(*": Flag = True
891:            Case stxt Like "*_SheetDeactivate(*": Flag = True
892:            Case stxt Like "*_SheetFollowHyperlink(*": Flag = True
893:            Case stxt Like "*_SheetPivotTableAfterValueChange(*": Flag = True
894:            Case stxt Like "*_SheetPivotTableBeforeAllocateChanges(*": Flag = True
895:            Case stxt Like "*_SheetPivotTableBeforeCommitChanges(*": Flag = True
896:            Case stxt Like "*_SheetPivotTableBeforeDiscardChanges(*": Flag = True
897:            Case stxt Like "*_SheetPivotTableChangeSync(*": Flag = True
898:            Case stxt Like "*_SheetPivotTableUpdate(*": Flag = True
899:            Case stxt Like "*_SheetSelectionChange(*": Flag = True
900:            Case stxt Like "*_Sync(*": Flag = True
901:            Case stxt Like "*_WindowActivate(*": Flag = True
902:            Case stxt Like "*_WindowDeactivate(*": Flag = True
903:            Case stxt Like "*_WindowResize(*": Flag = True
904:            Case stxt Like "*_NewWorkbook(*": Flag = True
905:            Case stxt Like "*_WorkbookActivate(*": Flag = True
906:            Case stxt Like "*_WorkbookAddinInstall(*": Flag = True
907:            Case stxt Like "*_WorkbookAddinUninstall(*": Flag = True
908:            Case stxt Like "*_WorkbookAfterSave(*": Flag = True
909:            Case stxt Like "*_WorkbookAfterXmlExport(*": Flag = True
910:            Case stxt Like "*_WorkbookAfterXmlImport(*": Flag = True
911:            Case stxt Like "*_WorkbookBeforeClose(*": Flag = True
912:            Case stxt Like "*_WorkbookBeforePrint(*": Flag = True
913:            Case stxt Like "*_WorkbookBeforeSave(*": Flag = True
914:            Case stxt Like "*_WorkbookBeforeXmlExport(*": Flag = True
915:            Case stxt Like "*_WorkbookBeforeXmlImport(*": Flag = True
916:            Case stxt Like "*_WorkbookDeactivate(*": Flag = True
917:            Case stxt Like "*_WorkbookModelChange(*": Flag = True
918:            Case stxt Like "*_WorkbookNewChart(*": Flag = True
919:            Case stxt Like "*_WorkbookNewSheet(*": Flag = True
920:            Case stxt Like "*_WorkbookOpen(*": Flag = True
921:            Case stxt Like "*_WorkbookPivotTableCloseConnection(*": Flag = True
922:            Case stxt Like "*_WorkbookPivotTableOpenConnection(*": Flag = True
923:            Case stxt Like "*_WorkbookRowsetComplete(*": Flag = True
924:            Case stxt Like "*_WorkbookSync(*": Flag = True
925:        End Select
926:    End If
927:    WorkBookAndSheetsEvents = Flag
928: End Function

Private Function UserFormsEvents(ByVal stxt As String, ByVal TypeModule As VBIDE.vbext_ComponentType) As Boolean
931:    Dim Flag        As Boolean
932:    Flag = False
933:    'только для событий юзер форм и класов
934:    If TypeModule = vbext_ct_MSForm Or TypeModule = vbext_ct_ClassModule Then
935:        Select Case True
            Case stxt Like "*_AfterUpdate(*": Flag = True
937:            Case stxt Like "*_BeforeDragOver(*": Flag = True
938:            Case stxt Like "*_BeforeDropOrPaste(*": Flag = True
939:            Case stxt Like "*_BeforeUpdate(*": Flag = True
940:            Case stxt Like "*_Change(*": Flag = True
941:            Case stxt Like "*_Click(*": Flag = True
942:            Case stxt Like "*_DblClick(*": Flag = True
943:            Case stxt Like "*_Deactivate(*": Flag = True
944:            Case stxt Like "*_DropButtonClick(*": Flag = True
945:            Case stxt Like "*_Enter(*": Flag = True
946:            Case stxt Like "*_Error(*": Flag = True
947:            Case stxt Like "*_Exit(*": Flag = True
948:            Case stxt Like "*_Initialize(*": Flag = True
949:            Case stxt Like "*_KeyDown(*": Flag = True
950:            Case stxt Like "*_KeyPress(*": Flag = True
951:            Case stxt Like "*_KeyUp(*": Flag = True
952:            Case stxt Like "*_Layout(*": Flag = True
953:            Case stxt Like "*_MouseDown(*": Flag = True
954:            Case stxt Like "*_MouseMove(*": Flag = True
955:            Case stxt Like "*_MouseUp(*": Flag = True
956:            Case stxt Like "*_QueryClose(*": Flag = True
957:            Case stxt Like "*_RemoveControl(*": Flag = True
958:            Case stxt Like "*_Resize(*": Flag = True
959:            Case stxt Like "*_Scroll(*": Flag = True
960:            Case stxt Like "*_Terminate(*": Flag = True
961:            Case stxt Like "*_Zoom(*": Flag = True
962:        End Select
963:    End If
964:    UserFormsEvents = Flag
End Function

