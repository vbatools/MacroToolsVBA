VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HiddenModule 
   Caption         =   "Скрыть модули VBA:"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8415
   OleObjectBlob   =   "HiddenModule.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HiddenModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : HiddenModule - скрытие модулей VBA
'* Created    : 12-02-2020 10:19
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

    Private Sub cmbCancel_Click()
11:    Unload Me
12: End Sub

    Private Sub cmbMain_Change()
15:    Call AddListCode
16: End Sub

    Private Sub lbCancel_Click()
19:    Call cmbCancel_Click
20: End Sub
    Private Sub CheckAll_Click()
22:    Dim i      As Integer
23:    With ListCode
24:        For i = 0 To .ListCount - 1
25:            .Selected(i) = CheckAll.Value
26:        Next i
27:    End With
28: End Sub

    Private Sub ListCode_Change()
31:    Dim i      As Integer
32:    With ListCode
33:        For i = 0 To .ListCount - 1
34:            If .Selected(i) Then
35:                lbMsg.visible = False
36:                lbOK.Enabled = True
37:                Call MsgSaveFile(cmbMain.Value)
38:                Exit Sub
39:            End If
40:        Next i
41:    End With
42:
43:    lbMsg.visible = True
44:    lbOK.Enabled = False
45: End Sub
    Private Sub UserForm_Activate()
47:    Me.StartUpPosition = 0
48:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
49:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
50:
51:    Dim WB     As Workbook
52:    On Error GoTo ErrorHandler
53:    If Workbooks.Count = 0 Then
54:        Unload Me
55:        Call MsgBox("Нет открытых " & Chr(34) & "Файлов Excel" & Chr(34) & "!", vbOKOnly + vbExclamation, "Ошибка:")
56:        Exit Sub
57:    End If
58:    With Me.cmbMain
59:        .Clear
60:        For Each WB In Workbooks
61:            .AddItem WB.Name
62:        Next
63:        .Value = ActiveWorkbook.Name
64:        Call MsgSaveFile(.Value)
65:    End With
66:    Call AddListCode
67:    lbOK.Enabled = False
68:
69:    Exit Sub
ErrorHandler:
71:    Unload Me
72:    Select Case Err.Number
        Case Else:
74:            Call MsgBox("Ошибка! в HiddenModule.UserForm_Activate" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbOKOnly + vbExclamation, "Ошибка:")
75:            Call WriteErrorLog("HiddenModule.UserForm_Activate")
76:    End Select
77:    Err.Clear
78: End Sub

    Private Sub MsgSaveFile(ByVal WBName As String)
81:    Dim WB     As Workbook
82:    Set WB = Workbooks(WBName)
83:    With WB
84:        If .Path = vbNullString Then
85:            lbSave.visible = True
86:            lbOK.Enabled = False
87:        Else
88:            lbSave.visible = False
89:            lbOK.Enabled = True
90:        End If
91:    End With
92: End Sub

     Private Sub AddListCode()
95:    Dim WB     As Workbook
96:    Dim iFile  As Integer
97:    Dim i      As Integer
98:    Set WB = Workbooks(cmbMain.Value)
99:    With ListCode
100:        .Clear
101:        If WB.VBProject.Protection <> vbext_pp_none Then
102:            Call MsgBox("VBA проект в книге - " & cmbMain.Value & " защищен, паролем!" & vbCrLf & "Снимите пароль!", vbCritical, "Ошибка:")
103:            Exit Sub
104:        End If
105:        For iFile = 1 To WB.VBProject.VBComponents.Count
106:            If WB.VBProject.VBComponents(iFile).Type = vbext_ct_StdModule Then
107:                .AddItem i + 1
108:                .List(i, 1) = WB.VBProject.VBComponents(iFile).Name
109:                i = i + 1
110:            End If
111:        Next iFile
112:    End With
113: End Sub
     Private Sub lbOK_Click()
115:    Dim WB     As Workbook
116:    Dim i      As Integer
117:    Dim strBinFile As String
118:    Dim strFulPathWB As String
119:    Dim strNameModules As String
120:    Dim strNameModulesAndChars As String
121:    Dim strNewNameFile As String
122:
123:    On Error GoTo ErrorHandler
124:
125:    Set WB = Workbooks(cmbMain.Value)
126:
127:    Application.ScreenUpdating = False
128:    Application.EnableEvents = False
129:    Application.DisplayAlerts = False
130:
131:    strNewNameFile = WB.Path & Application.PathSeparator & VBA.Split(WB.Name, ".")(0) & "_hidden"
132:    WB.SaveAs strNewNameFile
133:
134:    'создаю строку всех выделенных модулей и переименовываю
135:    For i = 0 To Me.ListCode.ListCount - 1
136:        If Me.ListCode.Selected(i) = True Then
137:            strNameModules = strNameModules & "||" & Me.ListCode.List(i, 1)
138:            strNameModulesAndChars = strNameModulesAndChars & "||" & VBA.Chr(10) & Me.ListCode.List(i, 1) & "="
139:        End If
140:    Next i
141:
142:    'создаю мусорные модули
143:    If chbAddModule Then
144:        For i = 1 To 5
145:            With WB.VBProject.VBComponents.Add(vbext_ct_StdModule)
146:                strNameModules = strNameModules & "||" & .Name
147:                strNameModulesAndChars = strNameModulesAndChars & "||" & VBA.Chr(10) & .Name & "="
148:            End With
149:        Next i
150:    End If
151:
152:    strNameModules = strNameModules & "||@@" & strNameModulesAndChars & "||"
153:
154:    strFulPathWB = WB.FullName
155:    WB.Save
156:    WB.Close
157:
158:    strBinFile = O_XML.OpenAndCloseExcelFile(bOpenFile:=True, bBackUp:=False, bShowMsg:=False, sFilePath:=strFulPathWB)
159:    Call WriteBinFileHidden(strBinFile, strNameModules)
160:    Call O_XML.OpenAndCloseExcelFile(bOpenFile:=False, bBackUp:=False, bShowMsg:=False, sFilePath:=strFulPathWB)
161:    Workbooks.Open strFulPathWB
162:
163:    Application.EnableEvents = True
164:    Application.DisplayAlerts = True
165:    Application.ScreenUpdating = True
166:    Unload Me
167:    Call MsgBox("Модули VBA скрыты!", vbInformation, "Скрыть модули VBA:")
168:
169:    Exit Sub
ErrorHandler:
171:    Unload Me
172:    Select Case Err.Number
        Case Else:
174:            Call MsgBox("Ошибка! в HiddenModule.lbOK_Click" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbOKOnly + vbExclamation, "Ошибка:")
175:            Call WriteErrorLog("HiddenModule.lbOK_Click")
176:    End Select
177:    Err.Clear
178:    Application.EnableEvents = True
179:    Application.DisplayAlerts = True
180:    Application.ScreenUpdating = True
181: End Sub

     Private Sub WriteBinFileHidden(ByVal WBName As String, ByVal strNameModule As String)
184:
185:    Dim strPath As String
186:    Dim i      As Long
187:    Dim j      As Long
188:    Dim ArrBin() As Byte
189:    Dim ArrBinNew() As Byte
190:    Dim ArrBinNewConst() As Byte
191:
192:    On Error GoTo ErrorHandler
193:
194:    'считывание бинарного файла
195:    strPath = WBName & Application.PathSeparator & "xl\vbaProject.bin"
196:    ArrBin = ByteArrayFromFile(strPath)
197:
198:    'нарезаю масив байт по строкам
199:    Dim arrVarColum() As Variant
200:    Dim arrVarRow() As Byte
201:    Dim k      As Long
202:    Dim bFlag  As Boolean
203:    Dim bFlag1 As Boolean
204:    j = 0: k = 0: bFlag1 = True
205:    For i = 0 To UBound(ArrBin) - 1
206:        'не изменая часть
207:        If i > 6 And bFlag1 Then
208:            If ArrBin(i) = 10 And ArrBin(i + 1) = 77 And ArrBin(i + 2) = 111 And ArrBin(i + 3) = 100 And ArrBin(i + 4) = 117 And ArrBin(i + 5) = 108 And ArrBin(i + 6) = 101 Then
209:                bFlag1 = False
210:            End If
211:        End If
212:        If bFlag1 Then
213:            ReDim Preserve ArrBinNewConst(0 To i)
214:            ArrBinNewConst(i) = ArrBin(i)
215:        Else
216:            'конец не изменая часть
217:            'нарезка
218:            ReDim Preserve arrVarRow(0 To j)
219:            arrVarRow(j) = ArrBin(i)
220:            j = j + 1
221:            bFlag = False
222:            If ArrBin(i) = 13 And ArrBin(i + 1) = 10 Then
223:                j = 0
224:                bFlag = True
225:            End If
226:            ReDim Preserve arrVarColum(0 To k)
227:            arrVarColum(k) = arrVarRow
228:            If bFlag Then
229:                ReDim arrVarRow(0 To 0)
230:                k = k + 1
231:            End If
232:            'конец нарезки
233:        End If
234:    Next i
235:
236:    ArrBinNew = ArrBinNewConst
237:
238:    j = UBound(ArrBinNew) + 1
239:    Dim ByteTemp() As Byte
240:    Dim strByteName As String
241:    Dim strByteNameAndChars As String
242:    Const sMODULE = "10||77||111||100||117||108||101||61||"
243:    Const sMODULE1 = "124||124||"
244:    Dim sName  As String
245:    Dim sByteTemp As String
246:
247:    strByteName = GetByteStringFromString(VBA.Split(strNameModule, "@@")(0)) & "||" & sMODULE1
248:    strByteNameAndChars = GetByteStringFromString(VBA.Split(strNameModule, "@@")(1))
249:
250:    For i = 0 To UBound(arrVarColum)
251:        ByteTemp = arrVarColum(i)
252:        sByteTemp = GetStringFromByte(ByteTemp)
253:
254:        If sByteTemp Like sMODULE & "*" Then
255:            sName = VBA.Right$(sByteTemp, Len(sByteTemp) - Len(sMODULE))
256:            sName = sMODULE1 & VBA.Left$(sName, VBA.Len(sName) - 2) & sMODULE1
257:            If strByteName Like "*" & sName & "*" Then
258:                For k = 0 To UBound(ByteTemp)
259:                    ByteTemp(k) = 0
260:                Next k
261:            End If
262:        End If
263:
264:        If sByteTemp Like "10||*||61*" Then
265:            If strByteNameAndChars Like "*" & Left(sByteTemp, InStr(1, sByteTemp, "||61") + 3) & "*" Then
266:                bFlag = True
267:                For k = 0 To UBound(ByteTemp)
268:                    If ByteTemp(k) = 61 Then bFlag = False
269:                    If bFlag Then
270:                        ByteTemp(k) = (ByteTemp(k) + VBA.Fix(VBA.Rnd(10) * 10)) Mod 256
271:                    End If
272:                Next k
273:            End If
274:        End If
275:
276:        For k = 0 To UBound(ByteTemp)
277:            ReDim Preserve ArrBinNew(0 To j)
278:            ArrBinNew(j) = ByteTemp(k)
279:            j = j + 1
280:        Next k
281:    Next i
282:
283:    'ArrBinNew = ReplaceByteFromArrayNameModule(ArrBinNew, VBA.Split(strNameModule, "@@")(0))
284:
285:    'удаление и создание файла пустого
286:    Call Kill(strPath)
287:    Call ByteFileAdd(strPath)
288:
289:    'загрузка массива байтов в файл
290:    Call ByteArrayToFile(strPath, ArrBinNew)
291:
292:    Exit Sub
ErrorHandler:
294:    Unload Me
295:    Select Case Err.Number
        Case Else:
297:            Call MsgBox("Ошибка! в HiddenModule.WriteBinFileHidden" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbOKOnly + vbExclamation, "Ошибка:")
298:            Call WriteErrorLog("HiddenModule.WriteBinFileHidden")
299:    End Select
300:    Err.Clear
301: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ReplaceByteFromArrayNameModule - побитное изменение названий модулей
'* Created    : 15-02-2020 14:00
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                     Description
'*
'* ByRef arrByte() As Byte        : - исходный массив байт
'* ByVal strNameModule As String  : - строка, названий модулей
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function ReplaceByteFromArrayNameModule(ByRef arrByte() As Byte, ByVal strNameModule As String) As Byte()
316:    Dim i As Long
317:    Dim j As Long
318:    Dim k As Long
319:    Dim m As Integer
320:    Dim byChr As Byte
321:    Dim arrByteNew() As Byte
322:    Dim arrByteName() As Byte
323:    Dim arrVar As Variant
324:    arrVar = VBA.Split(strNameModule, "||")
325:    arrByteNew = arrByte
326:    For i = 1 To UBound(arrVar) - 1
327:        k = VBA.InStrB(1, arrByteNew, GetByteFromString(arrVar(i)), vbBinaryCompare) - 1
328:        If k <> -1 Then
329:            For j = k To k + Len(arrVar(i)) - 1
330:                byChr = (arrByteNew(j) + VBA.Fix(VBA.Rnd(10) * 10)) Mod 256
331:                arrByteNew(j) = byChr
332:                m = m + 1
333:                arrByteNew(j + Len(arrVar(i)) + m) = byChr
334:            Next j
335:        End If
336:    Next i
337:    ReplaceByteFromArrayNameModule = arrByteNew
338: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetByteFromString - перевод строки в массив байт
'* Created    : 15-02-2020 12:45
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal strTxt As String : - стрка
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function GetByteFromString(ByVal strTxt As String) As Byte()
352:    GetByteFromString = StrConv(strTxt & VBA.Chr(0), vbFromUnicode) & strTxt
353: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetByteStringFromString - получение из строки, строку байтов с разделителем || ("Modu" -> "77||111||100||117")
'* Created    : 14-02-2020 08:43
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal strTxt As String : - строка
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function GetByteStringFromString(ByVal strTxt As String) As String
367:    Dim arrByte() As Byte
368:    Dim strTxtNew As String
369:    Dim i      As Long
370:    arrByte = StrConv(strTxt, vbFromUnicode)
371:    For i = 0 To UBound(arrByte)
372:        strTxtNew = strTxtNew & arrByte(i) & "||"
373:    Next i
374:    GetByteStringFromString = VBA.Left$(strTxtNew, VBA.Len(strTxtNew) - 2)
375: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetStringFromByte - получение из массива байтов строки байтов с разделителем ||
'* Created    : 14-02-2020 08:43
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):         Description
'*
'* arrByte() : - массив байтов
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function GetStringFromByte(ByRef arrByte() As Byte) As String
388:    Dim strTxtNew As String
389:    Dim i      As Long
390:    For i = 0 To UBound(arrByte)
391:        strTxtNew = strTxtNew & arrByte(i) & "||"
392:    Next i
393:    GetStringFromByte = VBA.Left$(strTxtNew, VBA.Len(strTxtNew) - 2)
394: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ByteFileAdd - сощздает пустой файл
'* Created    : 14-02-2020 08:42
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):         Description
'*
'* FilePath As String : - директория файла
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub ByteFileAdd(FilePath As String)
408:    Dim NumFile As Integer
409:    Dim s      As String
410:    'On Error GoTo eByteArrayToFile
411:    NumFile = FreeFile
412:    Open FilePath For Output As #NumFile
413:    Close #NumFile
414:    Exit Sub
eByteArrayToFile:
416:    s = "Ошибка открытия файла " & FilePath & "!"
417:    MsgBox s, 16, "ByteArrayToFile"
418: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ByteArrayToFile - Записывает в файл маcсив байтов
'* Created    : 14-02-2020 08:42
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):         Description
'*
'* FilePath As String : - директория файла
'* arrByte()          : - массив байтов
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub ByteArrayToFile(FilePath As String, arrByte() As Byte)
433:
434:    Dim NumFile As Integer
435:    Dim i As Long, LFile As Long
436:
437:    On Error GoTo ErrorHandler
438:
439:    NumFile = FreeFile
440:    LFile = FileLen(FilePath)
441:    Open FilePath For Binary As #NumFile
442:    For i = 1 To UBound(arrByte)
443:        Put #NumFile, LFile + i, arrByte(i)
444:    Next
445:    Close #NumFile
446:
447:    Exit Sub
ErrorHandler:
449:    Unload Me
450:    Select Case Err.Number
        Case Else:
452:            Call MsgBox("Ошибка! в HiddenModule.ByteArrayToFile" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbOKOnly + vbExclamation, "Ошибка:")
453:            Call WriteErrorLog("HiddenModule.ByteArrayToFile")
454:    End Select
455:    Err.Clear
456: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ByteArrayFromFile - Открывает бинарный файл и создает одномерный массив байтов
'* Created    : 14-02-2020 08:42
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):         Description
'*
'* FilePath As String : - директория файла
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ByteArrayFromFile(FilePath As String) As Byte()
470:    Dim NumFile As Integer
471:    Dim ArrBin() As Byte
472:    Dim sByte  As String
473:    Dim i As Long, LFile As Long
474:
475:    sByte = String(1, " ")
476:    LFile = FileLen(FilePath)
477:    ReDim ArrBin(LFile)
478:    NumFile = FreeFile
479:    Open FilePath For Binary As #NumFile
480:    For i = 1 To LFile
481:        Seek #NumFile, i
482:        Get #NumFile, , sByte
483:        ArrBin(i) = Asc(sByte)
484:    Next
485:    Close #NumFile
486:    ByteArrayFromFile = ArrBin
487:
End Function
