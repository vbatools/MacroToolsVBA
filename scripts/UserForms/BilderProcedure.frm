VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BilderProcedure 
   Caption         =   "Prosedure Bilder:"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17805
   OleObjectBlob   =   "BilderProcedure.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BilderProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : BilderProcedure - конструктор процедур
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
    Private Sub btnCancel_Click()
11:    Me.Hide
12: End Sub
    Private Sub lbCancel_Click()
14:    Call btnCancel_Click
15: End Sub

    Private Sub lbHelp_Click()
18:    Call URLLinks(C_Const.URL_BILD_PROC)
19: End Sub

    Private Sub UserForm_Initialize()
22:    With cmbFunc
23:        .AddItem "Boolean"
24:        .AddItem "String"
25:        .AddItem "Byte"
26:        .AddItem "Integer"
27:        .AddItem "Long"
28:        .AddItem "Single"
29:        .AddItem "Double"
30:        .AddItem "Currency"
31:        .AddItem "Variant"
32:        .AddItem "Date"
33:        .AddItem "Object"
34:    End With
35:    txtErroName.Text = "<- Поле ввода " & Chr(34) & Replace(lbName.Caption, "*:", vbNullString) & Chr(34) & " должно быть заполнено!"
36: End Sub
    Private Sub UserForm_Activate()
38:    chbAddMainProceure.Value = False
39:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
40: End Sub
    Private Sub chbAll_Change()
42:    Dim Flag        As Boolean
43:    Flag = chbAll.Value
44:    chbScreen.Value = Flag
45:    chbCalculations.Value = Flag
46:    chbAlerts.Value = Flag
47:    chbEvents.Value = Flag
48:    chbMsg.Value = Flag
49:    chbUseDefaultMsg.Value = Flag
50: End Sub
    Private Sub optTypeModif_Change()
52:    txtViewCode.Text = AddCode
53: End Sub
    Private Sub txtName_Change()
55:    Dim Txt         As String
56:    If txtName = vbNullString Then
57:        txtName.BorderColor = &HC0C0FF
58:    Else
59:        txtName.BorderColor = &H8000000D
60:    End If
61:    txtViewCode.Text = AddCode
62:    Txt = txtName.Text
63:    If VBA.Left$(Txt, 1) = "_" Then
64:        Txt = VBA.Right(Txt, VBA.Len(Txt) - 1)
65:        txtName.Text = Txt
66:    End If
67: End Sub
    Private Sub txtName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
69:    Dim sTemplate As String, Txt As String
70:    Txt = txtName.Text
71:    sTemplate = "!@#$%^&*+=.,'№/\|-:;{}[]() <>" & Chr(34)
72:    If InStr(1, sTemplate, ChrW(KeyAscii)) > 0 Then KeyAscii = 0
73:    If Txt = vbNullString Then
74:        Select Case KeyAscii
            Case 48 To 57: KeyAscii = 0
76:        End Select
77:    End If
78:    If VBA.Left$(Txt, 1) = "_" Then
79:        Txt = VBA.Right(Txt, VBA.Len(Txt) - 1)
80:        txtName.Text = Txt
81:    End If
82: End Sub
    Private Sub cmbFunc_Change()
84:    Call AddBackColorCombobox
85:    txtViewCode.Text = AddCode
86: End Sub
    Private Sub optTypeProcedure_Change()
88:    cmbFunc.Enabled = Not optTypeProcedure.Value
89:    chbArray.Enabled = Not optTypeProcedure.Value
90:    Call AddBackColorCombobox
91:    txtViewCode.Text = AddCode
92: End Sub
     Private Sub AddBackColorCombobox()
94:    cmbFunc.BorderColor = &H8000000D
95:    If (Not optTypeProcedure) Then
96:        If cmbFunc.Value = vbNullString Then
97:            cmbFunc.BorderColor = &HC0C0FF
98:        End If
99:    End If
100: End Sub
     Private Sub chbArray_Change()
102:    txtViewCode.Text = AddCode
103: End Sub
     Private Sub chbAlerts_Change()
105:    txtViewCode.Text = AddCode
106:    Call TernOffOn
107: End Sub
     Private Sub chbCalculations_Change()
109:    txtViewCode.Text = AddCode
110:    Call TernOffOn
111: End Sub
     Private Sub chbEvents_Change()
113:    txtViewCode.Text = AddCode
114:    Call TernOffOn
115: End Sub
     Private Sub chbMsg_Change()
117:    txtViewCode.Text = AddCode
118:
119:    chbUseDefaultMsg.Enabled = chbMsg.Value
120:    txtMsg.Enabled = chbMsg.Value
121: End Sub
     Private Sub chbScreen_Change()
123:    txtViewCode.Text = AddCode
124:    Call TernOffOn
125: End Sub
     Private Sub txtMsg_Change()
127:    txtViewCode.Text = AddCode
128: End Sub
     Private Sub txtDiscprition_Change()
130:    txtViewCode.Text = AddCode
131: End Sub
     Private Sub optDefaultError_Change()
133:    txtViewCode.Text = AddCode
134: End Sub
     Private Sub optResumNext_Change()
136:    txtViewCode.Text = AddCode
137: End Sub
     Private Sub chbUseDefaultMsg_Change()
139:    txtViewCode.Text = AddCode
140: End Sub
     Private Sub chbOffDiscription_Change()
142:    txtViewCode.Text = AddCode
143:    txtDiscprition.Enabled = chbOffDiscription.Value
144: End Sub
     Private Function AddCode() As String
146:    Dim strCode As String, strSpes As String, strEndLine As String
147:    Dim TypeModif As String, TypeProc As String, strDiscprition As String
148:    Dim TypeFunction As String, ResultDimFunc As String, ResultEndFunc As String
149:    Dim strMsg As String, strMsg1 As String, CustMsg As String
150:    Dim ErrorMsgFerst As String, ErrorMsgEnd As String
151:    Dim MsgStop     As String
152:    Dim ScreenUpdatingCalculationTrue As String, ScreenUpdatingCalculationFalse As String
153:    Dim txtArray    As String
154:
155:    If txtName.Text = vbNullString Then
156:        txtErroName.visible = True
157:        Exit Function
158:    Else
159:        txtErroName.visible = False
160:    End If
161:    If (Not optTypeProcedure) Then
162:        If cmbFunc.Value = vbNullString Then
163:            MsgStop = "Поле выбора типа данных функции должно быть заполнено!"
164:        End If
165:    End If
166:
167:    If MsgStop <> vbNullString Then
168:        Call MsgBox(MsgStop, vbOKOnly + vbCritical, "Ошибка:")
169:        Exit Function
170:    End If
171:
172:    strEndLine = vbNewLine & vbTab
173:    'отключение описание
174:    If chbOffDiscription Then
175:        strDiscprition = strEndLine & Chr(39) & "Описание: " & txtDiscprition.Text
176:        strDiscprition = strDiscprition & strEndLine & Chr(39) & "Дата создания: " & Format(Now(), "dddddd в  h:nn:ss")
177:        strDiscprition = strDiscprition & strEndLine & Chr(39) & "Автор: " & Environ("UserName")
178:    End If
179:    strSpes = Space(1)
180:    ScreenUpdatingCalculationTrue = "Call ScreenUpdatingCalculation(Screen:=True, Calculat:=True, Alerts:=True, Events:=True)"
181:
182:    'тип модификатора доступа
183:    If optTypeModif Then
184:        TypeModif = "Public"
185:    Else
186:        TypeModif = "Private"
187:    End If
188:
189:    'массив для функций
190:    If (Not optTypeProcedure.Value) And chbArray.Value Then
191:        txtArray = " ()"
192:    End If
193:
194:    'процедура или функция
195:    If optTypeProcedure Then
196:        TypeProc = "Sub"
197:        TypeFunction = vbNullString
198:    Else
199:        TypeProc = "Function"
200:        TypeFunction = " as " & cmbFunc.Value
201:        ResultDimFunc = vbNewLine & vbTab & "Dim Result" & txtArray & " as " & cmbFunc.Value
202:        ResultEndFunc = vbNewLine & vbTab & txtName.Text & " = Result"
203:    End If
204:
205:    'вывод сообщения по окончанию
206:    If chbMsg Then
207:        Dim txtNewLine As String
208:        If txtMsg.Text <> vbNullString Then txtNewLine = " & vbNewLine & "
209:        If chbUseDefaultMsg Then strMsg1 = Chr(34) & "Выполнение " & txtName.Text & " окнчено!" & Chr(34) & txtNewLine
210:        CustMsg = txtMsg.Text
211:        If CustMsg = vbNullString Then
212:            If chbUseDefaultMsg Then
213:                CustMsg = vbNullString
214:            Else
215:                CustMsg = Chr(34) & vbNullString & Chr(34)
216:            End If
217:        Else
218:            CustMsg = Replace(CustMsg, Chr(34), "| & Chr(34) & |")
219:            CustMsg = Chr(34) & Replace(CustMsg, "|", Chr(34)) & Chr(34)
220:        End If
221:        strMsg1 = strMsg1 & CustMsg
222:        strMsg = strEndLine & "Call MsgBox(" & strMsg1 & ", vbOKOnly + vbInformation," & Chr(34) & txtName.Text & Chr(34) & ")"
223:    End If
224:
225:    'обработка ошибок
226:    If optDefaultError Then
227:        ErrorMsgFerst = vbNullString
228:        ErrorMsgEnd = vbNullString
229:    End If
230:
231:    If optResumNext Then
232:        ErrorMsgFerst = strEndLine & "On Error Resume Next"
233:        ErrorMsgEnd = vbNullString
234:    End If
235:    ScreenUpdatingCalculationFalse = "Call ScreenUpdatingCalculation(Screen:=" & (Not chbScreen.Value) & ", Calculat:=" & (Not chbCalculations.Value) & ", Alerts:=" & (Not chbAlerts.Value) & ", Events:=" & (Not chbEvents.Value) & ")"
236:
237:    'если ни чего не отключено то не чего включать
238:    If ScreenUpdatingCalculationFalse = ScreenUpdatingCalculationTrue Then
239:        ScreenUpdatingCalculationTrue = vbNullString
240:        ScreenUpdatingCalculationFalse = vbNullString
241:    End If
242:
243:    If optErrorHandele Then
244:        ErrorMsgFerst = strEndLine & "On Error GoTo ErrorHandler"
245:        ErrorMsgEnd = strEndLine & "Exit " & TypeProc & vbNewLine & "ErrorHandler:" & strEndLine & ScreenUpdatingCalculationTrue
246:        ErrorMsgEnd = ErrorMsgEnd & strEndLine & "Select Case Err"
        ErrorMsgEnd = ErrorMsgEnd & strEndLine & vbTab & Chr(39) & "обработка ошибок для использования раскомментировать"
248:        ErrorMsgEnd = ErrorMsgEnd & strEndLine & vbTab & Chr(39) & "Case"
249:        ErrorMsgEnd = ErrorMsgEnd & strEndLine & vbTab & "Case Else:"
250:        ErrorMsgEnd = ErrorMsgEnd & strEndLine & vbTab & vbTab & "Debug.Print " & Chr(34) & "Произошла ошибка в " & txtName & Chr(34) & " & vbNewLine & Err.Number & vbNewLine & Err.Description"
251:        ErrorMsgEnd = ErrorMsgEnd & strEndLine & "End Select"
252:    End If
253:
254:    'формирование кода
255:    strCode = TypeModif & strSpes & TypeProc & strSpes & txtName & strSpes & "()" & TypeFunction & txtArray
256:    strCode = strCode & strDiscprition
257:    strCode = strCode & ResultDimFunc
258:    strCode = strCode & ErrorMsgFerst
259:    strCode = strCode & strEndLine & ScreenUpdatingCalculationFalse
260:    strCode = strCode & strEndLine & strEndLine & Chr(39) & "место для кода" & strEndLine
261:    strCode = strCode & ResultEndFunc
262:    strCode = strCode & strEndLine & ScreenUpdatingCalculationTrue
263:    strCode = strCode & strMsg
264:    strCode = strCode & ErrorMsgEnd
265:    strCode = strCode & vbNewLine & "End " & TypeProc
266:
267:    AddCode = strCode
268: End Function
     Private Function AddMainProceure() As String
270:    Dim txtCode     As String
271:    txtCode = txtViewCode.Text
272:
273:    'копирование процедуры ScreenUpdatingCalculation
274:    If chbAddMainProceure Then
275:        Dim snippets As ListObject
276:        Dim i_row   As Long
277:        Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
278:        i_row = snippets.ListColumns(3).DataBodyRange.Find(What:="cu.ScreenUpdatingCalculation", LookIn:=xlValues, LookAt:=xlWhole).Row
279:        txtCode = txtCode & vbNewLine & snippets.Range(i_row, 4)
280:    End If
281:    AddMainProceure = txtCode
282: End Function
     Private Sub btnCopyCode_Click()
284:    Dim sMsgBoxString As String, txtCode As String
285:    'получение кода
286:    txtCode = AddMainProceure()
287:
288:    If opbCliboard Then
289:        Call C_PublicFunctions.SetTextIntoClipboard(txtCode)
290:        sMsgBoxString = "Код скопирован в буфер обмена!" & vbNewLine & "Для вставки кода используйте " & Chr(34) & "Ctrl+V" & Chr(34)
291:        Call MsgBox(sMsgBoxString, vbInformation, "Копирование кода:")
292:    Else
293:        Debug.Print txtCode
294:    End If
295:    Me.Hide
296: End Sub
     Private Sub lbInsertCode_Click()
298:    Dim iLine       As Integer
299:    Dim txtCode As String, txtLine As String
300:    'получение кода
301:    txtCode = AddMainProceure()
302:    If txtCode = vbNullString Then Exit Sub
303:    txtLine = C_PublicFunctions.SelectedLineColumnProcedure
304:    If txtLine = vbNullString Then
305:        Me.Hide
306:        Exit Sub
307:    End If
308:    iLine = Split(txtLine, "|")(2)
309:
310:    With Application.VBE.ActiveCodePane
311:        .CodeModule.InsertLines iLine, txtCode
312:    End With
313:    Me.Hide
314: End Sub

Private Sub TernOffOn()
317:    If (chbScreen.Value + chbCalculations.Value + chbAlerts.Value + chbEvents.Value) <> 0 Then
318:        chbAddMainProceure.Enabled = True
319:    Else
320:        chbAddMainProceure.Enabled = False
321:    End If
End Sub
