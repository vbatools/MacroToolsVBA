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
10:    Me.Hide
11: End Sub
    Private Sub lbCancel_Click()
13:    Call btnCancel_Click
14: End Sub

    Private Sub lbHelp_Click()
17:    Call URLLinks(C_Const.URL_BILD_PROC)
18: End Sub

    Private Sub UserForm_Initialize()
21:    With cmbFunc
22:        .AddItem "Boolean"
23:        .AddItem "String"
24:        .AddItem "Byte"
25:        .AddItem "Integer"
26:        .AddItem "Long"
27:        .AddItem "Single"
28:        .AddItem "Double"
29:        .AddItem "Currency"
30:        .AddItem "Variant"
31:        .AddItem "Date"
32:        .AddItem "Object"
33:    End With
34:    txtErroName.Text = "< - Input field" & Chr(34) & Replace(lbName.Caption, "*:", vbNullString) & Chr(34) & "must be filled in!"
35: End Sub
    Private Sub UserForm_Activate()
37:    chbAddMainProceure.Value = False
38:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
39: End Sub
    Private Sub chbAll_Change()
41:    Dim Flag        As Boolean
42:    Flag = chbAll.Value
43:    chbScreen.Value = Flag
44:    chbCalculations.Value = Flag
45:    chbAlerts.Value = Flag
46:    chbEvents.Value = Flag
47:    chbMsg.Value = Flag
48:    chbUseDefaultMsg.Value = Flag
49: End Sub
    Private Sub optTypeModif_Change()
51:    txtViewCode.Text = AddCode
52: End Sub
    Private Sub txtName_Change()
54:    Dim Txt         As String
55:    If txtName = vbNullString Then
56:        txtName.BorderColor = &HC0C0FF
57:    Else
58:        txtName.BorderColor = &H8000000D
59:    End If
60:    txtViewCode.Text = AddCode
61:    Txt = txtName.Text
62:    If VBA.Left$(Txt, 1) = "_" Then
63:        Txt = VBA.Right(Txt, VBA.Len(Txt) - 1)
64:        txtName.Text = Txt
65:    End If
66: End Sub
    Private Sub txtName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
68:    Dim sTemplate As String, Txt As String
69:    Txt = txtName.Text
70:    sTemplate = "!@#$%^&*+=.,'№/\|-:;{}[]() <>" & Chr(34)
71:    If InStr(1, sTemplate, ChrW(KeyAscii)) > 0 Then KeyAscii = 0
72:    If Txt = vbNullString Then
73:        Select Case KeyAscii
            Case 48 To 57: KeyAscii = 0
75:        End Select
76:    End If
77:    If VBA.Left$(Txt, 1) = "_" Then
78:        Txt = VBA.Right(Txt, VBA.Len(Txt) - 1)
79:        txtName.Text = Txt
80:    End If
81: End Sub
    Private Sub cmbFunc_Change()
83:    Call AddBackColorCombobox
84:    txtViewCode.Text = AddCode
85: End Sub
    Private Sub optTypeProcedure_Change()
87:    cmbFunc.Enabled = Not optTypeProcedure.Value
88:    chbArray.Enabled = Not optTypeProcedure.Value
89:    Call AddBackColorCombobox
90:    txtViewCode.Text = AddCode
91: End Sub
    Private Sub AddBackColorCombobox()
93:    cmbFunc.BorderColor = &H8000000D
94:    If (Not optTypeProcedure) Then
95:        If cmbFunc.Value = vbNullString Then
96:            cmbFunc.BorderColor = &HC0C0FF
97:        End If
98:    End If
99: End Sub
     Private Sub chbArray_Change()
101:    txtViewCode.Text = AddCode
102: End Sub
     Private Sub chbAlerts_Change()
104:    txtViewCode.Text = AddCode
105:    Call TernOffOn
106: End Sub
     Private Sub chbCalculations_Change()
108:    txtViewCode.Text = AddCode
109:    Call TernOffOn
110: End Sub
     Private Sub chbEvents_Change()
112:    txtViewCode.Text = AddCode
113:    Call TernOffOn
114: End Sub
     Private Sub chbMsg_Change()
116:    txtViewCode.Text = AddCode
117:
118:    chbUseDefaultMsg.Enabled = chbMsg.Value
119:    txtMsg.Enabled = chbMsg.Value
120: End Sub
     Private Sub chbScreen_Change()
122:    txtViewCode.Text = AddCode
123:    Call TernOffOn
124: End Sub
     Private Sub txtMsg_Change()
126:    txtViewCode.Text = AddCode
127: End Sub
     Private Sub txtDiscprition_Change()
129:    txtViewCode.Text = AddCode
130: End Sub
     Private Sub optDefaultError_Change()
132:    txtViewCode.Text = AddCode
133: End Sub
     Private Sub optResumNext_Change()
135:    txtViewCode.Text = AddCode
136: End Sub
     Private Sub chbUseDefaultMsg_Change()
138:    txtViewCode.Text = AddCode
139: End Sub
     Private Sub chbOffDiscription_Change()
141:    txtViewCode.Text = AddCode
142:    txtDiscprition.Enabled = chbOffDiscription.Value
143: End Sub
     Private Function AddCode() As String
145:    Dim strCode As String, strSpes As String, strEndLine As String
146:    Dim TypeModif As String, TypeProc As String, strDiscprition As String
147: Dim TypeFunction As String, ResultDimFunc As String, ResultEndFunc As String
148:    Dim strMsg As String, strMsg1 As String, CustMsg As String
149:    Dim ErrorMsgFerst As String, ErrorMsgEnd As String
150:    Dim MsgStop     As String
151:    Dim ScreenUpdatingCalculationTrue As String, ScreenUpdatingCalculationFalse As String
152:    Dim txtArray    As String
153:
154:    If txtName.Text = vbNullString Then
155:        txtErroName.visible = True
156:        Exit Function
157:    Else
158:        txtErroName.visible = False
159:    End If
160:    If (Not optTypeProcedure) Then
161:        If cmbFunc.Value = vbNullString Then
162:            MsgStop = "The function data type selection field must be filled in!"
163:        End If
164:    End If
165:
166:    If MsgStop <> vbNullString Then
167:        Call MsgBox(MsgStop, vbOKOnly + vbCritical, "Error:")
168:        Exit Function
169:    End If
170:
171:    strEndLine = vbNewLine & vbTab
172:    'отключение описание
173:    If chbOffDiscription Then
174:        strDiscprition = strEndLine & Chr(39) & "Description:" & txtDiscprition.Text
175:        strDiscprition = strDiscprition & strEndLine & Chr(39) & "Дата создания: " & Format(Now(), "dddddd в  h:nn:ss")
176:        strDiscprition = strDiscprition & strEndLine & Chr(39) & "Author:" & Environ("UserName")
177:    End If
178:    strSpes = Space(1)
179:    ScreenUpdatingCalculationTrue = "Call ScreenUpdatingCalculation(Screen:=True, Calculat:=True, Alerts:=True, Events:=True)"
180:
181:    'тип модификатора доступа
182:    If optTypeModif Then
183:        TypeModif = "Public"
184:    Else
185:        TypeModif = "Private"
186:    End If
187:
188:    'массив для функций
189:    If (Not optTypeProcedure.Value) And chbArray.Value Then
190:        txtArray = " ()"
191:    End If
192:
193:    'процедура или функция
194:    If optTypeProcedure Then
195:        TypeProc = "Sub"
196: TypeFunction = vbNullString
197:    Else
198:        TypeProc = "Function"
199: TypeFunction = " as " & cmbFunc.Value
200:        ResultDimFunc = vbNewLine & vbTab & "Dim Result" & txtArray & " as " & cmbFunc.Value
201:        ResultEndFunc = vbNewLine & vbTab & txtName.Text & " = Result"
202:    End If
203:
204:    'вывод сообщения по окончанию
205:    If chbMsg Then
206:        Dim txtNewLine As String
207:        If txtMsg.Text <> vbNullString Then txtNewLine = " & vbNewLine & "
208:        If chbUseDefaultMsg Then strMsg1 = Chr(34) & "Accomplishment" & txtName.Text & "It's over!" & Chr(34) & txtNewLine
209:        CustMsg = txtMsg.Text
210:        If CustMsg = vbNullString Then
211:            If chbUseDefaultMsg Then
212:                CustMsg = vbNullString
213:            Else
214:                CustMsg = Chr(34) & vbNullString & Chr(34)
215:            End If
216:        Else
217:            CustMsg = Replace(CustMsg, Chr(34), "| & Chr(34) & |")
218:            CustMsg = Chr(34) & Replace(CustMsg, "|", Chr(34)) & Chr(34)
219:        End If
220:        strMsg1 = strMsg1 & CustMsg
221:        strMsg = strEndLine & "Call MsgBox(" & strMsg1 & ", vbOKOnly + vbInformation," & Chr(34) & txtName.Text & Chr(34) & ")"
222:    End If
223:
224:    'обработка ошибок
225:    If optDefaultError Then
226:        ErrorMsgFerst = vbNullString
227:        ErrorMsgEnd = vbNullString
228:    End If
229:
230:    If optResumNext Then
231:        ErrorMsgFerst = strEndLine & "On Error Resume Next"
232:        ErrorMsgEnd = vbNullString
233:    End If
234:    ScreenUpdatingCalculationFalse = "Call ScreenUpdatingCalculation(Screen:=" & (Not chbScreen.Value) & ", Calculat:=" & (Not chbCalculations.Value) & ", Alerts:=" & (Not chbAlerts.Value) & ", Events:=" & (Not chbEvents.Value) & ")"
235:
236:    'если ни чего не отключено то не чего включать
237:    If ScreenUpdatingCalculationFalse = ScreenUpdatingCalculationTrue Then
238:        ScreenUpdatingCalculationTrue = vbNullString
239:        ScreenUpdatingCalculationFalse = vbNullString
240:    End If
241:
242:    If optErrorHandele Then
243:        ErrorMsgFerst = strEndLine & "On Error GoTo ErrorHandler"
244:        ErrorMsgEnd = strEndLine & "Exit " & TypeProc & vbNewLine & "ErrorHandler:" & strEndLine & ScreenUpdatingCalculationTrue
245:        ErrorMsgEnd = ErrorMsgEnd & strEndLine & "Select Case Err"
        ErrorMsgEnd = ErrorMsgEnd & strEndLine & vbTab & Chr(39) & "error handling for using uncomment"
247:        ErrorMsgEnd = ErrorMsgEnd & strEndLine & vbTab & Chr(39) & "Case"
248:        ErrorMsgEnd = ErrorMsgEnd & strEndLine & vbTab & "Case Else:"
249:        ErrorMsgEnd = ErrorMsgEnd & strEndLine & vbTab & vbTab & "Debug.Print " & Chr(34) & "An error occurred in" & txtName & Chr(34) & " & vbNewLine & Err.Number & vbNewLine & Err.Description"
250:        ErrorMsgEnd = ErrorMsgEnd & strEndLine & "End Select"
251:    End If
252:
253:    'формирование кода
254: strCode = TypeModif & strSpes & TypeProc & strSpes & txtName & strSpes & "()" & TypeFunction & txtArray
255:    strCode = strCode & strDiscprition
256:    strCode = strCode & ResultDimFunc
257:    strCode = strCode & ErrorMsgFerst
258:    strCode = strCode & strEndLine & ScreenUpdatingCalculationFalse
259:    strCode = strCode & strEndLine & strEndLine & Chr(39) & "place for the code" & strEndLine
260:    strCode = strCode & ResultEndFunc
261:    strCode = strCode & strEndLine & ScreenUpdatingCalculationTrue
262:    strCode = strCode & strMsg
263:    strCode = strCode & ErrorMsgEnd
264:    strCode = strCode & vbNewLine & "End " & TypeProc
265:
266:    AddCode = strCode
267: End Function
     Private Function AddMainProceure() As String
269:    Dim txtCode     As String
270:    txtCode = txtViewCode.Text
271:
272:    'копирование процедуры ScreenUpdatingCalculation
273:    If chbAddMainProceure Then
274:        Dim snippets As ListObject
275:        Dim i_row   As Long
276:        Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
277:        i_row = snippets.ListColumns(3).DataBodyRange.Find(What:="cu.ScreenUpdatingCalculation", LookIn:=xlValues, LookAt:=xlWhole).Row
278:        txtCode = txtCode & vbNewLine & snippets.Range(i_row, 4)
279:    End If
280:    AddMainProceure = txtCode
281: End Function
     Private Sub btnCopyCode_Click()
283:    Dim sMsgBoxString As String, txtCode As String
284:    'получение кода
285:    txtCode = AddMainProceure()
286:
287:    If opbCliboard Then
288:        Call C_PublicFunctions.SetTextIntoClipboard(txtCode)
289:        sMsgBoxString = "The code is copied to the clipboard!" & vbNewLine & "To insert the code, use" & Chr(34) & "Ctrl+V" & Chr(34)
290:        Call MsgBox(sMsgBoxString, vbInformation, "Copying the code:")
291:    Else
292:        Debug.Print txtCode
293:    End If
294:    Me.Hide
295: End Sub
     Private Sub lbInsertCode_Click()
297:    Dim iLine       As Integer
298:    Dim txtCode As String, txtLine As String
299:    'получение кода
300:    txtCode = AddMainProceure()
301:    If txtCode = vbNullString Then Exit Sub
302:    txtLine = C_PublicFunctions.SelectedLineColumnProcedure
303:    If txtLine = vbNullString Then
304:        Me.Hide
305:        Exit Sub
306:    End If
307:    iLine = Split(txtLine, "|")(2)
308:
309:    With Application.VBE.ActiveCodePane
310:        .CodeModule.InsertLines iLine, txtCode
311:    End With
312:    Me.Hide
313: End Sub

Private Sub TernOffOn()
316:    If (chbScreen.Value + chbCalculations.Value + chbAlerts.Value + chbEvents.Value) <> 0 Then
317:        chbAddMainProceure.Enabled = True
318:    Else
319:        chbAddMainProceure.Enabled = False
320:    End If
End Sub
