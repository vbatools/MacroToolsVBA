VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ObfuscationCode 
   Caption         =   "Удаление форматирования:"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13335
   OleObjectBlob   =   "ObfuscationCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ObfuscationCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : ObfuscationCode - обфускация кода
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
    Private Sub cmbCancel_Click()
11:    Unload Me
12: End Sub

    Private Sub lbCancel_Click()
15:    Call cmbCancel_Click
16: End Sub

    Private Sub lbHelp_Click()
19:    Call URLLinks(C_Const.URL_FILE_OBFS)
20: End Sub

    Private Sub UserForm_Activate()
23:    Me.StartUpPosition = 0
24:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
25:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
26:
27:    Dim wb          As Workbook
28:    Dim vbProj      As VBIDE.VBProject
29:
30:    On Error GoTo ErrorHandler
31:    If Workbooks.Count = 0 Then
32:        Unload Me
33:        Call MsgBox("Нет открытых " & Chr(34) & "Файлов Excel" & Chr(34) & "!", vbOKOnly + vbExclamation, "Ошибка:")
34:        Exit Sub
35:    End If
36:    With Me.cmbMain
37:        .Clear
38:        On Error Resume Next
39:        For Each vbProj In Application.VBE.VBProjects
40:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
41:        Next
42:        On Error GoTo 0
43:        On Error GoTo ErrorHandler
44:        .Value = ActiveWorkbook.Name
45:    End With
46:    Call getWord(cmbMain)
47:    Call AddListCode
48:    lbMsg.visible = True
49:    lbOK.Enabled = False
50:
51:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
52:    Exit Sub
ErrorHandler:
54:    Unload Me
55:    Select Case Err.Number
        Case Else:
57:            Call MsgBox("Ошибка! в ObfuscationCode.UserForm_Activate" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbOKOnly + vbExclamation, "Ошибка:")
58:            Call WriteErrorLog("ObfuscationCode.UserForm_Activate")
59:    End Select
60:    Err.Clear
61: End Sub
    Private Sub getWord(ByRef oList As MSForms.ComboBox)
63:    On Error Resume Next
64:    Dim objW        As Object
65:    Dim vbProj      As VBIDE.VBProject
66:    Dim sVal        As String
67:    Set objW = GetObject(, "Word.Application")
68:    For Each vbProj In objW.VBE.VBProjects
69:        sVal = C_PublicFunctions.sGetFileName(vbProj.Filename)
70:        If sVal Like "*.docm" Or sVal Like "*.DOCM" Then oList.AddItem sVal
71:    Next
72: End Sub
    Private Sub CheckAll_Click()
74:    Dim i           As Integer
75:    With ListCode
76:        For i = 0 To .ListCount - 1
77:            .Selected(i) = CheckAll.Value
78:        Next i
79:    End With
80: End Sub
    Private Sub CheckAllParam_Click()
82:    If CheckAllParam.Value = True Then
83:        CheckAllParam.Value = True
84:        CheckDelComment.Value = True
85:        CheckDelEmptyLines.Value = True
86:        CheckDelExplicit.Value = True
87:        CheckDelFormat.Value = True
88:        CheckDelNomerLine.Value = True
89:        CheckDelBreaksLines.Value = True
90:    Else
91:        CheckAllParam.Value = False
92:        CheckDelComment.Value = False
93:        CheckDelEmptyLines.Value = False
94:        CheckDelExplicit.Value = False
95:        CheckDelFormat.Value = False
96:        CheckDelNomerLine.Value = False
97:        CheckDelBreaksLines.Value = False
98:    End If
99: End Sub

     Private Sub cmbMain_Change()
102:    If cmbMain.Value <> vbNullString Then
103:        Call AddListCode
104:    End If
105:    CheckAll.Value = False
106: End Sub
     Private Sub ListCode_Change()
108:    Dim i           As Integer
109:    With ListCode
110:        For i = 0 To .ListCount - 1
111:            If .Selected(i) Then
112:                lbMsg.visible = False
113:                lbOK.Enabled = True
114:                Exit Sub
115:            End If
116:        Next i
117:    End With
118:    lbMsg.visible = True
119:    lbOK.Enabled = False
120: End Sub
     Private Sub AddListCode()
122:    Dim wb          As Object
123:    Dim iFile       As Integer
124:    Dim Arr()       As Variant
125:    Dim sNameWB     As String
126:    sNameWB = cmbMain.Value
127:    If sNameWB = vbNullString Then Exit Sub
128:    If sNameWB Like "*.docm" Or sNameWB Like "*.DOCM" Then
129:        Dim objWrdApp As Object
130:        Set objWrdApp = GetObject(, "Word.Application")
131:        Set wb = objWrdApp.Documents(sNameWB)
132:    Else
133:        Set wb = Workbooks(sNameWB)
134:    End If
135:    With ListCode
136:        .Clear
137:        Dim vbProj  As Object
138:        Set vbProj = wb.VBProject
139:        If vbProj.Protection = 1 Then
140:            Call MsgBox("Проект VBA защищен паролем, снимите пароль с проекта!", vbCritical, "Удаление форматирования:")
141:            Exit Sub
142:        End If
143:        For iFile = 1 To vbProj.VBComponents.Count
144:            .AddItem iFile
145:            .List(iFile - 1, 1) = ComponentTypeToString(vbProj.VBComponents(iFile).Type)
146:            .List(iFile - 1, 2) = vbProj.VBComponents(iFile).Name
147:        Next iFile
148:        Arr = .List
149:        Call Sort2_asc(Arr, 1)
150:        .List = Arr
151:        For iFile = 0 To .ListCount - 1
152:            .List(iFile, 0) = iFile + 1
153:        Next iFile
154:    End With
155: End Sub

'сортировка массива
     Private Sub Sort2_asc(Arr(), col As Long)
159:    Dim temp()      As Variant
160:    Dim lb2 As Long, ub2 As Long, lTop As Long, lBot As Long
161:
162:    lTop = LBound(Arr, 1)
163:    lBot = UBound(Arr, 1)
164:    lb2 = LBound(Arr, 2)
165:    ub2 = UBound(Arr, 2)
166:    ReDim temp(lb2 To ub2)
167:
168:    Call QSort2_asc(Arr(), col, lTop, lBot, temp(), lb2, ub2)
169: End Sub
     Private Sub QSort2_asc(Arr(), C As Long, ByVal top As Long, ByVal bot As Long, temp(), lb2 As Long, ub2 As Long)
171:    Dim t As Long, LB As Long, MidItem, j As Long
172:
173:    MidItem = Arr((top + bot) \ 2, C)
174:    t = top: LB = bot
175:
176:    Do
177:        Do While Arr(t, C) < MidItem: t = t + 1: Loop
178:        Do While Arr(LB, C) > MidItem: LB = LB - 1: Loop
179:        If t < LB Then
180:            For j = lb2 To ub2: temp(j) = Arr(t, j): Next j
181:            For j = lb2 To ub2: Arr(t, j) = Arr(LB, j): Next j
182:            For j = lb2 To ub2: Arr(LB, j) = temp(j): Next j
183:            t = t + 1: LB = LB - 1
184:        ElseIf t = LB Then
185:            t = t + 1: LB = LB - 1
186:        End If
187:    Loop While t <= LB
188:
189:    If t < bot Then QSort2_asc Arr(), C, t, bot, temp(), lb2, ub2
190:    If top < LB Then QSort2_asc Arr(), C, top, LB, temp(), lb2, ub2
191:
192: End Sub
Private Sub lbOK_Click()
194:    Dim oldWbName   As String
195:    Dim i As Integer, j As Integer
196:    Dim vbProj      As VBIDE.VBProject
197:    Dim vbComp      As VBIDE.VBComponent
198:    Dim arrNameFile() As String
199:    Dim wb          As Object
200:    Dim sPath       As String
201:    Dim sNameWB     As String
202:    oldWbName = cmbMain.Value
203:
204:
205:    sNameWB = cmbMain.Value
206:    If sNameWB = vbNullString Then Exit Sub
207:    If sNameWB Like "*.docm" Or sNameWB Like "*.DOCM" Then
208:        Dim objWrdApp As Object
209:        Set objWrdApp = GetObject(, "Word.Application")
210:        Set wb = objWrdApp.Documents(sNameWB)
211:    Else
212:        Set wb = Workbooks(sNameWB)
213:    End If
214:
215:    Me.Hide
216:    If MsgBox("Вы полнить удаление форматирования кода ?", vbCritical + vbYesNo, "Удаление форматирования кода:") = vbYes Then
217:
218:        If Not wb.Name Like "*_obf_*" Then
219:            sPath = Left(wb.FullName, Len(wb.FullName) - Len(wb.Name))
220:            If sPath = vbNullString Then
221:                Call MsgBox("Файл не сохранен, для продолжения необходимо сохранить файл: [ " & wb.Name & " ]", vbInformation, "Ошибка:")
222:                Exit Sub
223:            End If
224:            arrNameFile = Split(wb.Name, ".")
225:            wb.SaveAs Filename:=sPath & arrNameFile(0) & "_obf_" & Replace(Now(), ":", ".") & "." & arrNameFile(1)    ', FileFormat:=wb.FileFormat
226:        End If
227:        j = -1
228:
229:        Set vbProj = wb.VBProject
230:        For i = 0 To ListCode.ListCount - 1
231:            If ListCode.Selected(i) = True Then
232:                Set vbComp = vbProj.VBComponents(ListCode.List(i, 2))
233:
234:                If CheckDelNomerLine Then
235:                    Call K_AddNumbersLine.RemoveLineNumbers(vbComp, vbLineNumbers_LabelTypes.vbLabelColon)
236:                    Call K_AddNumbersLine.RemoveLineNumbers(vbComp, vbLineNumbers_LabelTypes.vbLabelTab)
237:                End If
238:                If CheckDelComment Then
239:                    Call N_Obfuscation.Remove_Comments(vbComp.CodeModule)
240:                End If
241:                If CheckDelFormat Then
242:                    Call N_Obfuscation.TrimLinesTabAndSpase(vbComp.CodeModule)
243:                End If
244:                If CheckDelExplicit Then
245:                    Call N_Obfuscation.Remove_OptionExplicit(vbComp.CodeModule)
246:                End If
247:                If CheckDelEmptyLines Then
248:                    Call N_Obfuscation.Remove_EmptyLines(vbComp.CodeModule)
249:                End If
250:                If CheckDelBreaksLines Then
251:                    Call N_Obfuscation.RemoveBreaksLineInCode(vbComp.CodeModule)
252:                End If
253:            End If
254:        Next i
255:
256:        wb.Save
257:        Call MsgBox("Удаление форматирования [" & oldWbName & "] завершено!", vbInformation, "Удаление форматирования:")
258:        cmbMain.Value = wb.Name
259:        Me.Show
260:        Exit Sub
261:    End If
262:    cmbMain.Value = wb.Name
263:    Me.Show
End Sub
