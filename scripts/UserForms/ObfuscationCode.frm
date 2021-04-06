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
27:    Dim WB          As Workbook
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
46:    Call AddListCode
47:    lbMsg.visible = True
48:    lbOk.Enabled = False
49:
50:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
51:    Exit Sub
ErrorHandler:
53:    Unload Me
54:    Select Case Err.Number
        Case Else:
56:            Call MsgBox("Ошибка! в ObfuscationCode.UserForm_Activate" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbOKOnly + vbExclamation, "Ошибка:")
57:            Call WriteErrorLog("ObfuscationCode.UserForm_Activate")
58:    End Select
59:    Err.Clear
60: End Sub
    Private Sub CheckAll_Click()
62:    Dim i           As Integer
63:    With ListCode
64:        For i = 0 To .ListCount - 1
65:            .Selected(i) = CheckAll.Value
66:        Next i
67:    End With
68: End Sub
    Private Sub CheckAllParam_Click()
70:    If CheckAllParam.Value = True Then
71:        CheckAllParam.Value = True
72:        CheckDelComment.Value = True
73:        CheckDelEmptyLines.Value = True
74:        CheckDelExplicit.Value = True
75:        CheckDelFormat.Value = True
76:        CheckDelNomerLine.Value = True
77:        CheckDelBreaksLines.Value = True
78:    Else
79:        CheckAllParam.Value = False
80:        CheckDelComment.Value = False
81:        CheckDelEmptyLines.Value = False
82:        CheckDelExplicit.Value = False
83:        CheckDelFormat.Value = False
84:        CheckDelNomerLine.Value = False
85:        CheckDelBreaksLines.Value = False
86:    End If
87: End Sub

    Private Sub cmbMain_Change()
90:    If cmbMain.Value <> vbNullString Then
91:        Call AddListCode
92:    End If
93: End Sub
     Private Sub ListCode_Change()
95:    Dim i           As Integer
96:    With ListCode
97:        For i = 0 To .ListCount - 1
98:            If .Selected(i) Then
99:                lbMsg.visible = False
100:                lbOk.Enabled = True
101:                Exit Sub
102:            End If
103:        Next i
104:    End With
105:    lbMsg.visible = True
106:    lbOk.Enabled = False
107: End Sub
     Private Sub AddListCode()
109:    Dim WB          As Workbook
110:    Dim iFile       As Integer
111:    Dim Arr()       As Variant
112:    Set WB = Workbooks(cmbMain.Value)
113:    With ListCode
114:        .Clear
115:        For iFile = 1 To WB.VBProject.VBComponents.Count
116:            .AddItem iFile
117:            .List(iFile - 1, 1) = ComponentTypeToString(WB.VBProject.VBComponents(iFile).Type)
118:            .List(iFile - 1, 2) = WB.VBProject.VBComponents(iFile).Name
119:        Next iFile
120:        Arr = .List
121:        Call Sort2_asc(Arr, 1)
122:        .List = Arr
123:        For iFile = 0 To .ListCount - 1
124:            .List(iFile, 0) = iFile + 1
125:        Next iFile
126:    End With
127: End Sub
'сортировка массива
     Private Sub Sort2_asc(Arr(), col As Long)
130:    Dim temp()      As Variant
131:    Dim lb2 As Long, ub2 As Long, lTop As Long, lBot As Long
132:
133:    lTop = LBound(Arr, 1)
134:    lBot = UBound(Arr, 1)
135:    lb2 = LBound(Arr, 2)
136:    ub2 = UBound(Arr, 2)
137:    ReDim temp(lb2 To ub2)
138:
139:    Call QSort2_asc(Arr(), col, lTop, lBot, temp(), lb2, ub2)
140: End Sub
     Private Sub QSort2_asc(Arr(), C As Long, ByVal top As Long, ByVal bot As Long, temp(), lb2 As Long, ub2 As Long)
142:    Dim t As Long, LB As Long, MidItem, j As Long
143:
144:    MidItem = Arr((top + bot) \ 2, C)
145:    t = top: LB = bot
146:
147:    Do
148:        Do While Arr(t, C) < MidItem: t = t + 1: Loop
149:        Do While Arr(LB, C) > MidItem: LB = LB - 1: Loop
150:        If t < LB Then
151:            For j = lb2 To ub2: temp(j) = Arr(t, j): Next j
152:            For j = lb2 To ub2: Arr(t, j) = Arr(LB, j): Next j
153:            For j = lb2 To ub2: Arr(LB, j) = temp(j): Next j
154:            t = t + 1: LB = LB - 1
155:        ElseIf t = LB Then
156:            t = t + 1: LB = LB - 1
157:        End If
158:    Loop While t <= LB
159:
160:    If t < bot Then QSort2_asc Arr(), C, t, bot, temp(), lb2, ub2
161:    If top < LB Then QSort2_asc Arr(), C, top, LB, temp(), lb2, ub2
162:
163: End Sub
Private Sub lbOK_Click()
165:    Dim oldWbName   As String
166:    Dim i As Integer, j As Integer
167:    Dim vbProj      As VBIDE.VBProject
168:    Dim vbComp      As VBIDE.VBComponent
169:    Dim arrNameFile() As String
170:    Dim WB          As Workbook
171:    Dim sPath       As String
172:
173:    oldWbName = cmbMain.Value
174:    Set WB = Workbooks(oldWbName)
175:
176:    Me.Hide
177:    If MsgBox("Вы полнить удаление форматирования кода ?", vbCritical + vbYesNo, "Удаление форматирования кода:") = vbYes Then
178:
179:        If Not WB.Name Like "*_obf_*" Then
180:            sPath = Left(WB.FullName, Len(WB.FullName) - Len(WB.Name))
181:            If sPath = vbNullString Then
182:                Call MsgBox("Файл не сохранен, для продолжения необходимо сохранить файл: [ " & WB.Name & " ]", vbInformation, "Ошибка:")
183:                Exit Sub
184:            End If
185:            arrNameFile = Split(WB.Name, ".")
186:            WB.SaveAs Filename:=sPath & arrNameFile(0) & "_obf_" & Replace(Now(), ":", ".") & "." & arrNameFile(1), FileFormat:=WB.FileFormat
187:        End If
188:        j = -1
189:
190:        Set vbProj = WB.VBProject
191:        For i = 0 To ListCode.ListCount - 1
192:            If ListCode.Selected(i) = True Then
193:                Set vbComp = vbProj.VBComponents(ListCode.List(i, 2))
194:
195:                If CheckDelNomerLine Then
196:                    Call K_AddNumbersLine.RemoveLineNumbers(vbComp, vbLineNumbers_LabelTypes.vbLabelColon)
197:                    Call K_AddNumbersLine.RemoveLineNumbers(vbComp, vbLineNumbers_LabelTypes.vbLabelTab)
198:                End If
199:                If CheckDelComment Then
200:                    Call N_Obfuscation.Remove_Comments(vbComp.CodeModule)
201:                End If
202:                If CheckDelFormat Then
203:                    Call N_Obfuscation.TrimLinesTabAndSpase(vbComp.CodeModule)
204:                End If
205:                If CheckDelExplicit Then
206:                    Call N_Obfuscation.Remove_OptionExplicit(vbComp.CodeModule)
207:                End If
208:                If CheckDelEmptyLines Then
209:                    Call N_Obfuscation.Remove_EmptyLines(vbComp.CodeModule)
210:                End If
211:                If CheckDelBreaksLines Then
212:                    Call N_Obfuscation.RemoveBreaksLineInCode(vbComp.CodeModule)
213:                End If
214:            End If
215:        Next i
216:
217:        WB.Save
218:        Call MsgBox("Удаление форматирования [" & oldWbName & "] завершено!", vbInformation, "Удаление форматирования:")
219:        cmbMain.Value = WB.Name
220:        Me.Show
221:        Exit Sub
222:    End If
223:    cmbMain.Value = WB.Name
224:    Me.Show
End Sub
