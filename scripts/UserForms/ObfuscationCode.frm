VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ObfuscationCode 
   Caption         =   "Removing Formatting:"
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
10:    Unload Me
11: End Sub

    Private Sub lbCancel_Click()
14:    Call cmbCancel_Click
15: End Sub

    Private Sub lbHelp_Click()
18:    Call URLLinks(C_Const.URL_FILE_OBFS)
19: End Sub

    Private Sub UserForm_Activate()
22:    Me.StartUpPosition = 0
23:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
24:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
25:
26:    Dim WB          As Workbook
27:    Dim vbProj      As VBIDE.VBProject
28:
29:    On Error GoTo ErrorHandler
30:    If Workbooks.Count = 0 Then
31:        Unload Me
32:        Call MsgBox("No open ones" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
33:        Exit Sub
34:    End If
35:    With Me.cmbMain
36:        .Clear
37:        On Error Resume Next
38:        For Each vbProj In Application.VBE.VBProjects
39:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
40:        Next
41:        On Error GoTo 0
42:        On Error GoTo ErrorHandler
43:        .Value = ActiveWorkbook.Name
44:    End With
45:    Call AddListCode
46:    lbMsg.visible = True
47:    lbOK.Enabled = False
48:
49:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
50:    Exit Sub
ErrorHandler:
52:    Unload Me
53:    Select Case Err.Number
        Case Else:
55:            Call MsgBox("Error in the ObfuscationCode.UserForm_Activate" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbExclamation, "Error:")
56:            Call WriteErrorLog("ObfuscationCode.UserForm_Activate")
57:    End Select
58:    Err.Clear
59: End Sub
    Private Sub CheckAll_Click()
61:    Dim i           As Integer
62:    With ListCode
63:        For i = 0 To .ListCount - 1
64:            .Selected(i) = CheckAll.Value
65:        Next i
66:    End With
67: End Sub
    Private Sub CheckAllParam_Click()
69:    If CheckAllParam.Value = True Then
70:        CheckAllParam.Value = True
71:        CheckDelComment.Value = True
72:        CheckDelEmptyLines.Value = True
73:        CheckDelExplicit.Value = True
74:        CheckDelFormat.Value = True
75:        CheckDelNomerLine.Value = True
76:        CheckDelBreaksLines.Value = True
77:    Else
78:        CheckAllParam.Value = False
79:        CheckDelComment.Value = False
80:        CheckDelEmptyLines.Value = False
81:        CheckDelExplicit.Value = False
82:        CheckDelFormat.Value = False
83:        CheckDelNomerLine.Value = False
84:        CheckDelBreaksLines.Value = False
85:    End If
86: End Sub

    Private Sub cmbMain_Change()
89:    If cmbMain.Value <> vbNullString Then
90:        Call AddListCode
91:    End If
92: End Sub
     Private Sub ListCode_Change()
94:    Dim i           As Integer
95:    With ListCode
96:        For i = 0 To .ListCount - 1
97:            If .Selected(i) Then
98:                lbMsg.visible = False
99:                lbOK.Enabled = True
100:                Exit Sub
101:            End If
102:        Next i
103:    End With
104:    lbMsg.visible = True
105:    lbOK.Enabled = False
106: End Sub
     Private Sub AddListCode()
108:    Dim WB          As Workbook
109:    Dim iFile       As Integer
110:    Dim Arr()       As Variant
111:    Set WB = Workbooks(cmbMain.Value)
112:    With ListCode
113:        .Clear
114:        For iFile = 1 To WB.VBProject.VBComponents.Count
115:            .AddItem iFile
116:            .List(iFile - 1, 1) = ComponentTypeToString(WB.VBProject.VBComponents(iFile).Type)
117:            .List(iFile - 1, 2) = WB.VBProject.VBComponents(iFile).Name
118:        Next iFile
119:        Arr = .List
120:        Call Sort2_asc(Arr, 1)
121:        .List = Arr
122:        For iFile = 0 To .ListCount - 1
123:            .List(iFile, 0) = iFile + 1
124:        Next iFile
125:    End With
126: End Sub
'сортировка массива
     Private Sub Sort2_asc(Arr(), col As Long)
129:    Dim temp()      As Variant
130:    Dim lb2 As Long, ub2 As Long, lTop As Long, lBot As Long
131:
132:    lTop = LBound(Arr, 1)
133:    lBot = UBound(Arr, 1)
134:    lb2 = LBound(Arr, 2)
135:    ub2 = UBound(Arr, 2)
136:    ReDim temp(lb2 To ub2)
137:
138:    Call QSort2_asc(Arr(), col, lTop, lBot, temp(), lb2, ub2)
139: End Sub
     Private Sub QSort2_asc(Arr(), C As Long, ByVal top As Long, ByVal bot As Long, temp(), lb2 As Long, ub2 As Long)
141:    Dim t As Long, LB As Long, MidItem, j As Long
142:
143:    MidItem = Arr((top + bot) \ 2, C)
144:    t = top: LB = bot
145:
146:    Do
147:        Do While Arr(t, C) < MidItem: t = t + 1: Loop
148:        Do While Arr(LB, C) > MidItem: LB = LB - 1: Loop
149:        If t < LB Then
150:            For j = lb2 To ub2: temp(j) = Arr(t, j): Next j
151:            For j = lb2 To ub2: Arr(t, j) = Arr(LB, j): Next j
152:            For j = lb2 To ub2: Arr(LB, j) = temp(j): Next j
153:            t = t + 1: LB = LB - 1
154:        ElseIf t = LB Then
155:            t = t + 1: LB = LB - 1
156:        End If
157:    Loop While t <= LB
158:
159:    If t < bot Then QSort2_asc Arr(), C, t, bot, temp(), lb2, ub2
160:    If top < LB Then QSort2_asc Arr(), C, top, LB, temp(), lb2, ub2
161:
162: End Sub
Private Sub lbOK_Click()
164:    Dim oldWbName   As String
165:    Dim i As Integer, j As Integer
166:    Dim vbProj      As VBIDE.VBProject
167:    Dim vbComp      As VBIDE.VBComponent
168:    Dim arrNameFile() As String
169:    Dim WB          As Workbook
170:    Dim sPath       As String
171:
172:    oldWbName = cmbMain.Value
173:    Set WB = Workbooks(oldWbName)
174:
175:    Me.Hide
176:    If MsgBox("Perform code formatting removal ?", vbCritical + vbYesNo, "Removing the code formatting:") = vbYes Then
177:
178:        If Not WB.Name Like "*_obf_*" Then
179:            sPath = Left(WB.FullName, Len(WB.FullName) - Len(WB.Name))
180:            If sPath = vbNullString Then
181:                Call MsgBox("The file is not saved, you need to save the file to continue: [" & WB.Name & " ]", vbInformation, "Error:")
182:                Exit Sub
183:            End If
184:            arrNameFile = Split(WB.Name, ".")
185:            WB.SaveAs Filename:=sPath & arrNameFile(0) & "_obf_" & Replace(Now(), ":", ".") & "." & arrNameFile(1), FileFormat:=WB.FileFormat
186:        End If
187:        j = -1
188:
189:        Set vbProj = WB.VBProject
190:        For i = 0 To ListCode.ListCount - 1
191:            If ListCode.Selected(i) = True Then
192:                Set vbComp = vbProj.VBComponents(ListCode.List(i, 2))
193:
194:                If CheckDelNomerLine Then
195:                    Call K_AddNumbersLine.RemoveLineNumbers(vbComp, vbLineNumbers_LabelTypes.vbLabelColon)
196:                    Call K_AddNumbersLine.RemoveLineNumbers(vbComp, vbLineNumbers_LabelTypes.vbLabelTab)
197:                End If
198:                If CheckDelComment Then
199:                    Call N_Obfuscation.Remove_Comments(vbComp.CodeModule)
200:                End If
201:                If CheckDelFormat Then
202:                    Call N_Obfuscation.TrimLinesTabAndSpase(vbComp.CodeModule)
203:                End If
204:                If CheckDelExplicit Then
205:                    Call N_Obfuscation.Remove_OptionExplicit(vbComp.CodeModule)
206:                End If
207:                If CheckDelEmptyLines Then
208:                    Call N_Obfuscation.Remove_EmptyLines(vbComp.CodeModule)
209:                End If
210:                If CheckDelBreaksLines Then
211:                    Call N_Obfuscation.RemoveBreaksLineInCode(vbComp.CodeModule)
212:                End If
213:            End If
214:        Next i
215:
216:        WB.Save
217:        Call MsgBox("Removing formatting [" & oldWbName & "] completed!", vbInformation, "Delete formatting:")
218:        cmbMain.Value = WB.Name
219:        Me.Show
220:        Exit Sub
221:    End If
222:    cmbMain.Value = WB.Name
223:    Me.Show
End Sub
