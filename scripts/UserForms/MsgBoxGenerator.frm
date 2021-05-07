VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MsgBoxGenerator 
   Caption         =   "MsgBox Generator:"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8070
   OleObjectBlob   =   "MsgBoxGenerator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MsgBoxGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : MsgBoxGenerator - конструктор MsgBox
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
    Private Sub btnCancel_Click()
10:    Unload Me
11: End Sub
    Private Sub lbHelp_Click()
13:    Call URLLinks(C_Const.URL_BILD_MSG)
14: End Sub
    Private Sub UserForm_Activate()
16:    Const W    As Integer = 20
17:    Const H    As Integer = W
18:    With Application.CommandBars
19:        obtnCritical.Picture = .GetImageMso("CancelRequest", W, H)
20:        obtnQuestion.Picture = .GetImageMso("ButtonTaskSelfSupport", W, H)
21:        obtnCaution.Picture = .GetImageMso("LogicIncomplete", W, H)
22:        obtnInformation.Picture = .GetImageMso("Info", W, H)
23:    End With
24:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
25: End Sub
    Private Sub txtMsg_Change()
27:    Call UnLocet(txtMsg, txtMsg2, lbClear2, lbStr2)
28: End Sub
    Private Sub txtMsg2_Change()
30:    Call UnLocet(txtMsg2, txtMsg3, lbClear3, lbStr3)
31: End Sub
    Private Sub txtMsg3_Change()
33:    Call UnLocet(txtMsg3, txtMsg4, lbClear4, lbStr4)
34: End Sub
    Private Sub txtMsg4_Change()
36:    Call UnLocet(txtMsg4, txtMsg5, lbClear5, lbStr5)
37: End Sub
    Private Sub UnLocet(ByRef txtMain As MSForms.Textbox, ByRef txtChild As MSForms.Textbox, ByRef lbClear As MSForms.Label, ByRef lbStr As MSForms.Label)
39:
40:    With txtChild
41:        If txtMain.Value = vbNullString Then
42:            .Enabled = False
43:            .Value = vbNullString
44:            lbStr.visible = True
45:            lbClear.Enabled = False
46:        Else
47:            .Enabled = True
48:            lbStr.visible = False
49:            lbClear.Enabled = True
50:        End If
51:    End With
52:
53: End Sub
    Private Sub lbClearTitle_Click()
55:    txtTitel.Value = vbNullString
56: End Sub
    Private Sub lbClear1_Click()
58:    txtMsg.Value = vbNullString
59: End Sub
    Private Sub lbClear2_Click()
61:    txtMsg2.Value = vbNullString
62: End Sub
    Private Sub lbClear3_Click()
64:    txtMsg3.Value = vbNullString
65: End Sub
    Private Sub lbClear4_Click()
67:    txtMsg4.Value = vbNullString
68: End Sub
    Private Sub lbClear5_Click()
70:    txtMsg5.Value = vbNullString
71: End Sub
'preview
    Private Sub btnView_Click()
74:    Dim i      As Long
75:    Dim sSTR    As String
76:    sSTR = ButtonVal()
77:    i = CInt(Split(sSTR, "|")(0)) + CInt(Split(sSTR, "|")(1))
78:    If chbMsgBoxRtlReading Then i = i + vbMsgBoxRtlReading
79:    'Me.Hide
80:    Call MsgBox(AddStringMsg(), i, txtTitel)
81:    'Me.Show
82: End Sub
     Private Function ButtonVal() As String
84:    Dim iButton As Integer
85:    Dim iButton1 As Integer
86:
87:    iButton = vbOKOnly
88:    If obtnOKCancel Then iButton = vbOKCancel
89:    If obtnYesNo Then iButton = vbYesNo
90:    If obtnRepeatCancel Then iButton = vbRetryCancel
91:    If obtnYesNoCancel Then iButton = vbYesNoCancel
92:    If obtnObortRepeatIgnor Then iButton = vbAbortRetryIgnore
93:
94:    iButton1 = 0
95:    If obtnCritical Then iButton1 = vbCritical
96:    If obtnQuestion Then iButton1 = vbQuestion
97:    If obtnCaution Then iButton1 = vbExclamation
98:    If obtnInformation Then iButton1 = vbInformation
99:    ButtonVal = iButton & "|" & iButton1
100: End Function
     Private Function AddCodeText() As String
102:    Dim sSTR As String, sBtn As String, sVal As String, sTextMsg As String, sTitelMsg As String
103: Dim sFerstSub As String, sEndSub As String
104:    Const sChr As String = "||| & Chr(34) & |||"
105:
106:    sVal = ButtonVal()
107:
108:    Select Case CInt(Split(sVal, "|")(0))
        Case 0
110:            sBtn = "vbOKOnly"
111: sFerstSub = "Call "
112: sEndSub = vbNullString
113:        Case 1
114:            sBtn = "vbOKCancel"
115: sFerstSub = "Call "
116: sEndSub = vbNullString
117:        Case 4
118:            sBtn = "vbYesNo"
119: sFerstSub = "If "
120: sEndSub = " = vbYes Then" & vbNewLine & "End If"
121:        Case 5
122:            sBtn = "vbRetryCancel"
123: sFerstSub = "If "
124: sEndSub = " = vbRetry Then" & vbNewLine & "End If"
125:        Case 3
126:            sBtn = "vbYesNoCancel"
127: sFerstSub = "Select Case "
sEndSub = vbNewLine & vbTab & "Case vbYes" & vbNewLine & vbTab & "Case vbNo" & vbNewLine & vbTab & "Case vbCancel" & vbNewLine & "End Select"
129:        Case 2
130:            sBtn = "vbAbortRetryIgnore"
131: sFerstSub = "Select Case "
sEndSub = vbNewLine & vbTab & "Case vbRetry" & vbNewLine & vbTab & "Case vbIgnore" & vbNewLine & vbTab & "Case vbAbort" & vbNewLine & "End Select"
133:    End Select
134:
135:    Select Case CInt(Split(sVal, "|")(1))
        Case 0
137:            sBtn = sBtn
138:        Case 16
139:            sBtn = sBtn & "+vbCritical"
140:        Case 32
141:            sBtn = sBtn & "+vbQuestion"
142:        Case 48
143:            sBtn = sBtn & "+vbExclamation"
144:        Case 64
145:            sBtn = sBtn & "+vbInformation"
146:    End Select
147:
148:    If chbMsgBoxRtlReading Then sBtn = sBtn & "+vbMsgBoxRtlReading"
149:
150:    sTitelMsg = Replace(txtTitel.Text, Chr(34), sChr)
151:    sTitelMsg = Replace(sTitelMsg, "|||", Chr(34))
152:    sTextMsg = AddStringMsg(True)
153:
154: sSTR = sFerstSub & "MsgBox(" & Chr(34)
155:    sSTR = sSTR & sTextMsg & Chr(34) & ", "
156:    sSTR = sSTR & sBtn & ", "
157:    sSTR = sSTR & Chr(34) & sTitelMsg & Chr(34) & ")" & sEndSub
158:
159:    AddCodeText = sSTR
160: End Function
     Private Sub btnCopyCode_Click()
162:
163:    Dim sMsgBoxString As String, sSTR As String
164:
165:    sSTR = AddCodeText()
166:    If sSTR = vbNullString Then Exit Sub
167:    If opbCliboard.Value = True Then
168:        Call C_PublicFunctions.SetTextIntoClipboard(sSTR)
169:        sMsgBoxString = "The code is copied to the clipboard!" & vbNewLine & "To insert the code, use" & Chr(34) & "Ctrl+V" & Chr(34)
170:    Else
171:        Debug.Print sSTR
172:        sMsgBoxString = "The code is printed in the window:" & Chr(34) & "Immediate" & Chr(34)
173:    End If
174:
175:    Call MsgBox(sMsgBoxString, vbInformation, "Copying the code:")
176:    'Unload Me
177:    Me.Hide
178: End Sub
     Private Sub lbInsertCode_Click()
180:    Dim sSTR As String, txtLine As String
181:    Dim iLine  As Integer
182:
183:    sSTR = AddCodeText()
184:    If sSTR = vbNullString Then Exit Sub
185:    txtLine = C_PublicFunctions.SelectedLineColumnProcedure
186:    If txtLine = vbNullString Then
187:        Me.Hide
188:        Exit Sub
189:    End If
190:    iLine = Split(txtLine, "|")(2)
191:
192:    With Application.VBE.ActiveCodePane
193:        .CodeModule.InsertLines iLine, sSTR
194:    End With
195:    Me.Hide
196: End Sub
Private Function AddStringMsg(Optional bFlag As Boolean) As String
198:    Dim sTextMsg As String
199:    Dim sSTR    As String
200:    Const sChr As String = "||| & Chr(34) & |||"
201:
202:    If bFlag Then
203:        sSTR = "||| & vbNewLine & |||"
204:    Else
205:        sSTR = vbNewLine
206:    End If
207:
208:    sTextMsg = txtMsg.Text
209:    If txtMsg2.Value <> vbNullString Then sTextMsg = txtMsg.Value & sSTR & txtMsg2.Value
210:    If txtMsg3.Value <> vbNullString Then sTextMsg = sTextMsg & sSTR & txtMsg3.Value
211:    If txtMsg4.Value <> vbNullString Then sTextMsg = sTextMsg & sSTR & txtMsg4.Value
212:    If txtMsg5.Value <> vbNullString Then sTextMsg = sTextMsg & sSTR & txtMsg5.Value
213:
214:    If bFlag Then
215:        sTextMsg = Replace(sTextMsg, Chr(34), sChr)
216:        sTextMsg = Replace(sTextMsg, "|||", Chr(34))
217:    End If
218:
219:    AddStringMsg = sTextMsg
End Function
