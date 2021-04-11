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
4:    Unload Me
5: End Sub
  Private Sub lbHelp_Click()
8:    Call URLLinks(C_Const.URL_BILD_MSG)
9: End Sub
    Private Sub UserForm_Activate()
12:    Const W    As Integer = 20
13:    Const H    As Integer = W
14:    With Application.CommandBars
15:        obtnCritical.Picture = .GetImageMso("CancelRequest", W, H)
16:        obtnQuestion.Picture = .GetImageMso("ButtonTaskSelfSupport", W, H)
17:        obtnCaution.Picture = .GetImageMso("LogicIncomplete", W, H)
18:        obtnInformation.Picture = .GetImageMso("Info", W, H)
19:    End With
20:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
21: End Sub
    Private Sub txtMsg_Change()
23:    Call UnLocet(txtMsg, txtMsg2, lbClear2, lbStr2)
24: End Sub
    Private Sub txtMsg2_Change()
26:    Call UnLocet(txtMsg2, txtMsg3, lbClear3, lbStr3)
27: End Sub
    Private Sub txtMsg3_Change()
29:    Call UnLocet(txtMsg3, txtMsg4, lbClear4, lbStr4)
30: End Sub
    Private Sub txtMsg4_Change()
32:    Call UnLocet(txtMsg4, txtMsg5, lbClear5, lbStr5)
33: End Sub
    Private Sub UnLocet(ByRef txtMain As MSForms.Textbox, ByRef txtChild As MSForms.Textbox, ByRef lbClear As MSForms.Label, ByRef lbStr As MSForms.Label)
35:
36:    With txtChild
37:        If txtMain.Value = vbNullString Then
38:            .Enabled = False
39:            .Value = vbNullString
40:            lbStr.visible = True
41:            lbClear.Enabled = False
42:        Else
43:            .Enabled = True
44:            lbStr.visible = False
45:            lbClear.Enabled = True
46:        End If
47:    End With
48:
49: End Sub
    Private Sub lbClearTitle_Click()
51:    txtTitel.Value = vbNullString
52: End Sub
    Private Sub lbClear1_Click()
54:    txtMsg.Value = vbNullString
55: End Sub
    Private Sub lbClear2_Click()
57:    txtMsg2.Value = vbNullString
58: End Sub
    Private Sub lbClear3_Click()
60:    txtMsg3.Value = vbNullString
61: End Sub
    Private Sub lbClear4_Click()
63:    txtMsg4.Value = vbNullString
64: End Sub
    Private Sub lbClear5_Click()
66:    txtMsg5.Value = vbNullString
67: End Sub
'preview
    Private Sub btnView_Click()
70:    Dim i      As Long
71:    Dim sSTR    As String
72:    sSTR = ButtonVal()
73:    i = CInt(Split(sSTR, "|")(0)) + CInt(Split(sSTR, "|")(1))
74:    If chbMsgBoxRtlReading Then i = i + vbMsgBoxRtlReading
75:    'Me.Hide
76:    Call MsgBox(AddStringMsg(), i, txtTitel)
77:    'Me.Show
78: End Sub
    Private Function ButtonVal() As String
80:    Dim iButton As Integer
81:    Dim iButton1 As Integer
82:
83:    iButton = vbOKOnly
84:    If obtnOKCancel Then iButton = vbOKCancel
85:    If obtnYesNo Then iButton = vbYesNo
86:    If obtnRepeatCancel Then iButton = vbRetryCancel
87:    If obtnYesNoCancel Then iButton = vbYesNoCancel
88:    If obtnObortRepeatIgnor Then iButton = vbAbortRetryIgnore
89:
90:    iButton1 = 0
91:    If obtnCritical Then iButton1 = vbCritical
92:    If obtnQuestion Then iButton1 = vbQuestion
93:    If obtnCaution Then iButton1 = vbExclamation
94:    If obtnInformation Then iButton1 = vbInformation
95:    ButtonVal = iButton & "|" & iButton1
96: End Function
     Private Function AddCodeText() As String
98:    Dim sSTR As String, sBtn As String, sVal As String, sTextMsg As String, sTitelMsg As String
93:     Dim sFerstSub As String, sEndSub As String
100:    Const sChr As String = "||| & Chr(34) & |||"
101:
102:    sVal = ButtonVal()
103:
104:    Select Case CInt(Split(sVal, "|")(0))
        Case 0
106:            sBtn = "vbOKOnly"
107:            sFerstSub = "Call "
108:            sEndSub = vbNullString
109:        Case 1
110:            sBtn = "vbOKCancel"
105:             sFerstSub = "Call "
112:            sEndSub = vbNullString
113:        Case 4
114:            sBtn = "vbYesNo"
115:            sFerstSub = "If "
116:            sEndSub = " = vbYes Then" & vbNewLine & "End If"
117:        Case 5
118:            sBtn = "vbRetryCancel"
119:            sFerstSub = "If "
120:            sEndSub = " = vbRetry Then" & vbNewLine & "End If"
121:        Case 3
122:            sBtn = "vbYesNoCancel"
123:            sFerstSub = "Select Case "
            sEndSub = vbNewLine & vbTab & "Case vbYes" & vbNewLine & vbTab & "Case vbNo" & vbNewLine & vbTab & "Case vbCancel" & vbNewLine & "End Select"
125:        Case 2
126:            sBtn = "vbAbortRetryIgnore"
127:            sFerstSub = "Select Case "
            sEndSub = vbNewLine & vbTab & "Case vbRetry" & vbNewLine & vbTab & "Case vbIgnore" & vbNewLine & vbTab & "Case vbAbort" & vbNewLine & "End Select"
129:    End Select
130:
131:    Select Case CInt(Split(sVal, "|")(1))
        Case 0
133:            sBtn = sBtn
134:        Case 16
135:            sBtn = sBtn & "+vbCritical"
136:        Case 32
137:            sBtn = sBtn & "+vbQuestion"
138:        Case 48
139:            sBtn = sBtn & "+vbExclamation"
140:        Case 64
141:            sBtn = sBtn & "+vbInformation"
142:    End Select
143:
144:    If chbMsgBoxRtlReading Then sBtn = sBtn & "+vbMsgBoxRtlReading"
145:
146:    sTitelMsg = Replace(txtTitel.Text, Chr(34), sChr)
147:    sTitelMsg = Replace(sTitelMsg, "|||", Chr(34))
148:    sTextMsg = AddStringMsg(True)
149:
150:    sSTR = sFerstSub & "MsgBox(" & Chr(34)
151:    sSTR = sSTR & sTextMsg & Chr(34) & ", "
152:    sSTR = sSTR & sBtn & ", "
153:    sSTR = sSTR & Chr(34) & sTitelMsg & Chr(34) & ")" & sEndSub
154:
155:    AddCodeText = sSTR
156: End Function
     Private Sub btnCopyCode_Click()
158:
159:    Dim sMsgBoxString As String, sSTR As String
160:
161:    sSTR = AddCodeText()
162:    If sSTR = vbNullString Then Exit Sub
163:    If opbCliboard.Value = True Then
164:        Call C_PublicFunctions.SetTextIntoClipboard(sSTR)
165:        sMsgBoxString = "The code is copied to the clipboard!" & vbNewLine & "To insert the code, use" & Chr(34) & "Ctrl+V" & Chr(34)
166:    Else
167:        Debug.Print sSTR
168:        sMsgBoxString = "The code is printed in the window:" & Chr(34) & "Immediate" & Chr(34)
169:    End If
170:
171:    Call MsgBox(sMsgBoxString, vbInformation, "Copying the code:")
172:    'Unload Me
173:    Me.Hide
174: End Sub
     Private Sub lbInsertCode_Click()
176:    Dim sSTR As String, txtLine As String
177:    Dim iLine  As Integer
178:
179:    sSTR = AddCodeText()
180:    If sSTR = vbNullString Then Exit Sub
181:    txtLine = C_PublicFunctions.SelectedLineColumnProcedure
182:    If txtLine = vbNullString Then
183:        Me.Hide
184:        Exit Sub
185:    End If
186:    iLine = Split(txtLine, "|")(2)
187:
188:    With Application.VBE.ActiveCodePane
189:        .CodeModule.InsertLines iLine, sSTR
190:    End With
191:    Me.Hide
192: End Sub
Private Function AddStringMsg(Optional bFlag As Boolean) As String
194:    Dim sTextMsg As String
195:    Dim sSTR    As String
196:    Const sChr As String = "||| & Chr(34) & |||"
197:
198:    If bFlag Then
199:        sSTR = "||| & vbNewLine & |||"
200:    Else
201:        sSTR = vbNewLine
202:    End If
203:
204:    sTextMsg = txtMsg.Text
205:    If txtMsg2.Value <> vbNullString Then sTextMsg = txtMsg.Value & sSTR & txtMsg2.Value
206:    If txtMsg3.Value <> vbNullString Then sTextMsg = sTextMsg & sSTR & txtMsg3.Value
207:    If txtMsg4.Value <> vbNullString Then sTextMsg = sTextMsg & sSTR & txtMsg4.Value
208:    If txtMsg5.Value <> vbNullString Then sTextMsg = sTextMsg & sSTR & txtMsg5.Value
209:
210:    If bFlag Then
211:        sTextMsg = Replace(sTextMsg, Chr(34), sChr)
212:        sTextMsg = Replace(sTextMsg, "|||", Chr(34))
213:    End If
214:
215:    AddStringMsg = sTextMsg
End Function
