VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OptionsCodeFormat 
   Caption         =   "Settings:"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13950
   OleObjectBlob   =   "OptionsCodeFormat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OptionsCodeFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : OptionsCodeFormat - настройка форматирования кода, растановка отступов
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
    Private Sub cmbCancel_Click()
11:    ThisWorkbook.Save
12:    Unload Me
13: End Sub
    Private Sub lbCancel_Click()
15:    Call cmbCancel_Click
16: End Sub

    Private Sub lbHelp_Click()
19:    Call URLLinks(C_Const.URL_STYLE_STYLE)
20: End Sub

    Private Sub UserForm_Activate()
23:    Me.StartUpPosition = 0
24:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
25:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
26:
27:    Dim OptionsTb As ListObject
28:    Set OptionsTb = SHSNIPPETS.ListObjects(C_Const.TB_OPTIONSIDEDENT)
29:    Call UpdateCodeListBox
30:    With OptionsTb.ListColumns(2)
31:        SpinBtnTab.Value = .Range(2, 1)
32:        txtTab.Value = .Range(2, 1)
33:        chbIndentProc.Value = .Range(3, 1)
34:        chbIndentFirst.Value = .Range(4, 1)
35:        chbIndentDim.Value = .Range(5, 1)
36:        chbIndentCmt.Value = .Range(6, 1)
37:        chbIndentCase.Value = .Range(7, 1)
38:        chbAlignCont.Value = .Range(8, 1)
39:        chbAlignIgnoreOps.Value = .Range(9, 1)
40:        chbDebugCol1.Value = .Range(10, 1)
41:        chbAlignDim.Value = .Range(11, 1)
42:        SpinBtn.Value = .Range(12, 1)
43:        txtmiAlignDimCol.Value = .Range(12, 1)
44:        chbCompilerStuffCol1.Value = .Range(13, 1)
45:        chbIndentCompilerStuff.Value = .Range(14, 1)
46:        SpinBtnComment.Value = .Range(16, 1)
47:        txtComment.Value = .Range(16, 1)
48:
49:        Select Case .Range(15, 1)
            Case "Absolute":
51:                obtnAbsolute.Value = True
52:            Case "SameGap":
53:                obtnSameGap.Value = True
54:            Case "StandardGap":
55:                obtnStandardGap.Value = True
56:            Case "AlignInCol":
57:                obtnAlignInCol.Value = True
58:        End Select
59:    End With
60:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
61: End Sub
    Private Sub txtTab_Change()
63:    Call SetOptFromTable(2, txtTab.Value)
64: End Sub
    Private Sub chbIndentProc_Change()
66:    chbIndentFirst.Enabled = chbIndentProc.Value
67:    chbIndentDim.Enabled = chbIndentProc.Value
68:    Call SetOptFromTable(3, chbIndentProc.Value)
69: End Sub
    Private Sub chbIndentFirst_Change()
71:    Call SetOptFromTable(4, chbIndentFirst.Value)
72: End Sub
    Private Sub chbIndentDim_Change()
74:    Call SetOptFromTable(5, chbIndentDim.Value)
75: End Sub
    Private Sub chbIndentCmt_Change()
77:    Call SetOptFromTable(6, chbIndentCmt.Value)
78: End Sub
    Private Sub chbIndentCase_Change()
80:    Call SetOptFromTable(7, chbIndentCase.Value)
81: End Sub
    Private Sub chbAlignCont_Change()
83:    If chbAlignCont Then
84:        chbAlignIgnoreOps.Value = False
85:    Else
86:        chbAlignIgnoreOps.Value = True
87:    End If
88:    Call SetOptFromTable(8, chbAlignCont.Value)
89: End Sub
    Private Sub chbAlignIgnoreOps_Change()
91:    If chbAlignIgnoreOps Then chbAlignCont.Value = False
92:    Call SetOptFromTable(9, chbAlignIgnoreOps.Value)
93: End Sub
    Private Sub chbDebugCol1_Change()
95:    Call SetOptFromTable(10, chbDebugCol1.Value)
96: End Sub
     Private Sub chbAlignDim_Change()
98:    txtmiAlignDimCol.Enabled = chbAlignDim.Value
99:    SpinBtn.Enabled = chbAlignDim.Value
100:    Call SetOptFromTable(11, chbAlignDim.Value)
101: End Sub
     Private Sub txtmiAlignDimCol_Change()
103:    Call SetOptFromTable(12, txtmiAlignDimCol.Value)
104: End Sub
     Private Sub SpinBtn_SpinDown()
106:    Call SpinBtnChange(0, 30, Me.SpinBtn, Me.txtmiAlignDimCol)
107: End Sub
     Private Sub SpinBtn_SpinUp()
109:    Call SpinBtnChange(0, 30, Me.SpinBtn, Me.txtmiAlignDimCol)
110: End Sub
     Private Sub SpinBtnTab_SpinDown()
112:    Call SpinBtnChange(4, 8, Me.SpinBtnTab, Me.txtTab)
113: End Sub
     Private Sub SpinBtnTab_SpinUp()
115:    Call SpinBtnChange(4, 8, Me.SpinBtnTab, Me.txtTab)
116: End Sub
     Private Sub SpinBtnComment_SpinDown()
118:    Call SpinBtnChange(0, 100, Me.SpinBtnComment, Me.txtComment)
119: End Sub
     Private Sub SpinBtnComment_SpinUp()
121:    Call SpinBtnChange(0, 100, Me.SpinBtnComment, Me.txtComment)
122: End Sub
     Private Sub SpinBtnChange(ByVal iMin As Byte, ByVal iMax As Byte, ByRef objSpinBtn As MSForms.SpinButton, ByRef objTxt As MSForms.Textbox)
124:    With objSpinBtn
125:        If .Value < iMin Then .Value = iMin
126:        If .Value > iMax Then .Value = iMax
127:        objTxt.Text = .Value
128:    End With
129: End Sub
     Private Sub chbCompilerStuffCol1_Change()
131:    If chbCompilerStuffCol1 Then
132:        chbIndentCompilerStuff.Value = False
133:    Else
134:        chbIndentCompilerStuff.Value = True
135:    End If
136:    Call SetOptFromTable(13, chbCompilerStuffCol1.Value)
137: End Sub
     Private Sub chbIndentCompilerStuff_Change()
139:    If chbIndentCompilerStuff Then chbCompilerStuffCol1.Value = False
140:    Call SetOptFromTable(14, chbIndentCompilerStuff.Value)
141: End Sub
     Private Sub obtnAbsolute_Change()
143:    Call SetOptFromTable(15, obtnAbsolute.Tag)
144: End Sub
     Private Sub obtnAlignInCol_Change()
146:    txtComment.Enabled = obtnAlignInCol.Value
147:    SpinBtnComment.Enabled = obtnAlignInCol.Value
148:    Call SetOptFromTable(15, obtnAlignInCol.Tag)
149: End Sub
     Private Sub obtnSameGap_Change()
151:    Call SetOptFromTable(15, obtnSameGap.Tag)
152: End Sub
     Private Sub obtnStandardGap_Change()
154:    Call SetOptFromTable(15, obtnStandardGap.Tag)
155: End Sub
     Private Sub txtComment_Change()
157:    Call SetOptFromTable(16, txtComment.Value)
158: End Sub
     Private Sub UpdateCodeListBox()
160:
161:    Dim asCodeLines(1 To 30) As String
162:    Dim i      As Integer
163:
164:    'Define the example procedure code lines
165:    asCodeLines(1) = "' Example Procedure"
154:    asCodeLines(2) = "Sub ExampleProc()"
167:    asCodeLines(3) = ""
168:    asCodeLines(4) = "'надстройка " & C_Const.NAME_ADDIN
169:    asCodeLines(5) = "'© 2018-" & VBA.Year(Now()) & " by " & C_Const.NAME_ADDIN & " Ltd."
170:    asCodeLines(6) = ""
171:    asCodeLines(7) = "Dim iCount As Integer"
172:    asCodeLines(8) = "Static sName As String"
173:    asCodeLines(9) = ""
174:    asCodeLines(10) = "If YouWantMoreExamplesAndTools Then"
175:    asCodeLines(11) = "' Visit http://www.a.com"
176:    asCodeLines(12) = ""
177:    asCodeLines(13) = "Select Case X"
    asCodeLines(14) = "Case ""A"""
179:    asCodeLines(15) = "' If you have any comments or suggestions, _"
180:    asCodeLines(16) = " or find valid VBA code that isn't indented correctly,"
181:    asCodeLines(17) = ""
182:    asCodeLines(18) = "#If VBA6 Then"
183:    asCodeLines(19) = "MsgBox ""Contact A@A.com"""
184:    asCodeLines(20) = "#End If"
185:    asCodeLines(21) = ""
186:    asCodeLines(22) = "Case ""Continued strings and parameters can be"" _"
187:    asCodeLines(23) = "& ""lined up for easier reading, optionally ignoring"" _"
188:    asCodeLines(24) = ", ""any operators (&+, etc) at the start of the line."""
189:    asCodeLines(25) = ""
190:    asCodeLines(26) = "Debug.Print ""X<>1"""
191:    asCodeLines(27) = "End Select           'Case X"
192:    asCodeLines(28) = "End If               'More Tools?"
193:    asCodeLines(29) = ""
194:    asCodeLines(30) = "End Sub"
195:
196:
197:    'Run the array through the indenting code
198:    RebuildCodeArray asCodeLines, "", 0
199:
200:    'Put the procedure code in the list box.
201:
202:    txtCode.Text = vbNullString
203:    For i = LBound(asCodeLines) To UBound(asCodeLines)
204:        If i = UBound(asCodeLines) Then
205:            txtCode.Text = txtCode.Text & asCodeLines(i)
206:        Else
207:            txtCode.Text = txtCode.Text & asCodeLines(i) & vbNewLine
208:        End If
209:    Next
210: End Sub

Private Sub SetOptFromTable(ByVal iRow As Byte, ByVal iVal As Variant)
213:    Dim OptionsTb As ListObject
214:    Set OptionsTb = SHSNIPPETS.ListObjects(C_Const.TB_OPTIONSIDEDENT)
215:    OptionsTb.ListColumns(2).Range(iRow, 1) = iVal
216:    Call UpdateCodeListBox
End Sub
