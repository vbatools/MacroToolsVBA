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
10:    ThisWorkbook.Save
11:    Unload Me
12: End Sub
    Private Sub lbCancel_Click()
14:    Call cmbCancel_Click
15: End Sub

    Private Sub lbHelp_Click()
18:    Call URLLinks(C_Const.URL_STYLE_STYLE)
19: End Sub

    Private Sub UserForm_Activate()
22:    Me.StartUpPosition = 0
23:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
24:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
25:
26:    Dim OptionsTb As ListObject
27:    Set OptionsTb = SHSNIPPETS.ListObjects(C_Const.TB_OPTIONSIDEDENT)
28:    Call UpdateCodeListBox
29:    With OptionsTb.ListColumns(2)
30:        SpinBtnTab.Value = .Range(2, 1)
31:        txtTab.Value = .Range(2, 1)
32:        chbIndentProc.Value = .Range(3, 1)
33:        chbIndentFirst.Value = .Range(4, 1)
34:        chbIndentDim.Value = .Range(5, 1)
35:        chbIndentCmt.Value = .Range(6, 1)
36:        chbIndentCase.Value = .Range(7, 1)
37:        chbAlignCont.Value = .Range(8, 1)
38:        chbAlignIgnoreOps.Value = .Range(9, 1)
39:        chbDebugCol1.Value = .Range(10, 1)
40:        chbAlignDim.Value = .Range(11, 1)
41:        SpinBtn.Value = .Range(12, 1)
42:        txtmiAlignDimCol.Value = .Range(12, 1)
43:        chbCompilerStuffCol1.Value = .Range(13, 1)
44:        chbIndentCompilerStuff.Value = .Range(14, 1)
45:        SpinBtnComment.Value = .Range(16, 1)
46:        txtComment.Value = .Range(16, 1)
47:
48:        Select Case .Range(15, 1)
            Case "Absolute":
50:                obtnAbsolute.Value = True
51:            Case "SameGap":
52:                obtnSameGap.Value = True
53:            Case "StandardGap":
54:                obtnStandardGap.Value = True
55:            Case "AlignInCol":
56:                obtnAlignInCol.Value = True
57:        End Select
58:    End With
59:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
60: End Sub
    Private Sub txtTab_Change()
62:    Call SetOptFromTable(2, txtTab.Value)
63: End Sub
    Private Sub chbIndentProc_Change()
65:    chbIndentFirst.Enabled = chbIndentProc.Value
66:    chbIndentDim.Enabled = chbIndentProc.Value
67:    Call SetOptFromTable(3, chbIndentProc.Value)
68: End Sub
    Private Sub chbIndentFirst_Change()
70:    Call SetOptFromTable(4, chbIndentFirst.Value)
71: End Sub
    Private Sub chbIndentDim_Change()
73:    Call SetOptFromTable(5, chbIndentDim.Value)
74: End Sub
    Private Sub chbIndentCmt_Change()
76:    Call SetOptFromTable(6, chbIndentCmt.Value)
77: End Sub
    Private Sub chbIndentCase_Change()
79:    Call SetOptFromTable(7, chbIndentCase.Value)
80: End Sub
    Private Sub chbAlignCont_Change()
82:    If chbAlignCont Then
83:        chbAlignIgnoreOps.Value = False
84:    Else
85:        chbAlignIgnoreOps.Value = True
86:    End If
87:    Call SetOptFromTable(8, chbAlignCont.Value)
88: End Sub
    Private Sub chbAlignIgnoreOps_Change()
90:    If chbAlignIgnoreOps Then chbAlignCont.Value = False
91:    Call SetOptFromTable(9, chbAlignIgnoreOps.Value)
92: End Sub
    Private Sub chbDebugCol1_Change()
94:    Call SetOptFromTable(10, chbDebugCol1.Value)
95: End Sub
     Private Sub chbAlignDim_Change()
97:    txtmiAlignDimCol.Enabled = chbAlignDim.Value
98:    SpinBtn.Enabled = chbAlignDim.Value
99:    Call SetOptFromTable(11, chbAlignDim.Value)
100: End Sub
     Private Sub txtmiAlignDimCol_Change()
102:    Call SetOptFromTable(12, txtmiAlignDimCol.Value)
103: End Sub
     Private Sub SpinBtn_SpinDown()
105:    Call SpinBtnChange(0, 30, Me.SpinBtn, Me.txtmiAlignDimCol)
106: End Sub
     Private Sub SpinBtn_SpinUp()
108:    Call SpinBtnChange(0, 30, Me.SpinBtn, Me.txtmiAlignDimCol)
109: End Sub
     Private Sub SpinBtnTab_SpinDown()
111:    Call SpinBtnChange(4, 8, Me.SpinBtnTab, Me.txtTab)
112: End Sub
     Private Sub SpinBtnTab_SpinUp()
114:    Call SpinBtnChange(4, 8, Me.SpinBtnTab, Me.txtTab)
115: End Sub
     Private Sub SpinBtnComment_SpinDown()
117:    Call SpinBtnChange(0, 100, Me.SpinBtnComment, Me.txtComment)
118: End Sub
     Private Sub SpinBtnComment_SpinUp()
120:    Call SpinBtnChange(0, 100, Me.SpinBtnComment, Me.txtComment)
121: End Sub
     Private Sub SpinBtnChange(ByVal iMin As Byte, ByVal iMax As Byte, ByRef objSpinBtn As MSForms.SpinButton, ByRef objTxt As MSForms.Textbox)
123:    With objSpinBtn
124:        If .Value < iMin Then .Value = iMin
125:        If .Value > iMax Then .Value = iMax
126:        objTxt.Text = .Value
127:    End With
128: End Sub
     Private Sub chbCompilerStuffCol1_Change()
130:    If chbCompilerStuffCol1 Then
131:        chbIndentCompilerStuff.Value = False
132:    Else
133:        chbIndentCompilerStuff.Value = True
134:    End If
135:    Call SetOptFromTable(13, chbCompilerStuffCol1.Value)
136: End Sub
     Private Sub chbIndentCompilerStuff_Change()
138:    If chbIndentCompilerStuff Then chbCompilerStuffCol1.Value = False
139:    Call SetOptFromTable(14, chbIndentCompilerStuff.Value)
140: End Sub
     Private Sub obtnAbsolute_Change()
142:    Call SetOptFromTable(15, obtnAbsolute.Tag)
143: End Sub
     Private Sub obtnAlignInCol_Change()
145:    txtComment.Enabled = obtnAlignInCol.Value
146:    SpinBtnComment.Enabled = obtnAlignInCol.Value
147:    Call SetOptFromTable(15, obtnAlignInCol.Tag)
148: End Sub
     Private Sub obtnSameGap_Change()
150:    Call SetOptFromTable(15, obtnSameGap.Tag)
151: End Sub
     Private Sub obtnStandardGap_Change()
153:    Call SetOptFromTable(15, obtnStandardGap.Tag)
154: End Sub
     Private Sub txtComment_Change()
156:    Call SetOptFromTable(16, txtComment.Value)
157: End Sub
     Private Sub UpdateCodeListBox()
159:
160:    Dim asCodeLines(1 To 30) As String
161:    Dim i      As Integer
162:
163:    'Define the example procedure code lines
164:    asCodeLines(1) = "' Example Procedure"
165: asCodeLines(2) = "Sub ExampleProc()"
166:    asCodeLines(3) = ""
167:    asCodeLines(4) = "'надстройка " & C_Const.NAME_ADDIN
168:    asCodeLines(5) = "'© 2018-" & VBA.Year(Now()) & " by " & C_Const.NAME_ADDIN & " Ltd."
169:    asCodeLines(6) = ""
170:    asCodeLines(7) = "Dim iCount As Integer"
171:    asCodeLines(8) = "Static sName As String"
172:    asCodeLines(9) = ""
173:    asCodeLines(10) = "If YouWantMoreExamplesAndTools Then"
174:    asCodeLines(11) = "' Visit http://www.a.com"
175:    asCodeLines(12) = ""
176:    asCodeLines(13) = "Select Case X"
    asCodeLines(14) = "Case ""A"""
178:    asCodeLines(15) = "' If you have any comments or suggestions, _"
179:    asCodeLines(16) = " or find valid VBA code that isn't indented correctly,"
180:    asCodeLines(17) = ""
181:    asCodeLines(18) = "#If VBA6 Then"
182:    asCodeLines(19) = "MsgBox ""Contact A@A.com"""
183:    asCodeLines(20) = "#End If"
184:    asCodeLines(21) = ""
185:    asCodeLines(22) = "Case ""Continued strings and parameters can be"" _"
186:    asCodeLines(23) = "& ""lined up for easier reading, optionally ignoring"" _"
187:    asCodeLines(24) = ", ""any operators (&+, etc) at the start of the line."""
188:    asCodeLines(25) = ""
189:    asCodeLines(26) = "Debug.Print ""X<>1"""
190:    asCodeLines(27) = "End Select           'Case X"
191:    asCodeLines(28) = "End If               'More Tools?"
192:    asCodeLines(29) = ""
193:    asCodeLines(30) = "End Sub"
194:
195:
196:    'Run the array through the indenting code
197:    RebuildCodeArray asCodeLines, "", 0
198:
199:    'Put the procedure code in the list box.
200:
201:    txtCode.Text = vbNullString
202:    For i = LBound(asCodeLines) To UBound(asCodeLines)
203:        If i = UBound(asCodeLines) Then
204:            txtCode.Text = txtCode.Text & asCodeLines(i)
205:        Else
206:            txtCode.Text = txtCode.Text & asCodeLines(i) & vbNewLine
207:        End If
208:    Next
209: End Sub

Private Sub SetOptFromTable(ByVal iRow As Byte, ByVal iVal As Variant)
212:    Dim OptionsTb As ListObject
213:    Set OptionsTb = SHSNIPPETS.ListObjects(C_Const.TB_OPTIONSIDEDENT)
214:    OptionsTb.ListColumns(2).Range(iRow, 1) = iVal
215:    Call UpdateCodeListBox
End Sub
