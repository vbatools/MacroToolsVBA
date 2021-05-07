VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BilderFormat 
   Caption         =   "Format Constructor:"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10455
   OleObjectBlob   =   "BilderFormat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BilderFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************************************
' Author         : VBATools - конструктор форматов строк
' Date           : 15.09.2019
' Обратная связь : info@VBATools.ru
' Copyright      : VBATools.ru
'***********************************************************************************************************
Option Explicit
Private arrFormatDate As Variant
Private arrFormatDateDiscription As Variant
Private arrFormatValue As Variant
Private arrFormatValueDiscription As Variant
Private arrFormatDateCustom As Variant
Private arrFormatDateCustomDiscription As Variant

    Private Function AddCode() As String
16:    Dim sSTR As String, sErr As String
17:
18:    If obtnFormat Then
19:        sSTR = cmbFormat.Value
20:    Else
21:        sSTR = LTrim(txtCustom.Text)
22:    End If
23:
24:    sErr = vbNullString
25:    If lbView.Caption = vbNullString Then sErr = "The value input field is empty!" & vbNewLine
26:    If lbView.Caption Like "Error:" Then sErr = sErr & "Error, in the original format"
27:    If sErr <> vbNullString Then
28:        Call MsgBox(sErr, vbCritical, "Error:")
29:        AddCode = vbNullString
30:        Exit Function
31:    End If
32:
33:    sSTR = "VBA.Format$(" & Replace(txtValue.Value, ",", ".") & ", " & Chr(34) & sSTR & Chr(34) & ")"
34:    AddCode = sSTR
35: End Function

    Private Sub lbHelp_Click()
38:    Call URLLinks(C_Const.URL_BILD_FOFMAT)
39: End Sub

    Private Sub lbInsertCode_Click()
42:    Dim iLine  As Integer
43:    Dim txtCode As String, txtLine As String
44:    'получение кода
45:    txtCode = AddCode()
46:    If txtCode = vbNullString Then Exit Sub
47:    txtLine = C_PublicFunctions.SelectedLineColumnProcedure
48:    If txtLine = vbNullString Then
49:        Me.Hide
50:        Exit Sub
51:    End If
52:    iLine = Split(txtLine, "|")(2)
53:
54:    With Application.VBE.ActiveCodePane
55:        .CodeModule.InsertLines iLine, txtCode
56:    End With
57:    Me.Hide
58: End Sub
    Private Sub btnCopyCode_Click()
60:    Dim sSTR As String, sMsgBoxString As String
61:
62:    sSTR = AddCode()
63:    If sSTR = vbNullString Then Exit Sub
64:    Call C_PublicFunctions.SetTextIntoClipboard(sSTR)
65:
66:    sMsgBoxString = "The code is copied to the clipboard!" & vbNewLine & "To insert the code, use" & Chr(34) & "Ctrl+V" & Chr(34)
67:    Call MsgBox(sMsgBoxString, vbInformation, "Copying the code:")
68:
69:    Me.Hide
70: End Sub

    Private Sub cmbCancel_Click()
73:    Me.Hide
74: End Sub

    Private Sub cmbCustomFormat_Change()
77:    If cmbCustomFormat.ListIndex = -1 Then Exit Sub
78:    If obtnCustomFormat Then
79:        txtDiscription.Text = arrFormatDateCustomDiscription(cmbCustomFormat.ListIndex)
80:    End If
81: End Sub
    Private Sub txtCustom_Change()
83:    Call AddFormat(txtCustom.Text)
84: End Sub
    Private Sub lbClear_Click()
86:    txtCustom.Text = vbNullString
87: End Sub
    Private Sub cmbFormat_Change()
89:    If cmbFormat.ListIndex = -1 Then Exit Sub
90:    If obtnDate And obtnFormat Then
91:        txtDiscription.Text = arrFormatDateDiscription(cmbFormat.ListIndex)
92:    Else
93:        txtDiscription.Text = arrFormatValueDiscription(cmbFormat.ListIndex)
94:    End If
95:
96:    Call AddFormat(cmbFormat.Value)
97: End Sub

     Private Sub lbAddCustom_Click()
100:    If cmbCustomFormat.Value <> vbNullString Then
101:        txtCustom.Text = txtCustom & " " & cmbCustomFormat.Value
102:    End If
103: End Sub

     Private Sub obtnDate_Change()
106:    Call AddList
107: End Sub

     Private Sub AddList()
110:    If obtnFormat Then
111:        cmbCustomFormat.Clear
112:        If obtnDate Then
113:            cmbFormat.List = arrFormatDate
114:        Else
115:            cmbFormat.List = arrFormatValue
116:        End If
117:        obtnValue.visible = True
118:    Else
119:        cmbFormat.Clear
120:        cmbCustomFormat.List = arrFormatDateCustom
121:        obtnValue.visible = False
122:    End If
123:    Call ChengeFlag(obtnFormat)
124: End Sub
     Private Sub ChengeFlag(ByVal Flag As Boolean)
126:    obtnDate.visible = Flag
127:    cmbFormat.visible = Flag
128:    cmbCustomFormat.visible = (Not Flag)
129:    txtCustom.visible = (Not Flag)
130:    lbClear.visible = (Not Flag)
131:    lbAddCustom.visible = (Not Flag)
132:    If Flag Then
133:        Frame2.Left = lbAddCustom.Left - 5
134:    Else
135:        Frame2.Left = cmbCustomFormat.Left + cmbCustomFormat.Width + 5
136:    End If
137: End Sub
     Private Sub obtnFormat_Click()
139:    Call AddList
140: End Sub
     Private Sub obtnCustomFormat_Click()
142:    Call AddList
143: End Sub
     Private Sub txtValue_Change()
145:    If cmbFormat.Value = vbNullString Then
146:        Call AddFormat(cmbCustomFormat.Value)
147:    Else
148:        Call AddFormat(cmbFormat.Value)
149:    End If
150: End Sub
     Private Sub AddFormat(ByVal sSTR As String)
152:    On Error GoTo err_msg
153:    lbView.Caption = Format(CDbl(txtValue.Value), sSTR)
154:    lbView.ForeColor = &H8000000D
155:    Exit Sub
err_msg:
157:    Select Case Err.Number
        Case 6
159:            lbView.Caption = "Error: enter a smaller number!"
160:        Case 13
161:            lbView.Caption = vbNullString
162:        Case Else
163:            lbView.Caption = "Error:" & Err.Description & " " & Err.Number
164:    End Select
165:    lbView.ForeColor = &H8080FF
166:    Err.Clear
167: End Sub

     Private Sub txtValue_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
170:    Dim Txt    As String
171:    Txt = Me.txtValue    ' читаем текст из поля (для недопущения ввода двух и более запятых)
172:    Select Case KeyAscii
        Case 8:    ' нажат Backspace - ничего не делаем
174:        Case 44: KeyAscii = IIf(InStr(1, Txt, ",") > 0, 0, 44)    ' если запятая уже есть - отменяем ввод символа
175:        Case 46: KeyAscii = IIf(InStr(1, Txt, ",") > 0, 0, 44)    ' заменяем при вводе точку на запятую
176:        Case 48 To 57    ' если введена цифра  - ничего не делаем
177:        Case Else: KeyAscii = 0    ' иначе отменяем ввод символа
178:    End Select
179: End Sub

     Private Sub UserForm_Initialize()
182:    arrFormatDate = Array("General Date", "Long Date", "Medium Date", "Short Date", "Long Time", "Medium Time", "Short Time")
183:    arrFormatDateDiscription = Array("Displays the date and / or time, for example 4/3/93 05:34 PM. If there is no fractional part, only the date is displayed, for example, 4/3/93. If the whole part is missing, only the time is displayed, for example 05:34 PM. The date display is determined by the system parameters.", _
                    "Displays the date according to the long date format used in the system.", _
                    "Displays the date using the average date format corresponding to the language version of the host application.", _
                    "Displays the date according to the short date format used in the system.", _
                    "Displaying the time according to the long date format used in the system", _
                    "Displays the time in 12-hour format using hours, minutes, and the AM/PM pointer.", _
                    "Displays the time in a 24-hour format, such as 17:45.")
190:    arrFormatValue = Array("General Number", "Currency", "Fixed", "Standard", "Percent", "Scientific", "Yes/No", "True/False", "On/Off")
191:    arrFormatValueDiscription = Array("Displays an unsigned number separated by a group of digits.", _
                    "Displays a number using the digit group separator, if necessary; displays two digits to the right of the integer and fractional separator. The output is based on the system's language settings.", _
                    "Displays at least one digit to the left and two digits to the right of the integer and fractional separator.", _
                    "Displays a number using the digit group separator; displays at least one digit to the left and two digits to the right of the integer and fractional separator.", _
                    "Displaying a number multiplied by 100 with a percent sign ( % ) added to the right, always displaying two digits to the right of the integer and fractional separator.", _
                    "The standard exponential representation is used.", _
                    "No is displayed if the number is 0, otherwise Yes is displayed.", _
                    "False is displayed if the number is 0, otherwise True is displayed.", _
                    "It is displayed On if the number is 0, otherwise it is displayed Off.")
200:    arrFormatDateCustom = Array("c", "d", "dd", "ddd", "dddd", "ddddd", "dddddd", "w", "ww", "m", "mm", "mmm", "mmmm", "q", "y", "yy", "yyyy", "h", "hh", "n", "nn", "s", "ss", "ttttt")
201:    arrFormatDateCustomDiscription = Array("Date component separator. Some language standards may use other characters to represent the date component separator. This separator separates the day, month, and year when the date values are formatted. The character used as the date component separator in the formatted output is determined by the system parameters.", _
                    "Selected. days as a number without zero at the beginning (1 - 31).", _
                    "Select. days as a number with a zero at the beginning (01 - 31).", _
                    "Selected. days using abbreviations (Sun - Sat). Localized.", _
                    "Selected. of the day using the full name (Sunday - Saturday). Localized.", _
                    "Selected. dates using the full format (including day, month, and year) corresponding to the short date format in the system settings. The default short date format is m/d/yy.", _
                    "Select a number that represents a date using the full format (including day, month, and year) that corresponds to the long date format in the system settings. The default long date format is mmmm dd, yyyy.", _
                    "Selected. days of the week as a number (from 1 for Sunday to 7 for Saturday).", _
                    "Selected. weeks of the year as a number (1 - 54).", _
                    "Selected. month as a number without zero at the beginning (1 - 12). If m follows immediately after h or hh, the minute will be displayed instead of the month.", _
                    "Selected. month as a number with a zero at the beginning (01 - 12). If m follows immediately after h or hh, the minute will be displayed instead of the month.", _
                    "Selected. abbreviated name of the month (Jan  - Dec). Localized.", _
                    "Select the full name of the month (January - December). Localized.", _
                    "Select the quarter of the year as a number (1 - 4).", "Displaying the day of the year as a number (1 - 366).", "Select. years as a 2-digit number (00 - 99).", _
                    "Selected. years in the form of a 4-digit number (100 - 9999).", "Selected. hours as a number without zero at the beginning (0 - 23).", _
                    "Selected. hours as a number with a zero at the beginning (00 - 23).", "Select. minutes as a number without zero at the beginning (0 - 59).", _
                    "Selected. minutes as a number with a zero at the beginning (00 - 59).", _
                    "Select. seconds as a number without zero at the beginning (0 - 59).", _
                    "Selected. seconds as a number with zero at the beginning (00 - 59).", _
                    "Time is selected in the full format (including hour, minute, and second) using the time component separator defined in the time format specified in the system settings. Zero at the beginning is displayed if Zero at the beginning is selected and the time is earlier than 10:00 A.M. or P. M. The default time format is h:mm: ss.")
221:    cmbFormat.List = arrFormatDate
222:    Call ChengeFlag(obtnFormat)
223:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
224: End Sub

