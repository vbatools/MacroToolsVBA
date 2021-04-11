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
10:    Dim sSTR As String, sErr As String
11:
12:    If obtnFormat Then
13:        sSTR = cmbFormat.Value
14:    Else
15:        sSTR = LTrim(txtCustom.Text)
16:    End If
17:
18:    sErr = vbNullString
19:    If lbView.Caption = vbNullString Then sErr = "The value input field is empty!" & vbNewLine
20:    If lbView.Caption Like "Error:" Then sErr = sErr & "Error, in the original format"
21:    If sErr <> vbNullString Then
22:        Call MsgBox(sErr, vbCritical, "Error:")
23:        AddCode = vbNullString
24:        Exit Function
25:    End If
26:
27:    sSTR = "VBA.Format$(" & Replace(txtValue.Value, ",", ".") & ", " & Chr(34) & sSTR & Chr(34) & ")"
28:    AddCode = sSTR
29: End Function

    Private Sub lbHelp_Click()
32:    Call URLLinks(C_Const.URL_BILD_FOFMAT)
33: End Sub

    Private Sub lbInsertCode_Click()
36:    Dim iLine  As Integer
37:    Dim txtCode As String, txtLine As String
38:    'получение кода
39:    txtCode = AddCode()
40:    If txtCode = vbNullString Then Exit Sub
41:    txtLine = C_PublicFunctions.SelectedLineColumnProcedure
42:    If txtLine = vbNullString Then
43:        Me.Hide
44:        Exit Sub
45:    End If
46:    iLine = Split(txtLine, "|")(2)
47:
48:    With Application.VBE.ActiveCodePane
49:        .CodeModule.InsertLines iLine, txtCode
50:    End With
51:    Me.Hide
52: End Sub
    Private Sub btnCopyCode_Click()
54:    Dim sSTR As String, sMsgBoxString As String
55:
56:    sSTR = AddCode()
57:    If sSTR = vbNullString Then Exit Sub
58:    Call C_PublicFunctions.SetTextIntoClipboard(sSTR)
59:
60:    sMsgBoxString = "The code is copied to the clipboard!" & vbNewLine & "To insert the code, use" & Chr(34) & "Ctrl+V" & Chr(34)
61:    Call MsgBox(sMsgBoxString, vbInformation, "Copying the code:")
62:
63:    Me.Hide
64: End Sub

    Private Sub cmbCancel_Click()
67:    Me.Hide
68: End Sub

    Private Sub cmbCustomFormat_Change()
71:    If cmbCustomFormat.ListIndex = -1 Then Exit Sub
72:    If obtnCustomFormat Then
73:        txtDiscription.Text = arrFormatDateCustomDiscription(cmbCustomFormat.ListIndex)
74:    End If
75: End Sub
    Private Sub txtCustom_Change()
77:    Call AddFormat(txtCustom.Text)
78: End Sub
    Private Sub lbClear_Click()
80:    txtCustom.Text = vbNullString
81: End Sub
    Private Sub cmbFormat_Change()
83:    If cmbFormat.ListIndex = -1 Then Exit Sub
84:    If obtnDate And obtnFormat Then
85:        txtDiscription.Text = arrFormatDateDiscription(cmbFormat.ListIndex)
86:    Else
87:        txtDiscription.Text = arrFormatValueDiscription(cmbFormat.ListIndex)
88:    End If
89:
90:    Call AddFormat(cmbFormat.Value)
91: End Sub

    Private Sub lbAddCustom_Click()
94:    If cmbCustomFormat.Value <> vbNullString Then
95:        txtCustom.Text = txtCustom & " " & cmbCustomFormat.Value
96:    End If
97: End Sub

     Private Sub obtnDate_Change()
100:    Call AddList
101: End Sub

     Private Sub AddList()
104:    If obtnFormat Then
105:        cmbCustomFormat.Clear
106:        If obtnDate Then
107:            cmbFormat.List = arrFormatDate
108:        Else
109:            cmbFormat.List = arrFormatValue
110:        End If
111:        obtnValue.visible = True
112:    Else
113:        cmbFormat.Clear
114:        cmbCustomFormat.List = arrFormatDateCustom
115:        obtnValue.visible = False
116:    End If
117:    Call ChengeFlag(obtnFormat)
118: End Sub
     Private Sub ChengeFlag(ByVal Flag As Boolean)
120:    obtnDate.visible = Flag
121:    cmbFormat.visible = Flag
122:    cmbCustomFormat.visible = (Not Flag)
123:    txtCustom.visible = (Not Flag)
124:    lbClear.visible = (Not Flag)
125:    lbAddCustom.visible = (Not Flag)
126:    If Flag Then
127:        Frame2.Left = lbAddCustom.Left - 5
128:    Else
129:        Frame2.Left = cmbCustomFormat.Left + cmbCustomFormat.Width + 5
130:    End If
131: End Sub
     Private Sub obtnFormat_Click()
133:    Call AddList
134: End Sub
     Private Sub obtnCustomFormat_Click()
136:    Call AddList
137: End Sub
     Private Sub txtValue_Change()
139:    If cmbFormat.Value = vbNullString Then
140:        Call AddFormat(cmbCustomFormat.Value)
141:    Else
142:        Call AddFormat(cmbFormat.Value)
143:    End If
144: End Sub
     Private Sub AddFormat(ByVal sSTR As String)
146:    On Error GoTo err_msg
147:    lbView.Caption = Format(CDbl(txtValue.Value), sSTR)
148:    lbView.ForeColor = &H8000000D
149:    Exit Sub
err_msg:
151:    Select Case Err.Number
        Case 6
153:            lbView.Caption = "Error: enter a smaller number!"
154:        Case 13
155:            lbView.Caption = vbNullString
156:        Case Else
157:            lbView.Caption = "Error:" & Err.Description & " " & Err.Number
158:    End Select
159:    lbView.ForeColor = &H8080FF
160:    Err.Clear
161: End Sub

     Private Sub txtValue_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
164:    Dim Txt    As String
165:    Txt = Me.txtValue    ' читаем текст из поля (для недопущения ввода двух и более запятых)
166:    Select Case KeyAscii
        Case 8:    ' нажат Backspace - ничего не делаем
168:        Case 44: KeyAscii = IIf(InStr(1, Txt, ",") > 0, 0, 44)    ' если запятая уже есть - отменяем ввод символа
169:        Case 46: KeyAscii = IIf(InStr(1, Txt, ",") > 0, 0, 44)    ' заменяем при вводе точку на запятую
170:        Case 48 To 57    ' если введена цифра  - ничего не делаем
171:        Case Else: KeyAscii = 0    ' иначе отменяем ввод символа
172:    End Select
173: End Sub

     Private Sub UserForm_Initialize()
176:    arrFormatDate = Array("General Date", "Long Date", "Medium Date", "Short Date", "Long Time", "Medium Time", "Short Time")
177:    arrFormatDateDiscription = Array("Displays the date and / or time, for example 4/3/93 05:34 PM. If there is no fractional part, only the date is displayed, for example, 4/3/93. If the whole part is missing, only the time is displayed, for example 05:34 PM. The date display is determined by the system parameters.", _
                "Displays the date according to the long date format used in the system.", _
                "Displays the date using the average date format corresponding to the language version of the host application.", _
                "Displays the date according to the short date format used in the system.", _
                "Displaying the time according to the long date format used in the system", _
                "Displays the time in 12-hour format using hours, minutes, and the AM/PM pointer.", _
                "Displays the time in a 24-hour format, such as 17:45.")
184:    arrFormatValue = Array("General Number", "Currency", "Fixed", "Standard", "Percent", "Scientific", "Yes/No", "True/False", "On/Off")
185:    arrFormatValueDiscription = Array("Displays an unsigned number separated by a group of digits.", _
                "Displays a number using the digit group separator, if necessary; displays two digits to the right of the integer and fractional separator. The output is based on the system's language settings.", _
                "Displays at least one digit to the left and two digits to the right of the integer and fractional separator.", _
                "Displays a number using the digit group separator; displays at least one digit to the left and two digits to the right of the integer and fractional separator.", _
                "Displaying a number multiplied by 100 with a percent sign ( % ) added to the right, always displaying two digits to the right of the integer and fractional separator.", _
                "The standard exponential representation is used.", _
                "No is displayed if the number is 0, otherwise Yes is displayed.", _
                "False is displayed if the number is 0, otherwise True is displayed.", _
                "It is displayed On if the number is 0, otherwise it is displayed Off.")
194:    arrFormatDateCustom = Array("c", "d", "dd", "ddd", "dddd", "ddddd", "dddddd", "w", "ww", "m", "mm", "mmm", "mmmm", "q", "y", "yy", "yyyy", "h", "hh", "n", "nn", "s", "ss", "ttttt")
195:    arrFormatDateCustomDiscription = Array("Date component separator. Some language standards may use other characters to represent the date component separator. This separator separates the day, month, and year when the date values are formatted. The character used as the date component separator in the formatted output is determined by the system parameters.", _
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
215:    cmbFormat.List = arrFormatDate
216:    Call ChengeFlag(obtnFormat)
217:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
218: End Sub

