VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BilderFormat 
   Caption         =   "Конструктор форматов:"
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
19:    If lbView.Caption = vbNullString Then sErr = "В поле ввода значения пусто!" & vbNewLine
20:    If lbView.Caption Like "Ошибка: *" Then sErr = sErr & "Ошибка, в исходном формате"
21:    If sErr <> vbNullString Then
22:        Call MsgBox(sErr, vbCritical, "Ошибка:")
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
60:    sMsgBoxString = "Код скопирован в буфер обмена!" & vbNewLine & "Для вставки кода используйте " & Chr(34) & "Ctrl+V" & Chr(34)
61:    Call MsgBox(sMsgBoxString, vbInformation, "Копирование кода:")
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
153:            lbView.Caption = "Ошибка: ведите меньшее число!"
154:        Case 13
155:            lbView.Caption = vbNullString
156:        Case Else
157:            lbView.Caption = "Ошибка: " & Err.Description & " " & Err.Number
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
177:    arrFormatDateDiscription = Array("Отображение даты и/или времени, например 4/3/93 05:34 PM. Если дробная часть отсутствует, отображается только дата, например 4/3/93. Если отсутствует целая часть, отображается только время, например 05:34 PM. Отображение даты определяется параметрами системы.", _
                "Отображение даты в соответствии с длинным форматом даты, используемым в системе.", _
                "Отображение даты с использованием среднего формата даты, соответствующего языковой версии ведущего приложения.", _
                "Отображение даты в соответствии с кратким форматом даты, используемым в системе.", _
                "Отображение времени в соответствии с длинным форматом даты, используемым в системе; включает часы, минуты и секунды.", _
                "Отображение времени в 12-часовом формате с использованием часов, минут и указателя AM/PM.", _
                "Отображение времени в 24-часовом формате, например 17:45.")
184:    arrFormatValue = Array("General Number", "Currency", "Fixed", "Standard", "Percent", "Scientific", "Yes/No", "True/False", "On/Off")
185:    arrFormatValueDiscription = Array("Отображение числа без знака разделителя групп разрядов.", _
                "Отображение числа с использованием разделителя групп разрядов, если это необходимо; отображаются две цифры справа от разделителя целой и дробной части. Вывод основывается на языковых настройках системы.", _
                "Отображение по крайней мере одной цифры слева и двух цифр справа от разделителя целой и дробной части.", _
                "Отображение числа с использованием разделителя групп разрядов; отображаются по крайней мере одна цифра слева и две цифры справа от разделителя целой и дробной части.", _
                "Отображение числа, умноженного на 100 со знаком процента (%), добавляемого справа; всегда отображаются две цифры справа от разделителя целой и дробной части.", _
                "Используется стандартное экспоненциальное представление.", _
                "Отображается Нет, если число равняется 0; в противном случае отображается Да.", _
                "Отображается False, если число равняется 0; в противном случае отображается True.", _
                "Отображается Вкл, если число равняется 0; в противном случае отображается Выкл.")
194:    arrFormatDateCustom = Array("c", "d", "dd", "ddd", "dddd", "ddddd", "dddddd", "w", "ww", "m", "mm", "mmm", "mmmm", "q", "y", "yy", "yyyy", "h", "hh", "n", "nn", "s", "ss", "ttttt")
195:    arrFormatDateCustomDiscription = Array("Разделитель компонентов даты. В некоторых языковых стандартах могут использоваться другие знаки для представления разделителя компонентов даты. Этот разделитель отделяет день, месяц и год, когда значения даты форматируются. Символ, используемый в качестве разделителя компонентов даты в отформатированных выходных данных, определяется параметрами системы.", _
                "Отобp. дня в виде числа без нуля в начале (1–31).", _
                "Отобр. дня в виде числа с нулем в начале (01–31).", _
                "Отобр. дня с использованием сокращений (Вс–Сб). Локализовано.", _
                "Отобр. дня с использованием полного имени (Воскресенье–Суббота). Локализовано.", _
                "Отобр. даты с использованием полного формата (включая день, месяц и год), соответствующего краткому формату даты в настройках системы. Кратким форматом даты по умолчанию является m/d/yy.", _
                "Отобр. числа, представляющего дату, с использованием полного формата (включая день, месяц и год), соответствующего длинному формату даты в настройках системы. Длинным форматом даты по умолчанию является mmmm dd, yyyy.", _
                "Отобр. дня недели в виде числа (от 1 для воскресенья и до 7 для субботы).", _
                "Отобр. недели года в виде числа (1–54).", _
                "Отобр. месяца в виде числа без нуля в начале (1–12). Если m следует сразу же после h или hh, отображаться будет не месяц, а минута.", _
                "Отобр. месяца в виде числа с нулем в начале (01–12). Если m следует сразу же после h или hh, отображаться будет не месяц, а минута.", _
                "Отобр. сокращенного названия месяца (янв–дек). Локализовано.", _
                "Отобр. полного названия месяца (январь–декабрь). Локализовано.", _
                "Отобр. квартала года в виде числа (1–4).", "Отображение дня года в виде числа (1–366).", "Отобр. года в виде 2-значного числа (00–99).", _
                "Отобр. года в виде 4-значного числа (100–9999).", "Отобр. часа в виде числа без нуля в начале (0–23).", _
                "Отобр. часа в виде числа с нулем в начале (00–23).", "Отобр. минуты в виде числа без нуля в начале (0–59).", _
                "Отобр. минуты в виде числа с нулем в начале (00–59).", _
                "Отобр. секунды в виде числа без нуля в начале (0–59).", _
                "Отобр. секунды в виде числа с нулем в начале (00–59).", _
                "Отобр. времени в полном формате (включая час, минуту и секунду) с использованием разделителя компонентов времени, определенного в формате времени, указанного в настройках системы. Нуль в начале отображается, если выбран параметр Нуль в начале и время относится к интервалу ранее 10:00 A.M. или P.M. Форматом времени по умолчанию является h:mm:ss.")
215:    cmbFormat.List = arrFormatDate
216:    Call ChengeFlag(obtnFormat)
217:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
218: End Sub
