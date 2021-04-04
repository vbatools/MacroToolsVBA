VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProtectedSheets 
   Caption         =   "*****"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   OleObjectBlob   =   "ProtectedSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProtectedSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : ProtectedSheets - сн€тие паролей с листов Excel
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Public sPassword As String

    Private Function CalcSelected() As Integer
14:    Dim i      As Integer
15:    CalcSelected = 0
16:    With ListBox1
17:        For i = 0 To .ListCount - 1
18:            If .Selected(i) Then
19:                CalcSelected = CalcSelected + 1
20:            End If
21:        Next i
22:    End With
23: End Function

    Private Sub lbHelp_Click()
26:    Call URLLinks(C_Const.URL_FILE_PROTECT)
27: End Sub

    Private Sub UserForm_Initialize()
30:    Me.StartUpPosition = 0
31:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
32:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
33:
34:    Me.StartUpPosition = 0
35:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
36:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
37:    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
38:    Call RefreshListBoxSheets
39: End Sub
    Private Sub RefreshListBoxSheets()
41:    Dim n      As Integer
42:    Dim sh     As Worksheet
43:    n = 0
44:    With ListBox1
45:        .Clear
46:        .AddItem 0
47:        .List(n, 2) = " нига"
48:        .List(n, 1) = ActiveWorkbook.Name
49:        .List(n, 3) = "защиты нет"
50:        If ActiveWorkbook.ProtectStructure Or ActiveWorkbook.ProtectWindows Then
51:            .List(n, 3) = "защита есть"
52:        End If
53:        n = n + 1
54:        For Each sh In ActiveWorkbook.Worksheets
55:            .AddItem sh.Index
56:            .List(n, 2) = "Ћист"
57:            .List(n, 1) = sh.Name
58:            Select Case sh.ProtectContents
                Case True
60:                    .Selected(n) = True
61:                    .List(n, 3) = "защита есть"
62:                    .List(n, 4) = "***********"
63:                Case False
64:                    .List(n, 3) = "защиты нет"
65:            End Select
66:            n = n + 1
67:        Next
68:    End With
69: End Sub
    Private Sub CheckBox2_Click()
71:    Dim i      As Integer
72:    With ListBox1
73:        For i = 0 To .ListCount - 1
74:            If CheckBox2 And .List(i, 3) = "защита есть" Then
75:                .Selected(i) = True
76:            Else
77:                .Selected(i) = False
78:            End If
79:        Next i
80:    End With
81: End Sub
    Private Sub btnCellsProtrctedCansel_Click()
83:    Call btnCancel_Click
84: End Sub
    Private Sub btnCancel_Click()
86:    Unload Me
87: End Sub
     Private Sub btnCellsProtrcted_Click()
89:    Dim X As Integer, i As Integer
90:    Dim iSelected As Double
91:
92:    Me.Hide
93:    i = 0: iSelected = CalcSelected
94:    Application.ScreenUpdating = False
95:    Application.Calculation = xlCalculationManual
96:    With ListBox1
97:        For X = 0 To .ListCount - 1
98:            If .Selected(X) = True Then
99:                i = i + 1
100:                Application.ScreenUpdating = True
101:                Application.StatusBar = "ќбработан: " & .List(X, 1) & " выполнено: " & Format(i / iSelected, "0.00%")
102:                Application.ScreenUpdating = False
103:
104:                Call UnlockSheetsWorkbooks(X)
105:            End If
106:        Next X
107:    End With
108:    Application.ScreenUpdating = True
109:    Application.Calculation = xlCalculationAutomatic
110:    Application.StatusBar = False
111:    Me.Show
112:
113: End Sub
Private Sub UnlockSheetsWorkbooks(ByVal X As Integer)
115:    Dim i As Integer, j As Integer, k As Integer
116:    Dim l As Integer, m As Integer, n As Integer
117:    Dim i1 As Integer, i2 As Integer, i3 As Integer
118:    Dim i4 As Integer, i5 As Integer, i6 As Integer
119:
120:    Dim kennwort As String
121:
122:    On Error Resume Next
123:
124:    With ListBox1
125:
126:        If sPassword = vbNullString Then
TryNewPasword:
128:            For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
129:                        For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
130:                                    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
131:                                                For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
132:                                                            kennwort = Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
133:                                                            With ListBox1
134:                                                                If .List(X, 2) = " нига" Then
135:                                                                    ActiveWorkbook.Unprotect kennwort
136:                                                                    If ActiveWorkbook.ProtectStructure = False And ActiveWorkbook.ProtectWindows = False Then
137:                                                                        sPassword = kennwort
138:                                                                        .List(X, 3) = "защиты нет"
139:                                                                        .List(X, 4) = kennwort
140:                                                                        Exit Sub
141:                                                                    End If
142:                                                                Else
143:                                                                    ActiveWorkbook.Sheets(.List(X, 1)).Unprotect kennwort
144:                                                                    If ActiveWorkbook.Sheets(.List(X, 1)).ProtectContents = False Then
145:                                                                        sPassword = kennwort
146:                                                                        .List(X, 3) = "защиты нет"
147:                                                                        .List(X, 4) = kennwort
148:                                                                        Exit Sub
149:                                                                    End If
150:                                                                End If
151:                                                            End With
152:                                                        Next: Next: Next: Next: Next: Next
153:                                Next: Next: Next: Next: Next: Next
154:        Else
155:            If .List(X, 2) = " нига" Then
156:                Exit Sub
157:            Else
158:                ActiveWorkbook.Sheets(.List(X, 1)).Unprotect sPassword
159:                If ActiveWorkbook.Sheets(.List(X, 1)).ProtectContents = False Then
160:                    .List(X, 3) = "защиты нет"
161:                    .List(X, 4) = sPassword
162:                    Exit Sub
163:                End If
164:            End If
165:
166:            sPassword = vbNullString
167:            GoTo TryNewPasword
168:        End If
169:    End With
End Sub
