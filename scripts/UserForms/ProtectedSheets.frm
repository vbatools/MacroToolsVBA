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
'* Module     : ProtectedSheets - снятие паролей с листов Excel
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Public sPassword    As String

   Private Function CalcSelected() As Integer
13:     Dim i      As Integer
14:     CalcSelected = 0
15:     With ListBox1
16:         For i = 0 To .ListCount - 1
17:             If .Selected(i) Then
18:                 CalcSelected = CalcSelected + 1
19:                 End If
20:             Next i
21:         End With
22:     End Function

   Private Sub lbHelp_Click()
25:     Call URLLinks(C_Const.URL_FILE_PROTECT)
26:     End Sub

   Private Sub UserForm_Initialize()
29:     Me.StartUpPosition = 0
30:     Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
31:     Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
32:
33:     Me.StartUpPosition = 0
34:     Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
35:     Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
36:     Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
37:     Call RefreshListBoxSheets
38:     End Sub
   Private Sub RefreshListBoxSheets()
40:     Dim n      As Integer
41:     Dim SH     As Worksheet
42:     n = 0
43:     With ListBox1
44:         .Clear
45:         .AddItem 0
46:         .List(n, 2) = "Book"
47:         .List(n, 1) = ActiveWorkbook.Name
48:         .List(n, 3) = "of protection no"
49:         If ActiveWorkbook.ProtectStructure Or ActiveWorkbook.ProtectWindows Then
50:             .List(n, 3) = "protection is there"
51:             End If
52:         n = n + 1
53:         For Each SH In ActiveWorkbook.Worksheets
54:             .AddItem SH.Index
55:             .List(n, 2) = "Sheet"
56:             .List(n, 1) = SH.Name
57:             Select Case SH.ProtectContents
                Case True
59:                     .Selected(n) = True
60:                     .List(n, 3) = "protection is there"
61:                     .List(n, 4) = "***********"
62:                     Case False
63:                     .List(n, 3) = "of protection no"
64:                     End Select
65:             n = n + 1
66:             Next
67:         End With
68:     End Sub
   Private Sub CheckBox2_Click()
70:     Dim i      As Integer
71:     With ListBox1
72:         For i = 0 To .ListCount - 1
73:             If CheckBox2 And .List(i, 3) = "protection is there" Then
74:                 .Selected(i) = True
75:                 Else
76:                 .Selected(i) = False
77:                 End If
78:             Next i
79:         End With
80:     End Sub
   Private Sub btnCellsProtrctedCansel_Click()
82:     Call btnCancel_Click
83:     End Sub
   Private Sub btnCancel_Click()
85:     Unload Me
86:     End Sub
     Private Sub btnCellsProtrcted_Click()
88:     Dim X As Integer, i As Integer
89:     Dim iSelected As Double
90:
91:     Me.Hide
92:     i = 0: iSelected = CalcSelected
93:     Application.ScreenUpdating = False
94:     Application.Calculation = xlCalculationManual
95:     With ListBox1
96:         For X = 0 To .ListCount - 1
97:             If .Selected(X) = True Then
98:                 i = i + 1
99:                 Application.ScreenUpdating = True
100:                 Application.StatusBar = "Processed:" & .List(X, 1) & "done:" & Format(i / iSelected, "0.00%")
101:                 Application.ScreenUpdating = False
102:
103:                 Call UnlockSheetsWorkbooks(X)
104:                 End If
105:             Next X
106:         End With
107:     Application.ScreenUpdating = True
108:     Application.Calculation = xlCalculationAutomatic
109:     Application.StatusBar = False
110:     Me.Show
111:
112:     End Sub
Private Sub UnlockSheetsWorkbooks(ByVal X As Integer)
114:     Dim i As Integer, j As Integer, k As Integer
115:     Dim l As Integer, m As Integer, n As Integer
116:     Dim i1 As Integer, i2 As Integer, i3 As Integer
117:     Dim i4 As Integer, i5 As Integer, i6 As Integer
118:
119:     Dim kennwort As String
120:
121:     On Error Resume Next
122:
123:     With ListBox1
124:
125:         If sPassword = vbNullString Then
TryNewPasword:
127:             For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
128:                         For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
129:                                     For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
130:                                                 For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
131:                                                             kennwort = Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
132:                                                             With ListBox1
133:                                                                 If .List(X, 2) = "Book" Then
134:                                                                     ActiveWorkbook.Unprotect kennwort
135:                                                                     If ActiveWorkbook.ProtectStructure = False And ActiveWorkbook.ProtectWindows = False Then
136:                                                                         sPassword = kennwort
137:                                                                         .List(X, 3) = "of protection no"
138:                                                                         .List(X, 4) = kennwort
139:                                                                         Exit Sub
140:                                                                         End If
141:                                                                     Else
142:                                                                     ActiveWorkbook.Sheets(.List(X, 1)).Unprotect kennwort
143:                                                                     If ActiveWorkbook.Sheets(.List(X, 1)).ProtectContents = False Then
144:                                                                         sPassword = kennwort
145:                                                                         .List(X, 3) = "of protection no"
146:                                                                         .List(X, 4) = kennwort
147:                                                                         Exit Sub
148:                                                                         End If
149:                                                                     End If
150:                                                                 End With
151:                                                             Next: Next: Next: Next: Next: Next
152:                                     Next: Next: Next: Next: Next: Next
153:             Else
154:             If .List(X, 2) = "Book" Then
155:                 Exit Sub
156:                 Else
157:                 ActiveWorkbook.Sheets(.List(X, 1)).Unprotect sPassword
158:                 If ActiveWorkbook.Sheets(.List(X, 1)).ProtectContents = False Then
159:                     .List(X, 3) = "of protection no"
160:                     .List(X, 4) = sPassword
161:                     Exit Sub
162:                     End If
163:                 End If
164:
165:             sPassword = vbNullString
166:             GoTo TryNewPasword
167:             End If
168:         End With
End Sub
