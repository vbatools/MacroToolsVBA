VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CharsMonitor 
   Caption         =   "Character Monitor:"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9015
   OleObjectBlob   =   "CharsMonitor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CharsMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : CharsMonitor - Анализ символов в строке
'* Created    : 23-04-2020 14:27
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private m_clsAnchors As CAnchors

    Private Sub UserForm_Activate()
13:    Me.StartUpPosition = 0
14:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
15:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
16: End Sub
    Private Sub UserForm_Initialize()
18:    Set m_clsAnchors = New CAnchors
19:    Set m_clsAnchors.objParent = Me
20:
21:    m_clsAnchors.MinimumWidth = 454
22:    m_clsAnchors.MinimumHeight = 401
23:    With m_clsAnchors
24:        .funAnchor("lbClearForm").AnchorStyle = enumAnchorStyleRight
25:        .funAnchor("txtStr").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleLeft Or enumAnchorStyleTop Or enumAnchorStyleBottom
26:        .funAnchor("lbLoadStr").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
27:        .funAnchor("chStrAfter").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
28:        .funAnchor("lbMsg").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
29:        .funAnchor("Label3").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
30:        .funAnchor("ListChars").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight Or enumAnchorStyleBottom
31:        .funAnchor("lbExportStr").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
32:        .funAnchor("lbCancel").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
33:        .funAnchor("lbAddASIITable").AnchorStyle = enumAnchorStyleBottom
34:    End With
35: End Sub
    Private Sub UserForm_Terminate()
37:    Set m_clsAnchors = Nothing
38: End Sub
    Private Sub btnCancel_Click()
40:    Me.Hide
41: End Sub
    Private Sub lbCancel_Click()
43:    Call btnCancel_Click
44: End Sub

    Private Sub lbClearForm_Click()
47:    Me.txtStr = vbNullString
48:    Me.ListChars.Clear
49:    Me.lbMsg.Caption = vbNullString
50: End Sub
    Private Sub txtStr_Change()
52:    Call subStartParser
53: End Sub
    Private Sub subStartParser()
55:    Dim lRows       As Long
56:    Dim lWords      As Long
57:    Dim ar          As Variant
58:
59:    With Me.ListChars
60:        .Clear
61:        Me.lbMsg.Caption = vbNullString
62:        If Me.txtStr.Text <> vbNullString Then
63:            .List = funParseChars(Me.txtStr.Text)
64:            lRows = UBound(VBA.Split(Me.txtStr.Text, vbNewLine)) + 1
65:            lWords = UBound(VBA.Split(C_PublicFunctions.TrimSpace(VBA.Replace(Me.txtStr.Text, vbNewLine, VBA.Chr(32))), VBA.Chr(32))) + 1
66:            If lRows < 0 Then lRows = 0
67:            Me.lbMsg.Caption = "String length:" & VBA.Len(Me.txtStr.Text) & "this. Lines:" & lRows & "Words:" & lWords
68:        End If
69:    End With
70: End Sub

    Private Function funParseChars(ByVal sTxt As String) As Variant
73:
74:    Dim n           As Long
75:    Dim i           As Long
76:    Dim sChar       As String
77:
78:    On Error Resume Next
79:    n = Len(sTxt): ReDim Arr(1 To n, 1 To 5)
80:    For i = LBound(Arr) To UBound(Arr)
81:        Arr(i, 1) = i
82:        sChar = VBA.Mid$(sTxt, i, 1)
83:        Arr(i, 2) = sChar
84:        Arr(i, 3) = VBA.Asc(sChar)
85:        Arr(i, 4) = VBA.AscW(sChar)
86:        Arr(i, 5) = VBA.Hex$(VBA.Asc(sChar))
87:    Next i
88:    funParseChars = Arr
89: End Function
     Private Sub lbLoadStr_Click()
91:    Dim objRng      As Range
92:    Dim itemRng     As Range
93:    Dim sStrTemp    As String
94:    Dim iColCount   As Integer
95:    Dim i           As Integer
96:    Dim sChrDel     As String
97:
98:    i = 1
99:    sChrDel = VBA.Chr$(32)
100:    Me.Hide
101:    Set objRng = GetAddressCell()
102:    If objRng Is Nothing Then Exit Sub
103:    iColCount = objRng.Columns.Count
104:    For Each itemRng In objRng
105:        If chStrAfter.Value Then
106:            If i > iColCount Then
107:                i = 1
108:                sChrDel = vbNewLine
109:            Else
110:                sChrDel = VBA.Chr$(32)
111:            End If
112:        End If
113:        i = i + 1
114:        sStrTemp = sStrTemp & sChrDel & itemRng.Value
115:    Next itemRng
116:    Me.txtStr = VBA.Right$(sStrTemp, VBA.Len(sStrTemp) - VBA.Len(sChrDel))
117:    Call subStartParser
118:    Me.Show
119: End Sub
     Private Sub lbExportStr_Click()
121:    Dim objRng      As Range
122:    Me.Hide
123:    Set objRng = GetAddressCell("Select the cell to insert:")
124:    If objRng Is Nothing Then Exit Sub
125:    With Me.ListChars
126:        If .ListCount > 0 Then
127:            objRng.Cells(1, 1).Value = Me.txtStr
128:            objRng.Offset(1, 0).Resize(1, 5) = Array("№", "Char", "Asc", "AscW", "Hex")
129:            objRng.Offset(2, 0).Resize(.ListCount, 5) = .List
130:        End If
131:    End With
132:    Me.Show
133: End Sub
     Private Sub lbAddASIITable_Click()
135:    Dim objRng      As Range
136:    Dim i           As Integer
137:    ReDim Arr(1 To 256, 1 To 5)
138:    Me.Hide
139:    Set objRng = GetAddressCell("Select the cell to insert:")
140:    If objRng Is Nothing Then Exit Sub
141:    For i = 1 To 256
142:        Arr(i, 1) = i - 1
143:        Arr(i, 2) = VBA.Hex$(i - 1)
144:        Arr(i, 3) = VBA.Chr$(i - 1)
145:        Arr(i, 4) = VBA.AscW(Arr(i, 3))
146:        Arr(i, 5) = GetDiscriptionSpeshelChar(i - 1)
147:    Next i
148:    With objRng
149:        .Resize(1, 5).Value = Array("Dec/Asc", "Hex", "Char", "AscW", "Description")
150:        .Offset(1, 0).Resize(256, 5).Value = Arr
151:    End With
152: End Sub

     Private Function GetDiscriptionSpeshelChar(ByVal i As Byte) As String
155:    Select Case i
        Case 0: GetDiscriptionSpeshelChar = "NOP"
157:        Case 1: GetDiscriptionSpeshelChar = "SOH"
158:        Case 2: GetDiscriptionSpeshelChar = "STX"
159:        Case 3: GetDiscriptionSpeshelChar = "ETX"
160:        Case 4: GetDiscriptionSpeshelChar = "EOT"
161:        Case 5: GetDiscriptionSpeshelChar = "ENQ"
162:        Case 6: GetDiscriptionSpeshelChar = "ACK"
163:        Case 7: GetDiscriptionSpeshelChar = "BEL"
164:        Case 8: GetDiscriptionSpeshelChar = "BS"
165:        Case 9: GetDiscriptionSpeshelChar = "Tabulation"
166:        Case 10: GetDiscriptionSpeshelChar = "LF(Возвр. каретки)"
167:        Case 11: GetDiscriptionSpeshelChar = "VT"
168:        Case 12: GetDiscriptionSpeshelChar = "FF"
169:        Case 13: GetDiscriptionSpeshelChar = "CR(Новая строка)"
170:        Case 14: GetDiscriptionSpeshelChar = "SO"
171:        Case 15: GetDiscriptionSpeshelChar = "SI"
172:        Case 16: GetDiscriptionSpeshelChar = "DLE"
173:        Case 17: GetDiscriptionSpeshelChar = "DC1"
174:        Case 18: GetDiscriptionSpeshelChar = "DC2"
175:        Case 19: GetDiscriptionSpeshelChar = "DC3"
176:        Case 20: GetDiscriptionSpeshelChar = "DC4"
177:        Case 21: GetDiscriptionSpeshelChar = "NAK"
178:        Case 22: GetDiscriptionSpeshelChar = "SYN"
179:        Case 23: GetDiscriptionSpeshelChar = "ETB"
180:        Case 24: GetDiscriptionSpeshelChar = "CAN"
181:        Case 25: GetDiscriptionSpeshelChar = "EM"
182:        Case 26: GetDiscriptionSpeshelChar = "SUB"
183:        Case 27: GetDiscriptionSpeshelChar = "ESC"
184:        Case 28: GetDiscriptionSpeshelChar = "FS"
185:        Case 29: GetDiscriptionSpeshelChar = "GS"
186:        Case 30: GetDiscriptionSpeshelChar = "RS"
187:        Case 31: GetDiscriptionSpeshelChar = "US"
188:        Case 32: GetDiscriptionSpeshelChar = "SP (Пробел)"
189:    End Select
190: End Function
Private Function GetAddressCell(Optional sMsg As String = "Select a data range:") As Range
192:    Dim sDefault    As String
193:    On Error GoTo Canceled
194:    If TypeName(Selection) = "Range" Then
195:        sDefault = Selection.Address
196:    Else
197:        sDefault = vbNullString
198:    End If
199:    Set GetAddressCell = Application.InputBox(Prompt:=sMsg, Type:=8, Default:=sDefault)
200:    Exit Function
Canceled:
202:    Set GetAddressCell = Nothing
End Function
