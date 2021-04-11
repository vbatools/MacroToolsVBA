Attribute VB_Name = "Z_Clipboard"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : Z_Clipboard - модуль вставки обработки Unicode
'* Created    : 24-09-2020 14:12
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module

#If VBA7 Then
Declare PtrSafe Function OpenClipboard Lib "USER32" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function EmptyClipboard Lib "USER32" () As Long
Declare PtrSafe Function CloseClipboard Lib "USER32" () As Long
Declare PtrSafe Function IsClipboardFormatAvailable Lib "USER32" (ByVal wFormat As Long) As Long
Declare PtrSafe Function GetClipboardData Lib "USER32" (ByVal wFormat As Long) As LongPtr
Declare PtrSafe Function SetClipboardData Lib "USER32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As LongPtr

#Else
Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Declare Function EmptyClipboard Lib "user32.dll" () As Long
Declare Function CloseClipboard Lib "user32.dll" () As Long
Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
#End If

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SetInCipBoard - запись в буфер обмена
'* Created    : 24-09-2020 14:12
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub SetInCipBoard()
47:    Dim stxt        As String
48:    stxt = GetSelectedLineColumnInProcedure
49:    Call SetClipboard(stxt)
50:    Debug.Print "The data is copied to the clipboard!"
51: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : GetFromCipBoard - загрузка из буфера обмена
'* Created    : 24-09-2020 14:12
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub GetFromCipBoard()
61:
62:    Dim txtCode     As String
63:    Dim txtClearCode
64:    Dim SelectStr   As String
65:    Dim SelectLeft   As String
66:    Dim SelectRight   As String
67:    Dim StarCol As Long
68:    Dim EndCol As Long
69:    Dim EndLin As Long
70:    Dim StarLin As Long
71:    'On Error GoTo ErrorHandler
72:
73:    'получение текст
74:    txtCode = GetClipboard()
75:    txtClearCode = TrimSpace(txtCode) & VBA.Chr$(32)
76:    If VBA.Left$(txtCode, 1) = VBA.Chr$(32) Then
77:        txtClearCode = VBA.Chr$(32) & txtClearCode
78:    End If
79:    txtCode = txtClearCode
80:    If txtCode = vbNullString Then Exit Sub
81:
82:    With Application.VBE.ActiveCodePane
83:        .GetSelection StarLin, StarCol, EndLin, EndCol
84:        If StarLin = 0 Then Exit Sub
85:        SelectStr = .CodeModule.Lines(StarLin, 1)
86:        SelectLeft = VBA.Left$(SelectStr, EndCol - 1)
87:        SelectRight = VBA.Replace(SelectStr, SelectLeft, vbNullString)
88:        .CodeModule.InsertLines StarLin, SelectLeft & txtCode & SelectRight
89:        .SetSelection StarLin, VBA.Len(SelectLeft & txtCode) + 1, StarLin, VBA.Len(SelectLeft & txtCode) + 1
90:        If .CodeModule.CountOfLines >= StarLin + 1 Then .CodeModule.DeleteLines StarLin + 1
91:    End With
92:    Exit Sub
ErrorHandler:
94:    Select Case Err
        Case 91:
96:            Debug.Print "Error!, the module for inserting code is not activated!" & vbNewLine & Err.Number & vbNewLine & Err.Description
97:        Case Else:
98:            Debug.Print "An error occurred in Getfromclipboard" & vbNewLine & Err.Number & vbNewLine & Err.Description
99:    End Select
100: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetSelectedLineColumnInProcedure - получение выделеных строк в модуле
'* Created    : 24-09-2020 14:13
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function GetSelectedLineColumnInProcedure() As String
110:    Dim lStartLine  As Long
111:    Dim lStartColumn As Long
112:    Dim lEndLine    As Long
113:    Dim lEndColumn  As Long
114:
115:    On Error GoTo ErrorHandler
116:
117:    With Application.VBE.ActiveCodePane
118:        .GetSelection lStartLine, lStartColumn, lEndLine, lEndColumn
119:        GetSelectedLineColumnInProcedure = .CodeModule.Lines(lStartLine, lEndLine)
120:    End With
121:    Exit Function
ErrorHandler:
123:    Select Case Err
        Case 91:
125:            Debug.Print "Error!, the module for inserting code is not activated!" & vbNewLine & Err.Number & vbNewLine & Err.Description
126:        Case Else:
127:            Debug.Print "An error occurred in GESelectedLineColumnInProcedure" & vbNewLine & Err.Number & vbNewLine & Err.Description
128:    End Select
129: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SetClipboard - преобразование строк перед вставкой в буфер обмена
'* Created    : 24-09-2020 14:13
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):         Description
'*
'* sUniText As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub SetClipboard(sUniText As String)
143:    Dim iStrPtr, iLen As Long, iLock
144:    Const GMEM_MOVEABLE As Long = &H2
145:    Const GMEM_ZEROINIT As Long = &H40
146:    Const CF_UNICODETEXT As Long = &HD
147:    Call OpenClipboard(0&)
148:    EmptyClipboard
149:    iLen = LenB(sUniText) + 2&
150:    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
151:    iLock = GlobalLock(iStrPtr)
152:    lstrcpy iLock, StrPtr(sUniText)
153:    GlobalUnlock iStrPtr
154:    SetClipboardData CF_UNICODETEXT, iStrPtr
155:    CloseClipboard
156: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetClipboard - преобразование строк из буфера обмена
'* Created    : 24-09-2020 14:14
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function GetClipboard() As String
166:    Dim iStrPtr, iLen, iLock
167:    Dim sUniText    As String
168:    Const CF_UNICODETEXT As Long = 13&
169:    Call OpenClipboard(0&)
170:    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
171:        iStrPtr = GetClipboardData(CF_UNICODETEXT)
172:        If iStrPtr Then
173:            iLock = GlobalLock(iStrPtr)
174:            iLen = GlobalSize(iStrPtr)
175:            sUniText = String$(iLen \ 2& - 1&, vbNullChar)
176:            lstrcpy StrPtr(sUniText), iLock
177:            GlobalUnlock iStrPtr
178:        End If
179:        GetClipboard = sUniText
180:    End If
181:    CloseClipboard
End Function
