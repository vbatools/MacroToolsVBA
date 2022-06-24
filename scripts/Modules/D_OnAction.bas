Attribute VB_Name = "D_OnAction"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : D_OnAction - модуль сниппетов, вставки кода в модуль
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Private Module
Option Explicit
    Public Sub InsertCode()
12:    Dim str_search  As String
13:    Dim space_i     As Long
14:    Dim str_arr() As String, code_arr() As String
15:    Dim code_flag   As Boolean
16:    Dim snippets    As ListObject
17:    Dim i_row       As Long
18:    Dim code        As String
19:    code_flag = False
20:    Application.DisplayAlerts = False
21:    str_arr = Split(CreateLineProcedure, "|")
22:    str_search = str_arr(1)
23:    If str_search = vbNullString Then
24:        Debug.Print "Ничего не выбрано!"
25:        Exit Sub
26:    End If
27:    code_arr = Split(str_search, " ")
28:    If UBound(code_arr) > 0 Then
29:        str_search = code_arr(0)
30:        code_flag = True
31:    End If
32:    Set snippets = SHSNIPPETS.ListObjects(C_Const.TB_SNIPPETS)
33:    On Error GoTo errMsg
34:    i_row = snippets.ListColumns(2).DataBodyRange.Find(What:=str_search, LookIn:=xlValues, LookAt:=xlWhole).Row
35:    code = snippets.Range(i_row, 4)
36:    space_i = CInt(str_arr(2))
37:    If space_i > 0 Then code = AddSpaceCode(code, space_i)
38:    If code_flag Then
39:        code = Replace(code, "@1", AddCodeStr(code_arr))
40:    End If
41:    'вставка
42:    With Application.VBE.ActiveCodePane
43:        i_row = CLng(str_arr(0))
44:        .CodeModule.ReplaceLine i_row, code
45:        .SetSelection i_row + 1, 1, i_row + 1, 1
46:    End With
47:    Application.DisplayAlerts = True
48:    Exit Sub
errMsg:
50:    If Err.Number = 91 Then
51:        Debug.Print "Снипет не найден, выбрано: " & str_search
52:    Else
53:        Debug.Print "Ошибка в InsertCode!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
54:        Call WriteErrorLog("InsertCode")
55:    End If
56:    Err.Clear
57:    Application.DisplayAlerts = True
58: End Sub
    Private Function CreateLineProcedure() As String
60:    Dim lStartLine  As Long
61:    Dim lStartColumn As Long
62:    Dim lEndLine    As Long
63:    Dim lEndColumn  As Long
64:    Dim code        As String
65:    Dim i           As Long
66:    With Application.VBE.ActiveCodePane
67:        .GetSelection lStartLine, lStartColumn, lEndLine, lEndColumn
68:        code = .CodeModule.Lines(lStartLine, 1)
69:        If code Like "*.*" Then
70:            code = VBA.Right$(code, VBA.Len(code) - VBA.InStr(1, code, "."))
71:        End If
72:        i = Len(code) - Len(LTrim$(code))
73:        CreateLineProcedure = lStartLine & "|" & C_PublicFunctions.TrimSpace(Trim$(code)) & "|" & i
74:    End With
75: End Function
    Private Function AddSpaceCode(ByRef code As String, ByRef spac As Long) As String
77:    Dim str_arr()   As String
78:    Dim new_code As String, space_str As String
79:    Dim i           As Long
80:    new_code = vbNullString
81:    space_str = Space(spac)
82:    str_arr = Split(code, Chr$(10))
83:    For i = 0 To UBound(str_arr)
84:        If i = UBound(str_arr) Then
85:            new_code = new_code & space_str & str_arr(i)
86:        Else
87:            new_code = new_code & space_str & str_arr(i) & Chr$(10)
88:        End If
89:    Next i
90:    AddSpaceCode = new_code
91: End Function
Private Function AddCodeStr(ByRef strVar() As String) As String
93:    Dim i           As Long
94:    AddCodeStr = vbNullString
95:    For i = 1 To UBound(strVar)
96:        AddCodeStr = AddCodeStr & " " & strVar(i)
97:    Next
End Function

