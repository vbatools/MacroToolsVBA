Attribute VB_Name = "H_AddLegend"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : H_AddLegend - модуль создания легенды для сниппетов
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Private Module
Option Explicit
    Public Sub AddLegend()
12:    Dim str_legend As String, str1 As String, str2 As String
13:    Dim objLegend As ListObject
14:    Dim i      As Integer
15:    Dim LenLengh1 As Byte, LenLengh2 As Byte
16:    Set objLegend = SHSNIPPETS.ListObjects(C_Const.TB_DESCRIPTION)
17:    str_legend = vbNullString
18:    str1 = vbNullString
19:    str2 = vbNullString
20:    LenLengh1 = Len(objLegend.ListColumns(1).Range(1, 1))
21:    LenLengh2 = Len(objLegend.ListColumns(2).Range(1, 1))
22:    For i = 1 To objLegend.ListRows.Count + 1
23:        str1 = addString(objLegend.ListColumns(1).Range(i, 1), LenLengh1)
24:        str2 = addString(objLegend.ListColumns(2).Range(i, 1), LenLengh2)
25:        str_legend = str_legend & str1 & " | " & str2 & " | " & objLegend.ListColumns(1).Range(i, 3) & vbLf
26:    Next i
27:    Debug.Print str_legend
28: End Sub
Private Function addString(ByVal st As String, ByVal MaxLen As Byte) As String
30:    addString = st & Space(MaxLen - Len(st))
End Function
