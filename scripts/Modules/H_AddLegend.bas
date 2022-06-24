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
12:    Call AddLegendsFromTabel(C_Const.TB_DESCRIPTION)
13: End Sub
    Public Sub AddLegendHotKeys()
15:    Call AddLegendsFromTabel(C_Const.TB_HOT_KEYS)
16: End Sub

    Private Sub AddLegendsFromTabel(ByVal sTabelName As String)
19:    Dim str_legend As String, str1 As String, str2 As String
20:    Dim objLegend As ListObject
21:    Dim i      As Integer
22:    Dim LenLengh1 As Byte, LenLengh2 As Byte
23:    Set objLegend = SHSNIPPETS.ListObjects(sTabelName)
24:    str_legend = vbNullString
25:    str1 = vbNullString
26:    str2 = vbNullString
27:    LenLengh1 = Len(objLegend.ListColumns(1).Range(1, 1))
28:    LenLengh2 = Len(objLegend.ListColumns(2).Range(1, 1))
29:    For i = 1 To objLegend.ListRows.Count + 1
30:        str1 = addString(objLegend.ListColumns(1).Range(i, 1), LenLengh1)
31:        str2 = addString(objLegend.ListColumns(2).Range(i, 1), LenLengh2)
32:        str_legend = str_legend & str1 & " | " & str2 & " | " & objLegend.ListColumns(1).Range(i, 3) & vbLf
33:    Next i
34:    Debug.Print str_legend
35: End Sub

Private Function addString(ByVal st As String, ByVal MaxLen As Byte) As String
38:    addString = st & Space(MaxLen - Len(st))
End Function
