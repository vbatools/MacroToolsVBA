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
4:    Dim str_legend As String, str1 As String, str2 As String
5:    Dim objLegend As ListObject
6:    Dim i      As Integer
7:    Dim LenLengh1 As Byte, LenLengh2 As Byte
8:    Set objLegend = SHSNIPPETS.ListObjects(C_Const.TB_DESCRIPTION)
9:    str_legend = vbNullString
10:    str1 = vbNullString
11:    str2 = vbNullString
12:    LenLengh1 = Len(objLegend.ListColumns(1).Range(1, 1))
13:    LenLengh2 = Len(objLegend.ListColumns(2).Range(1, 1))
14:    For i = 1 To objLegend.ListRows.Count + 1
15:        str1 = addString(objLegend.ListColumns(1).Range(i, 1), LenLengh1)
16:        str2 = addString(objLegend.ListColumns(2).Range(i, 1), LenLengh2)
17:        str_legend = str_legend & str1 & " | " & str2 & " | " & objLegend.ListColumns(1).Range(i, 3) & vbLf
18:    Next i
19:    Debug.Print str_legend
20: End Sub
    Private Function addString(ByVal st As String, ByVal MaxLen As Byte) As String
22:    addString = st & Space(MaxLen - Len(st))
23: End Function
