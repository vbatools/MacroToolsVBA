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
15:    Dim sPatpApp    As String
16:    sPatpApp = ThisWorkbook.Path & Application.PathSeparator & FILE_NAME_HOT_KEYS
17:    If FileHave(sPatpApp) Then
18:        Call AddLegendsFromTabel(C_Const.TB_HOT_KEYS)
19:    Else
20:        Debug.Print "This function is not available, no file found:" & vbNewLine & sPatpApp
21:    End If
22: End Sub

    Private Sub AddLegendsFromTabel(ByVal sTabelName As String)
25:    Dim str_legend As String, str1 As String, str2 As String
26:    Dim objLegend As ListObject
27:    Dim i      As Integer
28:    Dim LenLengh1 As Byte, LenLengh2 As Byte
29:    Set objLegend = SHSNIPPETS.ListObjects(sTabelName)
30:    str_legend = vbNullString
31:    str1 = vbNullString
32:    str2 = vbNullString
33:    LenLengh1 = Len(objLegend.ListColumns(1).Range(1, 1))
34:    LenLengh2 = Len(objLegend.ListColumns(2).Range(1, 1))
35:    For i = 1 To objLegend.ListRows.Count + 1
36:        str1 = addString(objLegend.ListColumns(1).Range(i, 1), LenLengh1)
37:        str2 = addString(objLegend.ListColumns(2).Range(i, 1), LenLengh2)
38:        str_legend = str_legend & str1 & " | " & str2 & " | " & objLegend.ListColumns(1).Range(i, 3) & vbLf
39:    Next i
40:    Debug.Print str_legend
41: End Sub

Private Function addString(ByVal st As String, ByVal MaxLen As Byte) As String
44:    addString = st & Space(MaxLen - Len(st))
End Function
