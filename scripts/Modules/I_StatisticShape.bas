Attribute VB_Name = "I_StatisticShape"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : I_StatisticShape - формирование статистики привязанных макросов к формам на листе Excel
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module

Public Sub AddShapeStatistic()
5:    Dim fForm  As AddStatistic
6:    Dim wb_name As String, macro_name As String
7:    Dim shp    As Shape
8:    Dim SH     As Worksheet
9:    Dim i      As Integer
10:    Const SH_SHAPE As String = "Statistic_share"
11:
12:    On Error GoTo errmsg
13:
14:    Set fForm = New AddStatistic
15:    With fForm
16:        .lbOK.Caption = "TO CREATE"
17:        .Show
18:        wb_name = .cmbMain.Value
19:        If wb_name = vbNullString Then Exit Sub
20:    End With
21:    Application.ScreenUpdating = False
22:    ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
23:    With ActiveSheet
24:        .Name = SH_SHAPE
25:        .Cells(1, 1).Value = "Sheet name"
26:        .Cells(1, 2).Value = "Name of the shape"
27:        .Cells(1, 3).Value = "Shape text"
28:        .Cells(1, 4).Value = "Macro name"
29:        i = 1
30:        For Each SH In Workbooks(wb_name).Worksheets
31:            For Each shp In SH.Shapes
32:                i = i + 1
33:                .Hyperlinks.Add Anchor:=Cells(i, 1), Address:="", SubAddress:=SH.Name & "!A1", TextToDisplay:=SH.Name
34:                .Cells(i, 2).Value = shp.Name
35:
36:                Select Case shp.Type
                    Case msoAutoShape
38:                        .Cells(i, 3).Value = shp.TextFrame2.TextRange.Characters.Text
39:                    Case msoFormControl, msoOLEControlObject
40:                        .Cells(i, 3).Value = shp.AlternativeText
41:                    Case Else
42:                        .Cells(i, 3).Value = "no"
43:                End Select
44:
45:                macro_name = shp.OnAction
46:                If macro_name = vbNullString Then
47:                    .Cells(i, 4).Value = "no macro"
48:                Else
49:                    .Cells(i, 4).Value = Split(shp.OnAction, "!")(1)
50:                End If
51:            Next
52:        Next
53:        .Columns("A:D").EntireColumn.AutoFit
54:    End With
55:    Application.ScreenUpdating = True
56:    Exit Sub
errmsg:
58:    If Err.Number = 1004 Then
59:        Application.DisplayAlerts = False
60:        ActiveWorkbook.Sheets(SH_SHAPE).Delete
61:        Application.DisplayAlerts = True
62:        ActiveSheet.Name = SH_SHAPE
63:        Err.Clear
64:        Resume Next
65:    Else
66:        Application.ScreenUpdating = True
67:        Call MsgBox("Error in AddShapeStatistic" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Error:")
68:        Call WriteErrorLog("AddShapeStatistic")
69:    End If
End Sub
