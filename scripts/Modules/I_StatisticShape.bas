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
13:    Dim fForm  As AddStatistic
14:    Dim wb_name As String, macro_name As String
15:    Dim shp    As Shape
16:    Dim SH     As Worksheet
17:    Dim i      As Integer
18:    Const SH_SHAPE As String = "Statistic_share"
19:
20:    On Error GoTo errmsg
21:
22:    Set fForm = New AddStatistic
23:    With fForm
24:        .lbOK.Caption = "CREATE"
25:        .Show
26:        wb_name = .cmbMain.Value
27:        If wb_name = vbNullString Then Exit Sub
28:    End With
29:    Application.ScreenUpdating = False
30:    ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
31:    With ActiveSheet
32:        .Name = SH_SHAPE
33:        .Cells(1, 1).Value = "Sheet name"
34:        .Cells(1, 2).Value = "Name of the shape"
35:        .Cells(1, 3).Value = "Shape text"
36:        .Cells(1, 4).Value = "Macro name"
37:        i = 1
38:        For Each SH In Workbooks(wb_name).Worksheets
39:            For Each shp In SH.Shapes
40:                i = i + 1
41:                .Hyperlinks.Add Anchor:=Cells(i, 1), Address:="", SubAddress:=SH.Name & "!A1", TextToDisplay:=SH.Name
42:                .Cells(i, 2).Value = shp.Name
43:
44:                Select Case shp.Type
                    Case msoAutoShape
46:                        .Cells(i, 3).Value = shp.TextFrame2.TextRange.Characters.Text
47:                    Case msoFormControl, msoOLEControlObject
48:                        .Cells(i, 3).Value = shp.AlternativeText
49:                    Case Else
50:                        .Cells(i, 3).Value = "no"
51:                End Select
52:
53:                macro_name = shp.OnAction
54:                If macro_name = vbNullString Then
55:                    .Cells(i, 4).Value = "no macro"
56:                Else
57:                    .Cells(i, 4).Value = Split(shp.OnAction, "!")(1)
58:                End If
59:            Next
60:        Next
61:        .Columns("A:D").EntireColumn.AutoFit
62:    End With
63:    Application.ScreenUpdating = True
64:    Exit Sub
errmsg:
66:    If Err.Number = 1004 Then
67:        Application.DisplayAlerts = False
68:        ActiveWorkbook.Sheets(SH_SHAPE).Delete
69:        Application.DisplayAlerts = True
70:        ActiveSheet.Name = SH_SHAPE
71:        Err.Clear
72:        Resume Next
73:    Else
74:        Application.ScreenUpdating = True
75:        Call MsgBox("Error in AddShapeStatistic" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbCritical, "Error:")
76:        Call WriteErrorLog("AddShapeStatistic")
77:    End If
End Sub
