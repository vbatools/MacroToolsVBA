Attribute VB_Name = "J_EditCode"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : J_EditCode
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Private Module
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CutTab - удалить все Tab из кода VBA
'* Created    : 08-10-2020 14:08
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub CutTab()
20:    Dim cmb_txt As String
21:    Dim vbComp As VBIDE.VBComponent
22:
23:    On Error GoTo ErrorHandler
24:
25:    cmb_txt = B_CreateMenus.WhatIsTextInComboBoxHave
26:    Select Case cmb_txt
        Case C_Const.ALLVBAPROJECT:
28:            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
29:                Call N_Obfuscation.TrimLinesTabAndSpase(vbComp.CodeModule)
30:            Next vbComp
31:        Case C_Const.SELECTEDMODULE:
32:            Call N_Obfuscation.TrimLinesTabAndSpase(Application.VBE.ActiveCodePane.CodeModule)
33:    End Select
34:    Exit Sub
ErrorHandler:
36:    If Err.Number <> 91 Then
37:        Debug.Print "Error in CutTab" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
38:        Call WriteErrorLog("CutTab")
39:    End If
40:    Err.Clear
41: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CloseAllWindowsVBE - закрывает все окна VBE, кроме активного
'* Created    : 01-20-2020 14:32
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CloseAllWindowsVBE()
50:    Dim vbWin  As VBIDE.Window
51:    For Each vbWin In Application.VBE.Windows
52:        If (vbWin.Type = vbext_wt_CodeWindow Or vbWin.Type = vbext_wt_Designer) And Not vbWin Is Application.VBE.ActiveWindow Then
53:            vbWin.Close
54:        End If
55:    Next vbWin
56:    Application.VBE.ActiveWindow.WindowState = vbext_ws_Maximize
End Sub
