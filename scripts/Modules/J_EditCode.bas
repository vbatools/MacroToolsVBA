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
10:    Dim cmb_txt As String
11:    Dim vbComp As VBIDE.VBComponent
12:
13:    On Error GoTo ErrorHandler
14:
15:    cmb_txt = B_CreateMenus.WhatIsTextInComboBoxHave
16:    Select Case cmb_txt
        Case C_Const.ALLVBAPROJECT:
18:            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
19:                Call N_Obfuscation.TrimLinesTabAndSpase(vbComp.CodeModule)
20:            Next vbComp
21:        Case C_Const.SELECTEDMODULE:
22:            Call N_Obfuscation.TrimLinesTabAndSpase(Application.VBE.ActiveCodePane.CodeModule)
23:    End Select
24:    Exit Sub
ErrorHandler:
26:    If Err.Number <> 91 Then
27:        Debug.Print "Error in CutTab" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
28:        Call WriteErrorLog("CutTab")
29:    End If
30:    Err.Clear
31: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CloseAllWindowsVBE - закрывает все окна VBE, кроме активного
'* Created    : 01-20-2020 14:32
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub CloseAllWindowsVBE()
40:    Dim vbWin  As VBIDE.Window
41:    For Each vbWin In Application.VBE.Windows
42:        If (vbWin.Type = vbext_wt_CodeWindow Or vbWin.Type = vbext_wt_Designer) And Not vbWin Is Application.VBE.ActiveWindow Then
43:            vbWin.Close
44:        End If
45:    Next vbWin
46:    Application.VBE.ActiveWindow.WindowState = vbext_ws_Maximize
47: End Sub
