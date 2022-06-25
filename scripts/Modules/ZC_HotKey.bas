Attribute VB_Name = "ZC_HotKey"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : ZC_HotKey
'* Created    : 23-06-2022 16:21
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

    Public Sub hotKeysStart()
12:
13:    If Not ThisWorkbook.Name = C_Const.NAME_ADDIN & ".xlam" Then Exit Sub
14:
15:    On Error GoTo errMsg
16:    Dim sPatpApp    As String
17:    sPatpApp = ThisWorkbook.Path & Application.PathSeparator & FILE_NAME_HOT_KEYS
18:    If FileHave(sPatpApp) Then
19:        Call Shell(sPatpApp)
20:    Else
21:        Call MsgBox("File not found -" & FILE_NAME_HOT_KEYS, vbInformation, "HotKeys:")
22:    End If
23:    Exit Sub
errMsg:
25:    Call WriteErrorLog("hotKeysStart")
26: End Sub

Public Sub hotKeysStop()
29:    On Error GoTo errMsg
30:
31:    Dim sPatpApp    As String
32:    sPatpApp = ThisWorkbook.Path & Application.PathSeparator & FILE_NAME_HOT_KEYS
33:    If FileHave(sPatpApp) Then
34:        Dim WshShell As Object
35:        Set WshShell = CreateObject("WScript.Shell")
36:        If Not WshShell Is Nothing Then
37:            Dim WshExec As Object
38:            Set WshExec = WshShell.Exec(sPatpApp)
39:            If Not WshShell Is Nothing Then WshExec.Terminate
40:            Set WshExec = Nothing
41:        End If
42:        Set WshShell = Nothing
43:    End If
44:
45:    Exit Sub
errMsg:
47:    Call WriteErrorLog("hotKeysStop")
End Sub
