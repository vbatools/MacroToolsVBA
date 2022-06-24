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

Const FILE_NAME     As String = "MacroToolsHotKeys.exe"

Public Sub hotKeysStart()

    If Not ThisWorkbook.Name = C_Const.NAME_ADDIN & ".xlam" Then Exit Sub
    
    On Error GoTo errMsg
    Dim sPatpApp    As String
    sPatpApp = ThisWorkbook.Path & Application.PathSeparator & FILE_NAME
    If FileHave(sPatpApp) Then
        Call Shell(sPatpApp)
    Else
        Call MsgBox("Не найден файл - " & FILE_NAME, vbInformation, "HotKeys:")
    End If
    Exit Sub
errMsg:
    Call WriteErrorLog("hotKeysStart")
End Sub

Public Sub hotKeysStop()
    On Error GoTo errMsg

    Dim sPatpApp    As String
    sPatpApp = ThisWorkbook.Path & Application.PathSeparator & FILE_NAME
    If FileHave(sPatpApp) Then
        Dim WshShell As Object
        Set WshShell = CreateObject("WScript.Shell")
        If Not WshShell Is Nothing Then
            Dim WshExec As Object
            Set WshExec = WshShell.Exec(sPatpApp)
            If Not WshShell Is Nothing Then WshExec.Terminate
            Set WshExec = Nothing
        End If
        Set WshShell = Nothing
    End If

    Exit Sub
errMsg:
    Call WriteErrorLog("hotKeysStop")
End Sub
