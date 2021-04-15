Attribute VB_Name = "F_AddInInstall"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : F_AddInInstall - модуль установки надстройки
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
' Установка надстройки
    Public Sub InstallationAddMacro()
10:    Dim AddFolder As String
11:    On Error GoTo InstallationAdd_Err
12:    ' Проверяем имеется ли данная директория
13:    AddFolder = Replace(Application.UserLibraryPath & "\", "\\", "\")
14:    'проверка на наличие дириктории
15:    If Dir(AddFolder, vbDirectory) = vbNullString Then
16:        Call MsgBox("Unfortunately, the program cannot install the add-in on this computer." _
                      & vbCrLf & "The settings directory is missing." & vbCrLf & _
                      "Contact the program developer.", vbCritical, _
                      "Add-in installation failed")
20:        Exit Sub
21:    End If
22:    'Отключаем ранее установленую надстройку
23:    If FileHave(AddFolder & C_Const.NAME_ADDIN & ".xlam") Then AddIns(C_Const.NAME_ADDIN).Installed = False
24:    ' Проверяем открыта ли надстройка
25:    If WorkbookIsOpen(C_Const.NAME_ADDIN & ".xlam") Then
26:        Call MsgBox("The file with the add-in is already open." & vbCrLf & _
                      "It may have already been installed earlier.", vbCritical, _
                      "Program installation failed")
29:        Exit Sub
30:    End If
31:    ' Сохраняем как
32:    Application.EnableEvents = 0
33:    Application.DisplayAlerts = False
34:    If Workbooks.Count = 0 Then Workbooks.Add
35:    ThisWorkbook.SaveAs AddFolder & C_Const.NAME_ADDIN & ".xlam", FileFormat:=xlOpenXMLAddIn
36:    AddIns.Add Filename:=AddFolder & C_Const.NAME_ADDIN & ".xlam"
37:    AddIns(C_Const.NAME_ADDIN).Installed = True
38:    Application.EnableEvents = 1
39:    Application.DisplayAlerts = True
40:    Call MsgBox("The program is installed successfully!" & vbCrLf & _
                  "Just open or create a new document.", vbInformation, _
                  "Installing the add-in:" & C_Const.NAME_ADDIN)
43:    ThisWorkbook.Close False
44:    Exit Sub
InstallationAdd_Err:
46:    If Err.Number = 1004 Then
47:        MsgBox "To install the add-in, please close this file and run it again.", _
                      64, "Installation"
49:    Else
50:        MsgBox Err.Description & vbCrLf & "в F_AddInInstall.InstallationAdd " & vbCrLf & "in the line " & Erl, vbExclamation + vbOKOnly, "Error:"
51:        Call WriteErrorLog("F_AddInInstall.InstallationAdd")
52:    End If
53: End Sub

