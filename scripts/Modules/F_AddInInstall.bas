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
11:    Dim AddFolder As String
12:    On Error GoTo InstallationAdd_Err
13:    ' Проверяем имеется ли данная директория
14:    AddFolder = Replace(Application.UserLibraryPath & "\", "\\", "\")
15:    'проверка на наличие дириктории
16:    If Dir(AddFolder, vbDirectory) = vbNullString Then
17:        Call MsgBox("Unfortunately, the program cannot install the add-in on this computer." _
                         & vbCrLf & "The settings directory is missing." & vbCrLf & _
                         "Contact the program developer.", vbCritical, _
                         "Add-in installation failed")
21:        Exit Sub
22:    End If
23:    'Отключаем ранее установленую надстройку
24:    If FileHave(AddFolder & C_Const.NAME_ADDIN & ".xlam") Then AddIns(C_Const.NAME_ADDIN).Installed = False
25:    ' Проверяем открыта ли надстройка
26:    If WorkbookIsOpen(C_Const.NAME_ADDIN & ".xlam") Then
27:        Call MsgBox("The file with the add-in is already open." & vbCrLf & _
                         "It may have already been installed earlier.", vbCritical, _
                         "Program installation failed")
30:        Exit Sub
31:    End If
32:    ' Сохраняем как
33:    Application.EnableEvents = 0
34:    Application.DisplayAlerts = False
35:    If Workbooks.Count = 0 Then Workbooks.Add
36:    ThisWorkbook.SaveAs AddFolder & C_Const.NAME_ADDIN & ".xlam", FileFormat:=xlOpenXMLAddIn
37:    AddIns.Add Filename:=AddFolder & C_Const.NAME_ADDIN & ".xlam"
38:    AddIns(C_Const.NAME_ADDIN).Installed = True
39:    Application.EnableEvents = 1
40:    Application.DisplayAlerts = True
41:    Call MsgBox("The program is installed successfully!" & vbCrLf & _
                     "Just open or create a new document.", vbInformation, _
                     "Installing the add-in:" & C_Const.NAME_ADDIN)
44:    ThisWorkbook.Close False
45:    Exit Sub
InstallationAdd_Err:
47:    If Err.Number = 1004 Then
48:        MsgBox "To install the add-in, please close this file and run it again.", _
                         64, "Installation"
50:    Else
51:        MsgBox Err.Description & vbCrLf & "в F_AddInInstall.InstallationAdd " & vbCrLf & "in the line " & Erl, vbExclamation + vbOKOnly, "Error:"
52:        Call WriteErrorLog("F_AddInInstall.InstallationAdd")
53:    End If
54: End Sub

