VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InfoFile 
   Caption         =   "Свойства файла:"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13455
   OleObjectBlob   =   "InfoFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InfoFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : InfoFile - управление свойствами файла
'* Created    : 20-07-2020 15:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private Sub cmbMain_Change()
    On Error Resume Next
    Call UpdeteList(Me.ListCode, X_InfoFile.ShowProp(Workbooks(cmbMain.Value)))
    Call UpdeteList(Me.ListCustomProp, X_InfoFile.ShowCustomProp(Workbooks(cmbMain.Value)))
    On Error GoTo 0
End Sub
Private Sub UpdeteList(ByRef objList As MSForms.ListBox, ByVal Txt As String)
    Dim Arr         As Variant
    Dim i           As Byte
    objList.Clear
    If Txt <> vbNullString Then
        Arr = VBA.Split(Txt, vbNewLine)
        With objList
            For i = 0 To UBound(Arr)
                If Arr(i) <> vbNullString Then
                    .AddItem i + 1
                    .List(i, 1) = VBA.Split(Arr(i), ": ")(0)
                    .List(i, 2) = VBA.Split(Arr(i), ": ")(1)
                End If
            Next i
        End With
    End If
End Sub

Private Sub Label2_Click()
    Me.Hide
    Call InfoFile2.Show
    Call cmbMain_Change
    Me.Show
End Sub

Private Sub LbDelAllProper_Click()
    If MsgBox("Удалить ВСЕ свойства ?", vbYesNo + vbQuestion, "Удаление свойств:") = vbYes Then
        Dim iCount  As Byte
        iCount = X_InfoFile.DelAllProp(Workbooks(cmbMain.Value))
        Call cmbMain_Change
        Call MsgBox("Удалено свойств: " & iCount, vbInformation, "Удаление свойств:")
    End If
End Sub
Private Sub LbEdit_Click()
    Call EditProp
End Sub

Private Sub lbTemplete_Click()
    Dim tbData As Variant
    Dim i As Integer
    tbData = ThisWorkbook.Worksheets(C_Const.SH_SNIPPETS).ListObjects("TB_TEMPLETE").DataBodyRange.Value2
    tbData = ThisWorkbook.Worksheets(C_Const.SH_SNIPPETS).ListObjects("TB_TEMPLETE").DataBodyRange.Value2
    For i = 1 To UBound(tbData)
        Call X_InfoFile.AddOneCustomProp(Workbooks(cmbMain.Value), tbData(i, 1), tbData(i, 2))
    Next i
    Call cmbMain_Change
End Sub

Private Sub ListCode_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call EditProp
End Sub
Private Sub EditProp()
    Dim txtNew      As String
    Dim txtOld      As String
    Dim NameProp    As String
    With Me.ListCode
        If IsNumeric(.BoundValue) Then
            txtOld = VBA.Trim$(.List(CInt(.BoundValue) - 1, 2))
            NameProp = .List(CInt(.BoundValue) - 1, 1)
            txtNew = InputBox("Рактировать свойство [ " & NameProp & " ] ?", "Редактирование свойства:", txtOld)
            If txtNew <> txtOld Then
                Call X_InfoFile.WriteOneProp(Workbooks(cmbMain.Value), NameProp, txtNew)
                Call cmbMain_Change
            End If
        End If
    End With
End Sub

Private Sub lbAddCustProp_Click()
    Call AddCustProp(vbNullString, vbNullString)
End Sub

Private Sub lbEditCustProp_Click()

    Dim txtOld      As String
    Dim NameProp    As String
    With Me.ListCustomProp
        If IsNumeric(.BoundValue) Then
            txtOld = VBA.Trim$(.List(CInt(.BoundValue) - 1, 2))
            NameProp = .List(CInt(.BoundValue) - 1, 1)
            Call X_InfoFile.DelOneCustomProp(Workbooks(cmbMain.Value), NameProp)
            Call AddCustProp(NameProp, txtOld)
        End If
    End With
End Sub
Private Sub lbDelOneCustProp_Click()
    Dim NameProp    As String
    With Me.ListCustomProp
        If IsNumeric(.BoundValue) Then
            NameProp = .List(CInt(.BoundValue) - 1, 1)
            If MsgBox("Удалить свойство [ " & NameProp & " ] ?", vbYesNo + vbQuestion, "Удаление свойства:") = vbYes Then
                Call X_InfoFile.DelOneCustomProp(Workbooks(cmbMain.Value), NameProp)
                Call cmbMain_Change
            End If
        End If
    End With
End Sub
Private Sub AddCustProp(ByVal txtPropName As String, ByVal txtPropValue As String)
    txtPropName = InputBox("Ведите  название свойства", "Создание свойства:", txtPropName)
    If txtPropName <> vbNullString Then
        txtPropValue = InputBox("Ведите  значение свойства", "Создание свойства:", txtPropValue)
        If txtPropValue <> vbNullString Then
            Call X_InfoFile.AddOneCustomProp(Workbooks(cmbMain.Value), txtPropName, txtPropValue)
            Call cmbMain_Change
        End If
    End If
End Sub


Private Sub lbDelAllCustomProp_Click()
    If MsgBox("Удалить ВСЕ свойства ?", vbYesNo + vbQuestion, "Удаление свойств:") = vbYes Then
        Dim iCount  As Byte
        iCount = X_InfoFile.DelAllCustomProp(Workbooks(cmbMain.Value))
        Call cmbMain_Change
        Call MsgBox("Удалено свойств: " & iCount, vbInformation, "Удаление свойств:")
    End If
End Sub

Private Sub UserForm_Activate()
    Dim vbProj      As VBIDE.VBProject
    If Workbooks.Count = 0 Then
        Unload Me
        Call MsgBox("Нет открытых " & Chr(34) & "Файлов Excel" & Chr(34) & "!", vbOKOnly + vbExclamation, "Ошибка:")
        Exit Sub
    End If
    With Me.cmbMain
        .Clear
        On Error Resume Next
        For Each vbProj In Application.VBE.VBProjects
            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
        Next
        On Error GoTo 0
        .Value = ActiveWorkbook.Name
    End With
End Sub

Private Sub cmbCancel_Click()
    Unload Me
End Sub
Private Sub lbCancel_Click()
    Call cmbCancel_Click
End Sub
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
