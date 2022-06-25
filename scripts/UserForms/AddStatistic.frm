VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddStatistic 
   Caption         =   "Collecting VBA project statistics:"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8565
   OleObjectBlob   =   "AddStatistic.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddStatistic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : AddStatistic - פמנלא הכÿ גûבמנמג פאיכמג
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
    Private Sub cmbCancel_Click()
11:    cmbMain.Clear
12:    cmbMain.Value = vbNullString
13:    Unload Me
14: End Sub
    Private Sub lbCancel_Click()
16:    Call cmbCancel_Click
17: End Sub
    Private Sub lbOK_Click()
19:    Unload Me
20: End Sub
    Private Sub UserForm_Activate()
22:    Me.StartUpPosition = 0
23:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
24:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
25:
26:    Dim vbProj      As VBIDE.VBProject
27:    On Error Resume Next
28:    With cmbMain
29:        .Clear
30:        For Each vbProj In Application.VBE.VBProjects
31:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
32:        Next
33:        If lbWord.Caption = "1" Then Call getWord(cmbMain)
34:        .Value = ActiveWorkbook.Name
35:    End With
36:    Exit Sub
37: End Sub

    Private Sub getWord(ByRef oList As MSForms.ComboBox)
40:    On Error Resume Next
41:    Dim objW        As Object
42:    Dim vbProj      As VBIDE.VBProject
43:    Dim sVal        As String
44:    Set objW = GetObject(, "Word.Application")
45:    For Each vbProj In objW.VBE.VBProjects
46:        sVal = C_PublicFunctions.sGetFileName(vbProj.Filename)
47:        If sVal Like "*.docm" Or sVal Like "*.DOCM" Then oList.AddItem sVal
48:    Next
49: End Sub

