VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddStatistic 
   Caption         =   "Collecting Vba Project Statistics:"
   ClientHeight    =   1470
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
12:    cmbMain.Clear
13:    cmbMain.Value = vbNullString
14:    Unload Me
15: End Sub
    Private Sub lbCancel_Click()
17:    Call cmbCancel_Click
18: End Sub
    Private Sub lbOK_Click()
20:    Unload Me
21: End Sub
Private Sub UserForm_Activate()
23:    Me.StartUpPosition = 0
24:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
25:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
26:
27:    Dim vbProj      As VBIDE.VBProject
28:    On Error Resume Next
29:    With cmbMain
30:        .Clear
31:        For Each vbProj In Application.VBE.VBProjects
32:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
33:        Next
34:        .Value = ActiveWorkbook.Name
35:    End With
36:    Exit Sub
End Sub
