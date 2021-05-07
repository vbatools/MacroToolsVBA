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
10:    cmbMain.Clear
11:    cmbMain.Value = vbNullString
12:    Unload Me
13: End Sub
    Private Sub lbCancel_Click()
15:    Call cmbCancel_Click
16: End Sub
    Private Sub lbOK_Click()
18:    Unload Me
19: End Sub
Private Sub UserForm_Activate()
21:    Me.StartUpPosition = 0
22:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
23:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
24:
25:    Dim vbProj      As VBIDE.VBProject
26:    On Error Resume Next
27:    With cmbMain
28:        .Clear
29:        For Each vbProj In Application.VBE.VBProjects
30:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
31:        Next
32:        .Value = ActiveWorkbook.Name
33:    End With
34:    Exit Sub
End Sub
