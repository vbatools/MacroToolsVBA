VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertIconUserForm 
   Caption         =   "Icons:"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20265
   OleObjectBlob   =   "InsertIconUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "No"
End
Attribute VB_Name = "InsertIconUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1
  Private Sub Emitter_DblClick(control As Object, Cancel As MSForms.ReturnBoolean)
4:    Unload Me
5: End Sub
    Private Sub Label821_Click()
7:    Dim i           As Integer
8:    Dim cnt         As MSForms.control
9:
10:    For Each cnt In Me.Controls
11:        If TypeName(cnt) = "Label" Then
12:            i = i + 1
13:        End If
14:    Next cnt
15:    Debug.Print i - 6
16: End Sub
    Private Sub UserForm_Initialize()
18:    Set Emitter = New EventListenerEmitter
19:    With Emitter
20:        .AddEventListenerAll Me
21:    End With
22: End Sub
    Private Sub Emitter_MouseOut(control As Object)
24:    Call FrmBtnColor(control, vbWhite)
25: End Sub
    Private Sub Emitter_MouseOver(control As Object)
27:    Call FrmBtnColor(control, vbBlack, vbWhite, vbRed)
28:    If control.Tag <> "No" Then
29:        With control
30:            lbNameFont.Caption = .Font
31:            lbCapiton.Caption = .Caption
32:            lbASC.Caption = VBA.AscW(.Caption)
33:        End With
34:    End If
35: End Sub
    Public Sub FrmBtnColor(ByRef control As Object, ByVal BackColor As Long, Optional ByVal ForeColor As Long = vbBlack, Optional ByVal BorderColor As Long = vbBlack)
37:    With control
38:        If .Tag <> "No" Then
39:            With control
40:                .BackColor = BackColor
41:                .ForeColor = ForeColor
42:                .BorderColor = BorderColor
43:            End With
44:        End If
45:    End With
46: End Sub

