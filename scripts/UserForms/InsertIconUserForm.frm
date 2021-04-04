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
    Unload Me
End Sub
Private Sub Label821_Click()
    Dim i           As Integer
    Dim cnt         As MSForms.control

    For Each cnt In Me.Controls
        If TypeName(cnt) = "Label" Then
            i = i + 1
        End If
    Next cnt
    Debug.Print i - 6
End Sub
Private Sub UserForm_Initialize()
    Set Emitter = New EventListenerEmitter
    With Emitter
        .AddEventListenerAll Me
    End With
End Sub
Private Sub Emitter_MouseOut(control As Object)
    Call FrmBtnColor(control, vbWhite)
End Sub
Private Sub Emitter_MouseOver(control As Object)
    Call FrmBtnColor(control, vbBlack, vbWhite, vbRed)
    If control.Tag <> "No" Then
        With control
            lbNameFont.Caption = .Font
            lbCapiton.Caption = .Caption
            lbASC.Caption = VBA.AscW(.Caption)
        End With
    End If
End Sub
Public Sub FrmBtnColor(ByRef control As Object, ByVal BackColor As Long, Optional ByVal ForeColor As Long = vbBlack, Optional ByVal BorderColor As Long = vbBlack)
    With control
        If .Tag <> "No" Then
            With control
                .BackColor = BackColor
                .ForeColor = ForeColor
                .BorderColor = BorderColor
            End With
        End If
    End With
End Sub

