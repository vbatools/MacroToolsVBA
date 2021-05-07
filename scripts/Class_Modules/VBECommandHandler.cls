VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBECommandHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : VBECommandHandler - обратные вызовы для меню VBE
'* Created    : 01-20-2020 15:55
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Public WithEvents EvtHandler As CommandBarButton
Attribute EvtHandler.VB_VarHelpID = -1
Private Sub EvtHandler_Click(ByVal Ctrl As CommandBarButton, CancelDefault As Boolean)
12:    On Error GoTo ErrorHandler
13:    sTagNameConrol = Ctrl.Tag
14:    Application.Run Ctrl.OnAction
15:    Exit Sub
ErrorHandler:
17:    If Err.Number <> 0 Then
18:        Debug.Print "Error in Eventhandler_Click" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
19:        Call WriteErrorLog("EvtHandler_Click")
20:        Err.Clear
21:        Resume Next
22:    End If
End Sub
