Attribute VB_Name = "M_MoveControl"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : M_MoveControl - Микро подстройка элементов формы VBA и переименование элементов на форме вместе с кодом
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module
Public sTagNameConrol  As String
Public tpStyle      As ProperControlStyle
Type ProperControlStyle
    sError          As String

    snHeight        As Single
    snWidth         As Single

    bVisible        As Boolean
    bEnabled        As Boolean
    bLocked         As Boolean

    lBackColor      As Long
    lForeColor      As Long
    lBackStyle      As Long

    lBorderColor    As Long
    lBorderStyle    As Long

    bFontBold       As Boolean
    bFontItalic     As Boolean
    bFontStrikethru As Boolean
    bFontUnderline  As Boolean
    sFontName       As String
    cuFontSize      As Currency
End Type
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : MoveControl - Микроподстройка элементов формы
'* Created    : 08-10-2020 14:10
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub MoveControl()
    Dim myCommandBar As CommandBar
    Dim cntrl  As CommandBarControl
    Dim combox As CommandBarComboBox
    Dim sComBoxText As String
    Dim cnt    As control

    Set myCommandBar = Application.VBE.CommandBars(C_Const.MENUMOVECONTRL)
    For Each cntrl In myCommandBar.Controls
        If cntrl.Tag = C_Const.MTAGCOM Then
            Set combox = myCommandBar.Controls(cntrl.ID)
            sComBoxText = combox.Text
            Exit For
        End If
    Next cntrl

    Set cnt = TakeSelectControl
    If cnt Is Nothing Then Exit Sub
    Select Case sTagNameConrol
        Case C_Const.MTAG1:
            Call MoveCnt(cnt, 1, sComBoxText)
        Case C_Const.MTAG2:
            Call MoveCnt(cnt, 2, sComBoxText)
        Case C_Const.MTAG3:
            Call MoveCnt(cnt, 3, sComBoxText)
        Case C_Const.MTAG4:
            Call MoveCnt(cnt, 4, sComBoxText)
    End Select
End Sub
Private Sub MoveCnt(ByRef cnt As control, ByVal iVal As Integer, ByVal sComBoxText As String)
    Const Shag = 0.4
    With cnt
        Select Case sComBoxText
            Case C_Const.MOVECONT:
                Select Case iVal
                    Case 1:
                        .Left = .Left - Shag
                    Case 2:
                        .Left = .Left + Shag
                    Case 3:
                        .top = .top + Shag
                    Case 4:
                        .top = .top - Shag
                End Select
            Case C_Const.MOVECONTTOPLEFT:
                Select Case iVal
                    Case 1:
                        .Left = .Left - Shag
                        .Width = .Width + Shag
                    Case 2:
                        .Left = .Left + Shag
                        .Width = .Width - Shag
                    Case 3:
                        .top = .top + Shag
                        .Height = .Height - Shag
                    Case 4:
                        .top = .top - Shag
                        .Height = .Height + Shag
                End Select
            Case C_Const.MOVECONTBOTTOMRIGHT:
                Select Case iVal
                    Case 1:
                        .Width = .Width - Shag
                    Case 2:
                        .Width = .Width + Shag
                    Case 3:
                        .Height = .Height + Shag
                    Case 4:
                        .Height = .Height - Shag
                End Select
        End Select
    End With
End Sub

Private Function TakeSelectControl(Optional bUserForm As Boolean = False) As Object
    Dim W           As VBIDE.Window
    Dim strVar()    As String
    Dim cntName     As String

    On Error GoTo ErrorHandler

    If Application.VBE.ActiveWindow.Type = vbext_wt_Designer Then
        For Each W In Application.VBE.Windows
            If W.Type = vbext_wt_PropertyWindow Then
                strVar = Split(W.Caption, "-")
                cntName = Trim(strVar(1))
                Exit For
            End If
        Next

        Dim Form    As UserForm
        Set Form = Application.VBE.SelectedVBComponent.Designer
        Set TakeSelectControl = Form.Controls(cntName)
    End If
    Exit Function
ErrorHandler:
    If bUserForm And Not Form Is Nothing Then
        Err.Clear
        Set TakeSelectControl = Form
        Exit Function
    End If
    Select Case Err.Number
        Case -2147024809:
            Debug.Print "Error Select one object"
        Case 9:
            Debug.Print "To use the tool, open the View - > Properties Window"
        Case Else:
            Debug.Print "Error in TakeSelectControl" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
            Call WriteErrorLog("TakeSelectControl")
    End Select
    Err.Clear
End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RenameControl - переименование конторол на форме вместе скодом
'* Created    : 08-10-2020 14:11
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub RenameControl()
    Dim cnt    As control
    Dim sNewName As String
    Dim sOldName As String
    Dim NameModeCode As String
    Dim strVar As String
    Dim CodeMod As CodeModule

    On Error GoTo ErrorHandler

    Set cnt = TakeSelectControl
    If cnt Is Nothing Then Exit Sub
tryagin:
    sOldName = cnt.Name
    sNewName = InputBox("Enter a new name for Control", "Renaming Control:", sOldName)
    If sNewName = vbNullString Or sNewName = sOldName Then Exit Sub

    cnt.Name = sNewName
    Set CodeMod = Application.VBE.SelectedVBComponent.CodeModule
    With CodeMod
        strVar = .Lines(1, .CountOfLines)
        'strVar = Replace(strVar, sOldName, sNewName)
        strVar = ReplceCode(strVar, sOldName, sNewName)
        .DeleteLines StartLine:=1, Count:=.CountOfLines
        .InsertLines Line:=1, String:=strVar
    End With
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 40044:
            Call MsgBox("Error Invalid Control name entered [" & sNewName & "], set a different name!", vbCritical, "Invalid Control name entered:")
            Exit Sub
        Case -2147319764:
            Call MsgBox("This Control name is already in use [" & sNewName & "], set a different name!", vbCritical, "The name is ambiguous:")
            Exit Sub
        Case Else:
            Debug.Print "Error in RenameControl" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
            Call WriteErrorLog("RenameControl")
    End Select
    Err.Clear
End Sub
Public Sub CopyStyleControl()
    Dim cnt         As Object
    Set cnt = TakeSelectControl(True)
    If cnt Is Nothing Then Exit Sub

    'установка по умолчанию значений
    tpStyle.lBackStyle = 1
    tpStyle.lBorderColor = -2147483642
    tpStyle.lBorderStyle = 0
    tpStyle.bVisible = True
    tpStyle.bLocked = False
    tpStyle.bEnabled = True
    tpStyle.lBackStyle = 1

    On Error Resume Next
    With cnt
        tpStyle.bEnabled = .Enabled
        tpStyle.bFontBold = .Font.Bold
        tpStyle.bFontItalic = .Font.Italic
        tpStyle.bFontStrikethru = .Font.Strikethrough
        tpStyle.bFontUnderline = .Font.Underline
        tpStyle.bLocked = .Locked
        tpStyle.bVisible = .visible
        tpStyle.cuFontSize = .Font.Size
        tpStyle.lBackColor = .BackColor
        tpStyle.lForeColor = .ForeColor
        tpStyle.sFontName = .Font.Name
        tpStyle.snHeight = .Height
        tpStyle.snWidth = .Width

        tpStyle.lBackStyle = .BackStyle
        tpStyle.lBorderColor = .BorderColor
        tpStyle.lBorderStyle = .BorderStyle
    End With
End Sub
Public Sub PasteStyleControl()
    Dim cnt         As Object
    Set cnt = TakeSelectControl(True)
    If cnt Is Nothing Then Exit Sub
    On Error Resume Next
    With cnt
        .Enabled = tpStyle.bEnabled
        .Font.Bold = tpStyle.bFontBold
        .Font.Italic = tpStyle.bFontItalic
        .Font.Strikethrough = tpStyle.bFontStrikethru
        .Font.Underline = tpStyle.bFontUnderline
        .Locked = tpStyle.bLocked
        .visible = tpStyle.bVisible
        .Font.Size = tpStyle.cuFontSize
        .BackColor = tpStyle.lBackColor
        .ForeColor = tpStyle.lForeColor
        .Font.Name = tpStyle.sFontName
        If tpStyle.snHeight > 0 Then .Height = tpStyle.snHeight
        If tpStyle.snWidth > 0 Then .Width = tpStyle.snWidth

        .BackStyle = tpStyle.lBackStyle
        .BorderColor = tpStyle.lBorderColor
        .BorderStyle = tpStyle.lBorderStyle
    End With
End Sub
Public Sub AddIcon()
    Dim cnt    As control
    Dim objForm As InsertIconUserForm

    On Error GoTo ErrorHandler

    Set cnt = TakeSelectControl
    If cnt Is Nothing Then Exit Sub
    
    Set objForm = New InsertIconUserForm
    With objForm
        .Show
        cnt.Font.Name = .lbNameFont.Caption
        DoEvents
        If TypeName(cnt) = "Label" Then
            cnt.Caption = VBA.ChrW$(.lbASC.Caption)
        Else
            cnt.Value = VBA.ChrW$(.lbASC.Caption)
        End If
    End With

    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case Else:
            Debug.Print "Error in RenameControl" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
            Call WriteErrorLog("AddIcon")
    End Select
    Err.Clear
End Sub
Public Sub HelpMoveControl()
    Call URLLinks(C_Const.URL_MOVE_CNTR)
End Sub
