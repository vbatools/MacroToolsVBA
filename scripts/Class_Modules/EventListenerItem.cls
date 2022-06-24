VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventListenerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'@dependancy Reference to Microsoft Scripting Runtime (For Dictionarys)

'SET FROM AddEventListener - NEEDED TO EMIT EVENT BACK TO IT.
Private WithEvents pEmitter As EventListenerEmitter
Attribute pEmitter.VB_VarHelpID = -1

'CONTROLS THAT HAVE THE EVENTS
Private WithEvents Form As MSForms.UserForm
Attribute Form.VB_VarHelpID = -1
Private WithEvents Txt As MSForms.Textbox
Attribute Txt.VB_VarHelpID = -1
Private WithEvents Lbl As MSForms.Label
Attribute Lbl.VB_VarHelpID = -1
Private WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Private WithEvents Cmb As MSForms.ComboBox
Attribute Cmb.VB_VarHelpID = -1
Private WithEvents frm As MSForms.Frame
Attribute frm.VB_VarHelpID = -1

'PROPERTIES OF CONTROL
Private pControl As Object
Private IsHoveredControl As Boolean
Private IsFocusedControl As Boolean

'TODO
' - ADD EVENTS FOR ALL USERFORM CONTROLS
' - ADD MORE EVENTS THAN THE LIST BELOW
' - DOUBLE CHECK THAT EMITTS ARE ALL CURRENT

'CURRENT EVENTS (WILL ADD MORE AS I HAVE NEED)
' - MouseOver
' - MouseOut
' - MouseMove
' - Click
' - DblClick
' - Focus
' - Blur
' - KeyUp
' - KeyDown

'***********************************************************************************
' PUBLIC FUNCTIONS
'***********************************************************************************

'ONLY PUBLIC METHOD. CALLED FROM EVENTListener CLASS MODULE
Public Sub AddEventListener(ByRef ControlOrForm As Object, ByRef Emitter As EventListenerEmitter)
    
    'CAPTURE THE EMITTER CLASS. WILL USE THIS TO EMIT EVENTS FROM EACH CONTROL
    Set pEmitter = Emitter
    
    'USED TO COMPARE CHECK IF IT IS THE CONTROL TRIGGERING THE EVENT
    Set pControl = ControlOrForm
    
    'SET CONTROL(OR FORM) BASED ON IT'S TYPE
    Select Case TypeName(ControlOrForm)
        Case "CommandButton"
            Set btn = ControlOrForm
        
        Case "ComboBox"
            Set Cmb = ControlOrForm
            
        Case "Frame"
            Set frm = ControlOrForm
        
        Case "Label"
            Set Lbl = ControlOrForm
        
        Case "TextBox"
            Set Txt = ControlOrForm
        Case "MultiPage"
            
        Case Else
 
            If TypeOf ControlOrForm Is MSForms.UserForm Then
                Set Form = ControlOrForm
            Else
                Err.Raise 5, TypeName(Me), "Invalid control: Currently unable to listen to events on " & TypeName(ControlOrForm)
            End If

    End Select
    
End Sub

'***********************************************************************************
' PRIVATE HELPER FUNCTIONS
'***********************************************************************************

'CALLED ON MOUSEMOVE EVENT, THIS IS A WAY OF CREATING A MOUSEOVER AND MOUSEOUT EVENT
Private Sub CheckIfHoveredControl()
    
    If Not IsHoveredControl Then
        IsHoveredControl = True
        pEmitter.EmitEvent pControl, MouseOver
    End If
    
End Sub

'CALLED ON MOUSEMOVE EVENT, THIS IS A WAY OF CREATING A MOUSEOVER AND MOUSEOUT EVENT
Private Sub CheckIfFocusedControl()
    
    If Not IsFocusedControl Then
    
        If TypeName(pControl) = "Frame" Then
             pEmitter.SetFocusToControl pControl.ActiveControl
        Else
            IsFocusedControl = True
            pEmitter.EmitEvent pControl, Focus
        End If
        
    End If
    
End Sub

'SIMPLE Dictionary FACTORY - FOR EASE OF USE
Private Function ToDictionary(ParamArray KeyValuePairs() As Variant) As Scripting.Dictionary
    
    'CHECK IF THERE IS EVEN PARAMETERS
    Dim ArrayLenght As Long
    ArrayLenght = UBound(KeyValuePairs) - LBound(KeyValuePairs) + 1
    
    If ArrayLenght Mod 2 <> 0 Then
        Err.Raise 5, TypeName(Me), "Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
    End If
    
    Set ToDictionary = New Scripting.Dictionary
    Dim Index As Long
    For Index = LBound(KeyValuePairs) To UBound(KeyValuePairs) Step 2
        ToDictionary.Add KeyValuePairs(Index), KeyValuePairs(Index + 1)
    Next Index
    
End Function


'***********************************************************************************
' EVENTS
'***********************************************************************************

'----------------------------------------------------------------------------
' ONCE AN EVENT HAS EMMITED, EACH EVENTListenerITEM WILL LISTEN FOR THAT EVENT
' ак только событие будет отправлено, каждый EVENTListenerITEM будет прослушивать это событие
'----------------------------------------------------------------------------
Private Sub pEmitter_EmittedEvent(ByRef control As Object, ByVal EventName As EmittedEvent, ByRef EventParameters As Scripting.Dictionary)
    
    'CREATE A MOUSEOVER MOUSEOUT EVENTS
    Select Case EventName
    
        Case MouseOver
            If pControl.Name <> control.Name And IsHoveredControl Then
                IsHoveredControl = False
                pEmitter.EmitEvent pControl, MouseOut
            End If
            
        Case Focus
            If pControl.Name <> control.Name And IsFocusedControl Then
                IsFocusedControl = False
                pEmitter.EmitEvent pControl, Blur
            ElseIf pControl.Name = control.Name And IsFocusedControl = False Then
                IsFocusedControl = True
            End If
    
    End Select

End Sub

'------------------------------------------------------------------------
' USERFORM
'------------------------------------------------------------------------
Private Sub Form_Click()
    pEmitter.EmitEvent pControl, Click
End Sub

Private Sub Form_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, DoubleClick, ToDictionary("Cancel", Cancel)
End Sub

Private Sub Form_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   CheckIfHoveredControl
   pEmitter.EmitEvent pControl, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Form_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, MouseDown, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Form_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub


'------------------------------------------------------------------------
' COMMAND BUTTON
'------------------------------------------------------------------------
Private Sub Btn_Click()
    pEmitter.EmitEvent pControl, Click
End Sub

Private Sub Btn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, DoubleClick, ToDictionary("Cancel", Cancel)
End Sub

Private Sub Btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Btn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfFocusedControl
    pEmitter.EmitEvent pControl, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Btn_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, MouseDown, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Btn_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckIfFocusedControl
    pEmitter.EmitEvent pControl, KeyUp, ToDictionary("KeyCode", KeyCode, "Shift", Shift)
End Sub

Private Sub Btn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    pEmitter.EmitEvent pControl, KeyDown, ToDictionary("KeyCode", KeyCode, "Shift", Shift)
End Sub


'------------------------------------------------------------------------
' LABEL
'------------------------------------------------------------------------
Private Sub Lbl_Click()
    pEmitter.EmitEvent pControl, Click
End Sub

Private Sub Lbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, DoubleClick, ToDictionary("Cancel", Cancel)
End Sub

Private Sub Lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub lbl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, MouseDown, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub lbl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub


'------------------------------------------------------------------------
' Frame
'------------------------------------------------------------------------
Private Sub Frm_Click()
    pEmitter.EmitEvent pControl, Click
End Sub

Private Sub Frm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, DoubleClick, ToDictionary("Cancel", Cancel)
End Sub

Private Sub Frm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Frm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfFocusedControl 'FRAME DOESN'T TAKE FOCUS BUT ACTIVE CONTROL IN FRAME DOES
    pEmitter.EmitEvent pControl, MouseDown
End Sub

Private Sub Frm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub


'------------------------------------------------------------------------
' Textbox
'------------------------------------------------------------------------
Private Sub Txt_Click()
    pEmitter.EmitEvent pControl, Click
End Sub

Private Sub Txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, DoubleClick, ToDictionary("Cancel", Cancel)
End Sub

Private Sub Txt_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Txt_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfFocusedControl
    pEmitter.EmitEvent pControl, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Txt_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    pEmitter.EmitEvent pControl, MouseDown, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub

Private Sub Txt_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckIfFocusedControl
    pEmitter.EmitEvent pControl, KeyUp, ToDictionary("KeyCode", KeyCode, "Shift", Shift)
End Sub

Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    pEmitter.EmitEvent pControl, KeyDown, ToDictionary("KeyCode", KeyCode, "Shift", Shift)
End Sub


'------------------------------------------------------------------------
' Combobox
'------------------------------------------------------------------------
Private Sub Cmb_Click()
    pEmitter.EmitEvent pControl, Click
End Sub

Private Sub Cmb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pEmitter.EmitEvent pControl, DoubleClick, ToDictionary("Cancel", Cancel)
End Sub

Private Sub Cmb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CheckIfHoveredControl
    pEmitter.EmitEvent pControl, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
End Sub




