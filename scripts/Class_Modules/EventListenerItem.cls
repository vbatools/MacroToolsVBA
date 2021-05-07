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
44:
45:    'CAPTURE THE EMITTER CLASS. WILL USE THIS TO EMIT EVENTS FROM EACH CONTROL
46:    Set pEmitter = Emitter
47:
48:    'USED TO COMPARE CHECK IF IT IS THE CONTROL TRIGGERING THE EVENT
49:    Set pControl = ControlOrForm
50:
51:    'SET CONTROL(OR FORM) BASED ON IT'S TYPE
52:    Select Case TypeName(ControlOrForm)
        Case "CommandButton"
54:            Set btn = ControlOrForm
55:
56:        Case "ComboBox"
57:            Set Cmb = ControlOrForm
58:
59:        Case "Frame"
60:            Set frm = ControlOrForm
61:
62:        Case "Label"
63:            Set Lbl = ControlOrForm
64:
65:        Case "TextBox"
66:            Set Txt = ControlOrForm
67:        Case "MultiPage"
68:
69:        Case Else
70:
71:            If TypeOf ControlOrForm Is MSForms.UserForm Then
72:                Set Form = ControlOrForm
73:            Else
74:                Err.Raise 5, TypeName(Me), "Invalid control: Currently unable to listen to events on " & TypeName(ControlOrForm)
75:            End If
76:
77:    End Select
78:
79: End Sub

'***********************************************************************************
' PRIVATE HELPER FUNCTIONS
'***********************************************************************************

'CALLED ON MOUSEMOVE EVENT, THIS IS A WAY OF CREATING A MOUSEOVER AND MOUSEOUT EVENT
    Private Sub CheckIfHoveredControl()
87:
88:    If Not IsHoveredControl Then
89:        IsHoveredControl = True
90:        pEmitter.EmitEvent pControl, MouseOver
91:    End If
92:
93: End Sub

'CALLED ON MOUSEMOVE EVENT, THIS IS A WAY OF CREATING A MOUSEOVER AND MOUSEOUT EVENT
     Private Sub CheckIfFocusedControl()
97:
98:    If Not IsFocusedControl Then
99:
100:        If TypeName(pControl) = "Frame" Then
101:             pEmitter.SetFocusToControl pControl.ActiveControl
102:        Else
103:            IsFocusedControl = True
104:            pEmitter.EmitEvent pControl, Focus
105:        End If
106:
107:    End If
108:
109: End Sub

'SIMPLE Dictionary FACTORY - FOR EASE OF USE
     Private Function ToDictionary(ParamArray KeyValuePairs() As Variant) As Scripting.Dictionary
113:
114:    'CHECK IF THERE IS EVEN PARAMETERS
115:    Dim ArrayLenght As Long
116:    ArrayLenght = UBound(KeyValuePairs) - LBound(KeyValuePairs) + 1
117:
118:    If ArrayLenght Mod 2 <> 0 Then
119:        Err.Raise 5, TypeName(Me), "Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
120:    End If
121:
122:    Set ToDictionary = New Scripting.Dictionary
123:    Dim Index As Long
124:    For Index = LBound(KeyValuePairs) To UBound(KeyValuePairs) Step 2
125:        ToDictionary.Add KeyValuePairs(Index), KeyValuePairs(Index + 1)
126:    Next Index
127:
128: End Function


'***********************************************************************************
' EVENTS
'***********************************************************************************

'----------------------------------------------------------------------------
' ONCE AN EVENT HAS EMMITED, EACH EVENTListenerITEM WILL LISTEN FOR THAT EVENT
' ак только событие будет отправлено, каждый EVENTListenerITEM будет прослушивать это событие
'----------------------------------------------------------------------------
     Private Sub pEmitter_EmittedEvent(ByRef control As Object, ByVal EventName As EmittedEvent, ByRef EventParameters As Scripting.Dictionary)
140:
141:    'CREATE A MOUSEOVER MOUSEOUT EVENTS
142:    Select Case EventName
    
            Case MouseOver
145:            If pControl.Name <> control.Name And IsHoveredControl Then
146:                IsHoveredControl = False
147:                pEmitter.EmitEvent pControl, MouseOut
148:            End If
149:
150:        Case Focus
151:            If pControl.Name <> control.Name And IsFocusedControl Then
152:                IsFocusedControl = False
153:                pEmitter.EmitEvent pControl, Blur
154:            ElseIf pControl.Name = control.Name And IsFocusedControl = False Then
155:                IsFocusedControl = True
156:            End If
157:
158:    End Select
159:
160: End Sub

'------------------------------------------------------------------------
' USERFORM
'------------------------------------------------------------------------
     Private Sub Form_Click()
166:    pEmitter.EmitEvent pControl, Click
167: End Sub

     Private Sub Form_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
170:    pEmitter.EmitEvent pControl, DoubleClick, ToDictionary("Cancel", Cancel)
171: End Sub

     Private Sub Form_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
174:   CheckIfHoveredControl
175:   pEmitter.EmitEvent pControl, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
176: End Sub

     Private Sub Form_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
179:    pEmitter.EmitEvent pControl, MouseDown, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
180: End Sub

     Private Sub Form_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
183:    pEmitter.EmitEvent pControl, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
184: End Sub


'------------------------------------------------------------------------
' COMMAND BUTTON
'------------------------------------------------------------------------
     Private Sub Btn_Click()
191:    pEmitter.EmitEvent pControl, Click
192: End Sub

     Private Sub Btn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
195:    pEmitter.EmitEvent pControl, DoubleClick, ToDictionary("Cancel", Cancel)
196: End Sub

     Private Sub Btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
199:    CheckIfHoveredControl
200:    pEmitter.EmitEvent pControl, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
201: End Sub

     Private Sub Btn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
204:    CheckIfFocusedControl
205:    pEmitter.EmitEvent pControl, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
206: End Sub

     Private Sub Btn_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
209:    pEmitter.EmitEvent pControl, MouseDown, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
210: End Sub

     Private Sub Btn_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
213:    CheckIfFocusedControl
214:    pEmitter.EmitEvent pControl, KeyUp, ToDictionary("KeyCode", KeyCode, "Shift", Shift)
215: End Sub

     Private Sub Btn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
218:    pEmitter.EmitEvent pControl, KeyDown, ToDictionary("KeyCode", KeyCode, "Shift", Shift)
219: End Sub


'------------------------------------------------------------------------
' LABEL
'------------------------------------------------------------------------
     Private Sub Lbl_Click()
226:    pEmitter.EmitEvent pControl, Click
227: End Sub

     Private Sub Lbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
230:    pEmitter.EmitEvent pControl, DoubleClick, ToDictionary("Cancel", Cancel)
231: End Sub

     Private Sub Lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
234:    CheckIfHoveredControl
235:    pEmitter.EmitEvent pControl, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
236: End Sub

     Private Sub lbl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
239:    pEmitter.EmitEvent pControl, MouseDown, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
240: End Sub

     Private Sub lbl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
243:    pEmitter.EmitEvent pControl, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
244: End Sub


'------------------------------------------------------------------------
' Frame
'------------------------------------------------------------------------
     Private Sub Frm_Click()
251:    pEmitter.EmitEvent pControl, Click
252: End Sub

     Private Sub Frm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
255:    pEmitter.EmitEvent pControl, DoubleClick, ToDictionary("Cancel", Cancel)
256: End Sub

     Private Sub Frm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
259:    CheckIfHoveredControl
260:    pEmitter.EmitEvent pControl, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
261: End Sub

     Private Sub Frm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
264:    CheckIfFocusedControl 'FRAME DOESN'T TAKE FOCUS BUT ACTIVE CONTROL IN FRAME DOES
265:    pEmitter.EmitEvent pControl, MouseDown
266: End Sub

     Private Sub Frm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
269:    pEmitter.EmitEvent pControl, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
270: End Sub


'------------------------------------------------------------------------
' Textbox
'------------------------------------------------------------------------
     Private Sub Txt_Click()
277:    pEmitter.EmitEvent pControl, Click
278: End Sub

     Private Sub Txt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
281:    pEmitter.EmitEvent pControl, DoubleClick, ToDictionary("Cancel", Cancel)
282: End Sub

     Private Sub Txt_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
285:    CheckIfHoveredControl
286:    pEmitter.EmitEvent pControl, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
287: End Sub

     Private Sub Txt_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
290:    CheckIfFocusedControl
291:    pEmitter.EmitEvent pControl, MouseUp, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
292: End Sub

     Private Sub Txt_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
295:    pEmitter.EmitEvent pControl, MouseDown, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
296: End Sub

     Private Sub Txt_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
299:    CheckIfFocusedControl
300:    pEmitter.EmitEvent pControl, KeyUp, ToDictionary("KeyCode", KeyCode, "Shift", Shift)
301: End Sub

     Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
304:    pEmitter.EmitEvent pControl, KeyDown, ToDictionary("KeyCode", KeyCode, "Shift", Shift)
305: End Sub


'------------------------------------------------------------------------
' Combobox
'------------------------------------------------------------------------
     Private Sub Cmb_Click()
312:    pEmitter.EmitEvent pControl, Click
313: End Sub

     Private Sub Cmb_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
316:    pEmitter.EmitEvent pControl, DoubleClick, ToDictionary("Cancel", Cancel)
317: End Sub

     Private Sub Cmb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
320:    CheckIfHoveredControl
321:    pEmitter.EmitEvent pControl, MouseMove, ToDictionary("Button", Button, "Shift", Shift, "X", X, "Y", Y)
322: End Sub




