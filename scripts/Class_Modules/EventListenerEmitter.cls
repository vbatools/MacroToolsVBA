VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventListenerEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'ARRAY OF ALL THE DIFFERENT EVENT ListenerS FOR EVERY USERFORM CONTROL +FORM ITSELF
Private EventList() As New EventListenerItem

'ALL CURRENT POSSIBLE EVENTS THAT CAN BE EMITTED. NOTE, EMITTEDEVENT IS SENT FOR ALL EVENTS!
Public Event EmittedEvent(ByRef control As Object, ByVal EventName As EmittedEvent, ByRef EventParameters As Scripting.Dictionary)
Public Event Click(ByRef control As Object)
Public Event DblClick(ByRef control As Object, ByRef Cancel As MSForms.ReturnBoolean)
Public Event KeyUp(ByRef control As Object, ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
Public Event KeyDown(ByRef control As Object, ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
Public Event MouseOver(ByRef control As Object)
Public Event MouseOut(ByRef control As Object)
Public Event MouseMove(ByRef control As Object, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
Public Event Focus(ByRef control As Object)
Public Event Blur(ByRef control As Object)
Public Event Change(ByRef control As Object)

Public Enum EmittedEvent
    Click
    DoubleClick
    MouseMove
    MouseOut
    MouseOver
    MouseDown
    MouseUp
    KeyUp
    KeyDown
    Focus
    Blur
End Enum

'***********************************************************************************
' PUBLIC METHODS
'***********************************************************************************

'CALLED BY EVENTLISTENERCOLLECTION CLASS - MAIN ENTRYWAY OF EMITTING ALL EVENTS
    Public Sub EmitEvent(ByRef control As Object, ByVal EventName As EmittedEvent, Optional ByRef EventParameters As Scripting.Dictionary)
40:
41:    'EVENT RAISED FOR ALL EVENTS. THIS IS A WAY FOR THE USER TO COLLECT FROM A SINGLE LOCATION.
42:    RaiseEvent EmittedEvent(control, EventName, EventParameters)
43:
44:    'SPECIFIC EVENTS PER OBJECT-TYPE
45:    Select Case EventName
    
           Case Click
48:            RaiseEvent Click(control)
49:
50:        Case DoubleClick
51:            RaiseEvent DblClick(control, EventParameters("Cancel"))
52:
53:        Case KeyUp
54:            RaiseEvent KeyUp(control, EventParameters("KeyCode"), EventParameters("Shift"))
55:
56:        Case KeyDown
57:            RaiseEvent KeyDown(control, EventParameters("KeyCode"), EventParameters("Shift"))
58:
59:        Case MouseOver
60:            RaiseEvent MouseOver(control)
61:
62:        Case MouseOut
63:            RaiseEvent MouseOut(control)
64:
65:        Case Focus
66:            RaiseEvent Focus(control)
67:
68:        Case Blur
69:            RaiseEvent Blur(control)
70:
71:        Case MouseMove
72:            RaiseEvent MouseMove(control, EventParameters("Shift"), EventParameters("X"), EventParameters("Y"))
73:
74:    End Select
75:
76: End Sub

'MUST CALL THIS IF YOU WANT TO programmatically SET CONTROL! OTHERWISE, EVENT'S WILL BE OFF!
    Public Sub SetFocusToControl(ByRef control As Object)
80:
81:    'If the user was to set focus through VBA then this code will fall apart considering
82:    'it is unaware of that event occurring.
83:    If Not control Is Nothing Then
84:        control.setFocus
85:        EmitEvent control, Focus
86:    End If
87:
88: End Sub

'ADD EVENT ListenerS ON SPECIFIC CONTROLS - ALSO CALLED BY AddEventListenerAll
     Public Sub AddEventListener(ByRef control As Object)
92:
93:    'Events are stored in a private EventListenerItem array
94:    If IsArrayEmpty(EventList) Then
95:        ReDim EventList(0 To 0)
96:    Else
97:        ReDim Preserve EventList(0 To UBound(EventList) + 1)
98:    End If
99:
100:    'CALL AddEventListener IN EventListenerItem. THIS IS KEPT IN
101:    EventList(UBound(EventList)).AddEventListener control, Me
102:
103: End Sub

'ADD EVENT Listener TO ALL CONTROLS INCLUDING THE FORM
     Public Sub AddEventListenerAll(ByRef Form As Object)
107:
108:    AddEventListener Form
109:
110:    Dim Ctrl As MSForms.control
111:    For Each Ctrl In Form.Controls
112:        AddEventListener Ctrl
113:    Next Ctrl
114:
115: End Sub

     Private Function IsArrayEmpty(Arr As Variant) As Boolean
118: ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
119: ' CPEARSON: http://www.cpearson.com/excel/VBAArrays.htm
120: ' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
121: '
122: ' The VBA IsArray function indicates whether a variable is an array, but it does not
123: ' distinguish between allocated and unallocated arrays. It will return TRUE for both
124: ' allocated and unallocated arrays. This function tests whether the array has actually
125: ' been allocated.
126: '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
127:
128:    Err.Clear
129:    On Error Resume Next
130:    If IsArray(Arr) = False Then
131:        ' we weren't passed an array, return True
132:        IsArrayEmpty = True
133:    End If
134:
135:    ' Attempt to get the UBound of the array. If the array is
136:    ' unallocated, an error will occur.
137:    Dim ub As Long
138:    ub = UBound(Arr, 1)
139:    If (Err.Number <> 0) Then
140:        IsArrayEmpty = True
141:    Else
142:        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
143:        ' On rare occasion, under circumstances I cannot reliably replicate, Err.Number
144:        ' will be 0 for an unallocated, empty array. On these occasions, LBound is 0 and
145:        ' UBound is -1. To accommodate the weird behavior, test to see if LB > UB.
146:        ' If so, the array is not allocated.
147:        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
148:        Err.Clear
149:        Dim LB As Long
150:        LB = LBound(Arr)
151:        If LB > ub Then
152:            IsArrayEmpty = True
153:        Else
154:            IsArrayEmpty = False
155:        End If
156:    End If
157:
158: End Function


