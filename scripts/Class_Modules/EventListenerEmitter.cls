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
    
    'EVENT RAISED FOR ALL EVENTS. THIS IS A WAY FOR THE USER TO COLLECT FROM A SINGLE LOCATION.
    RaiseEvent EmittedEvent(control, EventName, EventParameters)
    
    'SPECIFIC EVENTS PER OBJECT-TYPE
    Select Case EventName
    
        Case Click
            RaiseEvent Click(control)
        
        Case DoubleClick
            RaiseEvent DblClick(control, EventParameters("Cancel"))
        
        Case KeyUp
            RaiseEvent KeyUp(control, EventParameters("KeyCode"), EventParameters("Shift"))
            
        Case KeyDown
            RaiseEvent KeyDown(control, EventParameters("KeyCode"), EventParameters("Shift"))
        
        Case MouseOver
            RaiseEvent MouseOver(control)
            
        Case MouseOut
            RaiseEvent MouseOut(control)
            
        Case Focus
            RaiseEvent Focus(control)
            
        Case Blur
            RaiseEvent Blur(control)
            
        Case MouseMove
            RaiseEvent MouseMove(control, EventParameters("Shift"), EventParameters("X"), EventParameters("Y"))
        
    End Select

End Sub

'MUST CALL THIS IF YOU WANT TO programmatically SET CONTROL! OTHERWISE, EVENT'S WILL BE OFF!
Public Sub SetFocusToControl(ByRef control As Object)
    
    'If the user was to set focus through VBA then this code will fall apart considering
    'it is unaware of that event occurring.
    If Not control Is Nothing Then
        control.setFocus
        EmitEvent control, Focus
    End If
    
End Sub

'ADD EVENT ListenerS ON SPECIFIC CONTROLS - ALSO CALLED BY AddEventListenerAll
Public Sub AddEventListener(ByRef control As Object)
    
    'Events are stored in a private EventListenerItem array
    If IsArrayEmpty(EventList) Then
        ReDim EventList(0 To 0)
    Else
        ReDim Preserve EventList(0 To UBound(EventList) + 1)
    End If
    
    'CALL AddEventListener IN EventListenerItem. THIS IS KEPT IN
    EventList(UBound(EventList)).AddEventListener control, Me
    
End Sub

'ADD EVENT Listener TO ALL CONTROLS INCLUDING THE FORM
Public Sub AddEventListenerAll(ByRef Form As Object)
    
    AddEventListener Form
    
    Dim Ctrl As MSForms.control
    For Each Ctrl In Form.Controls
        AddEventListener Ctrl
    Next Ctrl
    
End Sub

Private Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CPEARSON: http://www.cpearson.com/excel/VBAArrays.htm
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Err.Clear
    On Error Resume Next
    If IsArray(Arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    Dim ub As Long
    ub = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' On rare occasion, under circumstances I cannot reliably replicate, Err.Number
        ' will be 0 for an unallocated, empty array. On these occasions, LBound is 0 and
        ' UBound is -1. To accommodate the weird behavior, test to see if LB > UB.
        ' If so, the array is not allocated.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        Dim LB As Long
        LB = LBound(Arr)
        If LB > ub Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If

End Function


