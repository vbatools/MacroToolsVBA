VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTextBox_ContextMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : CTextBox_ContextMenu - создание контекст меню на форме
'* Created    : 15-09-2020 15:55
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Private Const mEDIT_CONTEXTMENU_NAME = "ajpiEditContextMenu"
Private Const mCUT_TAG = "DELETE"
Private Const mCOPY_TAG = "ADD"
Private Const mPASTE_TAG = "EDIT"
Private m_cbrContextMenu As CommandBar
Private WithEvents m_txtTBox As MSForms.ListBox
Attribute m_txtTBox.VB_VarHelpID = -1
Private WithEvents m_cbtCut As CommandBarButton
Attribute m_cbtCut.VB_VarHelpID = -1
Private WithEvents m_cbtCopy As CommandBarButton
Attribute m_cbtCopy.VB_VarHelpID = -1
Private WithEvents m_cbtPaste As CommandBarButton
Attribute m_cbtPaste.VB_VarHelpID = -1
Private m_objDataObject As DataObject
Private m_objParent As Object
    Private Function m_CreateEditContextMenu() As CommandBar
22:    '
23:    ' Build Context menu controls.
24:    '
25:    Dim cbrTemp As CommandBar
26:    Const CUT_MENUID = 9724
27:    Const COPY_MENUID = 462
28:    Const PASTE_MENUID = 1552
29:    Set cbrTemp = Application.CommandBars.Add(mEDIT_CONTEXTMENU_NAME, Position:=msoBarPopup)
30:    With cbrTemp
31:        With .Controls.Add(msoControlButton)
32:            .Caption = "DELETE"
33:            .FaceId = CUT_MENUID
34:            .Tag = mCUT_TAG
35:        End With
36:        With .Controls.Add(msoControlButton)
37:            .Caption = "TO CREATE"
38:            .FaceId = COPY_MENUID
39:            .Tag = mCOPY_TAG
40:        End With
41:        With .Controls.Add(msoControlButton)
42:            .Caption = "TO CHANGE"
43:            .FaceId = PASTE_MENUID
44:            .Tag = mPASTE_TAG
45:        End With
46:    End With
47:    Set m_CreateEditContextMenu = cbrTemp
48: End Function
    Private Sub m_DestroyEditContextMenu()
50:    On Error Resume Next
51:    Application.CommandBars(mEDIT_CONTEXTMENU_NAME).Delete
52: End Sub
    Private Function m_GetEditContextMenu() As CommandBar
54:    On Error Resume Next
55:    Set m_GetEditContextMenu = Application.CommandBars(mEDIT_CONTEXTMENU_NAME)
56:    If m_GetEditContextMenu Is Nothing Then
57:        Set m_GetEditContextMenu = m_CreateEditContextMenu
58:    End If
59: End Function
    Private Function m_ActiveTextbox() As Boolean
61:    '
62:    ' Make sure this instance is connected to active control
63:    ' May need to drill down through container controls to
64:    ' reach ActiveControl object
65:    '
66:    Dim objCtl As Object
67:    Set objCtl = m_objParent.ActiveControl
68:    Do While UCase$(TypeName(objCtl)) <> "LISTBOX"
69:        If UCase$(TypeName(objCtl)) = "MULTIPAGE" Then
70:            Set objCtl = objCtl.Pages(objCtl.Value).ActiveControl
71:        Else
72:            Set objCtl = objCtl.ActiveControl
73:        End If
74:    Loop
75:    m_ActiveTextbox = (StrComp(objCtl.Name, m_txtTBox.Name, vbTextCompare) = 0)
76: End Function
    Public Property Set prParent(ByRef RHS As Object)
78:    Set m_objParent = RHS
79: End Property
    Private Sub m_UseMenu()
81:    Dim lngIndex As Long
82:    For lngIndex = 1 To m_cbrContextMenu.Controls.Count
83:        Select Case m_cbrContextMenu.Controls(lngIndex).Tag
            Case mCUT_TAG
85:                Set m_cbtCut = m_cbrContextMenu.Controls(lngIndex)
86:            Case mCOPY_TAG
87:                Set m_cbtCopy = m_cbrContextMenu.Controls(lngIndex)
88:            Case mPASTE_TAG
89:                Set m_cbtPaste = m_cbrContextMenu.Controls(lngIndex)
90:        End Select
91:    Next
92: End Sub
    Public Property Set TBox(ByRef RHS As MSForms.ListBox)
94:    Set m_txtTBox = RHS
95: End Property
     Private Sub Class_Initialize()
97:    Set m_objDataObject = New DataObject
98:    Set m_cbrContextMenu = m_GetEditContextMenu
99:    If Not m_cbrContextMenu Is Nothing Then
100:        m_UseMenu
101:    End If
102: End Sub
     Private Sub Class_Terminate()
104:    Set m_objDataObject = Nothing
105:    m_DestroyEditContextMenu
106: End Sub
     Private Sub m_cbtCopy_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
108:    ' check active textbox is this instance of CTextBox_ContextMenu
109:    If m_ActiveTextbox() Then
110:        With m_objDataObject
111:            Call G_AddCodeViewForm.AddCode(m_txtTBox.ListCount)
112:        End With
113:    End If
114: End Sub
     Private Sub m_cbtCut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
116:    ' check active textbox is this instance of CTextBox_ContextMenu
117:    If m_ActiveTextbox() Then
118:        Call G_AddCodeViewForm.DeletRow(SelectedRow, m_txtTBox)
119:    End If
120: End Sub
     Private Sub m_cbtPaste_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
122:    ' check active textbox is this instance of CTextBox_ContextMenu
123:    On Error GoTo ErrPaste
124:    If m_ActiveTextbox() Then
125:        With m_objDataObject
126:            Call G_AddCodeViewForm.EditCode(SelectedRow, m_txtTBox)
127:        End With
128:    End If
ErrPaste:
130:    Exit Sub
131: End Sub
     Private Sub m_txtTBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
133:    If Button = 2 Then
134:        ' right click
135:        m_cbrContextMenu.ShowPopup
136:    End If
137: End Sub
Private Function SelectedRow() As Long
139:    Dim X      As Long
140:    SelectedRow = 0
141:    For X = 0 To m_txtTBox.ListCount - 1
142:        If m_txtTBox.Selected(X) Then
143:            SelectedRow = X + 1
144:            Exit Function
145:        End If
146:    Next X
End Function
