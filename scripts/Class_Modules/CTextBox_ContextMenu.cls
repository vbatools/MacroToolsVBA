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
14:    '
15:    ' Build Context menu controls.
16:    '
17:    Dim cbrTemp As CommandBar
18:    Const CUT_MENUID = 9724
19:    Const COPY_MENUID = 462
20:    Const PASTE_MENUID = 1552
21:    Set cbrTemp = Application.CommandBars.Add(mEDIT_CONTEXTMENU_NAME, Position:=msoBarPopup)
22:    With cbrTemp
23:        With .Controls.Add(msoControlButton)
24:            .Caption = "DELETE"
25:            .FaceId = CUT_MENUID
26:            .Tag = mCUT_TAG
27:        End With
28:        With .Controls.Add(msoControlButton)
29:            .Caption = "TO CREATE"
30:            .FaceId = COPY_MENUID
31:            .Tag = mCOPY_TAG
32:        End With
33:        With .Controls.Add(msoControlButton)
34:            .Caption = "TO CHANGE"
35:            .FaceId = PASTE_MENUID
36:            .Tag = mPASTE_TAG
37:        End With
38:    End With
39:    Set m_CreateEditContextMenu = cbrTemp
40: End Function
    Private Sub m_DestroyEditContextMenu()
42:    On Error Resume Next
43:    Application.CommandBars(mEDIT_CONTEXTMENU_NAME).Delete
44: End Sub
    Private Function m_GetEditContextMenu() As CommandBar
46:    On Error Resume Next
47:    Set m_GetEditContextMenu = Application.CommandBars(mEDIT_CONTEXTMENU_NAME)
48:    If m_GetEditContextMenu Is Nothing Then
49:        Set m_GetEditContextMenu = m_CreateEditContextMenu
50:    End If
51: End Function
    Private Function m_ActiveTextbox() As Boolean
53:    '
54:    ' Make sure this instance is connected to active control
55:    ' May need to drill down through container controls to
56:    ' reach ActiveControl object
57:    '
58:    Dim objCtl As Object
59:    Set objCtl = m_objParent.ActiveControl
60:    Do While UCase$(TypeName(objCtl)) <> "LISTBOX"
61:        If UCase$(TypeName(objCtl)) = "MULTIPAGE" Then
62:            Set objCtl = objCtl.Pages(objCtl.Value).ActiveControl
63:        Else
64:            Set objCtl = objCtl.ActiveControl
65:        End If
66:    Loop
67:    m_ActiveTextbox = (StrComp(objCtl.Name, m_txtTBox.Name, vbTextCompare) = 0)
68: End Function
    Public Property Set prParent(ByRef RHS As Object)
70:    Set m_objParent = RHS
71: End Property
    Private Sub m_UseMenu()
73:    Dim lngIndex As Long
74:    For lngIndex = 1 To m_cbrContextMenu.Controls.Count
75:        Select Case m_cbrContextMenu.Controls(lngIndex).Tag
            Case mCUT_TAG
77:                Set m_cbtCut = m_cbrContextMenu.Controls(lngIndex)
78:            Case mCOPY_TAG
79:                Set m_cbtCopy = m_cbrContextMenu.Controls(lngIndex)
80:            Case mPASTE_TAG
81:                Set m_cbtPaste = m_cbrContextMenu.Controls(lngIndex)
82:        End Select
83:    Next
84: End Sub
    Public Property Set TBox(ByRef RHS As MSForms.ListBox)
86:    Set m_txtTBox = RHS
87: End Property
    Private Sub Class_Initialize()
89:    Set m_objDataObject = New DataObject
90:    Set m_cbrContextMenu = m_GetEditContextMenu
91:    If Not m_cbrContextMenu Is Nothing Then
92:        m_UseMenu
93:    End If
94: End Sub
    Private Sub Class_Terminate()
96:    Set m_objDataObject = Nothing
97:    m_DestroyEditContextMenu
98: End Sub
     Private Sub m_cbtCopy_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
100:    ' check active textbox is this instance of CTextBox_ContextMenu
101:    If m_ActiveTextbox() Then
102:        With m_objDataObject
103:            Call G_AddCodeViewForm.AddCode(m_txtTBox.ListCount)
104:        End With
105:    End If
106: End Sub
     Private Sub m_cbtCut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
108:    ' check active textbox is this instance of CTextBox_ContextMenu
109:    If m_ActiveTextbox() Then
110:        Call G_AddCodeViewForm.DeletRow(SelectedRow, m_txtTBox)
111:    End If
112: End Sub
     Private Sub m_cbtPaste_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
114:    ' check active textbox is this instance of CTextBox_ContextMenu
115:    On Error GoTo ErrPaste
116:    If m_ActiveTextbox() Then
117:        With m_objDataObject
118:            Call G_AddCodeViewForm.EditCode(SelectedRow, m_txtTBox)
119:        End With
120:    End If
ErrPaste:
122:    Exit Sub
123: End Sub
     Private Sub m_txtTBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
125:    If Button = 2 Then
126:        ' right click
127:        m_cbrContextMenu.ShowPopup
128:    End If
129: End Sub
     Private Function SelectedRow() As Long
131:    Dim X      As Long
132:    SelectedRow = 0
133:    For X = 0 To m_txtTBox.ListCount - 1
134:        If m_txtTBox.Selected(X) Then
135:            SelectedRow = X + 1
136:            Exit Function
137:        End If
138:    Next X
139: End Function
