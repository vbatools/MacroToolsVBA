VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModuleTODO 
   Caption         =   "VBA Project Manager:"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13575
   OleObjectBlob   =   "ModuleTODO.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModuleTODO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : ModuleTODO - модуль поиска меток TODO
'* Created    : 01-20-2020 12:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Private m_clsAnchors As CAnchors

    Private Sub UserForm_Initialize()
12:    Set m_clsAnchors = New CAnchors
13:    Set m_clsAnchors.objParent = Me
14:    ' restrict minimum size of userform
15:    m_clsAnchors.MinimumWidth = 683.25
16:    m_clsAnchors.MinimumHeight = 405.75
17:    With m_clsAnchors
18:        .funAnchor("cmbMain").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft
19:        .funAnchor("ListCode").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop Or enumAnchorStyleLeft Or enumAnchorStyleBottom
20:        .funAnchor("lbCancel").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
21:    End With
22: End Sub
    Private Sub UserForm_Activate()
24:    Dim vbProj      As VBIDE.VBProject
25:    If Workbooks.Count = 0 Then
26:        Unload Me
27:        Call MsgBox("No open ones" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
28:        Exit Sub
29:    End If
30:    With Me.cmbMain
31:        .Clear
32:        On Error Resume Next
33:        For Each vbProj In Application.VBE.VBProjects
34:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
35:        Next
36:        On Error GoTo 0
37:        .Value = ActiveWorkbook.Name
38:        Call AddTODOList(.Value)
39:    End With
40: End Sub
    Private Sub UserForm_Terminate()
42:    Set m_clsAnchors = Nothing
43: End Sub

    Private Sub cmbCancel_Click()
46:    Unload Me
47: End Sub
    Private Sub lbCancel_Click()
49:    Call cmbCancel_Click
50: End Sub
    Private Sub cmbMain_Change()
52:    If cmbMain.Value = vbNullString Then Exit Sub
53:    Call AddTODOList(cmbMain.Value)
54: End Sub
    Private Sub ListCode_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
56:    Dim i           As Long
57:    Dim WB          As Workbook
58:    Dim VBC         As VBIDE.VBComponent
59:
60:    On Error GoTo ErrorHandler
61:
62:    If cmbMain.Value = vbNullString Then Exit Sub
63:    Set WB = Workbooks(cmbMain.Value)
64:    For i = 0 To ListCode.ListCount
65:        If ListCode.Selected(i) = True Then
66:            Set VBC = WB.VBProject.VBComponents(ListCode.List(i, 2))
67:            If VBC.Type = vbext_ct_MSForm Then
68:                VBC.CodeModule.CodePane.Show
69:            Else
70:                VBC.Activate
71:            End If
72:            Exit Sub
73:        End If
74:    Next i
75:    Exit Sub
ErrorHandler:
77:    Unload Me
78:    Select Case Err.Number
        Case Else:
80:            Call MsgBox("Error in Module TODO.List Code_DblClick" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in a row" & Erl, vbOKOnly + vbExclamation, "Error:")
81:            Call WriteErrorLog("ModuleTODO.ListCode_DblClick")
82:    End Select
83:    Err.Clear
84: End Sub

     Private Sub AddTODOList(sWb As String)    'As String
87:    Dim iFile       As Integer
88:    Dim WB          As Workbook
89:    On Error GoTo ErrorHandler
90:    Set WB = Workbooks(sWb)
91:    If WB.VBProject.Protection = vbext_pp_none Then
92:        ListCode.Clear
93:        For iFile = 1 To WB.VBProject.VBComponents.Count
94:            Call listLinesinModuleWhereFound(WB.VBProject.VBComponents(iFile), "'* TODO Created:")
95:        Next iFile
96:    Else
97:        ListCode.Clear
98:        Call MsgBox("VBA project in the book -" & WB.Name & "password protected!" & vbCrLf & "Remove the password!", vbCritical, "Error:")
99:    End If
100:    Exit Sub
ErrorHandler:
102:    Select Case Err.Number
        Case 4160:
104:            ListCode.Clear
105:            Call MsgBox("Error No access to the VBA project!", vbOKOnly + vbExclamation, "Error:")
106:        Case Else:
107:            Call MsgBox("Error in AddTODOList.UserForm_Activate" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in a row" & Erl, vbOKOnly + vbExclamation, "Error:")
108:            Call WriteErrorLog("AddTODOList.UserForm_Activate")
109:    End Select
110:    Err.Clear
111: End Sub

Sub listLinesinModuleWhereFound(ByVal oComponent As Object, ByVal sSearchTerm As String)
114:    Dim lTotalNoLines As Long
115:    Dim lLineNo     As Long
116:    Dim lListRow    As Long
117:
118:    On Error GoTo ErrorHandler
119:
120:    lLineNo = 1
121:    lListRow = ListCode.ListCount
122:    With oComponent
123:        lTotalNoLines = .CodeModule.CountOfLines
124:        Do While .CodeModule.Find(sSearchTerm, lLineNo, 1, -1, -1, False, False, False) = True
125:            ListCode.AddItem lListRow + 1
126:            ListCode.List(lListRow, 1) = ComponentTypeToString(.Type)
127:            ListCode.List(lListRow, 2) = .Name
128:            ListCode.List(lListRow, 3) = vbTab & "Line No:" & lLineNo
129:            ListCode.List(lListRow, 4) = Replace(Trim$(.CodeModule.Lines(lLineNo, 1)), "'*", vbNullString)
130:            ListCode.List(lListRow, 5) = Replace(Trim$(.CodeModule.Lines(lLineNo + 1, 1)), "'*", vbNullString)
131:            lLineNo = lLineNo + 1
132:            lListRow = lListRow + 1
133:        Loop
134:    End With
135:    Exit Sub
ErrorHandler:
137:    Unload Me
138:    Select Case Err.Number
        Case Else:
140:            Call MsgBox("Error in ModuleTODO.listLinesinModuleWhereFound" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in a row" & Erl, vbOKOnly + vbExclamation, "Error:")
141:            Call WriteErrorLog("ModuleTODO.listLinesinModuleWhereFound")
142:    End Select
143:    Err.Clear
End Sub
