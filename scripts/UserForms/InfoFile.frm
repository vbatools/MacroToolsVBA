VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InfoFile 
   Caption         =   "File Properties:"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13455
   OleObjectBlob   =   "InfoFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InfoFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : InfoFile - управление свойствами файла
'* Created    : 20-07-2020 15:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

    Private Sub cmbMain_Change()
11:    On Error Resume Next
12:    Call UpdeteList(Me.ListCode, X_InfoFile.ShowProp(Workbooks(cmbMain.Value)))
13:    Call UpdeteList(Me.ListCustomProp, X_InfoFile.ShowCustomProp(Workbooks(cmbMain.Value)))
14:    On Error GoTo 0
15: End Sub
    Private Sub UpdeteList(ByRef objList As MSForms.ListBox, ByVal Txt As String)
17:    Dim Arr         As Variant
18:    Dim i           As Byte
19:    objList.Clear
20:    If Txt <> vbNullString Then
21:        Arr = VBA.Split(Txt, vbNewLine)
22:        With objList
23:            For i = 0 To UBound(Arr)
24:                If Arr(i) <> vbNullString Then
25:                    .AddItem i + 1
26:                    .List(i, 1) = VBA.Split(Arr(i), ": ")(0)
27:                    .List(i, 2) = VBA.Split(Arr(i), ": ")(1)
28:                End If
29:            Next i
30:        End With
31:    End If
32: End Sub

    Private Sub Label2_Click()
35:    Me.Hide
36:    Call InfoFile2.Show
37:    Call cmbMain_Change
38:    Me.Show
39: End Sub

    Private Sub LbDelAllProper_Click()
42:    If MsgBox("Delete ALL properties ?", vbYesNo + vbQuestion, "Deleting properties:") = vbYes Then
43:        Dim iCount  As Byte
44:        iCount = X_InfoFile.DelAllProp(Workbooks(cmbMain.Value))
45:        Call cmbMain_Change
46:        Call MsgBox("Deleted properties:" & iCount, vbInformation, "Deleting properties:")
47:    End If
48: End Sub
    Private Sub LbEdit_Click()
50:    Call EditProp
51: End Sub

    Private Sub lbTemplete_Click()
54:    Dim tbData As Variant
55:    Dim i As Integer
56:    tbData = ThisWorkbook.Worksheets(C_Const.SH_SNIPPETS).ListObjects("TB_TEMPLETE").DataBodyRange.Value2
57:    tbData = ThisWorkbook.Worksheets(C_Const.SH_SNIPPETS).ListObjects("TB_TEMPLETE").DataBodyRange.Value2
58:    For i = 1 To UBound(tbData)
59:        Call X_InfoFile.AddOneCustomProp(Workbooks(cmbMain.Value), tbData(i, 1), tbData(i, 2))
60:    Next i
61:    Call cmbMain_Change
62: End Sub

    Private Sub ListCode_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
65:    Call EditProp
66: End Sub
    Private Sub EditProp()
68:    Dim txtNew      As String
69:    Dim txtOld      As String
70:    Dim NameProp    As String
71:    With Me.ListCode
72:        If IsNumeric(.BoundValue) Then
73:            txtOld = VBA.Trim$(.List(CInt(.BoundValue) - 1, 2))
74:            NameProp = .List(CInt(.BoundValue) - 1, 1)
75:            txtNew = InputBox("Edit a property [" & NameProp & " ] ?", "Edit a property:", txtOld)
76:            If txtNew <> txtOld Then
77:                Call X_InfoFile.WriteOneProp(Workbooks(cmbMain.Value), NameProp, txtNew)
78:                Call cmbMain_Change
79:            End If
80:        End If
81:    End With
82: End Sub

    Private Sub lbAddCustProp_Click()
85:    Call AddCustProp(vbNullString, vbNullString)
86: End Sub

     Private Sub lbEditCustProp_Click()
89:
90:    Dim txtOld      As String
91:    Dim NameProp    As String
92:    With Me.ListCustomProp
93:        If IsNumeric(.BoundValue) Then
94:            txtOld = VBA.Trim$(.List(CInt(.BoundValue) - 1, 2))
95:            NameProp = .List(CInt(.BoundValue) - 1, 1)
96:            Call X_InfoFile.DelOneCustomProp(Workbooks(cmbMain.Value), NameProp)
97:            Call AddCustProp(NameProp, txtOld)
98:        End If
99:    End With
100: End Sub
     Private Sub lbDelOneCustProp_Click()
102:    Dim NameProp    As String
103:    With Me.ListCustomProp
104:        If IsNumeric(.BoundValue) Then
105:            NameProp = .List(CInt(.BoundValue) - 1, 1)
106:            If MsgBox("Delete a property [" & NameProp & " ] ?", vbYesNo + vbQuestion, "Deleting a property:") = vbYes Then
107:                Call X_InfoFile.DelOneCustomProp(Workbooks(cmbMain.Value), NameProp)
108:                Call cmbMain_Change
109:            End If
110:        End If
111:    End With
112: End Sub
     Private Sub AddCustProp(ByVal txtPropName As String, ByVal txtPropValue As String)
114:    txtPropName = InputBox("Ведите  название свойства", "Creating a property:", txtPropName)
115:    If txtPropName <> vbNullString Then
116:        txtPropValue = InputBox("Ведите  значение свойства", "Creating a property:", txtPropValue)
117:        If txtPropValue <> vbNullString Then
118:            Call X_InfoFile.AddOneCustomProp(Workbooks(cmbMain.Value), txtPropName, txtPropValue)
119:            Call cmbMain_Change
120:        End If
121:    End If
122: End Sub


     Private Sub lbDelAllCustomProp_Click()
126:    If MsgBox("Delete ALL properties ?", vbYesNo + vbQuestion, "Deleting properties:") = vbYes Then
127:        Dim iCount  As Byte
128:        iCount = X_InfoFile.DelAllCustomProp(Workbooks(cmbMain.Value))
129:        Call cmbMain_Change
130:        Call MsgBox("Deleted properties:" & iCount, vbInformation, "Deleting properties:")
131:    End If
132: End Sub

     Private Sub UserForm_Activate()
135:    Dim vbProj      As VBIDE.VBProject
136:    If Workbooks.Count = 0 Then
137:        Unload Me
138:        Call MsgBox("No open ones" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
139:        Exit Sub
140:    End If
141:    With Me.cmbMain
142:        .Clear
143:        On Error Resume Next
144:        For Each vbProj In Application.VBE.VBProjects
145:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
146:        Next
147:        On Error GoTo 0
148:        .Value = ActiveWorkbook.Name
149:    End With
150: End Sub

     Private Sub cmbCancel_Click()
153:    Unload Me
154: End Sub
     Private Sub lbCancel_Click()
156:    Call cmbCancel_Click
157: End Sub
Private Sub UserForm_Initialize()
159:    Me.StartUpPosition = 0
160:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
161:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
