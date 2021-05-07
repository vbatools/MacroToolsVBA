VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAnchors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : CAnchors - модуль изменения размеров форм
'* Created    : 15-09-2019 15:53
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Private m_colAnchors As Collection
Private m_frmParent As Object
Public MinimumWidth As Single
Public MinimumHeight As Single
Public OrigLeft As Single
Public OrigTop As Single
Public OrigWidth As Single
Public OrigHeight As Single
Public UpdateWhilstDragging As Boolean
Private Const MRESIZEHANDLE = "ResizeGrabHandle"
Private m_sngLeftResizePos As Single
Private m_sngTopResizePos As Single
Private m_blnResizing As Boolean
Public WithEvents ResizeHandle As MSForms.Label
Attribute ResizeHandle.VB_VarHelpID = -1
    Public Property Set AddCntrl(ByRef RHS As MSForms.control)
25:    Dim clsTemp As CAnchor
26:    Set clsTemp = New CAnchor
27:    Set clsTemp.cnt = RHS
28:    With clsTemp
29:        .AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleTop
30:        .MinimumWidth = .cnt.Width
31:        .MinimumHeight = .cnt.Height
32:        .OrigLeft = .cnt.Left
33:        .OrigTop = .cnt.top
34:        .OrigWidth = .cnt.Width
35:        .OrigHeight = .cnt.Height
36:    End With
37:    m_colAnchors.Add clsTemp, clsTemp.cnt.Name
38: End Property
    Public Sub RemoveCntrl(ByRef varIndex As Variant)
40:    On Error Resume Next
41:    m_colAnchors.Remove varIndex
42: End Sub
    Private Sub m_AddResizer(ByRef objParent As Object)
44:    '
45:    ' add resizing control to bottom righthand corner of userform
46:    '
47:    Set ResizeHandle = objParent.Controls.Add("Forms.label.1", MRESIZEHANDLE, True)
48:    With ResizeHandle
49:        With .Font
50:            .Name = "Marlett"
51:            .Charset = 2
52:            .Size = 14
53:            .Bold = True
54:        End With
55:        .BackStyle = fmBackStyleTransparent
56:        .AutoSize = True
57:        .BorderStyle = fmBorderStyleNone
58:        .Caption = "o"
59:        .MousePointer = fmMousePointerSizeNWSE
60:        .ForeColor = &H8000000D
61:        .ZOrder
62:        .top = objParent.InsideHeight - .Height
63:        .Left = objParent.InsideWidth - .Width
64:    End With
65: End Sub
    Private Sub ResizeHandle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
67:    If Button = 1 Then
68:        m_sngLeftResizePos = X
69:        m_sngTopResizePos = Y
70:        m_blnResizing = True
71:    End If
72: End Sub
    Private Sub ResizeHandle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
74:    Dim sngSize As Single
75:    If Button = 1 Then
76:        With ResizeHandle
77:            .MOVE .Left + X - m_sngLeftResizePos, .top + Y - m_sngTopResizePos
78:            sngSize = m_frmParent.Width + X - m_sngLeftResizePos
79:            If sngSize < Me.MinimumWidth Then sngSize = MinimumWidth
80:            m_frmParent.Width = sngSize
81:            sngSize = m_frmParent.Height + Y - m_sngTopResizePos
82:            If sngSize < MinimumHeight Then sngSize = MinimumHeight
83:            m_frmParent.Height = sngSize
84:            .Left = m_frmParent.InsideWidth - .Width
85:            .top = m_frmParent.InsideHeight - .Height
86:            If UpdateWhilstDragging Then
87:                m_UpdateControls
88:            End If
89:        End With
90:    End If
91: End Sub
    Private Sub ResizeHandle_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
93:    If Button = 1 Then
94:        If Not UpdateWhilstDragging Then
95:            m_UpdateControls
96:        End If
97:        m_blnResizing = False
98:    End If
99: End Sub
     Public Function funAnchor(ByRef varIndex As Variant) As CAnchor
101:    ' access to specific anchored control
102:    On Error Resume Next
103:    Set funAnchor = m_colAnchors(varIndex)
104: End Function
     Public Function Anchors() As Collection
106:    ' access to the collection of anchored controls
107:    Set Anchors = m_colAnchors
108: End Function
     Public Property Set objParent(ByRef RHS As Object)
110:    '
111:    ' Use this to assign all default properties
112:    '
113:    Dim clsTemp As CAnchor
114:    Dim cntTemp As MSForms.control
115:    Set m_frmParent = RHS
116:    UpdateWhilstDragging = True    '!!
117:    With RHS
118:        MinimumWidth = .Width
119:        MinimumHeight = .Height
120:        OrigLeft = 1
121:        OrigTop = 1
122:        OrigWidth = .InsideWidth
123:        OrigHeight = .InsideHeight
124:    End With
125:    For Each cntTemp In m_frmParent.Controls
126:        Set clsTemp = New CAnchor
127:        Set clsTemp.cnt = cntTemp
128:        With clsTemp
129:            .AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleTop
130:            .MinimumWidth = cntTemp.Width
131:            .MinimumHeight = cntTemp.Height
132:            .OrigLeft = cntTemp.Left
133:            .OrigTop = cntTemp.top
134:            .OrigWidth = cntTemp.Width
135:            .OrigHeight = cntTemp.Height
136:        End With
137:        m_colAnchors.Add clsTemp, clsTemp.cnt.Name
138:    Next
139:    m_AddResizer RHS
140: End Property
     Private Sub Class_Initialize()
142:    Set m_colAnchors = New Collection
143: End Sub
     Private Sub Class_Terminate()
145:    Do While m_colAnchors.Count > 0
146:        m_colAnchors.Remove m_colAnchors.Count
147:    Loop
148:    Set m_colAnchors = Nothing
149:    m_frmParent.Controls.Remove MRESIZEHANDLE
150:    Set ResizeHandle = Nothing
151: End Sub
Private Sub m_UpdateControls()
153:    '
154:    ' Calculate New position of all controls
155:    '
156:    Dim clsAnchor As CAnchor
157:    Dim cntTemp As MSForms.control
158:    Dim sngLeft As Single
159:    Dim sngTop As Single
160:    Dim sngHeight As Single
161:    Dim sngWidth As Single
162:    For Each clsAnchor In m_colAnchors
163:        Set cntTemp = clsAnchor.cnt
164:        If clsAnchor.AnchorStyle = enumAnchorStyleNone Then
165:            ' do nothing with this control
166:        Else
167:            If ((clsAnchor.AnchorStyle And enumAnchorStyleTop) = enumAnchorStyleTop) And _
                            ((clsAnchor.AnchorStyle And enumAnchorStyleBottom) = enumAnchorStyleBottom) Then
169:                ' maintain gap between top and bottom edges by adjusting height
170:                sngHeight = m_frmParent.InsideHeight - (OrigHeight - clsAnchor.OrigTop - clsAnchor.OrigHeight) - clsAnchor.OrigTop
171:                If sngHeight < clsAnchor.MinimumHeight Then sngHeight = clsAnchor.MinimumHeight
172:                If sngHeight < 0 Then sngHeight = 0
173:                cntTemp.Height = sngHeight
174:            ElseIf (clsAnchor.AnchorStyle And enumAnchorStyleTop) = enumAnchorStyleTop Then
175:                ' maintain gap between top leave height alone
176:                ' does not require code
177:            ElseIf (clsAnchor.AnchorStyle And enumAnchorStyleBottom) = enumAnchorStyleBottom Then
178:                ' maintain gap between bottom leave height alone
179:                sngTop = m_frmParent.InsideHeight - (OrigHeight - clsAnchor.OrigTop - clsAnchor.OrigHeight) - clsAnchor.OrigHeight
180:                If sngTop < clsAnchor.MinimumTop Then sngTop = clsAnchor.MinimumTop
181:                If sngTop < 0 Then sngTop = 0
182:                cntTemp.top = sngTop
183:            End If
184:            If ((clsAnchor.AnchorStyle And enumAnchorStyleLeft) = enumAnchorStyleLeft) And _
                            ((clsAnchor.AnchorStyle And enumAnchorStyleRight) = enumAnchorStyleRight) Then
186:                ' maintain gap between left and right edges by adjusting Width
187:                sngWidth = m_frmParent.InsideWidth - (OrigWidth - clsAnchor.OrigLeft - clsAnchor.OrigWidth) - clsAnchor.OrigLeft
188:                If sngWidth < clsAnchor.MinimumWidth Then sngWidth = clsAnchor.MinimumWidth
189:                If sngWidth < 0 Then sngWidth = 0
190:                cntTemp.Width = sngWidth
191:            ElseIf (clsAnchor.AnchorStyle And enumAnchorStyleLeft) = enumAnchorStyleLeft Then
192:                ' maintain gap between left leave Width alone
193:                ' does not require code
194:            ElseIf (clsAnchor.AnchorStyle And enumAnchorStyleRight) = enumAnchorStyleRight Then
195:                ' maintain gap between Right leave Width alone
196:                sngLeft = m_frmParent.InsideWidth - (OrigWidth - clsAnchor.OrigLeft - clsAnchor.OrigWidth) - clsAnchor.OrigWidth
197:                If sngLeft < clsAnchor.MinimumLeft Then sngLeft = clsAnchor.MinimumLeft
198:                If sngLeft < 0 Then sngLeft = 0
199:                cntTemp.Left = sngLeft
200:            End If
201:        End If
202:    Next
203:    DoEvents
End Sub
