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
23:    Dim clsTemp As CAnchor
24:    Set clsTemp = New CAnchor
25:    Set clsTemp.cnt = RHS
26:    With clsTemp
27:        .AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleTop
28:        .MinimumWidth = .cnt.Width
29:        .MinimumHeight = .cnt.Height
30:        .OrigLeft = .cnt.Left
31:        .OrigTop = .cnt.top
32:        .OrigWidth = .cnt.Width
33:        .OrigHeight = .cnt.Height
34:    End With
35:    m_colAnchors.Add clsTemp, clsTemp.cnt.Name
36: End Property
    Public Sub RemoveCntrl(ByRef varIndex As Variant)
38:    On Error Resume Next
39:    m_colAnchors.Remove varIndex
40: End Sub
    Private Sub m_AddResizer(ByRef objParent As Object)
42:    '
43:    ' add resizing control to bottom righthand corner of userform
44:    '
45:    Set ResizeHandle = objParent.Controls.Add("Forms.label.1", MRESIZEHANDLE, True)
46:    With ResizeHandle
47:        With .Font
48:            .Name = "Marlett"
49:            .Charset = 2
50:            .Size = 14
51:            .Bold = True
52:        End With
53:        .BackStyle = fmBackStyleTransparent
54:        .AutoSize = True
55:        .BorderStyle = fmBorderStyleNone
56:        .Caption = "o"
57:        .MousePointer = fmMousePointerSizeNWSE
58:        .ForeColor = &H8000000D
59:        .ZOrder
60:        .top = objParent.InsideHeight - .Height
61:        .Left = objParent.InsideWidth - .Width
62:    End With
63: End Sub
    Private Sub ResizeHandle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
65:    If Button = 1 Then
66:        m_sngLeftResizePos = X
67:        m_sngTopResizePos = Y
68:        m_blnResizing = True
69:    End If
70: End Sub
    Private Sub ResizeHandle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
72:    Dim sngSize As Single
73:    If Button = 1 Then
74:        With ResizeHandle
75:            .MOVE .Left + X - m_sngLeftResizePos, .top + Y - m_sngTopResizePos
76:            sngSize = m_frmParent.Width + X - m_sngLeftResizePos
77:            If sngSize < Me.MinimumWidth Then sngSize = MinimumWidth
78:            m_frmParent.Width = sngSize
79:            sngSize = m_frmParent.Height + Y - m_sngTopResizePos
80:            If sngSize < MinimumHeight Then sngSize = MinimumHeight
81:            m_frmParent.Height = sngSize
82:            .Left = m_frmParent.InsideWidth - .Width
83:            .top = m_frmParent.InsideHeight - .Height
84:            If UpdateWhilstDragging Then
85:                m_UpdateControls
86:            End If
87:        End With
88:    End If
89: End Sub
    Private Sub ResizeHandle_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
91:    If Button = 1 Then
92:        If Not UpdateWhilstDragging Then
93:            m_UpdateControls
94:        End If
95:        m_blnResizing = False
96:    End If
97: End Sub
     Public Function funAnchor(ByRef varIndex As Variant) As CAnchor
99:    ' access to specific anchored control
100:    On Error Resume Next
101:    Set funAnchor = m_colAnchors(varIndex)
102: End Function
     Public Function Anchors() As Collection
104:    ' access to the collection of anchored controls
105:    Set Anchors = m_colAnchors
106: End Function
     Public Property Set objParent(ByRef RHS As Object)
108:    '
109:    ' Use this to assign all default properties
110:    '
111:    Dim clsTemp As CAnchor
112:    Dim cntTemp As MSForms.control
113:    Set m_frmParent = RHS
114:    UpdateWhilstDragging = True    '!!
115:    With RHS
116:        MinimumWidth = .Width
117:        MinimumHeight = .Height
118:        OrigLeft = 1
119:        OrigTop = 1
120:        OrigWidth = .InsideWidth
121:        OrigHeight = .InsideHeight
122:    End With
123:    For Each cntTemp In m_frmParent.Controls
124:        Set clsTemp = New CAnchor
125:        Set clsTemp.cnt = cntTemp
126:        With clsTemp
127:            .AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleTop
128:            .MinimumWidth = cntTemp.Width
129:            .MinimumHeight = cntTemp.Height
130:            .OrigLeft = cntTemp.Left
131:            .OrigTop = cntTemp.top
132:            .OrigWidth = cntTemp.Width
133:            .OrigHeight = cntTemp.Height
134:        End With
135:        m_colAnchors.Add clsTemp, clsTemp.cnt.Name
136:    Next
137:    m_AddResizer RHS
138: End Property
     Private Sub Class_Initialize()
140:    Set m_colAnchors = New Collection
141: End Sub
     Private Sub Class_Terminate()
143:    Do While m_colAnchors.Count > 0
144:        m_colAnchors.Remove m_colAnchors.Count
145:    Loop
146:    Set m_colAnchors = Nothing
147:    m_frmParent.Controls.Remove MRESIZEHANDLE
148:    Set ResizeHandle = Nothing
149: End Sub
     Private Sub m_UpdateControls()
151:    '
152:    ' Calculate New position of all controls
153:    '
154:    Dim clsAnchor As CAnchor
155:    Dim cntTemp As MSForms.control
156:    Dim sngLeft As Single
157:    Dim sngTop As Single
158:    Dim sngHeight As Single
159:    Dim sngWidth As Single
160:    For Each clsAnchor In m_colAnchors
161:        Set cntTemp = clsAnchor.cnt
162:        If clsAnchor.AnchorStyle = enumAnchorStyleNone Then
163:            ' do nothing with this control
164:        Else
165:            If ((clsAnchor.AnchorStyle And enumAnchorStyleTop) = enumAnchorStyleTop) And _
                        ((clsAnchor.AnchorStyle And enumAnchorStyleBottom) = enumAnchorStyleBottom) Then
167:                ' maintain gap between top and bottom edges by adjusting height
168:                sngHeight = m_frmParent.InsideHeight - (OrigHeight - clsAnchor.OrigTop - clsAnchor.OrigHeight) - clsAnchor.OrigTop
169:                If sngHeight < clsAnchor.MinimumHeight Then sngHeight = clsAnchor.MinimumHeight
170:                If sngHeight < 0 Then sngHeight = 0
171:                cntTemp.Height = sngHeight
172:            ElseIf (clsAnchor.AnchorStyle And enumAnchorStyleTop) = enumAnchorStyleTop Then
173:                ' maintain gap between top leave height alone
174:                ' does not require code
175:            ElseIf (clsAnchor.AnchorStyle And enumAnchorStyleBottom) = enumAnchorStyleBottom Then
176:                ' maintain gap between bottom leave height alone
177:                sngTop = m_frmParent.InsideHeight - (OrigHeight - clsAnchor.OrigTop - clsAnchor.OrigHeight) - clsAnchor.OrigHeight
178:                If sngTop < clsAnchor.MinimumTop Then sngTop = clsAnchor.MinimumTop
179:                If sngTop < 0 Then sngTop = 0
180:                cntTemp.top = sngTop
181:            End If
182:            If ((clsAnchor.AnchorStyle And enumAnchorStyleLeft) = enumAnchorStyleLeft) And _
                        ((clsAnchor.AnchorStyle And enumAnchorStyleRight) = enumAnchorStyleRight) Then
184:                ' maintain gap between left and right edges by adjusting Width
185:                sngWidth = m_frmParent.InsideWidth - (OrigWidth - clsAnchor.OrigLeft - clsAnchor.OrigWidth) - clsAnchor.OrigLeft
186:                If sngWidth < clsAnchor.MinimumWidth Then sngWidth = clsAnchor.MinimumWidth
187:                If sngWidth < 0 Then sngWidth = 0
188:                cntTemp.Width = sngWidth
189:            ElseIf (clsAnchor.AnchorStyle And enumAnchorStyleLeft) = enumAnchorStyleLeft Then
190:                ' maintain gap between left leave Width alone
191:                ' does not require code
192:            ElseIf (clsAnchor.AnchorStyle And enumAnchorStyleRight) = enumAnchorStyleRight Then
193:                ' maintain gap between Right leave Width alone
194:                sngLeft = m_frmParent.InsideWidth - (OrigWidth - clsAnchor.OrigLeft - clsAnchor.OrigWidth) - clsAnchor.OrigWidth
195:                If sngLeft < clsAnchor.MinimumLeft Then sngLeft = clsAnchor.MinimumLeft
196:                If sngLeft < 0 Then sngLeft = 0
197:                cntTemp.Left = sngLeft
198:            End If
199:        End If
200:    Next
201:    DoEvents
202: End Sub
