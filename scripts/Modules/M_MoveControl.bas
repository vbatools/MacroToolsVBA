Attribute VB_Name = "M_MoveControl"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : M_MoveControl - Микро подстройка элементов формы VBA и переименование элементов на форме вместе с кодом
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module
Public sTagNameConrol  As String
Public tpStyle      As ProperControlStyle
Type ProperControlStyle
    sError          As String

    snHeight        As Single
    snWidth         As Single

    bVisible        As Boolean
    bEnabled        As Boolean
    bLocked         As Boolean

    lBackColor      As Long
    lForeColor      As Long
    lBackStyle      As Long

    lBorderColor    As Long
    lBorderStyle    As Long

    bFontBold       As Boolean
    bFontItalic     As Boolean
    bFontStrikethru As Boolean
    bFontUnderline  As Boolean
    sFontName       As String
    cuFontSize      As Currency
End Type
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : MoveControl - Микроподстройка элементов формы
'* Created    : 08-10-2020 14:10
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub MoveControl()
45:    Dim myCommandBar As CommandBar
46:    Dim cntrl  As CommandBarControl
47:    Dim combox As CommandBarComboBox
48:    Dim sComBoxText As String
49:    Dim cnt    As control
50:
51:    Set myCommandBar = Application.VBE.CommandBars(C_Const.MENUMOVECONTRL)
52:    For Each cntrl In myCommandBar.Controls
53:        If cntrl.Tag = C_Const.MTAGCOM Then
54:            Set combox = myCommandBar.Controls(cntrl.ID)
55:            sComBoxText = combox.Text
56:            Exit For
57:        End If
58:    Next cntrl
59:
60:    Set cnt = TakeSelectControl
61:    If cnt Is Nothing Then Exit Sub
62:    Select Case sTagNameConrol
        Case C_Const.MTAG1:
64:            Call MoveCnt(cnt, 1, sComBoxText)
65:        Case C_Const.MTAG2:
66:            Call MoveCnt(cnt, 2, sComBoxText)
67:        Case C_Const.MTAG3:
68:            Call MoveCnt(cnt, 3, sComBoxText)
69:        Case C_Const.MTAG4:
70:            Call MoveCnt(cnt, 4, sComBoxText)
71:    End Select
72: End Sub
     Private Sub MoveCnt(ByRef cnt As control, ByVal iVal As Integer, ByVal sComBoxText As String)
74:    Const Shag = 0.4
75:    With cnt
76:        Select Case sComBoxText
            Case C_Const.MOVECONT:
78:                Select Case iVal
                    Case 1:
80:                        .Left = .Left - Shag
81:                    Case 2:
82:                        .Left = .Left + Shag
83:                    Case 3:
84:                        .top = .top + Shag
85:                    Case 4:
86:                        .top = .top - Shag
87:                End Select
88:            Case C_Const.MOVECONTTOPLEFT:
89:                Select Case iVal
                    Case 1:
91:                        .Left = .Left - Shag
92:                        .Width = .Width + Shag
93:                    Case 2:
94:                        .Left = .Left + Shag
95:                        .Width = .Width - Shag
96:                    Case 3:
97:                        .top = .top + Shag
98:                        .Height = .Height - Shag
99:                    Case 4:
100:                        .top = .top - Shag
101:                        .Height = .Height + Shag
102:                End Select
103:            Case C_Const.MOVECONTBOTTOMRIGHT:
104:                Select Case iVal
                    Case 1:
106:                        .Width = .Width - Shag
107:                    Case 2:
108:                        .Width = .Width + Shag
109:                    Case 3:
110:                        .Height = .Height + Shag
111:                    Case 4:
112:                        .Height = .Height - Shag
113:                End Select
114:        End Select
115:    End With
116: End Sub

     Private Function TakeSelectControl(Optional bUserForm As Boolean = False) As Object
119:    Dim W           As VBIDE.Window
120:    Dim strVar()    As String
121:    Dim cntName     As String
122:
123:    On Error GoTo ErrorHandler
124:
125:    If Application.VBE.ActiveWindow.Type = vbext_wt_Designer Then
126:        For Each W In Application.VBE.Windows
127:            If W.Type = vbext_wt_PropertyWindow Then
128:                strVar = Split(W.Caption, "-")
129:                cntName = Trim(strVar(1))
130:                Exit For
131:            End If
132:        Next
133:
134:        Dim Form    As UserForm
135:        Set Form = Application.VBE.SelectedVBComponent.Designer
136:        Set TakeSelectControl = Form.Controls(cntName)
137:    End If
138:    Exit Function
ErrorHandler:
140:    If bUserForm And Not Form Is Nothing Then
141:        Err.Clear
142:        Set TakeSelectControl = Form
143:        Exit Function
144:    End If
145:    Select Case Err.Number
        Case -2147024809:
147:            Debug.Print "Error Select one object"
148:        Case 9:
149:            Debug.Print "To use the tool, open the View - > Properties Window"
150:        Case Else:
151:            Debug.Print "Error in TakeSelectControl" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
152:            Call WriteErrorLog("TakeSelectControl")
153:    End Select
154:    Err.Clear
155: End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RenameControl - переименование конторол на форме вместе скодом
'* Created    : 08-10-2020 14:11
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub RenameControl()
164:    Dim cnt    As control
165:    Dim sNewName As String
166:    Dim sOldName As String
167:    Dim NameModeCode As String
168:    Dim strVar As String
169:    Dim CodeMod As CodeModule
170:
171:    On Error GoTo ErrorHandler
172:
173:    Set cnt = TakeSelectControl
174:    If cnt Is Nothing Then Exit Sub
tryagin:
176:    sOldName = cnt.Name
177:    sNewName = InputBox("Enter a new name for Control", "Renaming Control:", sOldName)
178:    If sNewName = vbNullString Or sNewName = sOldName Then Exit Sub
179:
180:    cnt.Name = sNewName
181:    Set CodeMod = Application.VBE.SelectedVBComponent.CodeModule
182:    With CodeMod
183:        strVar = .Lines(1, .CountOfLines)
184:        'strVar = Replace(strVar, sOldName, sNewName)
185:        strVar = ReplceCode(strVar, sOldName, sNewName)
186:        .DeleteLines StartLine:=1, Count:=.CountOfLines
187:        .InsertLines Line:=1, String:=strVar
188:    End With
189:    Exit Sub
ErrorHandler:
191:    Select Case Err.Number
        Case 40044:
193:            Call MsgBox("Error Invalid Control name entered [" & sNewName & "], set a different name!", vbCritical, "Invalid Control name entered:")
194:            Exit Sub
195:        Case -2147319764:
196:            Call MsgBox("This Control name is already in use [" & sNewName & "], set a different name!", vbCritical, "The name is ambiguous:")
197:            Exit Sub
198:        Case Else:
199:            Debug.Print "Error in RenameControl" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
200:            Call WriteErrorLog("RenameControl")
201:    End Select
202:    Err.Clear
203: End Sub
     Public Sub CopyStyleControl()
205:    Dim cnt         As Object
206:    Set cnt = TakeSelectControl(True)
207:    If cnt Is Nothing Then Exit Sub
208:
209:    'установка по умолчанию значений
210:    tpStyle.lBackStyle = 1
211:    tpStyle.lBorderColor = -2147483642
212:    tpStyle.lBorderStyle = 0
213:    tpStyle.bVisible = True
214:    tpStyle.bLocked = False
215:    tpStyle.bEnabled = True
216:    tpStyle.lBackStyle = 1
217:
218:    On Error Resume Next
219:    With cnt
220:        tpStyle.bEnabled = .Enabled
221:        tpStyle.bFontBold = .Font.Bold
222:        tpStyle.bFontItalic = .Font.Italic
223:        tpStyle.bFontStrikethru = .Font.Strikethrough
224:        tpStyle.bFontUnderline = .Font.Underline
225:        tpStyle.bLocked = .Locked
226:        tpStyle.bVisible = .visible
227:        tpStyle.cuFontSize = .Font.Size
228:        tpStyle.lBackColor = .BackColor
229:        tpStyle.lForeColor = .ForeColor
230:        tpStyle.sFontName = .Font.Name
231:        tpStyle.snHeight = .Height
232:        tpStyle.snWidth = .Width
233:
234:        tpStyle.lBackStyle = .BackStyle
235:        tpStyle.lBorderColor = .BorderColor
236:        tpStyle.lBorderStyle = .BorderStyle
237:    End With
238: End Sub
     Public Sub PasteStyleControl()
240:    Dim cnt         As Object
241:    Set cnt = TakeSelectControl(True)
242:    If cnt Is Nothing Then Exit Sub
243:    On Error Resume Next
244:    With cnt
245:        .Enabled = tpStyle.bEnabled
246:        .Font.Bold = tpStyle.bFontBold
247:        .Font.Italic = tpStyle.bFontItalic
248:        .Font.Strikethrough = tpStyle.bFontStrikethru
249:        .Font.Underline = tpStyle.bFontUnderline
250:        .Locked = tpStyle.bLocked
251:        .visible = tpStyle.bVisible
252:        .Font.Size = tpStyle.cuFontSize
253:        .BackColor = tpStyle.lBackColor
254:        .ForeColor = tpStyle.lForeColor
255:        .Font.Name = tpStyle.sFontName
256:        If tpStyle.snHeight > 0 Then .Height = tpStyle.snHeight
257:        If tpStyle.snWidth > 0 Then .Width = tpStyle.snWidth
258:
259:        .BackStyle = tpStyle.lBackStyle
260:        .BorderColor = tpStyle.lBorderColor
261:        .BorderStyle = tpStyle.lBorderStyle
262:    End With
263: End Sub
     Public Sub AddIcon()
265:    Dim cnt    As control
266:    Dim objForm As InsertIconUserForm
267:
268:    On Error GoTo ErrorHandler
269:
270:    Set cnt = TakeSelectControl
271:    If cnt Is Nothing Then Exit Sub
272:
273:    Set objForm = New InsertIconUserForm
274:    With objForm
275:        .Show
276:        cnt.Font.Name = .lbNameFont.Caption
277:        DoEvents
278:        If TypeName(cnt) = "Label" Then
279:            cnt.Caption = VBA.ChrW$(.lbASC.Caption)
280:        Else
281:            cnt.Value = VBA.ChrW$(.lbASC.Caption)
282:        End If
283:    End With
284:
285:    Exit Sub
ErrorHandler:
287:    Select Case Err.Number
        Case Else:
289:            Debug.Print "Error in RenameControl" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
290:            Call WriteErrorLog("AddIcon")
291:    End Select
292:    Err.Clear
293: End Sub
     Public Sub HelpMoveControl()
295:    Call URLLinks(C_Const.URL_MOVE_CNTR)
296: End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : UperTextInControl - изменение регистров у контроллов на форме
'* Created    : 13-04-2021 09:46
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub UperTextInControl()
305:    Dim oCont As Object
306:    Set oCont = TakeSelectControl
307:    If oCont Is Nothing Then Exit Sub
308:
309:    If PropertyIsCapiton(oCont, True) Then
310:        oCont.Caption = VBA.UCase$(oCont.Caption)
311:    End If
312:    If PropertyIsCapiton(oCont, False) Then
313:        oCont.Text = VBA.UCase$(oCont.Text)
314:    End If
315:
316: End Sub
     Public Sub LowerTextInControl()
318:    Dim oCont As Object
319:    Set oCont = TakeSelectControl
320:    If oCont Is Nothing Then Exit Sub
321:
322:    If PropertyIsCapiton(oCont, True) Then
323:        oCont.Caption = VBA.LCase$(oCont.Caption)
324:    End If
325:    If PropertyIsCapiton(oCont, False) Then
326:        oCont.Text = VBA.LCase$(oCont.Text)
327:    End If
328:
329: End Sub

     Public Sub UperTextInForm()
332:    Dim oVBComp As VBIDE.VBComponent
333:    Set oVBComp = Application.VBE.SelectedVBComponent
334:    With oVBComp
335:        If .Type = vbext_ct_MSForm Then
336:            .Properties("Caption") = VBA.UCase$(.Properties("Caption"))
337:        End If
338:    End With
339: End Sub
     Public Sub LowerTextInForm()
341:    Dim oVBComp As VBIDE.VBComponent
342:    Set oVBComp = Application.VBE.SelectedVBComponent
343:    With oVBComp
344:        If .Type = vbext_ct_MSForm Then
345:            .Properties("Caption") = VBA.LCase$(.Properties("Caption"))
346:        End If
347:    End With
348: End Sub

