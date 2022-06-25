Attribute VB_Name = "B_CreateMenus"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : B_CreateMenus - создание меню в VBE
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Private Module
Option Explicit
Public ToolContextEventHandlers As New Collection

#If Win64 Then
Private Declare PtrSafe Function GetKeyboardLayoutName Lib "USER32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare PtrSafe Function LoadKeyboardLayout Lib "USER32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
#Else
Private Declare Function GetKeyboardLayoutName Lib "USER32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare Function LoadKeyboardLayout Lib "USER32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
#End If

Private Const LANG_RUSSIAN = 419
Private Const LANG_ENGLISH = 409

    Private Sub Auto_Open()
24:    If VBAIsTrusted And ThisWorkbook.Name = C_Const.NAME_ADDIN & ".xlam" Then    '
25:        Call AddContextMenus
26:    End If
27: End Sub
    Public Sub AddContextMenus()
29:
30:    Call AddNewCommandBarMenu(C_Const.MENUMOVECONTRL)
31:    Call AddButtom(C_Const.MTAG5, 984, "Tool Reference", "HelpMoveControl", C_Const.MENUMOVECONTRL, False, True)
32:    Call AddButtom(C_Const.MTAG4, 38, "", "MoveControl", C_Const.MENUMOVECONTRL)
33:    Call AddButtom(C_Const.MTAG3, 40, "", "MoveControl", C_Const.MENUMOVECONTRL, False, True)
34:    Call AddButtom(C_Const.MTAG2, 39, "", "MoveControl", C_Const.MENUMOVECONTRL)
35:    Call AddButtom(C_Const.MTAG1, 41, "", "MoveControl", C_Const.MENUMOVECONTRL)
36:    Call AddComboBoxMove(C_Const.MENUMOVECONTRL)
37:
38:    Call AddNewCommandBarMenu(C_Const.TOOLSMENU)
39:    Call AddButtom(C_Const.TAG15, 984, "Add-in Help", "HelpMainAddin", C_Const.TOOLSMENU, False, True)
40:    Call AddButtom(C_Const.TAG29, 277, "Keyboard shortcuts", "AddLegendHotKeys", C_Const.TOOLSMENU, False, True)
41:    Call AddButtom(C_Const.TAG12, 0, "FormatBuilder", "subFormatBuilder", C_Const.TOOLSMENU, True, True)
42:    Call AddButtom(C_Const.TAG11, 0, "MsgBoxBuilder", "subMsgBoxGenerator", C_Const.TOOLSMENU, True, True)
43:    Call AddButtom(C_Const.TAG13, 0, "ProcedureBuilder", "subProcedureBuilder", C_Const.TOOLSMENU, True, True)
44:    Call AddButtom(C_Const.TAG23, 107, "Option's Explicit and Private Module", "insertOptionsExplicitAndPrivateModule", C_Const.TOOLSMENU, False, False)
45:    Call AddButtom(C_Const.TAG28, 0, "Option's", "subOptionsMenu", C_Const.TOOLSMENU, True, True)
46:
47:    Call AddButtom(C_Const.TAG24, 2045, "Copy", "SetInCipBoard", C_Const.TOOLSMENU, True, False)
48:    Call AddButtom(C_Const.TAG25, 22, "Paste", "GetFromCipBoard", C_Const.TOOLSMENU, True, True)
49:
50:    Call AddButtom(C_Const.TAG20, 1714, "Search for unused variables", "SerchVariableUnUsedInSelectedWorkBook", C_Const.TOOLSMENU, False, False)
51:    Call AddButtom(C_Const.TAG19, 3838, "Close all VBE windows", "CloseAllWindowsVBE", C_Const.TOOLSMENU, False, False)
52:    Call AddButtom(C_Const.TAG14, 22, "Insert the LogRecorder class", "AddLogRecorderClass", C_Const.TOOLSMENU, False, True)
53:
54:    Call AddButtom(C_Const.TAG19, 8, "TODO List", "ShowTODOList", C_Const.TOOLSMENU, False, False)
55:    Call AddButtom(C_Const.TAG18, 1972, "Create a TODO", "sysAddTODOTop", C_Const.TOOLSMENU, False, False)
56:    Call AddButtom(C_Const.TAG17, 456, "Create an update comment line", "sysAddModifiedTop", C_Const.TOOLSMENU, False, False)
57:    Call AddButtom(C_Const.TAG16, 1546, "Create a comment", "sysAddHeaderTop", C_Const.TOOLSMENU, False, True)
58:
59:    Call AddButtom(C_Const.TAG10, 3917, "Remove Code Formatting", "CutTab", C_Const.TOOLSMENU)
60:    Call AddButtom(C_Const.TAG9, 3919, "Format the Code", "ReBild", C_Const.TOOLSMENU, False, True)
61:    Call AddButtom(C_Const.TAG8, 12, "Remove line numbering", "RemoveLineNumbers_", C_Const.TOOLSMENU)
62:    Call AddButtom(C_Const.TAG7, 11, "Create line numbering", "AddLineNumbers_", C_Const.TOOLSMENU)
63:    Call AddComboBox(C_Const.TOOLSMENU)
64:    Call AddButtom(C_Const.TAG27, 210, "Sorting procedures alphabetically", "AlphabetizeProcedure", C_Const.TOOLSMENU, False, True)
65:    Call AddButtom(C_Const.TAG6, 47, "Clear the window [Immediate]", "ClearImmediateWindow", C_Const.TOOLSMENU, False, True)
66:    Call AddButtom(C_Const.TAG5, 2059, "Create a legend", "AddLegend", C_Const.TOOLSMENU)
67:    Call AddButtom(C_Const.TAG4, 21, "Delete a module", "DeleteSnippetEnumModule", C_Const.TOOLSMENU)
68:    Call AddButtom(C_Const.TAG3, 1753, "Insert a module", "AddSnippetEnumModule", C_Const.TOOLSMENU)
69:    Call AddButtom(C_Const.TAG2, 22, "Insert Code", "InsertCode", C_Const.TOOLSMENU, False, False)
70:
71:    Call AddButtom(C_Const.TAG26, 9634, "Swap the relative [=]", "SwapEgual", C_Const.POPMENU, True, False)
72:    Call AddButtom(C_Const.TAG21, 0, "UPPER Case", "toUpperCase", C_Const.POPMENU, True, False)
73:    Call AddButtom(C_Const.TAG22, 0, "lower Case", "toLowerCase", C_Const.POPMENU, True, False)
74:    Call AddButtom(C_Const.TAG1, 22, "Insert Code", "InsertCode", C_Const.POPMENU, True, False)
75:
76:    Call AddButtom(C_Const.RTAG1, 162, "ReName Control", "RenameControl", C_Const.RENAMEMENU, True)
77:    Call AddButtom(C_Const.RTAG2, 22, "Paste Style", "PasteStyleControl", C_Const.RENAMEMENU, True)
78:    Call AddButtom(C_Const.RTAG3, 1076, "Copy Style", "CopyStyleControl", C_Const.RENAMEMENU, True)
79:    Call AddButtom(C_Const.RTAG4, 704, "Paste Icon", "AddIcon", C_Const.RENAMEMENU, True, True)
80:    Call AddButtom(C_Const.RTAG5, 0, "UPPER Case", "UperTextInControl", C_Const.RENAMEMENU, True, False)
81:    Call AddButtom(C_Const.RTAG6, 0, "lower Case", "LowerTextInControl", C_Const.RENAMEMENU, True, False)
82:
83:    Call AddButtom(C_Const.CTAG1, 2045, "Copy Module", "CopyModyleVBE", C_Const.COPYMODULE, True, False)
84:
85:    Call AddButtom(C_Const.RTAG2, 22, "Paste Style", "PasteStyleControl", C_Const.mMSFORMS, True)
86:    Call AddButtom(C_Const.RTAG3, 1076, "Copy Style", "CopyStyleControl", C_Const.mMSFORMS, True)
87:    Call AddButtom(C_Const.RTAG5, 0, "UPPER Case", "UperTextInForm", C_Const.mMSFORMS, True, False)
88:    Call AddButtom(C_Const.RTAG6, 0, "lower Case", "LowerTextInForm", C_Const.mMSFORMS, True, False)
89: End Sub
     Private Sub AddNewCommandBarMenu(ByVal sNameCommandBar As String)
91:    Dim myCommandBar As CommandBar
92:    On Error GoTo AddNewCommandBar
93:    Set myCommandBar = Application.VBE.CommandBars(sNameCommandBar)
94:    If myCommandBar Is Nothing Then
AddNewCommandBar:
96:        Set myCommandBar = Application.VBE.CommandBars.Add(Name:=sNameCommandBar, Position:=msoBarTop)
97:        myCommandBar.visible = True
98:        myCommandBar.RowIndex = 3
99:    End If
100: End Sub
     Private Sub AddButtom( _
             ByVal sTag As String, _
             ByVal Face As Long, _
             ByVal Capitan As String, _
             ByVal sOnAction As String, _
             ByVal sMenu As String, _
             Optional ByRef VisibleCapiton As Boolean = False, _
             Optional ByVal Begin_Group As Boolean = False, _
             Optional ByVal ShortcutText As String = vbNullString, _
             Optional ByVal Before As Byte = 1)
111:    Dim btn         As CommandBarButton
112:    Dim evtContextMenu As VBECommandHandler
113:    Set btn = Application.VBE.CommandBars(sMenu).Controls.Add(Type:=msoControlButton, Before:=Before)
114:    With btn
115:        .FaceId = Face
116:        If VisibleCapiton Then .Caption = Capitan
117:        .TooltipText = Capitan
118:        .Tag = sTag
119:        .OnAction = "'" & ThisWorkbook.Name & "'!" & sOnAction
120:        .Style = msoButtonIconAndCaption
121:        .BeginGroup = Begin_Group
122:        .ShortcutText = ShortcutText
123:    End With
124:    Set evtContextMenu = New VBECommandHandler
125:    Set evtContextMenu.EvtHandler = btn
126:    ToolContextEventHandlers.Add evtContextMenu
127: End Sub
     Private Sub AddComboBox(ByVal sMenu As String)
129:    Dim combox      As CommandBarComboBox
130:    Set combox = Application.VBE.CommandBars(sMenu).Controls.Add(Type:=msoControlComboBox, Before:=1)
131:    With combox
132:        .Tag = C_Const.TAGCOM
133:        .AddItem C_Const.SELECTEDMODULE
134:        .AddItem C_Const.ALLVBAPROJECT
135:        .Text = C_Const.SELECTEDMODULE
136:    End With
137: End Sub
     Private Sub AddComboBoxMove(ByVal sMenu As String)
139:    Dim combox      As CommandBarComboBox
140:    Set combox = Application.VBE.CommandBars(sMenu).Controls.Add(Type:=msoControlComboBox, Before:=1)
141:    With combox
142:        .Tag = C_Const.MTAGCOM
143:        .AddItem C_Const.MOVECONT
144:        .AddItem C_Const.MOVECONTTOPLEFT
145:        .AddItem C_Const.MOVECONTBOTTOMRIGHT
146:        .Text = C_Const.MOVECONT
147:    End With
148: End Sub
     Private Sub Auto_Close()
150:    If VBAIsTrusted Then
151:        Call DeleteContextMenus
152:    End If
153: End Sub
     Public Sub DeleteContextMenus()
155:    Dim myCommandBar As CommandBar
156:    On Error GoTo ErrorHandler
157:
158:    Call DeleteButton(C_Const.TAG1, C_Const.POPMENU)
159:    Call DeleteButton(C_Const.TAG26, C_Const.POPMENU)
160:    Call DeleteButton(C_Const.TAG21, C_Const.POPMENU)
161:    Call DeleteButton(C_Const.TAG22, C_Const.POPMENU)
162:
163:    Call DeleteButton(C_Const.CTAG1, C_Const.COPYMODULE)
164:
165:    Call DeleteButton(C_Const.RTAG1, C_Const.RENAMEMENU)
166:    Call DeleteButton(C_Const.RTAG2, C_Const.RENAMEMENU)
167:    Call DeleteButton(C_Const.RTAG3, C_Const.RENAMEMENU)
168:    Call DeleteButton(C_Const.RTAG4, C_Const.RENAMEMENU)
169:    Call DeleteButton(C_Const.RTAG5, C_Const.RENAMEMENU)
170:    Call DeleteButton(C_Const.RTAG6, C_Const.RENAMEMENU)
171:
172:    Call DeleteButton(C_Const.RTAG2, C_Const.mMSFORMS)
173:    Call DeleteButton(C_Const.RTAG3, C_Const.mMSFORMS)
174:    Call DeleteButton(C_Const.RTAG5, C_Const.mMSFORMS)
175:    Call DeleteButton(C_Const.RTAG6, C_Const.mMSFORMS)
176:
177:    Set myCommandBar = Application.VBE.CommandBars(C_Const.TOOLSMENU)
178:    If Not myCommandBar Is Nothing Then
179:        myCommandBar.Delete
180:    End If
181:
182:    Set myCommandBar = Application.VBE.CommandBars(C_Const.MENUMOVECONTRL)
183:    If Not myCommandBar Is Nothing Then
184:        myCommandBar.Delete
185:    End If
186:
187:    'очистка колекции
188:    Do Until ToolContextEventHandlers.Count = 0
189:        ToolContextEventHandlers.Remove 1
190:    Loop
191:
192:    Exit Sub
ErrorHandler:
194:
195:    Select Case Err
        Case 5:
197:            Err.Clear
198:        Case Else:
199:            Debug.Print "Mistake! in DeleteContextMenus" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
200:            Call WriteErrorLog("DeleteContextMenus")
201:    End Select
202:    Err.Clear
203: End Sub
     Private Sub DeleteButton(ByRef sTag As String, ByVal sMenu As String)
205:    Dim Cbar        As CommandBar
206:    Dim Ctrl        As CommandBarControl
207:    On Error GoTo ErrorHandler
208:    Set Cbar = Application.VBE.CommandBars(sMenu)
209:    For Each Ctrl In Cbar.Controls
210:        If Ctrl.Tag = sTag Then
211:            Ctrl.Delete
212:            'Exit Sub
213:        End If
214:    Next Ctrl
215:    Exit Sub
ErrorHandler:
217:    Debug.Print "Mistake! in DeleteButton" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
218:    Call WriteErrorLog("DeleteButton")
219:    Err.Clear
220:    Resume Next
221: End Sub
     Public Function VBAIsTrusted() As Boolean
223:    On Error GoTo ErrorHandler
224:    Dim sTxt As String
225:    sTxt = Application.VBE.Version
226:    VBAIsTrusted = True
227:    Exit Function
ErrorHandler:
229:    Select Case Err.Number
        Case 1004:
231:            'If ThisWorkbook.Name = C_Const.NAME_ADDIN & ".xlam" Then
232:            Call MsgBox("Warning!" & C_Const.NAME_ADDIN & vbLf & vbNewLine & _
                        "Disabled: [Trust access to the VBE object model]" & vbLf & _
                        "To enable it, go to: File->Settings->Security Management Center->Macro Settings" & _
                        vbLf & vbNewLine & "And restart Excel", vbCritical, "Warning:")
236:        Case Else:
237:            Debug.Print "Mistake! in VBAIsTrusted" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
238:            Call WriteErrorLog("VBAIsTrusted")
239:    End Select
240:    Err.Clear
241:    VBAIsTrusted = False
242: End Function
     Public Function WhatIsTextInComboBoxHave() As String
244:    Dim myCommandBar As CommandBar
245:    Dim cntrl       As CommandBarControl
246:
247:    Set myCommandBar = Application.VBE.CommandBars(C_Const.TOOLSMENU)
248:    For Each cntrl In myCommandBar.Controls
249:        If cntrl.Tag = C_Const.TAGCOM Then
250:            WhatIsTextInComboBoxHave = cntrl.Text
251:            Exit Function
252:        End If
253:    Next cntrl
254: End Function
     Public Sub ClearImmediateWindow()
256:    Dim KeybLayoutName As String * 8
257:    KeybLayoutName = String(8, "0")
258:    GetKeyboardLayoutName KeybLayoutName
259:    KeybLayoutName = Val(KeybLayoutName)
260:
261:    Select Case Val(KeybLayoutName)
        Case LANG_ENGLISH
263:            Call ClearImmediateWindowFunction
264:            Call ClearImmediateWindowFunction
265:        Case LANG_RUSSIAN
266:            ' Переключение на английскую раскладку
267:            Call LoadKeyboardLayout("00000409", &H1)
268:            Call ClearImmediateWindowFunction
269:            Call LoadKeyboardLayout("00000419", &H1)
270:        Case Else
271:            Call MsgBox("Switch the keyboard layout to English!", vbInformation, "Switching the keyboard layout")
272:    End Select
273: End Sub
     Private Sub ClearImmediateWindowFunction()
275:    Call SendKeys("^g")
276:    Call SendKeys("^a")
277:    Call SendKeys("{DEL}")
278: End Sub
     Public Sub RefreshMenu()
280:    Call B_CreateMenus.DeleteContextMenus
281:    Call B_CreateMenus.AddContextMenus
282:    Call MsgBox("Rebooting the add-in" & C_Const.NAME_ADDIN & "passed!", vbInformation, "Rebooting the add-in" & C_Const.NAME_ADDIN & ":")
283: End Sub
     Private Sub subMsgBoxGenerator()
285:    MsgBoxGenerator.Show
286: End Sub
     Private Sub subFormatBuilder()
288:    BilderFormat.Show
289: End Sub
     Private Sub subProcedureBuilder()
291:    BilderProcedure.Show
292: End Sub
Private Sub subOptionsMenu()
294:    Call Y_Options.subOptions
End Sub

