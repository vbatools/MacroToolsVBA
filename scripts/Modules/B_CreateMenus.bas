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
40:    Call AddButtom(C_Const.TAG12, 0, "FormatBuilder", "subFormatBuilder", C_Const.TOOLSMENU, True, True)
41:    Call AddButtom(C_Const.TAG11, 0, "MsgBoxBuilder", "subMsgBoxGenerator", C_Const.TOOLSMENU, True, True)
42:    Call AddButtom(C_Const.TAG13, 0, "ProcedureBuilder", "subProcedureBuilder", C_Const.TOOLSMENU, True, True)
43:    Call AddButtom(C_Const.TAG23, 0, "Option's", "subOptionsMenu", C_Const.TOOLSMENU, True, True)
44:
45:    Call AddButtom(C_Const.TAG24, 2045, "Copy", "SetInCipBoard", C_Const.TOOLSMENU, True, False)
46:    Call AddButtom(C_Const.TAG25, 22, "Paste", "GetFromCipBoard", C_Const.TOOLSMENU, True, True)
47:
48:    Call AddButtom(C_Const.TAG20, 1714, "Finding unused variables", "SerchVariableUnUsedInSelectedWorkBook", C_Const.TOOLSMENU, False, False)
49:    Call AddButtom(C_Const.TAG19, 3838, "Close all VBE windows", "CloseAllWindowsVBE", C_Const.TOOLSMENU, False, False)
50:    Call AddButtom(C_Const.TAG14, 22, "Insert the LogRecorder class", "AddLogRecorderClass", C_Const.TOOLSMENU, False, True)
51:
52:    Call AddButtom(C_Const.TAG19, 8, "TODO List", "ShowTODOList", C_Const.TOOLSMENU, False, False)
53:    Call AddButtom(C_Const.TAG18, 1972, "Create a TODO", "sysAddTODOTop", C_Const.TOOLSMENU, False, False)
54:    Call AddButtom(C_Const.TAG17, 456, "Create an update comment line", "sysAddModifiedTop", C_Const.TOOLSMENU, False, False)
55:    Call AddButtom(C_Const.TAG16, 1546, "Create a comment", "sysAddHeaderTop", C_Const.TOOLSMENU, False, True)
56:
57:    Call AddButtom(C_Const.TAG10, 3917, "Remove Code Formatting", "CutTab", C_Const.TOOLSMENU)
58:    Call AddButtom(C_Const.TAG9, 3919, "Format The Code", "ReBild", C_Const.TOOLSMENU, False, True)
59:    Call AddButtom(C_Const.TAG8, 12, "Remove line numbering", "RemoveLineNumbers_", C_Const.TOOLSMENU)
60:    Call AddButtom(C_Const.TAG7, 11, "Create line numbering", "AddLineNumbers_", C_Const.TOOLSMENU)
61:    Call AddComboBox(C_Const.TOOLSMENU)
62:    Call AddButtom(C_Const.TAG27, 210, "Sorting procedures alphabetically", "AlphabetizeProcedure", C_Const.TOOLSMENU, False, True)
63:    Call AddButtom(C_Const.TAG6, 47, "Clear the window [Immediate]", "ClearImmediateWindow", C_Const.TOOLSMENU, False, True)
64:    Call AddButtom(C_Const.TAG5, 2059, "Create a legend", "AddLegend", C_Const.TOOLSMENU)
65:    Call AddButtom(C_Const.TAG4, 21, "Delete a module", "DeleteSnippetEnumModule", C_Const.TOOLSMENU)
66:    Call AddButtom(C_Const.TAG3, 1753, "Insert a module", "AddSnippetEnumModule", C_Const.TOOLSMENU)
67:    Call AddButtom(C_Const.TAG2, 22, "Insert code", "InsertCode", C_Const.TOOLSMENU, False, False)
68:
69:    Call AddButtom(C_Const.TAG26, 9634, "Swap the relation [=]", "SwapEgual", C_Const.POPMENU, True, False)
70:    Call AddButtom(C_Const.TAG21, 0, "UPPER Case", "toUpperCase", C_Const.POPMENU, True, False)
71:    Call AddButtom(C_Const.TAG22, 0, "lower Case", "toLowerCase", C_Const.POPMENU, True, False)
72:    Call AddButtom(C_Const.TAG1, 22, "Insert code", "InsertCode", C_Const.POPMENU, True, False)
73:
74:    Call AddButtom(C_Const.RTAG1, 162, "ReName Control", "RenameControl", C_Const.RENAMEMENU, True)
75:    Call AddButtom(C_Const.RTAG2, 22, "Paste Style", "PasteStyleControl", C_Const.RENAMEMENU, True)
76:    Call AddButtom(C_Const.RTAG3, 1076, "Copy Style", "CopyStyleControl", C_Const.RENAMEMENU, True)
77:    Call AddButtom(C_Const.RTAG5, 0, "UPPER Case", "UperTextInControl", C_Const.RENAMEMENU, True, False)
78:    Call AddButtom(C_Const.RTAG6, 0, "lower Case", "LowerTextInControl", C_Const.RENAMEMENU, True, False)
79:
80:    Call AddButtom(C_Const.CTAG1, 2045, "Copy Module", "CopyModyleVBE", C_Const.COPYMODULE, True, False)
81:
82:    Call AddButtom(C_Const.RTAG2, 22, "Paste Style", "PasteStyleControl", C_Const.mMSFORMS, True)
83:    Call AddButtom(C_Const.RTAG3, 1076, "Copy Style", "CopyStyleControl", C_Const.mMSFORMS, True)
84:    Call AddButtom(C_Const.RTAG5, 0, "UPPER Case", "UperTextInForm", C_Const.mMSFORMS, True, False)
85:    Call AddButtom(C_Const.RTAG6, 0, "lower Case", "LowerTextInForm", C_Const.mMSFORMS, True, False)
86: End Sub
    Private Sub AddNewCommandBarMenu(ByVal sNameCommandBar As String)
88:    Dim myCommandBar As CommandBar
89:    On Error GoTo AddNewCommandBar
90:    Set myCommandBar = Application.VBE.CommandBars(sNameCommandBar)
91:    If myCommandBar Is Nothing Then
AddNewCommandBar:
93:        Set myCommandBar = Application.VBE.CommandBars.Add(Name:=sNameCommandBar, Position:=msoBarTop)
94:        myCommandBar.visible = True
95:        myCommandBar.RowIndex = 3
96:    End If
97: End Sub
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
108:    Dim btn         As CommandBarButton
109:    Dim evtContextMenu As VBECommandHandler
110:    Set btn = Application.VBE.CommandBars(sMenu).Controls.Add(Type:=msoControlButton, Before:=Before)
111:    With btn
112:        .FaceId = Face
113:        If VisibleCapiton Then .Caption = Capitan
114:        .TooltipText = Capitan
115:        .Tag = sTag
116:        .OnAction = "'" & ThisWorkbook.Name & "'!" & sOnAction
117:        .Style = msoButtonIconAndCaption
118:        .BeginGroup = Begin_Group
119:        .ShortcutText = ShortcutText
120:    End With
121:    Set evtContextMenu = New VBECommandHandler
122:    Set evtContextMenu.EvtHandler = btn
123:    ToolContextEventHandlers.Add evtContextMenu
124: End Sub
     Private Sub AddComboBox(ByVal sMenu As String)
126:    Dim combox      As CommandBarComboBox
127:    Set combox = Application.VBE.CommandBars(sMenu).Controls.Add(Type:=msoControlComboBox, Before:=1)
128:    With combox
129:        .Tag = C_Const.TAGCOM
130:        .AddItem C_Const.SELECTEDMODULE
131:        .AddItem C_Const.ALLVBAPROJECT
132:        .Text = C_Const.SELECTEDMODULE
133:    End With
134: End Sub
     Private Sub AddComboBoxMove(ByVal sMenu As String)
136:    Dim combox      As CommandBarComboBox
137:    Set combox = Application.VBE.CommandBars(sMenu).Controls.Add(Type:=msoControlComboBox, Before:=1)
138:    With combox
139:        .Tag = C_Const.MTAGCOM
140:        .AddItem C_Const.MOVECONT
141:        .AddItem C_Const.MOVECONTTOPLEFT
142:        .AddItem C_Const.MOVECONTBOTTOMRIGHT
143:        .Text = C_Const.MOVECONT
144:    End With
145: End Sub
     Private Sub Auto_Close()
147:    If VBAIsTrusted Then
148:        Call DeleteContextMenus
149:    End If
150: End Sub
     Public Sub DeleteContextMenus()
152:    Dim myCommandBar As CommandBar
153:    On Error GoTo ErrorHandler
154:
155:    Call DeleteButton(C_Const.TAG1, C_Const.POPMENU)
156:    Call DeleteButton(C_Const.TAG26, C_Const.POPMENU)
157:    Call DeleteButton(C_Const.TAG21, C_Const.POPMENU)
158:    Call DeleteButton(C_Const.TAG22, C_Const.POPMENU)
159:
160:    Call DeleteButton(C_Const.CTAG1, C_Const.COPYMODULE)
161:
162:    Call DeleteButton(C_Const.RTAG1, C_Const.RENAMEMENU)
163:    Call DeleteButton(C_Const.RTAG2, C_Const.RENAMEMENU)
164:    Call DeleteButton(C_Const.RTAG3, C_Const.RENAMEMENU)
165:    Call DeleteButton(C_Const.RTAG4, C_Const.RENAMEMENU)
166:    Call DeleteButton(C_Const.RTAG5, C_Const.RENAMEMENU)
167:    Call DeleteButton(C_Const.RTAG6, C_Const.RENAMEMENU)
168:
169:    Call DeleteButton(C_Const.RTAG2, C_Const.mMSFORMS)
170:    Call DeleteButton(C_Const.RTAG3, C_Const.mMSFORMS)
171:    Call DeleteButton(C_Const.RTAG5, C_Const.mMSFORMS)
172:    Call DeleteButton(C_Const.RTAG6, C_Const.mMSFORMS)
173:
174:    Set myCommandBar = Application.VBE.CommandBars(C_Const.TOOLSMENU)
175:    If Not myCommandBar Is Nothing Then
176:        myCommandBar.Delete
177:    End If
178:
179:    Set myCommandBar = Application.VBE.CommandBars(C_Const.MENUMOVECONTRL)
180:    If Not myCommandBar Is Nothing Then
181:        myCommandBar.Delete
182:    End If
183:
184:    'очистка колекции
185:    Do Until ToolContextEventHandlers.Count = 0
186:        ToolContextEventHandlers.Remove 1
187:    Loop
188:
189:    Exit Sub
ErrorHandler:
191:
192:    Select Case Err
        Case 5:
194:            Err.Clear
195:        Case Else:
196:            Debug.Print "Error in DeleteContextMenus" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
197:            Call WriteErrorLog("DeleteContextMenus")
198:    End Select
199:    Err.Clear
200: End Sub
     Private Sub DeleteButton(ByRef sTag As String, ByVal sMenu As String)
202:    Dim Cbar        As CommandBar
203:    Dim Ctrl        As CommandBarControl
204:    On Error GoTo ErrorHandler
205:    Set Cbar = Application.VBE.CommandBars(sMenu)
206:    For Each Ctrl In Cbar.Controls
207:        If Ctrl.Tag = sTag Then
208:            Ctrl.Delete
209:            'Exit Sub
210:        End If
211:    Next Ctrl
212:    Exit Sub
ErrorHandler:
214:    Debug.Print "Error in DeleteButton" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
215:    Call WriteErrorLog("DeleteButton")
216:    Err.Clear
217:    Resume Next
218: End Sub
     Public Function VBAIsTrusted() As Boolean
220:    On Error GoTo ErrorHandler
221:    Dim sTxt As String
222:    sTxt = Application.VBE.Version
223:    VBAIsTrusted = True
224:    Exit Function
ErrorHandler:
226:    Select Case Err.Number
        Case 1004:
228:            'If ThisWorkbook.Name = C_Const.NAME_ADDIN & ".xlam" Then
229:             Call MsgBox("Warning! " & C_Const.NAME_ADDIN & vbLf & vbNewLine & _
                        "Disabled: [Trust access to the VBA object model]" & vbLf & _
                        "To enable it, go to: File->Settings->Security Management Center->Macro Settings" & _
                        vbLf & vbNewLine & "And restart Excel", vbCritical, "Warning:")
233:        Case Else:
234:            Debug.Print "Error! in VBA, isTrusted" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
235:            Call WriteErrorLog("VBAIsTrusted")
236:    End Select
237:    Err.Clear
238:    VBAIsTrusted = False
239: End Function
     Public Function WhatIsTextInComboBoxHave() As String
241:    Dim myCommandBar As CommandBar
242:    Dim cntrl       As CommandBarControl
243:
244:    Set myCommandBar = Application.VBE.CommandBars(C_Const.TOOLSMENU)
245:    For Each cntrl In myCommandBar.Controls
246:        If cntrl.Tag = C_Const.TAGCOM Then
247:            WhatIsTextInComboBoxHave = cntrl.Text
248:            Exit Function
249:        End If
250:    Next cntrl
251: End Function
     Public Sub ClearImmediateWindow()
253:    Dim KeybLayoutName As String * 8
254:    KeybLayoutName = String(8, "0")
255:    GetKeyboardLayoutName KeybLayoutName
256:    KeybLayoutName = Val(KeybLayoutName)
257:
258:    Select Case Val(KeybLayoutName)
        Case LANG_ENGLISH
260:            Call ClearImmediateWindowFunction
261:            Call ClearImmediateWindowFunction
262:        Case LANG_RUSSIAN
263:            ' Переключение на английскую раскладку
264:            Call LoadKeyboardLayout("00000409", &H1)
265:            Call ClearImmediateWindowFunction
266:            Call LoadKeyboardLayout("00000419", &H1)
267:        Case Else
268:            Call MsgBox("Switch your keyboard layout to English!", vbInformation, "Switching the keyboard layout")
269:    End Select
270: End Sub
     Private Sub ClearImmediateWindowFunction()
272:    Call SendKeys("^g")
273:    Call SendKeys("^a")
274:    Call SendKeys("{DEL}")
275: End Sub
     Public Sub RefreshMenu()
277:    Call B_CreateMenus.DeleteContextMenus
278:    Call B_CreateMenus.AddContextMenus
279:    Call MsgBox("The add-in " & C_Const.NAME_ADDIN & " - was rebooted!", vbInformation, "The add" & C_Const.NAME_ADDIN & ":")
280: End Sub
     Private Sub subMsgBoxGenerator()
282:    MsgBoxGenerator.Show
283: End Sub
     Private Sub subFormatBuilder()
285:    BilderFormat.Show
286: End Sub
     Private Sub subProcedureBuilder()
288:    BilderProcedure.Show
289: End Sub
     Private Sub subOptionsMenu()
291:    Call Y_Options.subOptions
292: End Sub

