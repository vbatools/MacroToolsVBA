Attribute VB_Name = "A_RibbonCallbacks"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : A_RibbonCallbacks - модуль обратных вызовов ленты управления Excel
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Private Module
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : MacroToolsLoadRibbon - при активайии проверить наличие обновлений
'* Created    : 08-10-2020 13:45
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByRef ribbon As IRibbonUI :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub MacroToolsLoadRibbon(ByRef ribbon As IRibbonUI)
23:    On Error Resume Next
24:    Call R_Update.StartUpdate
25: End Sub

    Private Sub RefrasBtn(ByRef control As IRibbonControl)
28:    If VBAIsTrusted Then
29:        Call B_CreateMenus.RefreshMenu
30:    Else
31:        Call MsgBox(C_Const.sMSGVBA1, vbCritical, C_Const.sMSGVBA2)
32:    End If
33: End Sub
    Private Sub ImportCodeBaseBtn(ByRef control As IRibbonControl)
35:    On Error GoTo ErrorHandler
36:    If VBAIsTrusted Then
37:        Workbooks(C_Const.NAME_ADDIN & ".xlam").Sheets(C_Const.SH_SNIPPETS).Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
38:        Call MsgBox("Выгрузка базы кода произведена", vbInformation, "Выгрузка базы кода:")
39:    Else
40:        Call MsgBox(C_Const.sMSGVBA1, vbCritical, C_Const.sMSGVBA2)
41:    End If
42:    Exit Sub
ErrorHandler:
44:    Select Case Err.Number
        Case 91:
46:            Call MsgBox("Нет открытых " & Chr(34) & "Файлов Excel" & Chr(34) & "!", vbOKOnly + vbExclamation, "Ошибка:")
47:        Case Else:
48:            Call MsgBox("Ошибка! в ImportCodeBaseBtn" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbOKOnly + vbExclamation, "Ошибка:")
49:            Call WriteErrorLog("ImportCodeBaseBtn")
50:    End Select
51:    Err.Clear
52: End Sub
    Private Sub AddCodeBaseBtn(ByRef control As IRibbonControl)
54:    If VBAIsTrusted Then
55:        Call AddCodeView.Show
56:    Else
57:        Call MsgBox(C_Const.sMSGVBA1, vbCritical, C_Const.sMSGVBA2)
58:    End If
59: End Sub
    Private Sub AddStatBtn(ByRef control As IRibbonControl)
61:    If VBAIsTrusted Then
62:        Call I_StatisticVBAProj.AddSheetStatistica
63:    Else
64:        Call MsgBox(C_Const.sMSGVBA1, vbCritical, C_Const.sMSGVBA2)
65:    End If
66: End Sub
    Sub btnHiddenModule(control As IRibbonControl)
68:    Call HiddenModule.Show
69: End Sub
    Private Sub AddInBtn(ByRef control As IRibbonControl)
71:    On Error GoTo ErrorHandler
72:    Application.Dialogs(xlDialogAddinManager).Show
73:    Exit Sub
ErrorHandler:
75:    Err.Clear
76:    Call MsgBox("Нет открытых " & Chr(34) & "Файлов Excel" & Chr(34) & "!", vbOKOnly + vbExclamation, "Ошибка:")
77: End Sub
    Private Sub VBABtn(ByRef control As IRibbonControl)
79:    Call VBAVBEOpen
80: End Sub
    Private Sub BtnExportVBA(control As IRibbonControl)
82:    Call VBAVBEOpen
83:    Call ModuleCommander.Show
84: End Sub
    Public Sub VBAVBEOpen()
86:    If C_PublicFunctions.Num_Not_Stable Then Call SendKeys("%{NUMLOCK}")
87:    Call SendKeys("%{F11}")
88: End Sub
    Private Sub BtnVSC(control As IRibbonControl)
90:    Call VersionSistemControls.Show
91: End Sub
     Private Sub onSwitcherReferenceStyle(ByRef control As IRibbonControl)
93:    With Application
94:        If .ReferenceStyle = xlR1C1 Then
95:            .ReferenceStyle = xlA1
96:        Else
97:            .ReferenceStyle = xlR1C1
98:        End If
99:    End With
100: End Sub
     Private Sub onOpenFileExcel(ByRef control As IRibbonControl)
102:    Call O_XML.OpenAndCloseExcelFileInFolder(bOpenFile:=True, bBackUp:=False)
103: End Sub
     Private Sub onCloseFileExcel(ByRef control As IRibbonControl)
105:    Call O_XML.OpenAndCloseExcelFileInFolder(bOpenFile:=False, bBackUp:=True)
106: End Sub
     Private Sub onUnProtectVBA(ByRef control As IRibbonControl)
108:    Call P_UnProtected.unprotected
109: End Sub
     Private Sub onUnProtectSheets(ByRef control As IRibbonControl)
111:    On Error GoTo ErrorHandler
112:    Call ProtectedSheets.Show
113:    Exit Sub
ErrorHandler:
115:    Select Case Err.Number
        Case 91:
117:            Call MsgBox("Нет открытых " & Chr(34) & "Файлов Excel" & Chr(34) & "!", vbOKOnly + vbExclamation, "Ошибка:")
118:        Case Else:
119:            Call MsgBox("Ошибка! в onUnUnProtectSheets" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbOKOnly + vbExclamation, "Ошибка:")
120:            Call WriteErrorLog("onUnUnProtectSheets")
121:    End Select
122:    Err.Clear
123: End Sub
     Private Sub onUnProtectVBAUnivable(control As IRibbonControl)
125:    Call P_UnProtected.DelPasswordVBAProjectUnivable
126: End Sub
     Private Sub onAddShapeStatistic(ByRef control As IRibbonControl)
128:    Call AddShapeStatistic
129: End Sub
     Private Sub onOptions(ByRef control As IRibbonControl)
131:    Call OptionsCodeFormat.Show
132: End Sub
     Private Sub onOptionsComments(ByRef control As IRibbonControl)
134:    Call SettingsAddCommentsProc.Show
135: End Sub
     Private Sub BtnHelpMain(control As IRibbonControl)
137:    Call HelpMainAddin
138: End Sub
     Private Sub BtnMainBuilders(control As IRibbonControl)
140:    Call URLLinks(C_Const.URL_BILD)
141: End Sub
     Private Sub BtnHelpControls(control As IRibbonControl)
143:    Call URLLinks(C_Const.URL_MOVE_CNTR)
144: End Sub
     Private Sub BtnHelpSnippets(control As IRibbonControl)
146:    Call URLLinks(C_Const.URL_STYLE)
147: End Sub
     Private Sub BtnHelpPass(control As IRibbonControl)
149:    Call URLLinks(C_Const.URL_FILE)
150: End Sub
     Private Sub BtnHelpContacts(control As IRibbonControl)
152:    Call URLLinks(C_Const.URL_CONTACT)
153: End Sub
     Private Sub HelpMainAddin()
155:    Call URLLinks(C_Const.URL_ADDIN)
156: End Sub
     Private Sub onInToFile(control As IRibbonControl)
158:    Call Q_InToFile.InToFile
159: End Sub
     Private Sub BtnOrderMacro(control As IRibbonControl)
161:    Call URLLinks(C_Const.URL_CONTACT)
162: End Sub
'Version
     Public Sub getVisible(control As IRibbonControl, ByRef visible)
165:    visible = C_Const.FlagVisible
166: End Sub
'btnVersion
     Public Sub onVisible(control As IRibbonControl)
169:    Call URLLinks(C_Const.URL_DOWNLOAD)
170: End Sub
     Private Sub BtnUnProtectSheetsXML(control As IRibbonControl)
172:    Call DeletePaswortSheets
173: End Sub
     Private Sub onProtectVBAUnivable(control As IRibbonControl)
175:    Call SetPasswordVBAProjectUnviewable
176: End Sub
     Private Sub BtnVK(control As IRibbonControl)
178:    Call URLLinks(C_Const.URL_VK)
179: End Sub
     Private Sub BtnFB(control As IRibbonControl)
181:    Call URLLinks(C_Const.URL_FB)
182: End Sub
'смена темы
     Private Sub onBlackTheme(control As IRibbonControl)
185:    Call V_BlackAndWiteTheme.ChangeColorDarkTheme
186: End Sub
     Private Sub onWhiteTheme(control As IRibbonControl)
188:    Call V_BlackAndWiteTheme.ChangeColorWhiteTheme
189: End Sub
     Private Sub onToolCharMonitor(control As IRibbonControl)
191:    Call CharsMonitor.Show
192: End Sub
'Регулярные выражения
     Private Sub onTestRegExp(control As IRibbonControl)
195:    Call W_RegExp.AddSheetTestRegExp
196: End Sub
     Private Sub onTempleteRegExp(control As IRibbonControl)
198:    Call RegExpTemplateManager.Show
199: End Sub
     Private Sub onRegExpFunValNumber(control As IRibbonControl)
201:    ActiveCell.FormulaR1C1 = "=РЕГВЫР_ПОЛУЧЗНАЧПОНОМЕРУ()"
202:    Call FunctionWizardShowExc
203: End Sub
     Private Sub onExpFunCount(control As IRibbonControl)
205:    ActiveCell.FormulaR1C1 = "=РЕГВЫР_СЧЁТ()"
206:    Call FunctionWizardShowExc
207: End Sub
     Private Sub onRegExpFunTest(control As IRibbonControl)
209:    ActiveCell.FormulaR1C1 = "=РЕГВЫР_ТЕСТ()"
210:    Call FunctionWizardShowExc
211: End Sub
     Private Sub onRegExpFunReplace(control As IRibbonControl)
213:    ActiveCell.FormulaR1C1 = "=РЕГВЫР_ЗАМЕНИТЬ()"
214:    Call FunctionWizardShowExc
215: End Sub
     Private Sub FunctionWizardShowExc()
217:    If Application.Dialogs(xlDialogFunctionWizard).Show = False Then
218:        ActiveCell.Clear
219:    End If
220:    Calculate
221: End Sub
     Private Sub onParserVBA(control As IRibbonControl)
223:    If VBAIsTrusted Then
224:        Call N_ObfParserVBA.StartParser
225:    Else
226:        Call MsgBox(C_Const.sMSGVBA1, vbCritical, C_Const.sMSGVBA2)
227:    End If
228: End Sub
     Private Sub onObfuscator(ByRef control As IRibbonControl)
230:    If VBAIsTrusted Then
231:        Call N_ObfMainNew.StartObfuscation
232:    Else
233:        Call MsgBox(C_Const.sMSGVBA1, vbCritical, C_Const.sMSGVBA2)
234:    End If
235: End Sub
     Private Sub onFormatsDel(control As IRibbonControl)
237:    If VBAIsTrusted Then
238:        Call ObfuscationCode.Show
239:    Else
240:        Call MsgBox(C_Const.sMSGVBA1, vbCritical, C_Const.sMSGVBA2)
241:    End If
242: End Sub
     Private Sub BtnInfoFile(control As IRibbonControl)
244:    Call InfoFile.Show
245: End Sub
     Sub ParserStrings(control As IRibbonControl)
247:    Call ZA_ParserString.ParserStringWB
248: End Sub

Sub ReNameParserString(control As IRibbonControl)
251:    Call ZA_ParserString.ReNameStr
End Sub
