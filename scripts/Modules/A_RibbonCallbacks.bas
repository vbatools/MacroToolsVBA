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
27:    If VBAIsTrusted Then
28:        Call B_CreateMenus.RefreshMenu
29:    End If
30: End Sub
    Private Sub ImportCodeBaseBtn(ByRef control As IRibbonControl)
32:    On Error GoTo ErrorHandler
33:    If VBAIsTrusted Then
34:        Workbooks(C_Const.NAME_ADDIN & ".xlam").Sheets(C_Const.SH_SNIPPETS).Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
35:        Call MsgBox("The code base has been unloaded", vbInformation, "Unloading the code base:")
36:    End If
37:    Exit Sub
ErrorHandler:
39:    Select Case Err.Number
        Case 91:
41:            Call MsgBox("No open" & Chr(34) & "Excel Files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Mistake:")
42:        Case Else:
43:            Call MsgBox("Mistake! in ImportCodeBaseBtn" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbExclamation, "Mistake:")
44:            Call WriteErrorLog("ImportCodeBaseBtn")
45:    End Select
46:    Err.Clear
47: End Sub
    Private Sub AddCodeBaseBtn(ByRef control As IRibbonControl)
49:    If VBAIsTrusted Then
50:        Call AddCodeView.Show
51:    End If
52: End Sub
    Private Sub AddStatBtn(ByRef control As IRibbonControl)
54:    If VBAIsTrusted Then
55:        Call I_StatisticVBAProj.AddSheetStatistica
56:    End If
57: End Sub
    Sub btnHiddenModule(control As IRibbonControl)
59:    Call HiddenModule.Show
60: End Sub
    Private Sub AddInBtn(ByRef control As IRibbonControl)
62:    On Error GoTo ErrorHandler
63:    Application.Dialogs(xlDialogAddinManager).Show
64:    Exit Sub
ErrorHandler:
66:    Err.Clear
67:    Call MsgBox("No open" & Chr(34) & "Excel Files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Mistake:")
68: End Sub
    Private Sub VBABtn(ByRef control As IRibbonControl)
70:    Call VBAVBEOpen
71: End Sub
    Private Sub BtnExportVBA(control As IRibbonControl)
73:    Call VBAVBEOpen
74:    Call ModuleCommander.Show
75: End Sub
    Public Sub VBAVBEOpen()
77:    If C_PublicFunctions.Num_Not_Stable Then Call SendKeys("%{NUMLOCK}")
78:    Call SendKeys("%{F11}")
79: End Sub
    Private Sub BtnVSC(control As IRibbonControl)
81:    Call VersionSistemControls.Show
82: End Sub
    Private Sub onSwitcherReferenceStyle(ByRef control As IRibbonControl)
84:    With Application
85:        If .ReferenceStyle = xlR1C1 Then
86:            .ReferenceStyle = xlA1
87:        Else
88:            .ReferenceStyle = xlR1C1
89:        End If
90:    End With
91: End Sub
    Private Sub onOpenFileExcel(ByRef control As IRibbonControl)
93:    Call O_XML.OpenAndCloseExcelFileInFolder(bOpenFile:=True, bBackUp:=False)
94: End Sub
    Private Sub onCloseFileExcel(ByRef control As IRibbonControl)
96:    Call O_XML.OpenAndCloseExcelFileInFolder(bOpenFile:=False, bBackUp:=True)
97: End Sub
     Private Sub onUnProtectVBA(ByRef control As IRibbonControl)
99:    Call P_UnProtected.unprotected
100: End Sub
     Private Sub onUnProtectSheets(ByRef control As IRibbonControl)
102:    On Error GoTo ErrorHandler
103:    Call ProtectedSheets.Show
104:    Exit Sub
ErrorHandler:
106:    Select Case Err.Number
        Case 91:
108:            Call MsgBox("No open" & Chr(34) & "Excel Files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Mistake:")
109:        Case Else:
110:            Call MsgBox("Mistake! in onUnUnProtectSheets" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbExclamation, "Mistake:")
111:            Call WriteErrorLog("onUnUnProtectSheets")
112:    End Select
113:    Err.Clear
114: End Sub
     Private Sub onUnProtectVBAUnivable(control As IRibbonControl)
116:    Call P_UnProtected.DelPasswordVBAProjectUnivable
117: End Sub
     Private Sub onAddShapeStatistic(ByRef control As IRibbonControl)
119:    Call AddShapeStatistic
120: End Sub
     Private Sub onOptions(ByRef control As IRibbonControl)
122:    Call OptionsCodeFormat.Show
123: End Sub
     Private Sub onOptionsComments(ByRef control As IRibbonControl)
125:    Call SettingsAddCommentsProc.Show
126: End Sub
     Private Sub BtnHelpMain(control As IRibbonControl)
128:    Call HelpMainAddin
129: End Sub
     Private Sub BtnMainBuilders(control As IRibbonControl)
131:    Call URLLinks(C_Const.URL_BILD)
132: End Sub
     Private Sub BtnHelpControls(control As IRibbonControl)
134:    Call URLLinks(C_Const.URL_MOVE_CNTR)
135: End Sub
     Private Sub BtnHelpSnippets(control As IRibbonControl)
137:    Call URLLinks(C_Const.URL_STYLE)
138: End Sub
     Private Sub BtnHelpPass(control As IRibbonControl)
140:    Call URLLinks(C_Const.URL_FILE)
141: End Sub
     Private Sub BtnHelpContacts(control As IRibbonControl)
143:    Call URLLinks(C_Const.URL_CONTACT)
144: End Sub
     Private Sub HelpMainAddin()
146:    Call URLLinks(C_Const.URL_ADDIN)
147: End Sub
     Private Sub onInToFile(control As IRibbonControl)
149:    Call Q_InToFile.InToFile
150: End Sub
     Private Sub BtnOrderMacro(control As IRibbonControl)
152:    Call URLLinks(C_Const.URL_CONTACT)
153: End Sub
'Version
     Public Sub getVisible(control As IRibbonControl, ByRef visible)
156:    visible = C_Const.FlagVisible
157: End Sub
'btnVersion
     Public Sub onVisible(control As IRibbonControl)
160:    Call URLLinks(C_Const.URL_DOWNLOAD)
161: End Sub
     Private Sub BtnUnProtectSheetsXML(control As IRibbonControl)
163:    Call DeletePaswortSheets
164: End Sub
     Private Sub onProtectVBAUnivable(control As IRibbonControl)
166:    Call SetPasswordVBAProjectUnviewable
167: End Sub
     Private Sub BtnVK(control As IRibbonControl)
169:    Call URLLinks(C_Const.URL_VK)
170: End Sub
     Private Sub BtnFB(control As IRibbonControl)
172:    Call URLLinks(C_Const.URL_FB)
173: End Sub
'смена темы
     Private Sub onBlackTheme(control As IRibbonControl)
176:    Call V_BlackAndWiteTheme.ChangeColorDarkTheme
177: End Sub
     Private Sub onWhiteTheme(control As IRibbonControl)
179:    Call V_BlackAndWiteTheme.ChangeColorWhiteTheme
180: End Sub
     Private Sub onToolCharMonitor(control As IRibbonControl)
182:    Call CharsMonitor.Show
183: End Sub
'Регулярные выражения
     Private Sub onTestRegExp(control As IRibbonControl)
186:    Call W_RegExp.AddSheetTestRegExp
187: End Sub
     Private Sub onTempleteRegExp(control As IRibbonControl)
189:    Call RegExpTemplateManager.Show
190: End Sub
     Private Sub onRegExpFunValNumber(control As IRibbonControl)
192:    ActiveCell.FormulaR1C1 = "=РЕГВЫР_ПОЛУЧЗНАЧПОНОМЕРУ()"
193:    Call FunctionWizardShowExc
194: End Sub
     Private Sub onExpFunCount(control As IRibbonControl)
196:    ActiveCell.FormulaR1C1 = "=РЕГВЫР_СЧЁТ()"
197:    Call FunctionWizardShowExc
198: End Sub
     Private Sub onRegExpFunTest(control As IRibbonControl)
200:    ActiveCell.FormulaR1C1 = "=РЕГВЫР_ТЕСТ()"
201:    Call FunctionWizardShowExc
202: End Sub
     Private Sub onRegExpFunReplace(control As IRibbonControl)
204:    ActiveCell.FormulaR1C1 = "=РЕГВЫР_ЗАМЕНИТЬ()"
205:    Call FunctionWizardShowExc
206: End Sub
     Private Sub FunctionWizardShowExc()
208:    If Application.Dialogs(xlDialogFunctionWizard).Show = False Then
209:        ActiveCell.Clear
210:    End If
211:    Calculate
212: End Sub
     Private Sub onParserVBA(control As IRibbonControl)
214:    If VBAIsTrusted Then
215:        Call N_ObfParserVBA.StartParser
216:    End If
217: End Sub
     Private Sub onObfuscator(ByRef control As IRibbonControl)
219:    If VBAIsTrusted Then
220:        Call N_ObfMainNew.StartObfuscation
221:    End If
222: End Sub
     Private Sub onFormatsDel(control As IRibbonControl)
224:    If VBAIsTrusted Then
225:        Call ObfuscationCode.Show
226:    End If
227: End Sub
     Private Sub BtnInfoFile(control As IRibbonControl)
229:    Call InfoFile.Show
230: End Sub
     Private Sub ParserStrings(control As IRibbonControl)
232:    Call ZA_ParserString.ParserStringWB
233: End Sub

     Private Sub ReNameParserString(control As IRibbonControl)
236:    Call ZA_ParserString.ReNameStr
237: End Sub
     Private Sub onDeleteAllLinks(control As IRibbonControl)
239:    Call ZB_DeleteLinksFile.deleteAllLinksInFile
240: End Sub
     Private Sub onAddListLinks(control As IRibbonControl)
242:    Call ZB_DeleteLinksFile.getListAllLinksInFile
243: End Sub
Private Sub onDeleteLinksOnList(control As IRibbonControl)
245:    Call ZB_DeleteLinksFile.deleteLinksOnList
End Sub
