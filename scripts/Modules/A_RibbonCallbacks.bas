Attribute VB_Name = "A_RibbonCallbacks"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : A_RibbonCallbacks - модуль обратных вызовов ленты управлени€ Excel
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
30:    End If
31: End Sub
    Private Sub ImportCodeBaseBtn(ByRef control As IRibbonControl)
33:    On Error GoTo ErrorHandler
34:    If VBAIsTrusted Then
35:        Workbooks(C_Const.NAME_ADDIN & ".xlam").Sheets(C_Const.SH_SNIPPETS).Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
36:        Call MsgBox("The code base has been uploaded", vbInformation, "Code Base Upload:")
37:    End If
38:    Exit Sub
ErrorHandler:
40:    Select Case Err.Number
        Case 91:
42:            Call MsgBox("No open files" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
43:        Case Else:
44:            Call MsgBox("Error in ImportCodeBaseBtn" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbExclamation, "Error:")
45:            Call WriteErrorLog("ImportCodeBaseBtn")
46:    End Select
47:    Err.Clear
48: End Sub
    Private Sub AddCodeBaseBtn(ByRef control As IRibbonControl)
50:    If VBAIsTrusted Then
51:        Call AddCodeView.Show
52:    End If
53: End Sub
    Private Sub AddStatBtn(ByRef control As IRibbonControl)
55:    If VBAIsTrusted Then
56:        Call I_StatisticVBAProj.AddSheetStatistica
57:    End If
58: End Sub
    Sub btnHiddenModule(control As IRibbonControl)
60:    Call HiddenModule.Show
61: End Sub
    Private Sub AddInBtn(ByRef control As IRibbonControl)
63:    On Error GoTo ErrorHandler
64:    Application.Dialogs(xlDialogAddinManager).Show
65:    Exit Sub
ErrorHandler:
67:    Err.Clear
68:    Call MsgBox("No open files" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
69: End Sub
    Private Sub VBABtn(ByRef control As IRibbonControl)
71:    Call VBAVBEOpen
72: End Sub
    Private Sub BtnExportVBA(control As IRibbonControl)
74:    Call VBAVBEOpen
75:    Call ModuleCommander.Show
76: End Sub
    Public Sub VBAVBEOpen()
78:    If C_PublicFunctions.Num_Not_Stable Then Call SendKeys("%{NUMLOCK}")
79:    Call SendKeys("%{F11}")
80: End Sub
    Private Sub BtnVSC(control As IRibbonControl)
82:    Call VersionSistemControls.Show
83: End Sub
    Private Sub onSwitcherReferenceStyle(ByRef control As IRibbonControl)
85:    With Application
86:        If .ReferenceStyle = xlR1C1 Then
87:            .ReferenceStyle = xlA1
88:        Else
89:            .ReferenceStyle = xlR1C1
90:        End If
91:    End With
92: End Sub
    Private Sub onOpenFileExcel(ByRef control As IRibbonControl)
94:    Call O_XML.OpenAndCloseExcelFileInFolder(bOpenFile:=True, bBackUp:=False)
95: End Sub
    Private Sub onCloseFileExcel(ByRef control As IRibbonControl)
97:    Call O_XML.OpenAndCloseExcelFileInFolder(bOpenFile:=False, bBackUp:=True)
98: End Sub
     Private Sub onUnProtectVBA(ByRef control As IRibbonControl)
100:    Call P_UnProtected.unprotected
101: End Sub
     Private Sub onUnProtectSheets(ByRef control As IRibbonControl)
103:    On Error GoTo ErrorHandler
104:    Call ProtectedSheets.Show
105:    Exit Sub
ErrorHandler:
107:    Select Case Err.Number
        Case 91:
109:            Call MsgBox("No open files" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
110:        Case Else:
111:            Call MsgBox("Error in onUnProtectSheets" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl, vbOKOnly + vbExclamation, "Error:")
112:            Call WriteErrorLog("onUnUnProtectSheets")
113:    End Select
114:    Err.Clear
115: End Sub
     Private Sub onUnProtectVBAUnivable(control As IRibbonControl)
117:    Call P_UnProtected.DelPasswordVBAProjectUnivable
118: End Sub
     Private Sub onAddShapeStatistic(ByRef control As IRibbonControl)
120:    Call AddShapeStatistic
121: End Sub
     Private Sub onOptions(ByRef control As IRibbonControl)
123:    Call OptionsCodeFormat.Show
124: End Sub
     Private Sub onOptionsComments(ByRef control As IRibbonControl)
126:    Call SettingsAddCommentsProc.Show
127: End Sub
     Private Sub BtnHelpMain(control As IRibbonControl)
129:    Call HelpMainAddin
130: End Sub
     Private Sub BtnMainBuilders(control As IRibbonControl)
132:    Call URLLinks(C_Const.URL_BILD)
133: End Sub
     Private Sub BtnHelpControls(control As IRibbonControl)
135:    Call URLLinks(C_Const.URL_MOVE_CNTR)
136: End Sub
     Private Sub BtnHelpSnippets(control As IRibbonControl)
138:    Call URLLinks(C_Const.URL_STYLE)
139: End Sub
     Private Sub BtnHelpPass(control As IRibbonControl)
141:    Call URLLinks(C_Const.URL_FILE)
142: End Sub
     Private Sub BtnHelpContacts(control As IRibbonControl)
144:    Call URLLinks(C_Const.URL_CONTACT)
145: End Sub
     Private Sub HelpMainAddin()
147:    Call URLLinks(C_Const.URL_ADDIN)
148: End Sub
     Private Sub onInToFile(control As IRibbonControl)
150:    Call Q_InToFile.InToFile
151: End Sub
     Private Sub BtnOrderMacro(control As IRibbonControl)
153:    Call URLLinks(C_Const.URL_CONTACT)
154: End Sub
'Version
     Public Sub getVisible(control As IRibbonControl, ByRef visible)
157:    visible = C_Const.FlagVisible
158: End Sub
'btnVersion
     Public Sub onVisible(control As IRibbonControl)
161:    Call URLLinks(C_Const.URL_DOWNLOAD)
162: End Sub
     Private Sub BtnUnProtectSheetsXML(control As IRibbonControl)
164:    Call DeletePaswortSheets
165: End Sub
     Private Sub onProtectVBAUnivable(control As IRibbonControl)
167:    Call SetPasswordVBAProjectUnviewable
168: End Sub
     Private Sub BtnVK(control As IRibbonControl)
170:    Call URLLinks(C_Const.URL_VK)
171: End Sub
     Private Sub BtnFB(control As IRibbonControl)
173:    Call URLLinks(C_Const.URL_FB)
174: End Sub
'смена темы
     Private Sub onBlackTheme(control As IRibbonControl)
177:    Call V_BlackAndWiteTheme.ChangeColorDarkTheme
178: End Sub
     Private Sub onWhiteTheme(control As IRibbonControl)
180:    Call V_BlackAndWiteTheme.ChangeColorWhiteTheme
181: End Sub
     Private Sub onToolCharMonitor(control As IRibbonControl)
183:    Call CharsMonitor.Show
184: End Sub
'–егул€рные выражени€
     Private Sub onTestRegExp(control As IRibbonControl)
187:    Call W_RegExp.AddSheetTestRegExp
188: End Sub
     Private Sub onTempleteRegExp(control As IRibbonControl)
190:    Call RegExpTemplateManager.Show
191: End Sub
     Private Sub onRegExpFunValNumber(control As IRibbonControl)
193:    ActiveCell.FormulaR1C1 = "=REG_GetValueByNumber()"
194:    Call FunctionWizardShowExc
195: End Sub
     Private Sub onExpFunCount(control As IRibbonControl)
197:    ActiveCell.FormulaR1C1 = "=REG_Count()"
198:    Call FunctionWizardShowExc
199: End Sub
     Private Sub onRegExpFunTest(control As IRibbonControl)
201:    ActiveCell.FormulaR1C1 = "=REG_Test()"
202:    Call FunctionWizardShowExc
203: End Sub
     Private Sub onRegExpFunReplace(control As IRibbonControl)
205:    ActiveCell.FormulaR1C1 = "=REG_Replace()"
206:    Call FunctionWizardShowExc
207: End Sub
     Private Sub FunctionWizardShowExc()
209:    If Application.Dialogs(xlDialogFunctionWizard).Show = False Then
210:        ActiveCell.Clear
211:    End If
212:    Calculate
213: End Sub
     Private Sub onParserVBA(control As IRibbonControl)
215:    If VBAIsTrusted Then
216:        Call N_ObfParserVBA.StartParser
217:    End If
218: End Sub
     Private Sub onObfuscator(ByRef control As IRibbonControl)
220:    If VBAIsTrusted Then
221:        Call N_ObfMainNew.StartObfuscation
222:    End If
223: End Sub
     Private Sub onFormatsDel(control As IRibbonControl)
225:    If VBAIsTrusted Then
226:        Call ObfuscationCode.Show
227:    End If
228: End Sub
     Private Sub BtnInfoFile(control As IRibbonControl)
230:    Call InfoFile.Show
231: End Sub
     Sub ParserStrings(control As IRibbonControl)
233:    Call ZA_ParserString.ParserStringWB
234: End Sub

Sub ReNameParserString(control As IRibbonControl)
237:    Call ZA_ParserString.ReNameStr
End Sub
