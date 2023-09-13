Attribute VB_Name = "N_ObfParserVBA"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : N_ObfParserVBA - VBA-Code-Parser
'* Created    : 08-10-2020 14:12
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 25-04-2023 10:20    CalDymos
'* Updated    : 13-09-2023 13:33    CalDymos            Parser functions changed / added

Option Explicit
Option Private Module

Private objCollUnical As New Collection
Private Const CHR_TO As String = "|XX|"

Private Type obfModule
    objName         As Scripting.Dictionary
    objNameGlobVar  As Scripting.Dictionary
    objContr        As Scripting.Dictionary
    objCtrlProps    As Scripting.Dictionary
    objSubFun       As Scripting.Dictionary
    objDimVar       As Scripting.Dictionary
    objTypeEnum     As Scripting.Dictionary
    objAPI          As Scripting.Dictionary
    objStringCode   As Scripting.Dictionary
End Type

Public Sub StartParser()
          Dim Form        As AddStatistic
          Dim sNameWB     As String
          Dim objWB       As Object

1         On Error GoTo ErrStartParser
2         Application.Calculation = xlCalculationManual
3         Set Form = New AddStatistic
4         With Form
5             .Caption = "Code base data collection:"
6             .lbOK.Caption = "Parse code"
7             .chQuestion.visible = True
8             .chQuestion2.visible = True 'added: 25.04.2023
9             .chQuestion.Value = True
10            .chQuestion2.Value = True 'added: 25.04.2023
11            .chQuestion.Caption = "Collect string values?"
12            .chQuestion2.Caption = "Use safe mode" 'added: 25.04.2023
13            .chQuestion2.ControlTipText = "Excel objects and APIs are excluded" 'changed: 25.04.2023
14            .lbWord.Caption = 1
15            .Show
16            sNameWB = .cmbMain.Value
17        End With
18        If sNameWB = vbNullString Then Exit Sub
19        If sNameWB Like "*.docm" Or sNameWB Like "*.DOCM" Then
              Dim objWrdApp As Object
20            Set objWrdApp = GetObject(, "Word.Application")
21            Set objWB = objWrdApp.Documents(sNameWB)
22        Else
23            Set objWB = Workbooks(sNameWB)
24        End If

25        Call MainObfParser(objWB, Form.chQuestion.Value, Form.chQuestion2.Value)
26        Set Form = Nothing
27        Application.Calculation = xlCalculationAutomatic
28        Exit Sub
ErrStartParser:
29        Application.Calculation = xlCalculationAutomatic
30        Application.ScreenUpdating = True
31        Call MsgBox("Error in N_ObfParserVBA.StartParser" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
32        Call WriteErrorLog("N_ObfParserVBA.StartParser")
End Sub


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : DelegateMainObfParser - Delegate Procedure for N_ObfMainNew.StartCompleteObfuscation
'* Created    : 25-04-2023 08:03
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):                                     Description
'*
'* ByRef objWB As Object                        :
'* Optional bEncodeStr As Boolean = False       :
'* Optional bSafeMode As Boolean = False        :
'* Optional bNoFinishMessage As Boolean = False :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub DelegateMainObfParser(ByRef objWB As Object, Optional bEncodeStr As Boolean = False, Optional bSafeMode As Boolean = False, Optional bNoFinishMessage As Boolean = False)
33        Call MainObfParser(objWB, bEncodeStr, bSafeMode, bNoFinishMessage)
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 25-04-2023 08:11    CalDymos            Safe mode option added
'* Updated    : 25-04-2023 08:11    CalDymos            Added option to disable finished message

Private Sub MainObfParser(ByRef objWB As Object, Optional bEncodeStr As Boolean = False, Optional bSafeMode As Boolean = False, Optional bNoFinishMessage As Boolean = False)
34        If objWB.VBProject.Protection = vbext_pp_locked Then
35            Call MsgBox("The project is protected, remove the password!", vbCritical, "The project is protected:")
36        Else
37            Call ParserProjectVBA(objWB, bEncodeStr, bSafeMode)
38            If Not bNoFinishMessage Then
39                Call MsgBox("Book code [" & objWB.Name & "] assembled!", vbInformation, "Code parsing:")
40            End If
41        End If
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserProjectVBA - The main code parser, collects module names and assigns ciphers to them
'* Created    : 27-03-2020 13:21
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef objWB As Workbook : selected / active Workbook
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 25-04-2023 08:09    CalDymos            Safe mode option added
'* Updated    : 30-05-2023 08:32    CalDymos            Del Worksheets NAME_SH and NAME_SH_STR before parsing
'* Updated    : 07-09-2023 08:16    CalDymos            modify for String Encryption
'* Updated    : 13-09-2023 13:35    CalDymos

Private Sub ParserProjectVBA(ByRef objWB As Object, Optional bEncodeStr As Boolean = False, Optional bSafeMode As Boolean = False)
          Dim objVBComp   As VBIDE.VBComponent
          Dim varModule   As obfModule
          Dim i           As Long
          Dim k           As Long
          Dim objDict     As Scripting.Dictionary
          Dim objTmpModuleName As Scripting.Dictionary
          Dim asSplitKey() As String
          Dim z1 As Integer
          Dim z2 As Integer
          
          'del old Data
42        If Not IsArrayEmpty(AddCtrlProp()) Then
43            For i = 0 To UBound(AddCtrlProp())
44                Set AddCtrlProp(i) = Nothing
45            Next
46            Erase AddCtrlProp()
47        End If
          
48        If bEncodeStr Then
49            If Not IsArrayEmpty(CryptFunc()) Then
50                For i = 0 To UBound(CryptFunc())
51                    Set CryptFunc(i) = Nothing
52                Next
53                Erase CryptFunc()
54            End If
55            If Not IsArrayEmpty(CryptKey()) Then
56                For i = 0 To UBound(CryptKey())
57                    Set CryptKey(i) = Nothing
58                Next
59                Erase CryptKey()
60            End If
              'Store function for string decryption in array
              'Here you must define the encryption function.
              'This should be individual, i.e. customized by the respective user.
61            ReDim CryptFunc(0)
62            Set CryptFunc(0) = New CAddProc
63            CryptFunc(0).BehavProcExists = enumBehavProcExistIgnoreCode
64            CryptFunc(0).Name = "MACROTools_DeCryptStr"
65            CryptFunc(0).AddCodeLine "Public Function " & CryptFunc(0).Name & "(ByVal Inp As String, sKey As String) As String"
66            CryptFunc(0).AddCodeLine "Dim strZ As String"
67            CryptFunc(0).AddCodeLine "Dim iCount As Integer, iPos As Integer"
68            CryptFunc(0).AddCodeLine "Dim cptZahl As Long, orgZahl As Long"
69            CryptFunc(0).AddCodeLine "Dim keyZahl As Long, cptString As String"
70            CryptFunc(0).AddCodeLine "For iCount = 1 To Len(Inp)"
71            CryptFunc(0).AddCodeLine "iPos = iPos + 1"
72            CryptFunc(0).AddCodeLine "If iPos > Len(sKey) Then iPos = 1"
73            CryptFunc(0).AddCodeLine "keyZahl = Asc(Mid(sKey, iPos, 1))"
74            CryptFunc(0).AddCodeLine "If iCount > Len(Inp) \ 2 Then Exit For"
75            CryptFunc(0).AddCodeLine "cptZahl = CByte(Chr$(38) & Chr$(72) & Mid$(Inp, iCount * 2 - 1, 2))"
76            CryptFunc(0).AddCodeLine "orgZahl = cptZahl Xor keyZahl"
77            CryptFunc(0).AddCodeLine "strZ = strZ & Chr$(orgZahl)"
78            CryptFunc(0).AddCodeLine "Next iCount"
79            CryptFunc(0).AddCodeLine CryptFunc(0).Name & " = strZ"
80            CryptFunc(0).AddCodeLine "End Function"
              
              'Generate and store random key
81            ReDim CryptKey(0)
82            Set CryptKey(0) = New CAddGlobalVar
83            CryptKey(0).IsConstant = True
84            CryptKey(0).DateType = enumDataTypeString
85            CryptKey(0).Name = "MACROTools_DeCryptKey"
86            CryptKey(0).Value = GenerateKey
87            CryptKey(0).Visibility = enumVisibilityPublic
88        End If
          'Delete old worksheets, otherwise they will be captured / parsed as well
89        DelWorksheet NAME_SH, objWB
90        DelWorksheet NAME_SH_STR, objWB
91        DelWorksheet NAME_SH_CTL, objWB
          
92        With varModule
              'главный парсер
93            Set .objName = AddNewDictionary(.objName)
94            Set .objDimVar = AddNewDictionary(.objDimVar)
95            Set .objSubFun = AddNewDictionary(.objSubFun)
96            Set .objContr = AddNewDictionary(.objContr)
97            Set .objCtrlProps = AddNewDictionary(.objCtrlProps)
98            Set .objTypeEnum = AddNewDictionary(.objTypeEnum)
99            Set .objNameGlobVar = AddNewDictionary(.objNameGlobVar)
100           Set .objStringCode = AddNewDictionary(.objStringCode)
101           Set .objAPI = AddNewDictionary(.objAPI)
              

              'Collect all control properties that contain text and all position properties
102           For Each objVBComp In objWB.VBProject.VBComponents
103               Call ParserPropertiesControlsForm(objVBComp.Name, objVBComp, AddCtrlProp(), .objCtrlProps)
104           Next
              
105           If bEncodeStr Then
                  'temporarily collect all module names
                  'to randomly place the MACROTools_DeCryptStr() function,
                  'which is needed to decrypt the strings, in a module.
106               Set objTmpModuleName = AddNewDictionary(objTmpModuleName)
                  
107               For Each objVBComp In objWB.VBProject.VBComponents
108                   If objVBComp.Type = vbext_ct_StdModule Then
109                       If Not objTmpModuleName.Exists(objVBComp.Name) Then objTmpModuleName.Add objVBComp.Name, 1
110                   End If
111               Next objVBComp

                  'If no module exist
112               If objTmpModuleName.Count = 0 Then objTmpModuleName.Add "Modul1", 1
                  
                  
113               Randomize Timer
114               z1 = Int(Rnd * objTmpModuleName.Count) 'Random number for the selection of the module for the function
115               z2 = Int(Rnd * objTmpModuleName.Count) 'Random number for the selection of the module for the key
                  
116               CryptFunc(0).ModuleName = objTmpModuleName.Keys(z1)
117               CryptKey(0).ModuleName = objTmpModuleName.Keys(z2)
                              
118           End If
119           For Each objVBComp In objWB.VBProject.VBComponents
                  'Collect module names
                  Dim skey As String
120               skey = objVBComp.Type & CHR_TO & objVBComp.Name
121               If Not .objName.Exists(skey) Then .objName.Add skey, 0
                  
                  'Collecting all controls in the forms
122               Call ParserNameControlsForm(objVBComp.Name, objVBComp, .objContr)
                  
                  'Capture procedures and functions
123               Call ParserNameSubFunc(objVBComp.Name, objVBComp, .objSubFun)
124               Call ParserNameSubFuncFromAddProc(objVBComp.Name, objVBComp, .objSubFun, AddCtrlProp())
125               If bEncodeStr Then Call ParserNameSubFuncFromAddProc(objVBComp.Name, objVBComp, .objSubFun, CryptFunc())
                  
                  'Capture of global variables
126               Call ParserNameGlobalVariable(objVBComp.Name, objVBComp, .objNameGlobVar, .objTypeEnum, .objAPI)
127               If bEncodeStr Then Call ParserNameGlobalVariableFromAddVar(objVBComp.Name, objVBComp, .objNameGlobVar, .objTypeEnum, .objAPI, CryptKey())
                  
                  'Collect variables in procedures and functions
128               Call ParserVariebleSubFunc(objVBComp, .objDimVar, .objStringCode)
129               Call ParserVariebleSubFuncFromAddProc(objVBComp, .objDimVar, .objStringCode, AddCtrlProp())
130               If bEncodeStr Then Call ParserVariebleSubFuncFromAddProc(objVBComp, .objDimVar, .objStringCode, CryptFunc())
                  
131           Next objVBComp
              'конец парсера
132       End With

          'Create a Sheet in the active workbook
133       Call AddSheetInWBook(NAME_SH, objWB)

134       ReDim arrRange(1 To varModule.objName.Count + varModule.objNameGlobVar.Count + varModule.objSubFun.Count + varModule.objContr.Count + varModule.objDimVar.Count + varModule.objTypeEnum.Count + varModule.objAPI.Count, 1 To 10) As String

135       Set objDict = New Scripting.Dictionary
          
          'Set comparison to insensitive, since var and Sub names in VBA are also insensitive,
          'i.e. no matter whether the var name is written in upper or lower case, it is the same var.
136       objDict.CompareMode = TextCompare 'added 04.09.2023 CalDymos

137       For i = 1 To varModule.objName.Count
138           arrRange(i, 1) = "Module"
139           arrRange(i, 2) = VBA.Split(varModule.objName.Keys(i - 1), CHR_TO)(0)
140           arrRange(i, 3) = VBA.Split(varModule.objName.Keys(i - 1), CHR_TO)(1)
141           arrRange(i, 4) = "Public"
142           arrRange(i, 8) = arrRange(i, 3)
143           arrRange(i, 9) = "yes"

144           If objDict.Exists(arrRange(i, 8)) = False Then
145               objDict.Add arrRange(i, 8), AddEncodeName()
146           End If
147           arrRange(i, 10) = objDict.Item(arrRange(i, 8))
148       Next i
149       k = i
150       Application.StatusBar = "Data collection: Module names, completed:" & VBA.Format(1 / 7, "Percent")
151       For i = 1 To varModule.objNameGlobVar.Count
152           arrRange(k, 1) = "Global variable"
153           arrRange(k, 2) = varModule.objNameGlobVar.Items(i - 1)
154           arrRange(k, 3) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(0)
155           arrRange(k, 4) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(1)
156           arrRange(k, 6) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(2)
157           arrRange(k, 7) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(3)
158           arrRange(k, 8) = arrRange(k, 7)
159           arrRange(k, 9) = "yes"

160           If objDict.Exists(arrRange(k, 8)) = False Then
161               objDict.Add arrRange(k, 8), AddEncodeName()
162           End If
163           arrRange(k, 10) = objDict.Item(arrRange(k, 8))
              
              'store the cipher for the constant that contains the key
164           If bEncodeStr Then
165               If arrRange(k, 1) <> "Global variable" Then
166               ElseIf arrRange(k, 2) <> 1 Then
167               ElseIf arrRange(k, 6) <> "Const" Then
168               ElseIf arrRange(k, 7) = CryptKey(0).Name Then
169                   CryptKey(0).CipherName = arrRange(k, 10)
170               End If
171           End If
172           k = k + 1
173       Next i

174       Application.StatusBar = "Data collection: Global variables, completed:" & VBA.Format(2 / 7, "Percent")
175       For i = 1 To varModule.objSubFun.Count
176           arrRange(k, 1) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(1)
177           arrRange(k, 2) = varModule.objSubFun.Items(i - 1)
178           arrRange(k, 3) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(0)
179           arrRange(k, 4) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(2)
180           arrRange(k, 5) = arrRange(k, 1)
181           arrRange(k, 6) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(3)
182           arrRange(k, 8) = arrRange(k, 6)
183           arrRange(k, 9) = "yes"

184           If objDict.Exists(arrRange(k, 8)) = False Then
185               objDict.Add arrRange(k, 8), AddEncodeName()
186           End If
187           arrRange(k, 10) = objDict.Item(arrRange(k, 8))
              
              'store the cipher for the function for the string encryption
188           If bEncodeStr Then
189               If arrRange(k, 1) <> "Sub" And arrRange(k, 1) <> "Function" And Left$(arrRange(k, 1), 8) <> "Property" Then
190               ElseIf arrRange(k, 2) <> 1 Then
191               ElseIf arrRange(k, 6) = GetNameSubFromString(CStr(CryptFunc(0).GetCodeLine(0))) Then
192                   CryptFunc(0).CipherName = arrRange(k, 10)
193               End If
194           End If
              
195           k = k + 1
196       Next i

197       Application.StatusBar = "Data collection: Procedure names, completed:" & VBA.Format(3 / 7, "Percent")
198       For i = 1 To varModule.objContr.Count
199           arrRange(k, 1) = "Control"
200           arrRange(k, 2) = varModule.objContr.Items(i - 1)
201           arrRange(k, 3) = VBA.Split(varModule.objContr.Keys(i - 1), CHR_TO)(0)
202           arrRange(k, 4) = "Private"
203           arrRange(k, 6) = VBA.Split(varModule.objContr.Keys(i - 1), CHR_TO)(1)
204           arrRange(k, 8) = arrRange(k, 6)
205           arrRange(k, 9) = "yes"

206           If objDict.Exists(arrRange(k, 8)) = False Then
207               objDict.Add arrRange(k, 8), AddEncodeName()
208           End If
209           arrRange(k, 10) = objDict.Item(arrRange(k, 8))
210           k = k + 1
211       Next i

212       Application.StatusBar = "Data collection: Names of controls, completed:" & VBA.Format(4 / 7, "Percent")
213       For i = 1 To varModule.objDimVar.Count
214           arrRange(k, 1) = "Variable"
215           arrRange(k, 2) = varModule.objDimVar.Items(i - 1)
216           arrRange(k, 3) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(0)
217           arrRange(k, 4) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(3)
218           arrRange(k, 5) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(1)
219           arrRange(k, 6) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(2)
220           arrRange(k, 7) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(4)
221           arrRange(k, 8) = arrRange(k, 7)
222           arrRange(k, 9) = "yes"

223           If objDict.Exists(arrRange(k, 8)) = False Then
224               objDict.Add arrRange(k, 8), AddEncodeName()
225           End If
226           arrRange(k, 10) = objDict.Item(arrRange(k, 8))
227           k = k + 1
228           If i Mod 50 = 0 Then
229               Application.StatusBar = "Data collection: Names of controls, completed:" & VBA.Format(i / varModule.objDimVar.Count, "Percent")
230               DoEvents
231           End If
232       Next i

233       Application.StatusBar = "Data collection: Variable names, completed:" & VBA.Format(5 / 7, "Percent")
234       For i = 1 To varModule.objTypeEnum.Count
235           arrRange(k, 1) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(2)
236           arrRange(k, 2) = varModule.objTypeEnum.Items(i - 1)
237           arrRange(k, 3) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(0)
238           arrRange(k, 4) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(1)
239           arrRange(k, 6) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(3)
240           arrRange(k, 8) = arrRange(k, 6)
241           arrRange(k, 9) = "yes"

242           If objDict.Exists(arrRange(k, 8)) = False Then
243               objDict.Add arrRange(k, 8), AddEncodeName()
244           End If
245           arrRange(k, 10) = objDict.Item(arrRange(k, 8))
246           k = k + 1
247       Next i

248       Application.StatusBar = "Data collection: Names of enumerations and types, completed:" & VBA.Format(6 / 7, "Percent")
249       For i = 1 To varModule.objAPI.Count
250           arrRange(k, 1) = "API"
251           arrRange(k, 2) = varModule.objAPI.Items(i - 1)
252           arrRange(k, 3) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(0)
253           arrRange(k, 4) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(1)
254           arrRange(k, 5) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(2)
255           arrRange(k, 6) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(3)
256           arrRange(k, 8) = arrRange(k, 6)
257           arrRange(k, 9) = "yes"

258           If objDict.Exists(arrRange(k, 8)) = False Then
259               objDict.Add arrRange(k, 8), AddEncodeName()
260           End If
261           arrRange(k, 10) = objDict.Item(arrRange(k, 8))
262           k = k + 1
263       Next i
264       Application.StatusBar = "Data collection: API names, completed:" & VBA.Format(7 / 7, "Percent")

265       With ActiveSheet
266           Application.StatusBar = "Application of formats"
267           .Cells.ClearContents
268           .Cells(1, 1).Value = "Type"
269           .Cells(1, 2).Value = "Module type"
270           .Cells(1, 3).Value = "Module name"
271           .Cells(1, 4).Value = "Access Modifiers"
272           .Cells(1, 5).Value = "Percentage type. and funk."
273           .Cells(1, 6).Value = "The name of the percentage. and funk."
274           .Cells(1, 7).Value = "Name of variables"
275           .Cells(1, 8).Value = "Encryption Object"
276           .Cells(1, 9).Value = "Encrypt yes/No"
277           .Cells(1, 10).Value = "Code"
278           .Cells(1, 11).Value = "Mistakes"

279           .Cells(2, 1).Resize(UBound(arrRange), 10) = arrRange

280           .Range(.Cells(2, 11), .Cells(k, 11)).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3]," & SHSNIPPETS.ListObjects(C_Const.TB_SERVICEWORDS).DataBodyRange.Address(ReferenceStyle:=xlR1C1, External:=True) & ",1,0),"""")"
281           .Range(.Cells(2, 9), .Cells(k, 9)).FormulaR1C1 = "=IF(RC[2]="""",""yes"",""no"")"
282           .Columns("A:K").AutoFilter
283           .Columns("A:K").EntireColumn.AutoFit
284           .Range(Cells(2, 9), Cells(UBound(arrRange) + 1, 9)).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="YES, NO"
              
              '
285           If bSafeMode Then
286               For i = 2 To UBound(arrRange) + 1
                      'exclude Excel Objects
287                   If .Cells(i, 2).Value = "100" Then
                          'Debug.Print .Cells(i, 5).Value
288                       If .Cells(i, 1).Value = "Module" And _
                              .Cells(i, 5).Value = "" Then 'changed 31.08 CalDymos
289                           .Cells(i, 9).Value = "NO"
290                       End If
                          'exclude API
291                   ElseIf .Cells(i, 1).Value = "API" Then
292                       .Cells(i, 9).Value = "NO"
293                   End If
294               Next i
295           End If
296           Application.StatusBar = "Application of formats, finished"
297       End With

          'Load String Vars
298       If bEncodeStr Then
299           Call AddSheetInWBook(NAME_SH_STR, objWB)
300           Application.StatusBar = "Collecting String variables"
301           If varModule.objStringCode.Count <> 0 Then
302               ReDim arrRange(1 To varModule.objStringCode.Count, 1 To 8) As String
303               For i = 1 To varModule.objStringCode.Count
304                   asSplitKey() = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)
305                   arrRange(i, 1) = varModule.objStringCode.Items(i - 1)
306                   arrRange(i, 2) = asSplitKey(0)
307                   arrRange(i, 3) = asSplitKey(1)
308                   arrRange(i, 4) = asSplitKey(2)
309                   arrRange(i, 5) = asSplitKey(3)
310                   arrRange(i, 6) = asSplitKey(4)
311                   arrRange(i, 7) = LCase$(asSplitKey(5)) ' changed 13.09
312                   arrRange(i, 8) = AddEncodeName()

313                   If i Mod 50 = 0 Then
314                       Application.StatusBar = "Collecting String variables, completed:" & VBA.Format(i / varModule.objStringCode.Count, "Percent")
315                       DoEvents
316                   End If
317               Next i
318               Application.StatusBar = "Collecting String variables, completed"
319               With ActiveSheet
320                   .Cells(1, 1).Value = "Module type"
321                   .Cells(1, 2).Value = "Module name"
322                   .Cells(1, 3).Value = "Type Sub or Fun"
323                   .Cells(1, 4).Value = "Name Sub or Fun"
324                   .Cells(1, 5).Value = "Line"
325                   .Cells(1, 6).Value = "Array Strings"
326                   .Cells(1, 7).Value = "Encrypt yes/No"
327                   .Cells(1, 8).Value = "Code"
328                   .Cells(1, 9).Value = "Module cipher"
                      
329                   .Cells(1, 11).Value = "The cipher of the Const module"
330                   .Cells(2, 11).Value = AddEncodeName()
                      
                      'Add additional information for string encryption
331                   .Cells(1, 12).Value = "The name of the Key constant"
332                   .Cells(2, 12).Value = CryptKey(0).Name
333                   .Cells(1, 13).Value = "The Key value"
334                   .Cells(2, 13).Value = CryptKey(0).Value ' Key for string encryption
335                   .Cells(1, 14).Value = "The module for the key"
336                   .Cells(2, 14).Value = CryptKey(0).ModuleName
337                   .Cells(1, 15).Value = "The cipher of Key constant"
338                   .Cells(2, 15).Value = CryptKey(0).CipherName
339                   .Cells(1, 16).Value = "Name of Crypt Func"
340                   .Cells(2, 16).Value = CryptFunc(0).Name
341                   .Cells(1, 17).Value = "The module for die Crypt Func"
342                   .Cells(2, 17).Value = CryptFunc(0).ModuleName
343                   .Cells(1, 18).Value = "The cipher of Crypt Func"
344                   .Cells(2, 18).Value = CryptFunc(0).CipherName

345                   .Cells(2, 1).Resize(UBound(arrRange), 8) = arrRange

346                   .Range(Cells(2, 7), Cells(UBound(arrRange) + 1, 7)).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="YES, NO"
347                   .Range(Cells(2, 9), Cells(UBound(arrRange) + 1, 9)).FormulaR1C1 = "=IF(RC1*1=100,RC2,VLOOKUP(RC2,DATA_OBF_VBATools!R2C3:R" & k & "C10,8,0))"
348                   .Columns("A:I").AutoFilter
349                   .Columns("A:D").EntireColumn.AutoFit
350                   .Columns("E").ColumnWidth = 60
351                   .Columns("F:S").EntireColumn.AutoFit
352                   .Columns("A:S").HorizontalAlignment = xlCenter
353                   .Rows("2:" & UBound(arrRange) + 1).RowHeight = 12
354               End With
355           End If
356       End If
          
          'Loading the control properties
357       Call AddSheetInWBook(NAME_SH_CTL, objWB)
358       Application.StatusBar = "Collecting Ctrl Properties"
359       If varModule.objStringCode.Count <> 0 Then
360           ReDim arrRange(1 To varModule.objCtrlProps.Count, 1 To 5) As String
361           For i = 1 To varModule.objCtrlProps.Count
362               asSplitKey() = VBA.Split(varModule.objCtrlProps.Keys(i - 1), CHR_TO)
363               arrRange(i, 1) = varModule.objCtrlProps.Items(i - 1)
364               arrRange(i, 2) = asSplitKey(0)
365               arrRange(i, 3) = asSplitKey(1)
366               arrRange(i, 4) = asSplitKey(2)
367               arrRange(i, 5) = asSplitKey(3)

368               If i Mod 50 = 0 Then
369                   Application.StatusBar = "Collecting Ctrl Properties, completed:" & VBA.Format(i / varModule.objStringCode.Count, "Percent")
370                   DoEvents
371               End If
372           Next i
373           Application.StatusBar = "Collecting String variables, completed"
374           With objWB.ActiveSheet
375               .Cells(1, 1).Value = "Control Type"
376               .Cells(1, 2).Value = "Module name"
377               .Cells(1, 3).Value = "Name of Control"
378               .Cells(1, 4).Value = "Name of Property"
379               .Cells(1, 5).Value = "Value of Property"
                  
380               .Cells(2, 1).Resize(UBound(arrRange), 5) = arrRange
381               .Columns("A:E").NumberFormat = "@"
382               .Columns("A:E").AutoFilter
383               .Columns("A:E").EntireColumn.AutoFit
384               .Columns("A:E").HorizontalAlignment = xlCenter
385               .Rows("2:" & UBound(arrRange) + 1).RowHeight = 12
386           End With
387       End If
              
388       objWB.Worksheets(NAME_SH).Activate

389       Application.StatusBar = False
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddShhetInWBook - Creating a Sheet in active Workbook
'* Created    : 22-03-2023 16:15
'* Author     : VBATools
'* Contacts   : https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal WSheetName As String :
'* ByRef wb As Workbook       :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddSheetInWBook(ByVal WSheetName As String, ByRef wb As Workbook)
          'Creating a Sheet in active Workbook
390       Application.DisplayAlerts = False
391       On Error Resume Next
392       wb.Worksheets(WSheetName).Delete
393       On Error GoTo 0
394       Application.DisplayAlerts = True
395       wb.Sheets.Add Before:=ActiveSheet
396       ActiveSheet.Name = WSheetName
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserVariebleSubFunc - Sammlung von Variablen aus VBA-Code
'* Created    : 22-03-2023 16:16
'* Author     : VBATools
'* Contacts   : https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objVBC As VBIDE.VBComponent             :
'* ByRef objDic As Scripting.Dictionary          :
'* ByRef objDicStr As Scripting.Dictionary       :
'* Optional varAddProc As Variant = Nothing :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 10:19    CalDymos
'* Updated    : 13-09-2023 13:38    CalDymos
Private Sub ParserVariebleSubFunc(ByRef objVBC As VBIDE.VBComponent, ByRef objDic As Scripting.Dictionary, ByRef objDicStr As Scripting.Dictionary)
          Dim lLine       As Long
          Dim sCode       As String
          Dim sVar        As String
          Dim sSubName    As String
          Dim sNumTypeName As String
          Dim sType       As String
          Dim arrStrCode  As Variant
          Dim arrEnum     As Variant
          Dim itemArr     As Variant
          Dim itemVar     As Variant
          Dim arrVar      As Variant
          Dim i As Long
          Dim k As Long
          
397       With objVBC.CodeModule
398           lLine = .CountOfLines
399           If lLine > 0 Then
400               sCode = .Lines(1, lLine)
401               If sCode <> vbNullString Then
                      'remove the line breaks.
402                   sCode = VBA.Replace(sCode, " _" & vbNewLine, vbNullString)
403                   arrStrCode = VBA.Split(sCode, vbNewLine)
404                   For Each itemArr In arrStrCode
405                       itemArr = C_PublicFunctions.TrimSpace(itemArr)
406                       If itemArr <> vbNullString And VBA.Left$(itemArr, 1) <> "'" Then
407                           sVar = vbNullString
                              'If the code contains a comment, delete it.
408                           itemArr = DeleteCommentString(itemArr)
                              'Extract from the declaration clause and determine what is included in the process
409                           If (itemArr Like "* Sub *(*)*" Or itemArr Like "* Function *(*)*" Or itemArr Like "* Property Let *(*)*" Or itemArr Like "* Property Set *(*)*" Or itemArr Like "* Property Get *(*)*" Or _
                                  itemArr Like "Sub *(*)*" Or itemArr Like "Function *(*)*" Or itemArr Like "Property Let *(*)*" Or itemArr Like "Property Set *(*)*" Or itemArr Like "Property Get *(*)*") _
                                  And (Not itemArr Like "*As IRibbonControl*" And Not itemArr Like "* Declare *(*)*") Then

410                               sSubName = TypeProcedure(VBA.CStr(itemArr))
411                               sSubName = sSubName & CHR_TO & GetNameSubFromString(itemArr)
412                               sVar = ParserStrDimConst(itemArr, sSubName, .Name)

413                           End If
                              'If in enumeration or in data type
414                           If itemArr Like "Private Enum *" Or itemArr Like "Public Enum *" Or itemArr Like "Enum *" Or itemArr Like "Private Type *" Or itemArr Like "Public Type *" Or itemArr Like "Type *" Then
415                               arrEnum = VBA.Split(itemArr, " ")
416                               If VBA.CStr(itemArr) Like "Private *" Then
417                                   sNumTypeName = "Private"
418                               Else
419                                   sNumTypeName = "Public"
420                               End If
421                               sNumTypeName = arrEnum(UBound(arrEnum)) & CHR_TO & sNumTypeName
422                               If itemArr Like "* Enum *" Or itemArr Like "Enum *" Then
423                                   sType = "Enum"
424                               Else
425                                   sType = "Type"
426                               End If
427                           End If
                              'go out of the process or enumeration
428                           If itemArr Like "*End Sub" Or itemArr Like "*End Function" Or itemArr Like "*End Property" Or itemArr Like "*End Enum" Or itemArr Like "*End Type" Then
429                               sSubName = vbNullString
430                               sNumTypeName = vbNullString
431                           End If
                              'If within a type or enumeration
432                           If sNumTypeName <> vbNullString And Not itemArr Like "* Enum *" And Not itemArr Like "Enum *" And Not itemArr Like "* Type *" And Not itemArr Like "Type *" Then
433                               arrEnum = VBA.Split(VBA.Trim$(itemArr), " ")
434                               sVar = arrEnum(0)
435                               If sVar Like "*(*" Then sVar = VBA.Left$(sVar, VBA.InStr(1, sVar, "(") - 1)
436                               sVar = .Name & CHR_TO & sType & CHR_TO & sNumTypeName & CHR_TO & ReplaceType(sVar)
437                           End If
                              'when we are only inside the procedure
438                           If (itemArr Like "* Dim *" Or itemArr Like "* Const *" Or itemArr Like "Dim *" Or itemArr Like "Const *") And sSubName <> vbNullString Then
439                               sVar = ParserStrDimConst(itemArr, sSubName, .Name)
440                           End If
441                           arrVar = VBA.Split(sVar, vbNewLine)
442                           For Each itemVar In arrVar
443                               If itemVar <> vbNullString And objDic.Exists(itemVar) = False Then
444                                   objDic.Add itemVar, objVBC.Type
445                               End If
446                           Next itemVar
447                           Call ParserStringInCode(itemArr, sSubName, objVBC, objDicStr)
448                       End If
449                   Next itemArr
450               End If
451           End If
452       End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserVariebleSubFuncFromAddProc
'* Created    : 13-09-2023 13:40
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):                             Description
'*
'* ByRef objVBC As VBIDE.VBComponent       :
'* ByRef objDic As Scripting.Dictionary    :
'* ByRef objDicStr As Scripting.Dictionary :
'* AddProcs(                               :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ParserVariebleSubFuncFromAddProc(ByRef objVBC As VBIDE.VBComponent, ByRef objDic As Scripting.Dictionary, ByRef objDicStr As Scripting.Dictionary, AddProcs() As CAddProc)
          Dim lLine       As Long
          Dim sCode       As String
          Dim sVar        As String
          Dim sSubName    As String
          Dim sNumTypeName As String
          Dim sType       As String
          Dim arrStrCode  As Variant
          Dim arrEnum     As Variant
          Dim itemArr     As Variant
          Dim itemVar     As Variant
          Dim arrVar      As Variant
          Dim i As Long
          Dim k As Long
          
          
453       With objVBC.CodeModule
454           If Not IsArrayEmpty(AddProcs()) Then
455               For i = 0 To UBound(AddProcs())
456                   If AddProcs(i).ModuleName = .Name Then
457                       For Each itemArr In AddProcs(i).CodeLines
458                           itemArr = C_PublicFunctions.TrimSpace(itemArr)
459                           If itemArr <> vbNullString And VBA.Left$(itemArr, 1) <> "'" Then
460                               sVar = vbNullString
                                  'If the code contains a comment, delete it.
461                               itemArr = DeleteCommentString(itemArr)
                                  'aus der Deklarationsklausel entnehmen und feststellen, was in das Verfahren aufgenommen wurde
462                               If (itemArr Like "* Sub *(*)*" Or itemArr Like "* Function *(*)*" Or itemArr Like "* Property Let *(*)*" Or itemArr Like "* Property Set *(*)*" Or itemArr Like "* Property Get *(*)*" Or _
                                      itemArr Like "Sub *(*)*" Or itemArr Like "Function *(*)*" Or itemArr Like "Property Let *(*)*" Or itemArr Like "Property Set *(*)*" Or itemArr Like "Property Get *(*)*") _
                                      And (Not itemArr Like "*As IRibbonControl*" And Not itemArr Like "* Declare *(*)*") Then

463                                   sSubName = TypeProcedure(VBA.CStr(itemArr))
464                                   sSubName = sSubName & CHR_TO & GetNameSubFromString(itemArr)
465                                   sVar = ParserStrDimConst(itemArr, sSubName, .Name)

466                               End If
                                  'Wenn in der Aufzдhlung und im Datentyp
467                               If itemArr Like "Private Enum *" Or itemArr Like "Public Enum *" Or itemArr Like "Enum *" Or itemArr Like "Private Type *" Or itemArr Like "Public Type *" Or itemArr Like "Type *" Then
468                                   arrEnum = VBA.Split(itemArr, " ")
469                                   If VBA.CStr(itemArr) Like "Private *" Then
470                                       sNumTypeName = "Private"
471                                   Else
472                                       sNumTypeName = "Public"
473                                   End If
474                                   sNumTypeName = arrEnum(UBound(arrEnum)) & CHR_TO & sNumTypeName
475                                   If itemArr Like "* Enum *" Or itemArr Like "Enum *" Then
476                                       sType = "Enum"
477                                   Else
478                                       sType = "Type"
479                                   End If
480                               End If
                                  'aus dem Prozess oder der Aufzдhlung herausgehen
481                               If itemArr Like "*End Sub" Or itemArr Like "*End Function" Or itemArr Like "*End Property" Or itemArr Like "*End Enum" Or itemArr Like "*End Type" Then
482                                   sSubName = vbNullString
483                                   sNumTypeName = vbNullString
484                               End If
                                  'Falls innerhalb des Typs oder der Aufzдhlung
485                               If sNumTypeName <> vbNullString And Not itemArr Like "* Enum *" And Not itemArr Like "Enum *" And Not itemArr Like "* Type *" And Not itemArr Like "Type *" Then
486                                   arrEnum = VBA.Split(VBA.Trim$(itemArr), " ")
487                                   sVar = arrEnum(0)
488                                   If sVar Like "*(*" Then sVar = VBA.Left$(sVar, VBA.InStr(1, sVar, "(") - 1)
489                                   sVar = .Name & CHR_TO & sType & CHR_TO & sNumTypeName & CHR_TO & ReplaceType(sVar)
490                               End If
                                  'wenn wir uns nur innerhalb der Prozedur befinden
491                               If (itemArr Like "* Dim *" Or itemArr Like "* Const *" Or itemArr Like "Dim *" Or itemArr Like "Const *") And sSubName <> vbNullString Then
492                                   sVar = ParserStrDimConst(itemArr, sSubName, .Name)
493                               End If
494                               arrVar = VBA.Split(sVar, vbNewLine)
495                               For Each itemVar In arrVar
496                                   If itemVar <> vbNullString And objDic.Exists(itemVar) = False Then
497                                       objDic.Add itemVar, objVBC.Type
498                                   End If
499                               Next itemVar
500                               Call ParserStringInCode(itemArr, sSubName, objVBC, objDicStr)
501                           End If
502                       Next itemArr
503                   End If
504               Next
505           End If
506       End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : GetNameSubFromString - Get the procedure name from the string
'* Created    : 20-04-2020 18:19
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal sStrCode As String : строка
'*
'* Modified   : Date and Time       Author              Description
'* Updated    : 13-09-2023 13:41    CalDymos
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function GetNameSubFromString(ByVal sStrCode As String) As String
          Dim sTemp       As String
          Dim Pos1 As Long
          
507       Pos1 = VBA.InStr(1, sStrCode, "(")
          
508       If Pos1 <> 0 Then
509           sTemp = VBA.Trim$(VBA.Left$(sStrCode, Pos1 - 1))
510           Select Case True
                  Case sTemp Like "*Sub *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Sub ") - 3)
511               Case sTemp Like "*Function *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Function ") - 8)
512               Case sTemp Like "*Property Let *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Let ") - 12)
513               Case sTemp Like "*Property Set *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Set ") - 12)
514               Case sTemp Like "*Property Get *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Get ") - 12)
515           End Select
516       End If
517       GetNameSubFromString = VBA.Trim$(sTemp)
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserStringInCode - Sammlung von String-Konstanten aus dem Code
'* Created    : 22-03-2023 16:18
'* Author     : VBATools
'* Contacts   : https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByVal sSTR As String                    :
'* ByVal sNameSub As String                :
'* ByRef objVBC As VBIDE.VBComponent       :
'* ByRef objDicStr As Scripting.Dictionary :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 13-09-2023 13:41    CalDymos
'* Updated    : 31-08-2023 08:07    CalDymos
Private Sub ParserStringInCode(ByVal sSTR As String, ByVal sNameSub As String, ByRef objVBC As VBIDE.VBComponent, ByRef objDicStr As Scripting.Dictionary)
          Dim sTxt        As String
          Dim arrStr      As Variant
          Dim arr         As Variant
          Dim sReplace    As String
          Dim i           As Integer
          Dim sArray      As String
          Dim sYesNo      As String
          Const CHAR_REPLACE As String = "XXXXX"
          

518       sSTR = VBA.Trim$(sSTR)

519       If sSTR Like "*" & VBA.Chr$(34) & "*" And _
              sSTR <> vbNullString And _
              Not sSTR Like "*Declare * Lib *(*)*" Then

520           If sSTR Like "*Const *" Or sSTR Like "Const *" Then 'changed 13.09 - Constants must not be encrypted
521               sYesNo = "No"
522           Else
523               sYesNo = "Yes"
524           End If
525           sTxt = VBA.Right$(sSTR, VBA.Len(sSTR) - VBA.InStr(1, sSTR, VBA.Chr$(34)) + 1)
526           sTxt = VBA.Replace(sTxt, VBA.Chr$(34) & VBA.Chr$(34), CHAR_REPLACE)
527           arrStr = VBA.Split(sTxt, VBA.Chr$(34))

528           sArray = VBA.Left$(sSTR, VBA.InStr(1, sSTR, VBA.Chr$(34)) - 1)
529           If sArray Like "* = Array(" Then
530               sArray = VBA.Replace(sArray, " = Array(", vbNullString)
531               arr = VBA.Split(sArray, " ")
532               sArray = arr(UBound(arr))
533           Else
534               sArray = vbNullString
535           End If
536           For i = 1 To UBound(arrStr) Step 2
537               If arrStr(i) <> vbNullString Then
538                   If sNameSub = vbNullString Then sNameSub = "Declaration" & CHR_TO

539                   sReplace = VBA.Replace(arrStr(i), CHAR_REPLACE, VBA.Chr$(34) & VBA.Chr$(34))
540                   sTxt = objVBC.Name & CHR_TO & sNameSub & CHR_TO & VBA.Chr$(34) & sReplace & VBA.Chr$(34) & CHR_TO & sArray & CHR_TO & sYesNo 'changed 13.09
541                   If arrStr(i + 1) Like "*: * = *" Then sArray = vbNullString
542                   If arrStr(i + 1) Like "*: * = Array(*" Then
543                       sArray = VBA.Replace(arrStr(i + 1), ": ", vbNullString)
544                       sArray = VBA.Replace(sArray, " = Array(", vbNullString)
545                       sArray = VBA.Replace(sArray, ")", vbNullString)
546                   End If
547                   If objDicStr.Exists(sTxt) = False Then objDicStr.Add sTxt, objVBC.Type
548               End If
549           Next i
550           sArray = vbNullString
551       End If
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ParserStrDimConst - String parser for initialization of variables and constants
'* Created    : 14-04-2020 22:45
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sTxt As String : - code line
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function ParserStrDimConst(ByVal sTxt As String, ByVal sNameSub As String, ByVal sNameMod As String) As String
          Dim sTemp       As String
          Dim sWord       As String
          Dim sWordTemp   As String
          Dim arrStr      As Variant
          Dim itemArr     As Variant
          Dim arrWord     As Variant
          Dim sType       As String

552       sTemp = C_PublicFunctions.TrimSpace(sTxt)
553       sType = "Dim"
554       If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
              'If there is a comment in the code string, delete it.
555           sTemp = DeleteCommentString(sTemp)
556           If sTemp Like "*Sub *(*)*" Or sTemp Like "*Function *(*)*" Or sTemp Like "*Property Let *(*)*" Or sTemp Like "*Property Set *(*)*" Or sTemp Like "*Property Get *(*)*" Then
557               If VBA.InStr(1, sTemp, ")") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, ")") - 1)
558               If VBA.InStr(1, sTemp, " = ") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ") - 1)
559               If VBA.Len(sTemp) - VBA.InStr(1, sTemp, "(") >= 0 Then
560                   sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "("))
561               End If
562           ElseIf sTemp Like "* Dim *" Or sTemp Like Chr$(68) & "im *" Then
563               sType = "Dim"
564               If VBA.InStr(1, sTemp, "Dim ") >= 3 Then sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Dim ") - 3)
565           ElseIf sTemp Like "* Const *" Or sTemp Like Chr$(67) & "onst *" Then
566               sType = "Const"
567               If VBA.InStr(1, sTemp, "Const ") >= 5 Then sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Const ") - 5)
568               If VBA.InStr(1, sTemp, " = ") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ") - 1)
569           Else
570               sTemp = vbNullString
571           End If
572       End If

573       If sTemp Like "*: *" Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, ": ") - 1)
574       If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
575           arrStr = VBA.Split(sTemp, ",")
576           For Each itemArr In arrStr
577               If itemArr Like "*(*" Then itemArr = VBA.Left$(itemArr, VBA.InStr(1, itemArr, "(") - 1)
578               If Not itemArr Like "*)*" And Not itemArr Like "* To *" Then
579                   arrWord = VBA.Split(itemArr, " As ")
580                   arrWord = VBA.Split(VBA.Trim$(arrWord(0)), " ")
581                   If UBound(arrWord) = -1 Then
582                       sWord = vbNullString
583                   Else
584                       sWordTemp = VBA.Trim$(arrWord(UBound(arrWord)))
585                       sWordTemp = ReplaceType(sWordTemp)
586                       sWord = sWord & vbNewLine & sNameMod & CHR_TO & sNameSub & CHR_TO & sType & CHR_TO & sWordTemp
587                   End If
588               End If
589           Next itemArr
590       End If
591       sWord = VBA.Trim$(sWord)
592       If VBA.Len(sWord) = 0 Then
593           sWord = vbNullString
594       Else
595           sWord = VBA.Trim$(VBA.Right$(sWord, VBA.Len(sWord) - 2))
596       End If
597       ParserStrDimConst = sWord
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserNameSubFunc - сбор названий процедур и функций
'* Created    : 27-03-2020 13:20
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objCodeModule As VBIDE.CodeModule : объект модуль
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 10:22    CalDymos
'* Updated    : 13-09-2023 13:42    CalDymos

Private Sub ParserNameSubFunc(ByVal sNameVBC As String, ByRef objVBC As VBIDE.VBComponent, ByRef varSubFun As Scripting.Dictionary)
          Dim ProcKind    As VBIDE.vbext_ProcKind
          Dim lLine       As Long
          Dim lineOld     As Long
          Dim sNameSub    As String
          Dim strFunctionBody As String
          Dim skey As String
          
598       With objVBC.CodeModule
599           If .CountOfLines > 0 Then
600               lLine = .CountOfDeclarationLines
601               If lLine = 0 Then lLine = 2
602               Do Until lLine >= .CountOfLines

                      'Sammeln von Namen von Prozeduren und Funktionen
603                   sNameSub = .ProcOfLine(lLine, ProcKind)
604                   If sNameSub <> vbNullString Then
605                       strFunctionBody = C_PublicFunctions.TrimSpace(.Lines(lLine - 1, .ProcCountLines(sNameSub, ProcKind)))
                          'Debug.Print strFunctionBody
                          'Debug.Print .Lines(lLine - 1, .ProcCountLines(sNameSub, ProcKind))
606                       If (Not strFunctionBody Like "*As IRibbonControl*") And _
                              (Not strFunctionBody Like "*As IRibbonUI*") And _
                              (Not WorkBookAndSheetsEvents(strFunctionBody, objVBC.Type)) And _
                              (Not (strFunctionBody Like "* UserForm_*" And objVBC.Type = vbext_ct_MSForm)) And _
                              (Not UserFormsEvents(strFunctionBody, objVBC.Type)) Then
607                           skey = sNameVBC & CHR_TO & TypeProcedure(strFunctionBody) & CHR_TO & TypeOfAccessModifier(strFunctionBody) & CHR_TO & sNameSub
                              'Debug.Print skey
608                           If Not varSubFun.Exists(skey) Then
609                               varSubFun.Add skey, objVBC.Type
610                           End If
611                       End If
612                       lLine = .ProcStartLine(sNameSub, ProcKind) + .ProcCountLines(sNameSub, ProcKind) + 1
613                   Else
614                       lLine = lLine + 1
615                   End If
616                   If lineOld > lLine Then Exit Do
617                   lineOld = lLine
618               Loop
619           End If
620       End With
            
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserNameSubFuncFromAddProc
'* Created    : 13-09-2023 13:42
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):                             Description
'*
'* ByVal sNameVBC As String                :
'* ByRef objVBC As VBIDE.VBComponent       :
'* ByRef varSubFun As Scripting.Dictionary :
'* AddProcs(                               :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ParserNameSubFuncFromAddProc(ByVal sNameVBC As String, ByRef objVBC As VBIDE.VBComponent, ByRef varSubFun As Scripting.Dictionary, AddProcs() As CAddProc)
          Dim skey As String
          Dim i As Long
          
             
621       If Not IsArrayEmpty(AddProcs()) Then
622           For i = 0 To UBound(AddProcs())
623               If AddProcs(i).ModuleName = sNameVBC Then
624                   If Not UserFormsEvents(AddProcs(i).GetCodeLine(0), objVBC.Type) Then
625                       skey = sNameVBC & CHR_TO & TypeProcedure(CStr(AddProcs(i).GetCodeLine(0))) & CHR_TO & _
                              TypeOfAccessModifier(CStr(AddProcs(i).GetCodeLine(0))) & CHR_TO & AddProcs(i).Name
                          'Debug.Print skey
626                       If Not varSubFun.Exists(skey) Then
627                           varSubFun.Add skey, objVBC.Type
628                       End If
629                   End If
630               End If
631           Next
632       End If
        
End Sub

Private Sub ParserNameControlsForm(ByVal sNameVBC As String, ByRef objVBC As VBIDE.VBComponent, ByRef obfNewDict As Scripting.Dictionary)
          Dim objCont     As MSForms.control
633       If Not objVBC.Designer Is Nothing Then
634           With objVBC.Designer
635               For Each objCont In .Controls
                      'Debug.Print sNameVBC & CHR_TO & objCont.Name, objVBC.Type
636                   obfNewDict.Add sNameVBC & CHR_TO & objCont.Name, objVBC.Type
637               Next objCont
638           End With
639       End If
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserPropertiesControlsForm
'* Created    : 13-09-2023 13:42
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):                         Description
'*
'* ByVal sNameVBC As String          :
'* ByRef objVBC As VBIDE.VBComponent :
'* ByRef varAddCtrlProp(             :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ParserPropertiesControlsForm(ByVal sNameVBC As String, ByRef objVBC As VBIDE.VBComponent, ByRef varAddCtrlProp() As CAddProc, ByRef obfNewDict As Scripting.Dictionary)
          Dim objCtl     As MSForms.control
          Dim objTxtBox As MSForms.TextBox
          Dim objLbl As MSForms.Label
          Dim objCmdBtn As MSForms.CommandButton
          Dim objFrame As MSForms.Frame
          Dim objChkBox As MSForms.CheckBox
          Dim objOptBtn As MSForms.OptionButton
          Dim objTglBtn As MSForms.ToggleButton

          Dim i1 As Long
          Dim i2 As Long

640       If Not objVBC.Designer Is Nothing And objVBC.Type = vbext_ct_MSForm Then
641           If Not IsArrayEmpty(varAddCtrlProp()) Then
642               i1 = UBound(varAddCtrlProp()) + 1
643               ReDim Preserve varAddCtrlProp(i1)
644           Else
645               ReDim varAddCtrlProp(0)
646               i1 = 0
647           End If
648           Set varAddCtrlProp(i1) = New CAddProc
              
649           i2 = 0
                    
650           varAddCtrlProp(i1).ModuleName = sNameVBC
651           varAddCtrlProp(i1).BehavProcExists = enumBehavProcExistInsCodeAtBegin
652           varAddCtrlProp(i1).Name = "UserForm_Initialize"
653           varAddCtrlProp(i1).AddCodeLine "Private Sub " & varAddCtrlProp(i1).Name & "()"
654           i2 = i2 + 1

655           If objVBC.Properties("Caption") <> "" Then
656               varAddCtrlProp(i1).AddCodeLine "Me.Caption = " & Chr$(34) & objVBC.Properties("Caption") & Chr$(34)
657               i2 = i2 + 1
658               obfNewDict.Add sNameVBC & CHR_TO & "" & CHR_TO & "Caption" & CHR_TO & objVBC.Properties("Caption"), "UserForm"
659           End If
660           If objVBC.Properties("Tag") <> "" Then
661               varAddCtrlProp(i1).AddCodeLine "Me.Tag= " & Chr$(34) & objVBC.Properties("Tag") & Chr$(34)
662               i2 = i2 + 1
663               obfNewDict.Add sNameVBC & CHR_TO & "" & CHR_TO & "Tag" & CHR_TO & objVBC.Properties("Tag"), "UserForm"
664           End If
665           With objVBC.Designer
666               For Each objCtl In .Controls
667                   If TypeOf objCtl Is MSForms.TextBox Then
668                       Set objTxtBox = objCtl
669                       If objTxtBox.Text <> "" Then
670                           varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".Text = " & Chr$(34) & objTxtBox.Text & Chr$(34)
671                           i2 = i2 + 1
672                           obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "Text" & CHR_TO & objTxtBox.Text, TypeName(objCtl)
673                       End If
674                   ElseIf TypeOf objCtl Is MSForms.Label Then
675                       Set objLbl = objCtl
676                       If objLbl.Caption <> "" Then
677                           varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".Caption = " & Chr$(34) & objLbl.Caption & Chr$(34)
678                           i2 = i2 + 1
679                           obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "Caption" & CHR_TO & objLbl.Caption, TypeName(objCtl)
680                       End If
681                   ElseIf TypeOf objCtl Is MSForms.CommandButton Then
682                       Set objCmdBtn = objCtl
683                       If objCmdBtn.Caption <> "" Then
684                           varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".Caption = " & Chr$(34) & objCmdBtn.Caption & Chr$(34)
685                           i2 = i2 + 1
686                           obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "Caption" & CHR_TO & objCmdBtn.Caption, TypeName(objCtl)
687                       End If
688                   ElseIf TypeOf objCtl Is MSForms.Frame Then
689                       Set objFrame = objCtl
690                       If objFrame.Caption <> "" Then
691                           varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".Caption = " & Chr$(34) & objFrame.Caption & Chr$(34)
692                           i2 = i2 + 1
693                           obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "Caption" & CHR_TO & objFrame.Caption, TypeName(objCtl)
694                       End If
695                   ElseIf TypeOf objCtl Is MSForms.CheckBox Then
696                       Set objChkBox = objCtl
697                       If objChkBox.Caption <> "" Then
698                           varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".Caption = " & Chr$(34) & objChkBox.Caption & Chr$(34)
699                           i2 = i2 + 1
700                           obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "Caption" & CHR_TO & objChkBox.Caption, TypeName(objCtl)
701                       End If
702                   ElseIf TypeOf objCtl Is MSForms.OptionButton Then
703                       Set objOptBtn = objCtl
704                       If objOptBtn.Caption <> "" Then
705                           varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".Caption = " & Chr$(34) & objOptBtn.Caption & Chr$(34)
706                           i2 = i2 + 1
707                           obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "Caption" & CHR_TO & objOptBtn.Caption, TypeName(objCtl)
708                       End If
709                   ElseIf TypeOf objCtl Is MSForms.ToggleButton Then
710                       Set objTglBtn = objCtl
711                       If objTglBtn.Caption <> "" Then
712                           varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".Caption = " & Chr$(34) & objTglBtn.Caption & Chr$(34)
713                           i2 = i2 + 1
714                           obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "Caption" & CHR_TO & objTglBtn.Caption, TypeName(objCtl)
715                       End If
716                   End If
717                   If objCtl.Tag <> "" Then
718                       varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".Tag = " & Chr$(34) & objCtl.Tag & Chr$(34)
719                       i2 = i2 + 1
720                       obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "Tag" & CHR_TO & objCtl.Tag, TypeName(objCtl)
721                   End If
722                   If objCtl.ControlTipText <> "" Then
723                       varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".ControlTipText = " & Chr$(34) & objCtl.ControlTipText & Chr$(34)
724                       i2 = i2 + 1
725                       obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "ControlTipText" & CHR_TO & objCtl.ControlTipText, TypeName(objCtl)
726                   End If
727                   If objCtl.Height > 0 Then
728                       varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".Height = " & Replace$(CStr(objCtl.Height), ",", ".")
729                       i2 = i2 + 1
730                       obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "Height" & CHR_TO & objCtl.Height, TypeName(objCtl)
731                   End If
732                   If objCtl.Width > 0 Then
733                       varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".Width = " & Replace$(CStr(objCtl.Width), ",", ".")
734                       i2 = i2 + 1
735                       obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "Width" & CHR_TO & objCtl.Width, TypeName(objCtl)
736                   End If
737                   If objCtl.Left > 0 Then
738                       varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".Left = " & Replace$(CStr(objCtl.Left), ",", ".")
739                       i2 = i2 + 1
740                       obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "Left" & CHR_TO & objCtl.Left, TypeName(objCtl)
741                   End If
742                   If objCtl.top > 0 Then
743                       varAddCtrlProp(i1).AddCodeLine "Me." & objCtl.Name & ".Top = " & Replace$(CStr(objCtl.top), ",", ".")
744                       i2 = i2 + 1
745                       obfNewDict.Add sNameVBC & CHR_TO & objCtl.Name & CHR_TO & "Top" & CHR_TO & objCtl.top, TypeName(objCtl)
746                   End If
                          
747               Next objCtl
748               varAddCtrlProp(i1).AddCodeLine "End Sub"
749               varAddCtrlProp(i1).SetCodeLinesSize (i2)
750           End With
751       End If
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserNameGlobalVariable - сбор глобальных переменных
'* Created    : 27-03-2020 15:38
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                         Description
'*
'* ByVal sDeclarationLines As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 10:23    CalDymos
'* Updated    : 13-09-2023 13:43    CalDymos

Private Sub ParserNameGlobalVariable(ByVal sNameVBC As String, ByRef objVBC As VBIDE.VBComponent, ByRef dicGloblVar As Scripting.Dictionary, ByRef dicTypeEnum As Scripting.Dictionary, ByRef dicAPI As Scripting.Dictionary)
          Dim varArr      As Variant
          Dim varArrWord  As Variant
          Dim varStr      As Variant
          Dim itemVarStr  As Variant
          Dim varAPI      As Variant
          Dim sTemp       As String
          Dim sTempArr    As String
          Dim i           As Long
          Dim bFlag       As Boolean
          Dim j           As Byte
          Dim itemArr     As Byte
          
752       bFlag = True
753       If objVBC.CodeModule.CountOfDeclarationLines <> 0 Then
754           sTemp = objVBC.CodeModule.Lines(1, objVBC.CodeModule.CountOfDeclarationLines)
755           sTemp = VBA.Replace(sTemp, " _" & vbNewLine, vbNullString)
756           If sTemp <> vbNullString Then
757               varArr = VBA.Split(sTemp, vbNewLine)
758               For i = 0 To UBound(varArr)
759                   sTemp = C_PublicFunctions.TrimSpace(DeleteCommentString(varArr(i)))
760                   If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
761                       If sTemp Like "* Type *" Or sTemp Like "* Enum *" Or sTemp Like "Type *" Or sTemp Like "Enum *" Then
762                           varArrWord = VBA.Split(sTemp, " ")
763                           If UBound(varArrWord) = 2 Then
764                               sTemp = VBA.Trim$(varArrWord(0)) & CHR_TO & VBA.Trim$(varArrWord(1)) & CHR_TO & VBA.Trim$(varArrWord(2))
765                           ElseIf UBound(varArrWord) = 1 Then
766                               sTemp = "Public" & CHR_TO & VBA.Trim$(varArrWord(0)) & CHR_TO & VBA.Trim$(varArrWord(1))
767                           End If
768                           sTemp = sNameVBC & CHR_TO & sTemp
                              'Debug.Print sTemp
769                           If Not dicTypeEnum.Exists(sTemp) Then dicTypeEnum.Add sTemp, objVBC.Type
770                           bFlag = False
771                       End If
772                       If bFlag And Not (sTemp Like "Implements *" Or sTemp Like "Option *" Or VBA.Left$(sTemp, 1) = "'" Or sTemp = vbNullString Or VBA.Left$(sTemp, 1) = "#" Or sTemp Like "*Declare *(*)*" Or sTemp Like "*Event *(*)") Then

773                           If sTemp Like "* = *" Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ", vbTextCompare) + 2)
774                           If sTemp Like "* *(* To *) *" Then
775                               sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, "(", vbTextCompare) - 1)
776                           End If
777                           varStr = VBA.Split(sTemp, ",")
778                           For Each itemVarStr In varStr
779                               sTemp = VBA.Trim$(itemVarStr)
780                               varArrWord = VBA.Split(sTemp, " As ")
781                               varArrWord = VBA.Split(varArrWord(0), " = ")
782                               sTemp = varArrWord(0)
783                               varArrWord = VBA.Split(sTemp, " ")

784                               j = UBound(varArrWord)
785                               If j > 1 Then
786                                   If varArrWord(0) = "Dim" Or varArrWord(0) = "Const" Then
787                                       sTemp = "Private" & CHR_TO & varArrWord(0) & CHR_TO
788                                       sTempArr = varArrWord(1)
789                                   ElseIf (varArrWord(0) = "Private" Or varArrWord(0) = "Public") And (varArrWord(1) = "Dim" Or varArrWord(1) = "Const" Or varArrWord(1) = "WithEvents") Then
790                                       sTemp = varArrWord(0) & CHR_TO & varArrWord(1) & CHR_TO
791                                       sTempArr = varArrWord(2)
792                                   ElseIf (varArrWord(0) = "Private" Or varArrWord(0) = "Public") And Not (varArrWord(1) = "Dim" Or varArrWord(1) = "Const" Or varArrWord(1) = "WithEvents") Then
793                                       sTemp = varArrWord(0) & CHR_TO & "Dim" & CHR_TO
794                                       sTempArr = varArrWord(1)
795                                   End If
796                               ElseIf j = 1 And varArrWord(0) = "Global" Then
797                                   sTemp = "Public" & CHR_TO & varArrWord(0) & CHR_TO
798                                   sTempArr = varArrWord(1)
799                               ElseIf j = 1 And (varArrWord(0) = "Private" Or varArrWord(0) = "Public") Then
800                                   sTemp = varArrWord(0) & CHR_TO & "Dim" & CHR_TO
801                                   sTempArr = varArrWord(1)
802                               ElseIf j = 1 And (varArrWord(0) = "Dim" Or varArrWord(0) = "Const") Then
803                                   sTemp = "Private" & CHR_TO & varArrWord(0) & CHR_TO
804                                   sTempArr = varArrWord(1)
805                               ElseIf j = 0 Then
806                                   sTemp = "Private" & CHR_TO & " Dim" & CHR_TO
807                                   sTempArr = varArrWord(0)
808                               End If

809                               sTempArr = ReplaceType(sTempArr)
810                               If sTempArr Like "*(*" Then sTempArr = VBA.Left$(sTempArr, VBA.InStr(1, sTempArr, "(") - 1)
811                               sTemp = sNameVBC & CHR_TO & sTemp & sTempArr
                                  'Debug.Print sTemp
812                               If Not dicGloblVar.Exists(sTemp) Then dicGloblVar.Add sTemp, objVBC.Type

813                               sTemp = vbNullString
814                           Next itemVarStr
815                           sTemp = vbNullString
816                       End If
817                       If sTemp Like "*End Type" Or sTemp Like "*End Enum" Then
818                           bFlag = True
819                       End If
820                       If sTemp Like "*Declare * Lib " & VBA.Chr$(34) & "*" & VBA.Chr$(34) & " (*)*" Then
821                           sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " Lib ", vbTextCompare) - 1)
822                           varAPI = VBA.Split(sTemp, VBA.Chr$(32))
823                           itemArr = UBound(varAPI)
824                           sTemp = CHR_TO & varAPI(itemArr - 1) & CHR_TO & varAPI(itemArr)
825                           If varAPI(1) = "Declare" Then
826                               sTemp = sNameVBC & CHR_TO & varAPI(0) & sTemp
827                           Else
828                               sTemp = sNameVBC & CHR_TO & "Private" & sTemp
829                           End If
830                           If Not dicAPI.Exists(sTemp) Then dicAPI.Add sTemp, objVBC.Type
831                       End If
832                       If sTemp Like "*Event *(*)" Then
833                           sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, "(", vbTextCompare) - 1)
834                           varAPI = VBA.Split(sTemp, VBA.Chr$(32))
835                           itemArr = UBound(varAPI)
836                           sTemp = CHR_TO & varAPI(itemArr - 1) & CHR_TO & varAPI(itemArr)
837                           If varAPI(1) = "Event" Then
838                               sTemp = sNameVBC & CHR_TO & varAPI(0) & sTemp
839                           Else
840                               sTemp = sNameVBC & CHR_TO & "Private" & sTemp
841                           End If
                              'Debug.Print sTemp
842                           If Not dicAPI.Exists(sTemp) Then dicAPI.Add sTemp, objVBC.Type
843                       End If
844                   End If
845               Next i
846           End If
847       End If
          
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserNameGlobalVariableFromAddVar
'* Created    : 13-09-2023 13:43
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):                                 Description
'*
'* ByVal sNameVBC As String                  :
'* ByRef objVBC As VBIDE.VBComponent         :
'* ByRef dicGloblVar As Scripting.Dictionary :
'* ByRef dicTypeEnum As Scripting.Dictionary :
'* ByRef dicAPI As Scripting.Dictionary      :
'* AddGlobalVars(                            :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ParserNameGlobalVariableFromAddVar(ByVal sNameVBC As String, ByRef objVBC As VBIDE.VBComponent, ByRef dicGloblVar As Scripting.Dictionary, ByRef dicTypeEnum As Scripting.Dictionary, ByRef dicAPI As Scripting.Dictionary, AddGlobalVars() As CAddGlobalVar)
          Dim varArr      As Variant
          Dim varArrWord  As Variant
          Dim varStr      As Variant
          Dim itemVarStr  As Variant
          Dim varAPI      As Variant
          Dim sTemp       As String
          Dim sTempArr    As String
          Dim i           As Long
          Dim bFlag       As Boolean
          Dim j           As Byte
          Dim itemArr     As Byte
              
848       If Not IsArrayEmpty(AddGlobalVars()) Then
849           For i = 0 To UBound(AddGlobalVars())
850               If AddGlobalVars(i).ModuleName = sNameVBC Then
851                   sTemp = sNameVBC & CHR_TO
852                   Select Case AddGlobalVars(i).Visibility
                          Case enumVisibility.enumVisibilityPublic
853                           sTemp = sTemp & "Public" & CHR_TO
854                       Case enumVisibility.enumVisibilityPrivate
855                           sTemp = sTemp & "Private" & CHR_TO
856                   End Select
857                   If AddGlobalVars(i).IsConstant Then
858                       sTemp = sTemp & "Const" & CHR_TO
859                   Else
860                       sTemp = sTemp & "Dim" & CHR_TO
861                   End If
862                   sTemp = sTemp & AddGlobalVars(i).Name
863                   If Not dicGloblVar.Exists(sTemp) Then dicGloblVar.Add sTemp, objVBC.Type
864                   sTemp = vbNullString
865               End If
866           Next
867       End If
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddNewDictionary -функция инициализации словаря
'* Created    : 27-03-2020 13:21
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objDict As Scripting.Dictionary :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function AddNewDictionary(ByRef objDict As Scripting.Dictionary) As Scripting.Dictionary
868       Set objDict = Nothing
869       Set objDict = New Scripting.Dictionary
870       Set AddNewDictionary = objDict
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DeleteCommentString - удаление в строке комментария
'* Created    : 20-04-2020 18:18
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sWord As String : строка
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function DeleteCommentString(ByVal sWord As String) As String
          'есть '
          Dim sTemp       As String
871       sTemp = sWord
872       If VBA.InStr(1, sTemp, "'") <> 0 Then
873           If VBA.InStr(1, sTemp, VBA.Chr(34)) <> 0 Then
                  'есть "
874               If VBA.InStr(1, sTemp, "'") < VBA.InStr(1, sTemp, VBA.Chr(34)) Then
                      'если так -> '"
875                   sTemp = VBA.Trim$(VBA.Left$(sTemp, VBA.InStr(1, sTemp, "'") - 1))
876               End If
877           Else
                  'нет " -> '
878               sTemp = VBA.Trim$(VBA.Left$(sTemp, VBA.InStr(1, sTemp, "'") - 1))
879           End If
880       End If
881       DeleteCommentString = sTemp
End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddEncodeName - Die Funktion der zufдlligen Zuweisung eines zufдllig vergebenen Namens
'* Created    : 27-03-2020 13:22
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function AddEncodeName() As String
          Const CharCount As Integer = 20 ' Possible names = 2^19 = 524288
          Dim i           As Integer
          Dim sName       As String

          Const FIRST_CODE_SIGN As String = "1"
          Const SECOND_CODE_SIGN As String = "0"
tryAgain:
882       Err.Clear
883       sName = vbNullString
884       Randomize
885       sName = "o"
886       For i = 2 To CharCount
887           If (VBA.Round(VBA.Rnd() * 1000)) Mod 2 = 1 Then sName = sName & FIRST_CODE_SIGN Else sName = sName & SECOND_CODE_SIGN
888       Next i
889       On Error Resume Next
          'add a name to the collection, if the name exists, then
          'an error is generated, which restarts the generation of the name
890       objCollUnical.Add sName, sName
891       If Err.Number <> 0 Then GoTo tryAgain
892       AddEncodeName = sName
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GenerateKey
'* Created    : 07-09-2023 10:25
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'This function should or can also be customized depending on the user,
'to increase the security
Private Function GenerateKey() As String

          Dim z1 As Integer
          Dim z2 As Integer
          Dim Z3 As Integer
          Dim i As Integer
          Dim skey As String

893       Randomize Timer

894       For i = 1 To 18
895           z1 = Int(1 + Rnd * (26 - 1 + 1))
              'Debug.Print Z1
896           z2 = Int(Rnd * (9 + 1))
              'Debug.Print Z2
897           Z3 = Int(1 + Rnd * (2 - 1 + 1))
              'Debug.Print Z3

898           If Z3 = 1 Then
899               skey = skey & Chr$(64 + z1)
900           Else
901               skey = skey & CStr(z2)
902           End If

903       Next i
          'Debug.Print skey
904       GenerateKey = skey

End Function

'взято
Private Function TypeOfAccessModifier(ByRef StrDeclarationProcedure As String) As String
905       If StrDeclarationProcedure Like "*Private *(*)*" Then
906           TypeOfAccessModifier = "Private"
907       Else
908           TypeOfAccessModifier = "Public"
909       End If
End Function
Private Function TypeProcedure(ByRef StrDeclarationProcedure As String) As String
910       If StrDeclarationProcedure Like "*Sub *" Then
911           TypeProcedure = "Sub"
912       ElseIf StrDeclarationProcedure Like "*Function *" Then
913           TypeProcedure = "Function"
914       ElseIf StrDeclarationProcedure Like "*Property Set *" Then
915           TypeProcedure = "Property Set"
916       ElseIf StrDeclarationProcedure Like "*Property Get *" Then
917           TypeProcedure = "Property Get"
918       ElseIf StrDeclarationProcedure Like "*Property Let *" Then
919           TypeProcedure = "Property Let"
920       Else
921           TypeProcedure = "Unknown Type"
922       End If
End Function

Private Function ReplaceType(ByVal sVar As String) As String
923       sVar = Replace(sVar, "%", vbNullString)     'Integer
924       sVar = Replace(sVar, "&", vbNullString)     'Long
925       sVar = Replace(sVar, "$", vbNullString)     'String
926       sVar = Replace(sVar, "!", vbNullString)     'Single
927       sVar = Replace(sVar, "#", vbNullString)     'Double
928       sVar = Replace(sVar, "@", vbNullString)     'Currency
929       ReplaceType = sVar
End Function

Private Function WorkBookAndSheetsEvents(ByVal sTxt As String, ByVal TypeModule As VBIDE.vbext_ComponentType) As Boolean
          Dim Flag        As Boolean
930       Flag = False
          'nur fьr Sheets, Workbooks und Klassenmodule
931       If TypeModule = vbext_ct_Document Or TypeModule = vbext_ct_ClassModule Then
932           Select Case True
                  Case sTxt Like "*_Activate(*": Flag = True
933               Case sTxt Like "*_AddinInstall(*": Flag = True
934               Case sTxt Like "*_AddinUninstall(*": Flag = True
935               Case sTxt Like "*_AfterSave(*": Flag = True
936               Case sTxt Like "*_AfterXmlExport(*": Flag = True
937               Case sTxt Like "*_AfterXmlImport(*": Flag = True
938               Case sTxt Like "*_BeforeClose(*": Flag = True
939               Case sTxt Like "*_BeforeDoubleClick(*": Flag = True
940               Case sTxt Like "*_BeforePrint(*": Flag = True
941               Case sTxt Like "*_BeforeRightClick(*": Flag = True
942               Case sTxt Like "*_BeforeSave(*": Flag = True
943               Case sTxt Like "*_BeforeXmlExport(*": Flag = True
944               Case sTxt Like "*_BeforeXmlImport(*": Flag = True
945               Case sTxt Like "*_Calculate(*": Flag = True
946               Case sTxt Like "*_Change(*": Flag = True
947               Case sTxt Like "*_Deactivate(*": Flag = True
948               Case sTxt Like "*_FollowHyperlink(*": Flag = True
949               Case sTxt Like "*_MouseDown(*": Flag = True
950               Case sTxt Like "*_MouseMove(*": Flag = True
951               Case sTxt Like "*_MouseUp(*": Flag = True
952               Case sTxt Like "*_NewChart(*": Flag = True
953               Case sTxt Like "*_NewSheet(*": Flag = True
954               Case sTxt Like "*_Open(*": Flag = True
955               Case sTxt Like "*_PivotTableAfterValueChange(*": Flag = True
956               Case sTxt Like "*_PivotTableBeforeAllocateChanges(*": Flag = True
957               Case sTxt Like "*_PivotTableBeforeCommitChanges(*": Flag = True
958               Case sTxt Like "*_PivotTableBeforeDiscardChanges(*": Flag = True
959               Case sTxt Like "*_PivotTableChangeSync(*": Flag = True
960               Case sTxt Like "*_PivotTableCloseConnection(*": Flag = True
961               Case sTxt Like "*_PivotTableOpenConnection(*": Flag = True
962               Case sTxt Like "*_PivotTableUpdate(*": Flag = True
963               Case sTxt Like "*_Resize(*": Flag = True
964               Case sTxt Like "*_RowsetComplete(*": Flag = True
965               Case sTxt Like "*_SelectionChange(*": Flag = True
966               Case sTxt Like "*_SeriesChange(*": Flag = True
967               Case sTxt Like "*_SheetActivate(*": Flag = True
968               Case sTxt Like "*_SheetBeforeDoubleClick(*": Flag = True
969               Case sTxt Like "*_SheetBeforeRightClick(*": Flag = True
970               Case sTxt Like "*_SheetCalculate(*": Flag = True
971               Case sTxt Like "*_SheetChange(*": Flag = True
972               Case sTxt Like "*_SheetDeactivate(*": Flag = True
973               Case sTxt Like "*_SheetFollowHyperlink(*": Flag = True
974               Case sTxt Like "*_SheetPivotTableAfterValueChange(*": Flag = True
975               Case sTxt Like "*_SheetPivotTableBeforeAllocateChanges(*": Flag = True
976               Case sTxt Like "*_SheetPivotTableBeforeCommitChanges(*": Flag = True
977               Case sTxt Like "*_SheetPivotTableBeforeDiscardChanges(*": Flag = True
978               Case sTxt Like "*_SheetPivotTableChangeSync(*": Flag = True
979               Case sTxt Like "*_SheetPivotTableUpdate(*": Flag = True
980               Case sTxt Like "*_SheetSelectionChange(*": Flag = True
981               Case sTxt Like "*_Sync(*": Flag = True
982               Case sTxt Like "*_WindowActivate(*": Flag = True
983               Case sTxt Like "*_WindowDeactivate(*": Flag = True
984               Case sTxt Like "*_WindowResize(*": Flag = True
985               Case sTxt Like "*_NewWorkbook(*": Flag = True
986               Case sTxt Like "*_WorkbookActivate(*": Flag = True
987               Case sTxt Like "*_WorkbookAddinInstall(*": Flag = True
988               Case sTxt Like "*_WorkbookAddinUninstall(*": Flag = True
989               Case sTxt Like "*_WorkbookAfterSave(*": Flag = True
990               Case sTxt Like "*_WorkbookAfterXmlExport(*": Flag = True
991               Case sTxt Like "*_WorkbookAfterXmlImport(*": Flag = True
992               Case sTxt Like "*_WorkbookBeforeClose(*": Flag = True
993               Case sTxt Like "*_WorkbookBeforePrint(*": Flag = True
994               Case sTxt Like "*_WorkbookBeforeSave(*": Flag = True
995               Case sTxt Like "*_WorkbookBeforeXmlExport(*": Flag = True
996               Case sTxt Like "*_WorkbookBeforeXmlImport(*": Flag = True
997               Case sTxt Like "*_WorkbookDeactivate(*": Flag = True
998               Case sTxt Like "*_WorkbookModelChange(*": Flag = True
999               Case sTxt Like "*_WorkbookNewChart(*": Flag = True
1000              Case sTxt Like "*_WorkbookNewSheet(*": Flag = True
1001              Case sTxt Like "*_WorkbookOpen(*": Flag = True
1002              Case sTxt Like "*_WorkbookPivotTableCloseConnection(*": Flag = True
1003              Case sTxt Like "*_WorkbookPivotTableOpenConnection(*": Flag = True
1004              Case sTxt Like "*_WorkbookRowsetComplete(*": Flag = True
1005              Case sTxt Like "*_WorkbookSync(*": Flag = True
1006          End Select
1007      End If
1008      WorkBookAndSheetsEvents = Flag
End Function

Private Function UserFormsEvents(ByVal sTxt As String, ByVal TypeModule As VBIDE.vbext_ComponentType) As Boolean
          Dim Flag        As Boolean
1009      Flag = False
          'Nur fьr Events, der Forms und Klassen
1010      If TypeModule = vbext_ct_MSForm Or TypeModule = vbext_ct_ClassModule Then
1011          Select Case True
                  Case sTxt Like "*_AfterUpdate(*": Flag = True
1012              Case sTxt Like "*_BeforeDragOver(*": Flag = True
1013              Case sTxt Like "*_BeforeDropOrPaste(*": Flag = True
1014              Case sTxt Like "*_BeforeUpdate(*": Flag = True
1015              Case sTxt Like "*_Change(*": Flag = True
1016              Case sTxt Like "*_Click(*": Flag = True
1017              Case sTxt Like "*_DblClick(*": Flag = True
1018              Case sTxt Like "*_Deactivate(*": Flag = True
1019              Case sTxt Like "*_DropButtonClick(*": Flag = True
1020              Case sTxt Like "*_Enter(*": Flag = True
1021              Case sTxt Like "*_Error(*": Flag = True
1022              Case sTxt Like "*_Exit(*": Flag = True
1023              Case sTxt Like "*_Initialize(*": Flag = True
1024              Case sTxt Like "*_KeyDown(*": Flag = True
1025              Case sTxt Like "*_KeyPress(*": Flag = True
1026              Case sTxt Like "*_KeyUp(*": Flag = True
1027              Case sTxt Like "*_Layout(*": Flag = True
1028              Case sTxt Like "*_MouseDown(*": Flag = True
1029              Case sTxt Like "*_MouseMove(*": Flag = True
1030              Case sTxt Like "*_MouseUp(*": Flag = True
1031              Case sTxt Like "*_QueryClose(*": Flag = True
1032              Case sTxt Like "*_RemoveControl(*": Flag = True
1033              Case sTxt Like "*_Resize(*": Flag = True
1034              Case sTxt Like "*_Scroll(*": Flag = True
1035              Case sTxt Like "*_Terminate(*": Flag = True
1036              Case sTxt Like "*_Zoom(*": Flag = True
1037          End Select
1038      End If
1039      UserFormsEvents = Flag
End Function

