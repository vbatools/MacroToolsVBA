'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : N_ObfParserVBA - VBA-Code-Parser
'* Created    : 08-10-2020 14:12
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module

Private objCollUnical As New Collection
Private Const CHR_TO As String = "|XX|"

Private Type obfModule
    objName         As Scripting.Dictionary
    objNameGlobVar  As Scripting.Dictionary
    objContr        As Scripting.Dictionary
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

1      On Error GoTo ErrStartParser
2      Application.Calculation = xlCalculationManual
3      Set Form = New AddStatistic
4      With Form
5          .Caption = "Code base data collection:"
6          .lbOK.Caption = "Parse code"
7          .chQuestion.visible = True
8          .chQuestion2.visible = True 'added: 25.04.2023
9          .chQuestion.Value = False
10         .chQuestion2.Value = True 'added: 25.04.2023
11         .chQuestion.Caption = "Collect string values?"
12         .chQuestion2.Caption = "Use safe mode" 'added: 25.04.2023
13         .chQuestion2.ControlTipText = "Excel objects and APIs are excluded" 'changed: 25.04.2023
14         .lbWord.Caption = 1
15         .Show
16         sNameWB = .cmbMain.Value
17     End With
18     If sNameWB = vbNullString Then Exit Sub
19     If sNameWB Like "*.docm" Or sNameWB Like "*.DOCM" Then
           Dim objWrdApp As Object
20         Set objWrdApp = GetObject(, "Word.Application")
21         Set objWB = objWrdApp.Documents(sNameWB)
22     Else
23         Set objWB = Workbooks(sNameWB)
24     End If

25     Call MainObfParser(objWB, Form.chQuestion.Value, Form.chQuestion2.Value)
26     Set Form = Nothing
27     Application.Calculation = xlCalculationAutomatic
28     Exit Sub
ErrStartParser:
29     Application.Calculation = xlCalculationAutomatic
30     Application.ScreenUpdating = True
31     Call MsgBox("Error in N_ObfParserVBA.StartParser" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
32     Call WriteErrorLog("N_ObfParserVBA.StartParser")
34    End Sub


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
44    End Sub

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

Private Sub ParserProjectVBA(ByRef objWB As Object, Optional bEncodeStr As Boolean = False, Optional bSafeMode As Boolean = False)
          Dim objVBComp   As VBIDE.VBComponent
          Dim varModule   As obfModule
          Dim i           As Long
          Dim k           As Long
          Dim objDict     As Scripting.Dictionary
          Dim objTmpName As Scripting.Dictionary
          Dim z1 As Integer
          Dim z2 As Integer
          Dim lSize As Long
          Dim strCryptFuncCipher As String
          Dim strCryptKeyCipher As String
          
          'del old Data
42        If IsArray(varStrCryptFunc) Then varStrCryptFunc = Empty
43        Erase asCryptKey()
          
          'Store function for string decryption in array
          'Here you must define the encryption function.
          'This should be individual, i.e. customized by the user.
44        varStrCryptFunc = Array("Public Function MACROTools_DeCryptStr(ByVal Inp As String, sKey As String) As String", _
              "Dim strEnc as String", _
              "strEnc = Inp", _
              "'code line", _
              "......", _
              "'.....", _
              "'.....", _
              "'code line", _
              "MACROTools_DeCryptStr = strEnc", _
              "End Function")
              
          'Generate and store random key
45        asCryptKey(0) = "MACROTools_DeCryptKey"
46        asCryptKey(1) = GenerateKey
            
          'Delete old worksheets, otherwise they will be captured / parsed as well
47        DelWorksheet NAME_SH, ActiveWorkbook
48        DelWorksheet NAME_SH_STR, ActiveWorkbook
          
49        With varModule
              'ãëàâíûé ïàðñåð
50            Set .objName = AddNewDictionary(.objName)
51            Set .objDimVar = AddNewDictionary(.objDimVar)
52            Set .objSubFun = AddNewDictionary(.objSubFun)
53            Set .objContr = AddNewDictionary(.objContr)
54            Set .objTypeEnum = AddNewDictionary(.objTypeEnum)
55            Set .objNameGlobVar = AddNewDictionary(.objNameGlobVar)
56            Set .objStringCode = AddNewDictionary(.objStringCode)
57            Set .objAPI = AddNewDictionary(.objAPI)

58            If bEncodeStr Then
                  'temporarily collect all module names
                  'to randomly place the MACROTools_DeCryptStr() function,
                  'which is needed to decrypt the strings, in a module.
59                Set objTmpName = AddNewDictionary(objTmpName)
                  
60                For Each objVBComp In objWB.VBProject.VBComponents
61                    If objVBComp.Type = vbext_ct_StdModule Then
62                        If Not objTmpName.Exists(objVBComp.Name) Then objTmpName.Add objVBComp.Name, 1
63                    End If
64                Next objVBComp

                  'If no module exist
65                If objTmpName.Count = 0 Then objTmpName.Add "Modul1", 1
                  
                  
66                Randomize Timer
67                z1 = Int(Rnd * objTmpName.Count) 'Random number for the selection of the module for the function
68                z2 = Int(Rnd * objTmpName.Count) 'Random number for the selection of the module for the key
                  
                              
69            End If
70            For Each objVBComp In objWB.VBProject.VBComponents
                  'Collect module names
                  Dim skey As String
71                skey = objVBComp.Type & CHR_TO & objVBComp.Name
72                If Not .objName.Exists(skey) Then .objName.Add skey, 0
                  'Collecting all controls in the forms
73                Call ParserNameControlsForm(objVBComp.Name, objVBComp, .objContr)
                  
74                If Not bEncodeStr Then
                      'Capture procedures and functions
75                    Call ParserNameSubFunc(objVBComp.Name, objVBComp, .objSubFun)
                      'Capture of global variables
76                    Call ParserNameGlobalVariable(objVBComp.Name, objVBComp, .objNameGlobVar, .objTypeEnum, .objAPI)
                      'Collect variables in procedures and functions
77                    Call ParserVariebleSubFunc(objVBComp, .objDimVar, .objStringCode)
78                Else
                      'Capture procedures and functions
79                    If objVBComp.Name = objTmpName.Keys(z1) Then
80                        Call ParserNameSubFunc(objVBComp.Name, objVBComp, .objSubFun, varStrCryptFunc)
81                    Else
82                        Call ParserNameSubFunc(objVBComp.Name, objVBComp, .objSubFun)
83                    End If
                      'Capture of global variables
84                    If objVBComp.Name = objTmpName.Keys(z2) Then
85                        Call ParserNameGlobalVariable(objVBComp.Name, objVBComp, .objNameGlobVar, .objTypeEnum, .objAPI, asCryptKey())
86                    Else
87                        Call ParserNameGlobalVariable(objVBComp.Name, objVBComp, .objNameGlobVar, .objTypeEnum, .objAPI)
88                    End If
                      'Collect variables in procedures and functions
89                    If objVBComp.Name = objTmpName.Keys(z1) Then
90                        Call ParserVariebleSubFunc(objVBComp, .objDimVar, .objStringCode, varStrCryptFunc)
91                    Else
92                        Call ParserVariebleSubFunc(objVBComp, .objDimVar, .objStringCode)
93                    End If
94                End If
95            Next objVBComp
              'êîíåö ïàðñåðà
96        End With

          'Erstellen einer Liste in der aktiven Arbeitsmappe
97        Call AddSheetInWBook(NAME_SH, ActiveWorkbook)

98        ReDim arrRange(1 To varModule.objName.Count + varModule.objNameGlobVar.Count + varModule.objSubFun.Count + varModule.objContr.Count + varModule.objDimVar.Count + varModule.objTypeEnum.Count + varModule.objAPI.Count, 1 To 10) As String

99        Set objDict = New Scripting.Dictionary
          
          'Set comparison to insensitive, since var and Sub names in VBA are also insensitive,
          'i.e. no matter whether the var name is written in upper or lower case, it is the same var.
100       objDict.CompareMode = TextCompare 'added 04.09.2023 CalDymos

101       For i = 1 To varModule.objName.Count
102           arrRange(i, 1) = "Module"
103           arrRange(i, 2) = VBA.Split(varModule.objName.Keys(i - 1), CHR_TO)(0)
104           arrRange(i, 3) = VBA.Split(varModule.objName.Keys(i - 1), CHR_TO)(1)
105           arrRange(i, 4) = "Public"
106           arrRange(i, 8) = arrRange(i, 3)
107           arrRange(i, 9) = "yes"

108           If objDict.Exists(arrRange(i, 8)) = False Then
109               objDict.Add arrRange(i, 8), AddEncodeName()
110           End If
111           arrRange(i, 10) = objDict.Item(arrRange(i, 8))
112       Next i
113       k = i
114       Application.StatusBar = "Data collection: Module names, completed:" & VBA.Format(1 / 7, "Percent")
115       For i = 1 To varModule.objNameGlobVar.Count
116           arrRange(k, 1) = "Global variable"
117           arrRange(k, 2) = varModule.objNameGlobVar.Items(i - 1)
118           arrRange(k, 3) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(0)
119           arrRange(k, 4) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(1)
120           arrRange(k, 6) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(2)
121           arrRange(k, 7) = VBA.Split(varModule.objNameGlobVar.Keys(i - 1), CHR_TO)(3)
122           arrRange(k, 8) = arrRange(k, 7)
123           arrRange(k, 9) = "yes"

124           If objDict.Exists(arrRange(k, 8)) = False Then
125               objDict.Add arrRange(k, 8), AddEncodeName()
126           End If
127           arrRange(k, 10) = objDict.Item(arrRange(k, 8))
              
              'store temporarily the cipher for the constant that contains the key
128           If bEncodeStr Then
129               If arrRange(k, 1) <> "Global variable" Then
130               ElseIf arrRange(k, 2) <> 1 Then
131               ElseIf arrRange(k, 6) <> "Const" Then
132               ElseIf arrRange(k, 7) = asCryptKey(0) Then
133                   strCryptKeyCipher = arrRange(k, 10)
134               End If
135           End If
136           k = k + 1
137       Next i

138       Application.StatusBar = "Data collection: Global variables, completed:" & VBA.Format(2 / 7, "Percent")
139       For i = 1 To varModule.objSubFun.Count
140           arrRange(k, 1) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(1)
141           arrRange(k, 2) = varModule.objSubFun.Items(i - 1)
142           arrRange(k, 3) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(0)
143           arrRange(k, 4) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(2)
144           arrRange(k, 5) = arrRange(k, 1)
145           arrRange(k, 6) = VBA.Split(varModule.objSubFun.Keys(i - 1), CHR_TO)(3)
146           arrRange(k, 8) = arrRange(k, 6)
147           arrRange(k, 9) = "yes"

148           If objDict.Exists(arrRange(k, 8)) = False Then
149               objDict.Add arrRange(k, 8), AddEncodeName()
150           End If
151           arrRange(k, 10) = objDict.Item(arrRange(k, 8))
              
              'store temporarily the cipher for the function for the string encryption
152           If bEncodeStr Then
153               If arrRange(k, 1) <> "Sub" And arrRange(k, 1) <> "Function" And Left$(arrRange(k, 1), 8) <> "Property" Then
154               ElseIf arrRange(k, 2) <> 1 Then
155               ElseIf arrRange(k, 6) = GetNameSubFromString(CStr(varStrCryptFunc(0))) Then
156                   strCryptFuncCipher = arrRange(k, 10)
157               End If
158           End If
              
159           k = k + 1
160       Next i

161       Application.StatusBar = "Data collection: Procedure names, completed:" & VBA.Format(3 / 7, "Percent")
162       For i = 1 To varModule.objContr.Count
163           arrRange(k, 1) = "Control"
164           arrRange(k, 2) = varModule.objContr.Items(i - 1)
165           arrRange(k, 3) = VBA.Split(varModule.objContr.Keys(i - 1), CHR_TO)(0)
166           arrRange(k, 4) = "Private"
167           arrRange(k, 6) = VBA.Split(varModule.objContr.Keys(i - 1), CHR_TO)(1)
168           arrRange(k, 8) = arrRange(k, 6)
169           arrRange(k, 9) = "yes"

170           If objDict.Exists(arrRange(k, 8)) = False Then
171               objDict.Add arrRange(k, 8), AddEncodeName()
172           End If
173           arrRange(k, 10) = objDict.Item(arrRange(k, 8))
174           k = k + 1
175       Next i

176       Application.StatusBar = "Data collection: Names of controls, completed:" & VBA.Format(4 / 7, "Percent")
177       For i = 1 To varModule.objDimVar.Count
178           arrRange(k, 1) = "Variable"
179           arrRange(k, 2) = varModule.objDimVar.Items(i - 1)
180           arrRange(k, 3) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(0)
181           arrRange(k, 4) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(3)
182           arrRange(k, 5) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(1)
183           arrRange(k, 6) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(2)
184           arrRange(k, 7) = VBA.Split(varModule.objDimVar.Keys(i - 1), CHR_TO)(4)
185           arrRange(k, 8) = arrRange(k, 7)
186           arrRange(k, 9) = "yes"

187           If objDict.Exists(arrRange(k, 8)) = False Then
188               objDict.Add arrRange(k, 8), AddEncodeName()
189           End If
190           arrRange(k, 10) = objDict.Item(arrRange(k, 8))
191           k = k + 1
192           If i Mod 50 = 0 Then
193               Application.StatusBar = "Data collection: Names of controls, completed:" & VBA.Format(i / varModule.objDimVar.Count, "Percent")
194               DoEvents
195           End If
196       Next i

197       Application.StatusBar = "Data collection: Variable names, completed:" & VBA.Format(5 / 7, "Percent")
198       For i = 1 To varModule.objTypeEnum.Count
199           arrRange(k, 1) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(2)
200           arrRange(k, 2) = varModule.objTypeEnum.Items(i - 1)
201           arrRange(k, 3) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(0)
202           arrRange(k, 4) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(1)
203           arrRange(k, 6) = VBA.Split(varModule.objTypeEnum.Keys(i - 1), CHR_TO)(3)
204           arrRange(k, 8) = arrRange(k, 6)
205           arrRange(k, 9) = "yes"

206           If objDict.Exists(arrRange(k, 8)) = False Then
207               objDict.Add arrRange(k, 8), AddEncodeName()
208           End If
209           arrRange(k, 10) = objDict.Item(arrRange(k, 8))
210           k = k + 1
211       Next i

212       Application.StatusBar = "Data collection: Names of enumerations and types, completed:" & VBA.Format(6 / 7, "Percent")
213       For i = 1 To varModule.objAPI.Count
214           arrRange(k, 1) = "API"
215           arrRange(k, 2) = varModule.objAPI.Items(i - 1)
216           arrRange(k, 3) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(0)
217           arrRange(k, 4) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(1)
218           arrRange(k, 5) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(2)
219           arrRange(k, 6) = VBA.Split(varModule.objAPI.Keys(i - 1), CHR_TO)(3)
220           arrRange(k, 8) = arrRange(k, 6)
221           arrRange(k, 9) = "yes"

222           If objDict.Exists(arrRange(k, 8)) = False Then
223               objDict.Add arrRange(k, 8), AddEncodeName()
224           End If
225           arrRange(k, 10) = objDict.Item(arrRange(k, 8))
226           k = k + 1
227       Next i
228       Application.StatusBar = "Data collection: API names, completed:" & VBA.Format(7 / 7, "Percent")

229       With ActiveSheet
230           Application.StatusBar = "Application of formats"
231           .Cells.ClearContents
232           .Cells(1, 1).Value = "Type"
233           .Cells(1, 2).Value = "Module type"
234           .Cells(1, 3).Value = "Module name"
235           .Cells(1, 4).Value = "Access Modifiers"
236           .Cells(1, 5).Value = "Percentage type. and funk."
237           .Cells(1, 6).Value = "The name of the percentage. and funk."
238           .Cells(1, 7).Value = "Name of variables"
239           .Cells(1, 8).Value = "Encryption Object"
240           .Cells(1, 9).Value = "Encrypt yes/No"
241           .Cells(1, 10).Value = "Code"
242           .Cells(1, 11).Value = "Mistakes"

243           .Cells(2, 1).Resize(UBound(arrRange), 10) = arrRange

244           .Range(.Cells(2, 11), .Cells(k, 11)).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3]," & SHSNIPPETS.ListObjects(C_Const.TB_SERVICEWORDS).DataBodyRange.Address(ReferenceStyle:=xlR1C1, External:=True) & ",1,0),"""")"
245           .Range(.Cells(2, 9), .Cells(k, 9)).FormulaR1C1 = "=IF(RC[2]="""",""yes"",""no"")"
246           .Columns("A:K").AutoFilter
247           .Columns("A:K").EntireColumn.AutoFit
248           .Range(Cells(2, 9), Cells(UBound(arrRange) + 1, 9)).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="YES, NO"
              
              '
249           If bSafeMode Then
250               For i = 2 To UBound(arrRange) + 1
                      'exclude Excel Objects
251                   If .Cells(i, 2).Value = "100" Then
                          'Debug.Print .Cells(i, 5).Value
252                       If .Cells(i, 1).Value = "Module" And _
                              .Cells(i, 5).Value = "" Then 'changed 31.08 CalDymos
253                           .Cells(i, 9).Value = "NO"
254                       End If
                      'exclude API
255                   ElseIf .Cells(i, 1).Value = "API" Then
256                       .Cells(i, 9).Value = "NO"
257                   End If
258               Next i
259           End If
260           Application.StatusBar = "Application of formats, finished"
261       End With

          'Laden der String-Variablen
262       If bEncodeStr Then
263           Call AddSheetInWBook(NAME_SH_STR, ActiveWorkbook)
264           Application.StatusBar = "Collecting String variables"
265           If varModule.objStringCode.Count <> 0 Then
266               ReDim arrRange(1 To varModule.objStringCode.Count, 1 To 8) As String
267               For i = 1 To varModule.objStringCode.Count
268                   arrRange(i, 1) = varModule.objStringCode.Items(i - 1)
269                   arrRange(i, 2) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(0)
270                   arrRange(i, 3) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(1)
271                   arrRange(i, 4) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(2)
272                   arrRange(i, 5) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(3)
273                   arrRange(i, 6) = VBA.Split(varModule.objStringCode.Keys(i - 1), CHR_TO)(4)
274                   arrRange(i, 7) = "yes"
275                   arrRange(i, 8) = AddEncodeName() ' Modulname für String Konstanten

276                   If i Mod 50 = 0 Then
277                       Application.StatusBar = "Collecting String variables, completed:" & VBA.Format(i / varModule.objStringCode.Count, "Percent")
278                       DoEvents
279                   End If
280               Next i
281               Application.StatusBar = "Collecting String variables, completed"
282               With ActiveSheet
283                   .Cells(1, 1).Value = "Module type"
284                   .Cells(1, 2).Value = "Module name"
285                   .Cells(1, 3).Value = "Type Sub or Fun"
286                   .Cells(1, 4).Value = "Name Sub or Fun"
287                   .Cells(1, 5).Value = "Line"
288                   .Cells(1, 6).Value = "Array Strings"
289                   .Cells(1, 7).Value = "Encrypt yes/No"
290                   .Cells(1, 8).Value = "Code"
291                   .Cells(1, 9).Value = "Module cipher"
                      
292                   .Cells(1, 11).Value = "The cipher of the Const module"
293                   .Cells(2, 11).Value = AddEncodeName()
                      
                      'Add additional information for string encryption
294                   .Cells(1, 12).Value = "The name of the Key constant"
295                   .Cells(2, 12).Value = asCryptKey(0) ' Schlüssel für die Stringverschlüsselung
296                   .Cells(1, 13).Value = "The Key value"
297                   .Cells(2, 13).Value = asCryptKey(1) ' Schlüssel für die Stringverschlüsselung
298                   .Cells(1, 14).Value = "The module for the key"
299                   .Cells(2, 14).Value = objTmpName.Keys(z2)
300                   .Cells(1, 15).Value = "The cipher of Key constant"
301                   .Cells(2, 15).Value = strCryptKeyCipher
302                   .Cells(1, 16).Value = "Name of Crypt Func"
303                   .Cells(2, 16).Value = GetNameSubFromString(CStr(varStrCryptFunc(0)))
304                   .Cells(1, 17).Value = "The module for die Crypt Func"
305                   .Cells(2, 17).Value = objTmpName.Keys(z1)
306                   .Cells(1, 18).Value = "The cipher of Crypt Func"
307                   .Cells(2, 18).Value = strCryptFuncCipher

308                   .Cells(2, 1).Resize(UBound(arrRange), 8) = arrRange

309                   .Range(Cells(2, 7), Cells(UBound(arrRange) + 1, 7)).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="YES, NO"
310                   .Range(Cells(2, 9), Cells(UBound(arrRange) + 1, 9)).FormulaR1C1 = "=IF(RC1*1=100,RC2,VLOOKUP(RC2,DATA_OBF_VBATools!R2C3:R" & k & "C10,8,0))"
311                   .Columns("A:I").AutoFilter
312                   .Columns("A:D").EntireColumn.AutoFit
313                   .Columns("E").ColumnWidth = 60
314                   .Columns("F:S").EntireColumn.AutoFit
315                   .Columns("A:S").HorizontalAlignment = xlCenter
316                   .Rows("2:" & UBound(arrRange) + 1).RowHeight = 12
317               End With
318           End If
319       End If
320       ActiveWorkbook.Worksheets(NAME_SH).Activate

321       Application.StatusBar = False
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
322       Application.DisplayAlerts = False
323       On Error Resume Next
324       wb.Worksheets(WSheetName).Delete
325       On Error GoTo 0
326       Application.DisplayAlerts = True
327       wb.Sheets.Add Before:=ActiveSheet
328       ActiveSheet.Name = WSheetName
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
'* Optional varStrCryptFunc As Variant = Nothing :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 10:19    CalDymos
Private Sub ParserVariebleSubFunc(ByRef objVBC As VBIDE.VBComponent, ByRef objDic As Scripting.Dictionary, ByRef objDicStr As Scripting.Dictionary, Optional varStrCryptFunc As Variant = Nothing)
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

329       With objVBC.CodeModule
330           lLine = .CountOfLines
331           If lLine > 0 Then
332               sCode = .Lines(1, lLine)
333               If sCode <> vbNullString Then
                      'remove the line breaks.
334                   sCode = VBA.Replace(sCode, " _" & vbNewLine, vbNullString)
335                   arrStrCode = VBA.Split(sCode, vbNewLine)
336                   For Each itemArr In arrStrCode
337                       itemArr = C_PublicFunctions.TrimSpace(itemArr)
338                       If itemArr <> vbNullString And VBA.Left$(itemArr, 1) <> "'" Then
339                           sVar = vbNullString
                              'If the code contains a comment, delete it.
340                           itemArr = DeleteCommentString(itemArr)
                              'Extract from the declaration clause and determine what is included in the process
341                           If (itemArr Like "* Sub *(*)*" Or itemArr Like "* Function *(*)*" Or itemArr Like "* Property Let *(*)*" Or itemArr Like "* Property Set *(*)*" Or itemArr Like "* Property Get *(*)*" Or _
                                  itemArr Like "Sub *(*)*" Or itemArr Like "Function *(*)*" Or itemArr Like "Property Let *(*)*" Or itemArr Like "Property Set *(*)*" Or itemArr Like "Property Get *(*)*") _
                                  And (Not itemArr Like "*As IRibbonControl*" And Not itemArr Like "* Declare *(*)*") Then

342                               sSubName = TypeProcedure(VBA.CStr(itemArr))
343                               sSubName = sSubName & CHR_TO & GetNameSubFromString(itemArr)
344                               sVar = ParserStrDimConst(itemArr, sSubName, .Name)

345                           End If
                              'Wenn in der Aufzählung und im Datentyp
346                           If itemArr Like "Private Enum *" Or itemArr Like "Public Enum *" Or itemArr Like "Enum *" Or itemArr Like "Private Type *" Or itemArr Like "Public Type *" Or itemArr Like "Type *" Then
347                               arrEnum = VBA.Split(itemArr, " ")
348                               If VBA.CStr(itemArr) Like "Private *" Then
349                                   sNumTypeName = "Private"
350                               Else
351                                   sNumTypeName = "Public"
352                               End If
353                               sNumTypeName = arrEnum(UBound(arrEnum)) & CHR_TO & sNumTypeName
354                               If itemArr Like "* Enum *" Or itemArr Like "Enum *" Then
355                                   sType = "Enum"
356                               Else
357                                   sType = "Type"
358                               End If
359                           End If
                              'aus dem Prozess oder der Aufzählung herausgehen
360                           If itemArr Like "*End Sub" Or itemArr Like "*End Function" Or itemArr Like "*End Property" Or itemArr Like "*End Enum" Or itemArr Like "*End Type" Then
361                               sSubName = vbNullString
362                               sNumTypeName = vbNullString
363                           End If
                              'Falls innerhalb des Typs oder der Aufzählung
364                           If sNumTypeName <> vbNullString And Not itemArr Like "* Enum *" And Not itemArr Like "Enum *" And Not itemArr Like "* Type *" And Not itemArr Like "Type *" Then
365                               arrEnum = VBA.Split(VBA.Trim$(itemArr), " ")
366                               sVar = arrEnum(0)
367                               If sVar Like "*(*" Then sVar = VBA.Left$(sVar, VBA.InStr(1, sVar, "(") - 1)
368                               sVar = .Name & CHR_TO & sType & CHR_TO & sNumTypeName & CHR_TO & ReplaceType(sVar)
369                           End If
                              'wenn wir uns nur innerhalb der Prozedur befinden
370                           If (itemArr Like "* Dim *" Or itemArr Like "* Const *" Or itemArr Like "Dim *" Or itemArr Like "Const *") And sSubName <> vbNullString Then
371                               sVar = ParserStrDimConst(itemArr, sSubName, .Name)
372                           End If
373                           arrVar = VBA.Split(sVar, vbNewLine)
374                           For Each itemVar In arrVar
375                               If itemVar <> vbNullString And objDic.Exists(itemVar) = False Then
376                                   objDic.Add itemVar, objVBC.Type
377                               End If
378                           Next itemVar
379                           Call ParserStringInCode(itemArr, sSubName, objVBC, objDicStr)
380                       End If
381                   Next itemArr
382               End If
383           End If
              
384           If IsArray(varStrCryptFunc) Then
385               For Each itemArr In varStrCryptFunc
386                   itemArr = C_PublicFunctions.TrimSpace(itemArr)
387                   If itemArr <> vbNullString And VBA.Left$(itemArr, 1) <> "'" Then
388                       sVar = vbNullString
                          'If the code contains a comment, delete it.
389                       itemArr = DeleteCommentString(itemArr)
                          'aus der Deklarationsklausel entnehmen und feststellen, was in das Verfahren aufgenommen wurde
390                       If (itemArr Like "* Sub *(*)*" Or itemArr Like "* Function *(*)*" Or itemArr Like "* Property Let *(*)*" Or itemArr Like "* Property Set *(*)*" Or itemArr Like "* Property Get *(*)*" Or _
                              itemArr Like "Sub *(*)*" Or itemArr Like "Function *(*)*" Or itemArr Like "Property Let *(*)*" Or itemArr Like "Property Set *(*)*" Or itemArr Like "Property Get *(*)*") _
                              And (Not itemArr Like "*As IRibbonControl*" And Not itemArr Like "* Declare *(*)*") Then

391                           sSubName = TypeProcedure(VBA.CStr(itemArr))
392                           sSubName = sSubName & CHR_TO & GetNameSubFromString(itemArr)
393                           sVar = ParserStrDimConst(itemArr, sSubName, .Name)

394                       End If
                          'Wenn in der Aufzählung und im Datentyp
395                       If itemArr Like "Private Enum *" Or itemArr Like "Public Enum *" Or itemArr Like "Enum *" Or itemArr Like "Private Type *" Or itemArr Like "Public Type *" Or itemArr Like "Type *" Then
396                           arrEnum = VBA.Split(itemArr, " ")
397                           If VBA.CStr(itemArr) Like "Private *" Then
398                               sNumTypeName = "Private"
399                           Else
400                               sNumTypeName = "Public"
401                           End If
402                           sNumTypeName = arrEnum(UBound(arrEnum)) & CHR_TO & sNumTypeName
403                           If itemArr Like "* Enum *" Or itemArr Like "Enum *" Then
404                               sType = "Enum"
405                           Else
406                               sType = "Type"
407                           End If
408                       End If
                          'aus dem Prozess oder der Aufzählung herausgehen
409                       If itemArr Like "*End Sub" Or itemArr Like "*End Function" Or itemArr Like "*End Property" Or itemArr Like "*End Enum" Or itemArr Like "*End Type" Then
410                           sSubName = vbNullString
411                           sNumTypeName = vbNullString
412                       End If
                          'Falls innerhalb des Typs oder der Aufzählung
413                       If sNumTypeName <> vbNullString And Not itemArr Like "* Enum *" And Not itemArr Like "Enum *" And Not itemArr Like "* Type *" And Not itemArr Like "Type *" Then
414                           arrEnum = VBA.Split(VBA.Trim$(itemArr), " ")
415                           sVar = arrEnum(0)
416                           If sVar Like "*(*" Then sVar = VBA.Left$(sVar, VBA.InStr(1, sVar, "(") - 1)
417                           sVar = .Name & CHR_TO & sType & CHR_TO & sNumTypeName & CHR_TO & ReplaceType(sVar)
418                       End If
                          'wenn wir uns nur innerhalb der Prozedur befinden
419                       If (itemArr Like "* Dim *" Or itemArr Like "* Const *" Or itemArr Like "Dim *" Or itemArr Like "Const *") And sSubName <> vbNullString Then
420                           sVar = ParserStrDimConst(itemArr, sSubName, .Name)
421                       End If
422                       arrVar = VBA.Split(sVar, vbNewLine)
423                       For Each itemVar In arrVar
424                           If itemVar <> vbNullString And objDic.Exists(itemVar) = False Then
425                               objDic.Add itemVar, objVBC.Type
426                           End If
427                       Next itemVar
428                       Call ParserStringInCode(itemArr, sSubName, objVBC, objDicStr)
429                   End If
430               Next itemArr
431           End If
432       End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : GetNameSubFromString - Get the procedure name from the string
'* Created    : 20-04-2020 18:19
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal sStrCode As String : ñòðîêà
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function GetNameSubFromString(ByVal sStrCode As String) As String
        Dim sTemp       As String
433     sTemp = VBA.Trim$(VBA.Left$(sStrCode, VBA.InStr(1, sStrCode, "(") - 1))
434     Select Case True
        Case sTemp Like "*Sub *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Sub ") - 3)
435         Case sTemp Like "*Function *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Function ") - 8)
436         Case sTemp Like "*Property Let *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Let ") - 12)
437         Case sTemp Like "*Property Set *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Set ") - 12)
438         Case sTemp Like "*Property Get *": sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Property Get ") - 12)
439     End Select
440     GetNameSubFromString = VBA.Trim$(sTemp)
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
'* Updated    : 31-08-2023 08:07    CalDymos
     Private Sub ParserStringInCode(ByVal sSTR As String, ByVal sNameSub As String, ByRef objVBC As VBIDE.VBComponent, ByRef objDicStr As Scripting.Dictionary)
        Dim sTxt        As String
        Dim arrStr      As Variant
        Dim Arr         As Variant
        Dim sReplace    As String
        Dim i           As Integer
        Dim sArray      As String
        Const CHAR_REPLACE As String = "XXXXX"

441     sSTR = VBA.Trim$(sSTR)

442     If sSTR Like "*" & VBA.Chr$(34) & "*" And _
           sSTR <> vbNullString And _
           Not sSTR Like "*Declare * Lib *(*)*" And _
           Not sSTR Like "* Const * = " & VBA.Chr$(34) & "*" Then 'changed : am 31.08 CalDymos

443         sTxt = VBA.Right$(sSTR, VBA.Len(sSTR) - VBA.InStr(1, sSTR, VBA.Chr$(34)) + 1)
444         sTxt = VBA.Replace(sTxt, VBA.Chr$(34) & VBA.Chr$(34), CHAR_REPLACE)
445         arrStr = VBA.Split(sTxt, VBA.Chr$(34))

446         sArray = VBA.Left$(sSTR, VBA.InStr(1, sSTR, VBA.Chr$(34)) - 1)
447         If sArray Like "* = Array(" Then
448             sArray = VBA.Replace(sArray, " = Array(", vbNullString)
449             Arr = VBA.Split(sArray, " ")
450             sArray = Arr(UBound(Arr))
451         Else
452             sArray = vbNullString
453         End If
454         For i = 1 To UBound(arrStr) Step 2
455             If arrStr(i) <> vbNullString Then
456                 If sNameSub = vbNullString Then sNameSub = "Declaration" & CHR_TO

457                 sReplace = VBA.Replace(arrStr(i), CHAR_REPLACE, VBA.Chr$(34) & VBA.Chr$(34))
458                 sTxt = objVBC.Name & CHR_TO & sNameSub & CHR_TO & VBA.Chr$(34) & sReplace & VBA.Chr$(34) & CHR_TO & sArray    '& CHR_TO & sYesNo
459                 If arrStr(i + 1) Like "*: * = *" Then sArray = vbNullString
460                 If arrStr(i + 1) Like "*: * = Array(*" Then
461                     sArray = VBA.Replace(arrStr(i + 1), ": ", vbNullString)
462                     sArray = VBA.Replace(sArray, " = Array(", vbNullString)
463                     sArray = VBA.Replace(sArray, ")", vbNullString)
464                 End If
465                 If objDicStr.Exists(sTxt) = False Then objDicStr.Add sTxt, objVBC.Type
466             End If
467         Next i
468         sArray = vbNullString
469     End If
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

470     sTemp = C_PublicFunctions.TrimSpace(sTxt)
471     sType = "Dim"
472     If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
            'åñëè åñòü êîìåíòàðèé â ñòðîêå êîäà òî óäàëÿåì åãî
473         sTemp = DeleteCommentString(sTemp)
474         If sTemp Like "*Sub *(*)*" Or sTemp Like "*Function *(*)*" Or sTemp Like "*Property Let *(*)*" Or sTemp Like "*Property Set *(*)*" Or sTemp Like "*Property Get *(*)*" Then
475             If VBA.InStr(1, sTemp, ")") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, ")") - 1)
476             If VBA.InStr(1, sTemp, " = ") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ") - 1)
477             If VBA.Len(sTemp) - VBA.InStr(1, sTemp, "(") >= 0 Then
478                 sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "("))
479             End If
480         ElseIf sTemp Like "* Dim *" Or sTemp Like Chr$(68) & "im *" Then
481             sType = "Dim"
482             If VBA.InStr(1, sTemp, "Dim ") >= 3 Then sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Dim ") - 3)
483         ElseIf sTemp Like "* Const *" Or sTemp Like Chr$(67) & "onst *" Then
484             sType = "Const"
485             If VBA.InStr(1, sTemp, "Const ") >= 5 Then sTemp = VBA.Right$(sTemp, VBA.Len(sTemp) - VBA.InStr(1, sTemp, "Const ") - 5)
486             If VBA.InStr(1, sTemp, " = ") >= 1 Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ") - 1)
487         Else
488             sTemp = vbNullString
489         End If
490     End If

491     If sTemp Like "*: *" Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, ": ") - 1)
492     If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
493         arrStr = VBA.Split(sTemp, ",")
494         For Each itemArr In arrStr
495             If itemArr Like "*(*" Then itemArr = VBA.Left$(itemArr, VBA.InStr(1, itemArr, "(") - 1)
496             If Not itemArr Like "*)*" And Not itemArr Like "* To *" Then
497                 arrWord = VBA.Split(itemArr, " As ")
498                 arrWord = VBA.Split(VBA.Trim$(arrWord(0)), " ")
499                 If UBound(arrWord) = -1 Then
500                     sWord = vbNullString
501                 Else
502                     sWordTemp = VBA.Trim$(arrWord(UBound(arrWord)))
503                     sWordTemp = ReplaceType(sWordTemp)
504                     sWord = sWord & vbNewLine & sNameMod & CHR_TO & sNameSub & CHR_TO & sType & CHR_TO & sWordTemp
505                 End If
506             End If
507         Next itemArr
508     End If
509     sWord = VBA.Trim$(sWord)
510     If VBA.Len(sWord) = 0 Then
511         sWord = vbNullString
512     Else
513         sWord = VBA.Trim$(VBA.Right$(sWord, VBA.Len(sWord) - 2))
514     End If
515     ParserStrDimConst = sWord
   End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserNameSubFunc - ñáîð íàçâàíèé ïðîöåäóð è ôóíêöèé
'* Created    : 27-03-2020 13:20
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objCodeModule As VBIDE.CodeModule : îáúåêò ìîäóëü
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 10:22    CalDymos

     Private Sub ParserNameSubFunc(ByVal sNameVBC As String, ByRef objVBC As VBIDE.VBComponent, ByRef varSubFun As Scripting.Dictionary, Optional varStrCryptFunc As Variant = Nothing)
          Dim ProcKind    As VBIDE.vbext_ProcKind
          Dim lLine       As Long
          Dim lineOld     As Long
          Dim sNameSub    As String
          Dim strFunctionBody As String
                        Dim skey As String
                        
516       With objVBC.CodeModule
517           If .CountOfLines > 0 Then
518               lLine = .CountOfDeclarationLines
519               If lLine = 0 Then lLine = 2
520               Do Until lLine >= .CountOfLines

                      'Sammeln von Namen von Prozeduren und Funktionen
521                   sNameSub = .ProcOfLine(lLine, ProcKind)
522                   If sNameSub <> vbNullString Then
523                       strFunctionBody = C_PublicFunctions.TrimSpace(.Lines(lLine - 1, .ProcCountLines(sNameSub, ProcKind)))
                            'Debug.Print strFunctionBody
                            'Debug.Print .Lines(lLine - 1, .ProcCountLines(sNameSub, ProcKind))
524                       If (Not strFunctionBody Like "*As IRibbonControl*") And _
                              (Not strFunctionBody Like "*As IRibbonUI*") And _
                              (Not WorkBookAndSheetsEvents(strFunctionBody, objVBC.Type)) And _
                              (Not (strFunctionBody Like "* UserForm_*" And objVBC.Type = vbext_ct_MSForm)) And _
                              (Not UserFormsEvents(strFunctionBody, objVBC.Type)) Then
525                           skey = sNameVBC & CHR_TO & TypeProcedure(strFunctionBody) & CHR_TO & TypeOfAccessModifier(strFunctionBody) & CHR_TO & sNameSub
                              'Debug.Print skey
526                           If Not varSubFun.Exists(skey) Then
527                               varSubFun.Add skey, objVBC.Type
528                           End If
529                       End If
530                       lLine = .ProcStartLine(sNameSub, ProcKind) + .ProcCountLines(sNameSub, ProcKind) + 1
531                   Else
532                       lLine = lLine + 1
533                   End If
534                   If lineOld > lLine Then Exit Do
535                   lineOld = lLine
536               Loop
537           End If
538       End With
539       If IsArray(varStrCryptFunc) Then

540               skey = sNameVBC & CHR_TO & TypeProcedure(CStr(varStrCryptFunc(0))) & CHR_TO & TypeOfAccessModifier(CStr(varStrCryptFunc(0))) & CHR_TO & GetNameSubFromString(CStr(varStrCryptFunc(0)))
                  'Debug.Print skey
541                 If Not varSubFun.Exists(skey) Then
542                     varSubFun.Add skey, objVBC.Type
543                 End If
544       End If
        
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ParserNameControlsForm - ñáîð íàçâàíèé êîíòðîëîâ þçåðôîðì
'* Created    : 27-03-2020 13:50
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                         Description
'*
'* ByRef objVBC As VBIDE.VBComponent :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub ParserNameControlsForm(ByVal sNameVBC As String, ByRef objVBC As VBIDE.VBComponent, ByRef obfNewDict As Scripting.Dictionary)
        Dim objCont     As MSForms.control
545     If Not objVBC.Designer Is Nothing Then
546         With objVBC.Designer
547             For Each objCont In .Controls
                    'Debug.Print sNameVBC & CHR_TO & objCont.Name, objVBC.Type
548                 obfNewDict.Add sNameVBC & CHR_TO & objCont.Name, objVBC.Type
549             Next objCont
550         End With
551     End If
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserNameGlobalVariable - ñáîð ãëîáàëüíûõ ïåðåìåííûõ
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

     Private Sub ParserNameGlobalVariable(ByVal sNameVBC As String, ByRef objVBC As VBIDE.VBComponent, ByRef dicGloblVar As Scripting.Dictionary, ByRef dicTypeEnum As Scripting.Dictionary, ByRef dicAPI As Scripting.Dictionary, Optional asCryptKey As Variant = Nothing)
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
552       bFlag = True
553       If objVBC.CodeModule.CountOfDeclarationLines <> 0 Then
554           sTemp = objVBC.CodeModule.Lines(1, objVBC.CodeModule.CountOfDeclarationLines)
555           sTemp = VBA.Replace(sTemp, " _" & vbNewLine, vbNullString)
556           If sTemp <> vbNullString Then
557               varArr = VBA.Split(sTemp, vbNewLine)
558               For i = 0 To UBound(varArr)
559                   sTemp = C_PublicFunctions.TrimSpace(DeleteCommentString(varArr(i)))
560                   If sTemp <> vbNullString And VBA.Left$(sTemp, 1) <> "'" Then
561                       If sTemp Like "* Type *" Or sTemp Like "* Enum *" Or sTemp Like "Type *" Or sTemp Like "Enum *" Then
562                           varArrWord = VBA.Split(sTemp, " ")
563                           If UBound(varArrWord) = 2 Then
564                               sTemp = VBA.Trim$(varArrWord(0)) & CHR_TO & VBA.Trim$(varArrWord(1)) & CHR_TO & VBA.Trim$(varArrWord(2))
565                           ElseIf UBound(varArrWord) = 1 Then
566                               sTemp = "Public" & CHR_TO & VBA.Trim$(varArrWord(0)) & CHR_TO & VBA.Trim$(varArrWord(1))
567                           End If
568                           sTemp = sNameVBC & CHR_TO & sTemp
569                           If Not dicTypeEnum.Exists(sTemp) Then dicTypeEnum.Add sTemp, objVBC.Type
570                           bFlag = False
571                       End If
572                       If bFlag And Not (sTemp Like "Implements *" Or sTemp Like "Option *" Or VBA.Left$(sTemp, 1) = "'" Or sTemp = vbNullString Or VBA.Left$(sTemp, 1) = "#" Or sTemp Like "*Declare *(*)*" Or sTemp Like "*Event *(*)") Then

573                           If sTemp Like "* = *" Then sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " = ", vbTextCompare) + 2)
574                           If sTemp Like "* *(* To *) *" Then
575                               sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, "(", vbTextCompare) - 1)
576                           End If
577                           varStr = VBA.Split(sTemp, ",")
578                           For Each itemVarStr In varStr
579                               sTemp = VBA.Trim$(itemVarStr)
580                               varArrWord = VBA.Split(sTemp, " As ")
581                               varArrWord = VBA.Split(varArrWord(0), " = ")
582                               sTemp = varArrWord(0)
583                               varArrWord = VBA.Split(sTemp, " ")

584                               j = UBound(varArrWord)
585                               If j > 1 Then
586                                   If varArrWord(0) = "Dim" Or varArrWord(0) = "Const" Then
587                                       sTemp = "Private" & CHR_TO & varArrWord(0) & CHR_TO
588                                       sTempArr = varArrWord(1)
589                                   ElseIf (varArrWord(0) = "Private" Or varArrWord(0) = "Public") And (varArrWord(1) = "Dim" Or varArrWord(1) = "Const" Or varArrWord(1) = "WithEvents") Then
590                                       sTemp = varArrWord(0) & CHR_TO & varArrWord(1) & CHR_TO
591                                       sTempArr = varArrWord(2)
592                                   ElseIf (varArrWord(0) = "Private" Or varArrWord(0) = "Public") And Not (varArrWord(1) = "Dim" Or varArrWord(1) = "Const" Or varArrWord(1) = "WithEvents") Then
593                                       sTemp = varArrWord(0) & CHR_TO & "Dim" & CHR_TO
594                                       sTempArr = varArrWord(1)
595                                   End If
596                               ElseIf j = 1 And varArrWord(0) = "Global" Then
597                                   sTemp = "Public" & CHR_TO & varArrWord(0) & CHR_TO
598                                   sTempArr = varArrWord(1)
599                               ElseIf j = 1 And (varArrWord(0) = "Private" Or varArrWord(0) = "Public") Then
600                                   sTemp = varArrWord(0) & CHR_TO & "Dim" & CHR_TO
601                                   sTempArr = varArrWord(1)
602                               ElseIf j = 1 And (varArrWord(0) = "Dim" Or varArrWord(0) = "Const") Then
603                                   sTemp = "Private" & CHR_TO & varArrWord(0) & CHR_TO
604                                   sTempArr = varArrWord(1)
605                               ElseIf j = 0 Then
606                                   sTemp = "Private" & CHR_TO & " Dim" & CHR_TO
607                                   sTempArr = varArrWord(0)
608                               End If

609                               sTempArr = ReplaceType(sTempArr)
610                               If sTempArr Like "*(*" Then sTempArr = VBA.Left$(sTempArr, VBA.InStr(1, sTempArr, "(") - 1)
611                               sTemp = sNameVBC & CHR_TO & sTemp & sTempArr
612                               If Not dicGloblVar.Exists(sTemp) Then dicGloblVar.Add sTemp, objVBC.Type

613                               sTemp = vbNullString
614                           Next itemVarStr
615                           sTemp = vbNullString
616                       End If
617                       If sTemp Like "*End Type" Or sTemp Like "*End Enum" Then
618                           bFlag = True
619                       End If
620                       If sTemp Like "*Declare * Lib " & VBA.Chr$(34) & "*" & VBA.Chr$(34) & " (*)*" Then
621                           sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, " Lib ", vbTextCompare) - 1)
622                           varAPI = VBA.Split(sTemp, VBA.Chr$(32))
623                           itemArr = UBound(varAPI)
624                           sTemp = CHR_TO & varAPI(itemArr - 1) & CHR_TO & varAPI(itemArr)
625                           If varAPI(1) = "Declare" Then
626                               sTemp = sNameVBC & CHR_TO & varAPI(0) & sTemp
627                           Else
628                               sTemp = sNameVBC & CHR_TO & "Private" & sTemp
629                           End If
630                           If Not dicAPI.Exists(sTemp) Then dicAPI.Add sTemp, objVBC.Type
631                       End If
632                       If sTemp Like "*Event *(*)" Then
633                           sTemp = VBA.Left$(sTemp, VBA.InStr(1, sTemp, "(", vbTextCompare) - 1)
634                           varAPI = VBA.Split(sTemp, VBA.Chr$(32))
635                           itemArr = UBound(varAPI)
636                           sTemp = CHR_TO & varAPI(itemArr - 1) & CHR_TO & varAPI(itemArr)
637                           If varAPI(1) = "Event" Then
638                               sTemp = sNameVBC & CHR_TO & varAPI(0) & sTemp
639                           Else
640                               sTemp = sNameVBC & CHR_TO & "Private" & sTemp
641                           End If
642                           If Not dicAPI.Exists(sTemp) Then dicAPI.Add sTemp, objVBC.Type
643                       End If
644                   End If
645               Next i
646           End If
647       End If
648       If IsArray(asCryptKey) Then
649           sTemp = sNameVBC & CHR_TO & "Public" & CHR_TO & "Const" & CHR_TO & CStr(asCryptKey(0))
650           If Not dicGloblVar.Exists(sTemp) Then dicGloblVar.Add sTemp, objVBC.Type

651           sTemp = vbNullString
652       End If
   End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddNewDictionary -ôóíêöèÿ èíèöèàëèçàöèè ñëîâàðÿ
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
653     Set objDict = Nothing
654     Set objDict = New Scripting.Dictionary
655     Set AddNewDictionary = objDict
   End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DeleteCommentString - óäàëåíèå â ñòðîêå êîììåíòàðèÿ
'* Created    : 20-04-2020 18:18
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sWord As String : ñòðîêà
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function DeleteCommentString(ByVal sWord As String) As String
        'åñòü '
        Dim sTemp       As String
656     sTemp = sWord
657     If VBA.InStr(1, sTemp, "'") <> 0 Then
658         If VBA.InStr(1, sTemp, VBA.Chr(34)) <> 0 Then
                'åñòü "
659             If VBA.InStr(1, sTemp, "'") < VBA.InStr(1, sTemp, VBA.Chr(34)) Then
                    'åñëè òàê -> '"
660                 sTemp = VBA.Trim$(VBA.Left$(sTemp, VBA.InStr(1, sTemp, "'") - 1))
661             End If
662         Else
                'íåò " -> '
663             sTemp = VBA.Trim$(VBA.Left$(sTemp, VBA.InStr(1, sTemp, "'") - 1))
664         End If
665     End If
666     DeleteCommentString = sTemp
   End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddEncodeName - Die Funktion der zufälligen Zuweisung eines zufällig vergebenen Namens
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
667     Err.Clear
668     sName = vbNullString
669     Randomize
670     sName = "o"
671     For i = 2 To CharCount
672         If (VBA.Round(VBA.Rnd() * 1000)) Mod 2 = 1 Then sName = sName & FIRST_CODE_SIGN Else sName = sName & SECOND_CODE_SIGN
673     Next i
674     On Error Resume Next
       'add a name to the collection, if the name exists, then
        'an error is generated, which restarts the generation of the name
675     objCollUnical.Add sName, sName
676     If Err.Number <> 0 Then GoTo tryAgain
677     AddEncodeName = sName
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

678   Randomize Timer

679   For i = 1 To 18
680   z1 = Int(1 + Rnd * (26 - 1 + 1))
      'Debug.Print Z1
681   z2 = Int(Rnd * (9 + 1))
      'Debug.Print Z2
682   Z3 = Int(1 + Rnd * (2 - 1 + 1))
      'Debug.Print Z3

683   If Z3 = 1 Then
684       skey = skey & Chr$(64 + z1)
685   Else
686       skey = skey & CStr(z2)
687   End If

688   Next i
      'Debug.Print skey
689   GenerateKey = skey

End Function

'âçÿòî
     Private Function TypeOfAccessModifier(ByRef StrDeclarationProcedure As String) As String
690     If StrDeclarationProcedure Like "*Private *(*)*" Then
691         TypeOfAccessModifier = "Private"
692     Else
693         TypeOfAccessModifier = "Public"
694     End If
   End Function
     Private Function TypeProcedure(ByRef StrDeclarationProcedure As String) As String
695     If StrDeclarationProcedure Like "*Sub *" Then
696         TypeProcedure = "Sub"
697     ElseIf StrDeclarationProcedure Like "*Function *" Then
698         TypeProcedure = "Function"
699     ElseIf StrDeclarationProcedure Like "*Property Set *" Then
700         TypeProcedure = "Property Set"
701     ElseIf StrDeclarationProcedure Like "*Property Get *" Then
702         TypeProcedure = "Property Get"
703     ElseIf StrDeclarationProcedure Like "*Property Let *" Then
704         TypeProcedure = "Property Let"
705     Else
706         TypeProcedure = "Unknown Type"
707     End If
   End Function

     Private Function ReplaceType(ByVal sVar As String) As String
708     sVar = Replace(sVar, "%", vbNullString)     'Integer
709     sVar = Replace(sVar, "&", vbNullString)     'Long
710     sVar = Replace(sVar, "$", vbNullString)     'String
711     sVar = Replace(sVar, "!", vbNullString)     'Single
712     sVar = Replace(sVar, "#", vbNullString)     'Double
713     sVar = Replace(sVar, "@", vbNullString)     'Currency
714     ReplaceType = sVar
   End Function

Private Function WorkBookAndSheetsEvents(ByVal sTxt As String, ByVal TypeModule As VBIDE.vbext_ComponentType) As Boolean
          Dim Flag        As Boolean
715       Flag = False
          'nur für Sheets, Workbooks und Klassenmodule
716       If TypeModule = vbext_ct_Document Or TypeModule = vbext_ct_ClassModule Then
717           Select Case True
                  Case sTxt Like "*_Activate(*": Flag = True
718               Case sTxt Like "*_AddinInstall(*": Flag = True
719               Case sTxt Like "*_AddinUninstall(*": Flag = True
720               Case sTxt Like "*_AfterSave(*": Flag = True
721               Case sTxt Like "*_AfterXmlExport(*": Flag = True
722               Case sTxt Like "*_AfterXmlImport(*": Flag = True
723               Case sTxt Like "*_BeforeClose(*": Flag = True
724               Case sTxt Like "*_BeforeDoubleClick(*": Flag = True
725               Case sTxt Like "*_BeforePrint(*": Flag = True
726               Case sTxt Like "*_BeforeRightClick(*": Flag = True
727               Case sTxt Like "*_BeforeSave(*": Flag = True
728               Case sTxt Like "*_BeforeXmlExport(*": Flag = True
729               Case sTxt Like "*_BeforeXmlImport(*": Flag = True
730               Case sTxt Like "*_Calculate(*": Flag = True
731               Case sTxt Like "*_Change(*": Flag = True
732               Case sTxt Like "*_Deactivate(*": Flag = True
733               Case sTxt Like "*_FollowHyperlink(*": Flag = True
734               Case sTxt Like "*_MouseDown(*": Flag = True
735               Case sTxt Like "*_MouseMove(*": Flag = True
736               Case sTxt Like "*_MouseUp(*": Flag = True
737               Case sTxt Like "*_NewChart(*": Flag = True
738               Case sTxt Like "*_NewSheet(*": Flag = True
739               Case sTxt Like "*_Open(*": Flag = True
740               Case sTxt Like "*_PivotTableAfterValueChange(*": Flag = True
741               Case sTxt Like "*_PivotTableBeforeAllocateChanges(*": Flag = True
742               Case sTxt Like "*_PivotTableBeforeCommitChanges(*": Flag = True
743               Case sTxt Like "*_PivotTableBeforeDiscardChanges(*": Flag = True
744               Case sTxt Like "*_PivotTableChangeSync(*": Flag = True
745               Case sTxt Like "*_PivotTableCloseConnection(*": Flag = True
746               Case sTxt Like "*_PivotTableOpenConnection(*": Flag = True
747               Case sTxt Like "*_PivotTableUpdate(*": Flag = True
748               Case sTxt Like "*_Resize(*": Flag = True
749               Case sTxt Like "*_RowsetComplete(*": Flag = True
750               Case sTxt Like "*_SelectionChange(*": Flag = True
751               Case sTxt Like "*_SeriesChange(*": Flag = True
752               Case sTxt Like "*_SheetActivate(*": Flag = True
753               Case sTxt Like "*_SheetBeforeDoubleClick(*": Flag = True
754               Case sTxt Like "*_SheetBeforeRightClick(*": Flag = True
755               Case sTxt Like "*_SheetCalculate(*": Flag = True
756               Case sTxt Like "*_SheetChange(*": Flag = True
757               Case sTxt Like "*_SheetDeactivate(*": Flag = True
758               Case sTxt Like "*_SheetFollowHyperlink(*": Flag = True
759               Case sTxt Like "*_SheetPivotTableAfterValueChange(*": Flag = True
760               Case sTxt Like "*_SheetPivotTableBeforeAllocateChanges(*": Flag = True
761               Case sTxt Like "*_SheetPivotTableBeforeCommitChanges(*": Flag = True
762               Case sTxt Like "*_SheetPivotTableBeforeDiscardChanges(*": Flag = True
763               Case sTxt Like "*_SheetPivotTableChangeSync(*": Flag = True
764               Case sTxt Like "*_SheetPivotTableUpdate(*": Flag = True
765               Case sTxt Like "*_SheetSelectionChange(*": Flag = True
766               Case sTxt Like "*_Sync(*": Flag = True
767               Case sTxt Like "*_WindowActivate(*": Flag = True
768               Case sTxt Like "*_WindowDeactivate(*": Flag = True
769               Case sTxt Like "*_WindowResize(*": Flag = True
770               Case sTxt Like "*_NewWorkbook(*": Flag = True
771               Case sTxt Like "*_WorkbookActivate(*": Flag = True
772               Case sTxt Like "*_WorkbookAddinInstall(*": Flag = True
773               Case sTxt Like "*_WorkbookAddinUninstall(*": Flag = True
774               Case sTxt Like "*_WorkbookAfterSave(*": Flag = True
775               Case sTxt Like "*_WorkbookAfterXmlExport(*": Flag = True
776               Case sTxt Like "*_WorkbookAfterXmlImport(*": Flag = True
777               Case sTxt Like "*_WorkbookBeforeClose(*": Flag = True
778               Case sTxt Like "*_WorkbookBeforePrint(*": Flag = True
779               Case sTxt Like "*_WorkbookBeforeSave(*": Flag = True
780               Case sTxt Like "*_WorkbookBeforeXmlExport(*": Flag = True
781               Case sTxt Like "*_WorkbookBeforeXmlImport(*": Flag = True
782               Case sTxt Like "*_WorkbookDeactivate(*": Flag = True
783               Case sTxt Like "*_WorkbookModelChange(*": Flag = True
784               Case sTxt Like "*_WorkbookNewChart(*": Flag = True
785               Case sTxt Like "*_WorkbookNewSheet(*": Flag = True
786               Case sTxt Like "*_WorkbookOpen(*": Flag = True
787               Case sTxt Like "*_WorkbookPivotTableCloseConnection(*": Flag = True
788               Case sTxt Like "*_WorkbookPivotTableOpenConnection(*": Flag = True
789               Case sTxt Like "*_WorkbookRowsetComplete(*": Flag = True
790               Case sTxt Like "*_WorkbookSync(*": Flag = True
791           End Select
792       End If
793       WorkBookAndSheetsEvents = Flag
End Function

Private Function UserFormsEvents(ByVal sTxt As String, ByVal TypeModule As VBIDE.vbext_ComponentType) As Boolean
        Dim Flag        As Boolean
794     Flag = False
        'Nur für Events, der Forms und Klassen
795     If TypeModule = vbext_ct_MSForm Or TypeModule = vbext_ct_ClassModule Then
796         Select Case True
            Case sTxt Like "*_AfterUpdate(*": Flag = True
797             Case sTxt Like "*_BeforeDragOver(*": Flag = True
798             Case sTxt Like "*_BeforeDropOrPaste(*": Flag = True
799             Case sTxt Like "*_BeforeUpdate(*": Flag = True
800             Case sTxt Like "*_Change(*": Flag = True
801             Case sTxt Like "*_Click(*": Flag = True
802             Case sTxt Like "*_DblClick(*": Flag = True
803             Case sTxt Like "*_Deactivate(*": Flag = True
804             Case sTxt Like "*_DropButtonClick(*": Flag = True
805             Case sTxt Like "*_Enter(*": Flag = True
806             Case sTxt Like "*_Error(*": Flag = True
807             Case sTxt Like "*_Exit(*": Flag = True
808             Case sTxt Like "*_Initialize(*": Flag = True
809             Case sTxt Like "*_KeyDown(*": Flag = True
810             Case sTxt Like "*_KeyPress(*": Flag = True
811             Case sTxt Like "*_KeyUp(*": Flag = True
812             Case sTxt Like "*_Layout(*": Flag = True
813             Case sTxt Like "*_MouseDown(*": Flag = True
814             Case sTxt Like "*_MouseMove(*": Flag = True
815             Case sTxt Like "*_MouseUp(*": Flag = True
816             Case sTxt Like "*_QueryClose(*": Flag = True
817             Case sTxt Like "*_RemoveControl(*": Flag = True
818             Case sTxt Like "*_Resize(*": Flag = True
819             Case sTxt Like "*_Scroll(*": Flag = True
820             Case sTxt Like "*_Terminate(*": Flag = True
821             Case sTxt Like "*_Zoom(*": Flag = True
822         End Select
823     End If
824     UserFormsEvents = Flag
End Function

