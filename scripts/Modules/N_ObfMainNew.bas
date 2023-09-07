'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : N_ObfMainNew - Modul zur Code-Verschleierung
'* Created    : 08-10-2020 14:11
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 11:26    CalDymos

Option Explicit
Option Private Module

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : IsSubOrFunc
'* Created    : 07-09-2023 07:25
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):         Description
'*
'* strLine As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function IsSubOrFunc(strLine As String) As Boolean

1         If strLine <> "" Then
2             If strLine Like "*Sub *(*" Then
3                 IsSubOrFunc = True
4             ElseIf strLine Like "*Function *(*" Then
5                 IsSubOrFunc = True
6             End If
7         End If

End Function

Private Sub Sort2_asc(Arr(), col As Long)
          Dim temp()      As Variant
          Dim lb2 As Long, ub2 As Long, lTop As Long, lBot As Long

8         lTop = LBound(Arr, 1)
9         lBot = UBound(Arr, 1)
10        lb2 = LBound(Arr, 2)
11        ub2 = UBound(Arr, 2)
12        ReDim temp(lb2 To ub2)

13        Call QSort2_asc(Arr(), col, lTop, lBot, temp(), lb2, ub2)
End Sub
Private Sub QSort2_asc(Arr(), C As Long, ByVal top As Long, ByVal bot As Long, temp(), lb2 As Long, ub2 As Long)
          Dim t As Long, LB As Long, MidItem, j As Long

14        MidItem = Arr((top + bot) \ 2, C)
15        t = top: LB = bot

16        Do
17            Do While Arr(t, C) < MidItem: t = t + 1: Loop
18            Do While Arr(LB, C) > MidItem: LB = LB - 1: Loop
19            If t < LB Then
20                For j = lb2 To ub2: temp(j) = Arr(t, j): Next j
21                For j = lb2 To ub2: Arr(t, j) = Arr(LB, j): Next j
22                For j = lb2 To ub2: Arr(LB, j) = temp(j): Next j
23                t = t + 1: LB = LB - 1
24            ElseIf t = LB Then
25                t = t + 1: LB = LB - 1
26            End If
27        Loop While t <= LB

28        If t < bot Then QSort2_asc Arr(), C, t, bot, temp(), lb2, ub2
29        If top < LB Then QSort2_asc Arr(), C, top, LB, temp(), lb2, ub2

End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : StartCompleteObfuscation - Performs a complete obfuscation, including removing all comments and formatting.
'* Created    : 07-09-2023 07:02
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub StartCompleteObfuscation()
          Dim Form        As AddStatistic
          Dim sNameWB     As String
          Dim objWB       As Workbook
          Dim sPath       As String
          Dim arrNameFile() As String
          Dim i As Integer, j As Integer
          Dim vbComp      As VBIDE.VBComponent
          Dim objWkSh As Worksheet

30        On Error GoTo ErrCompleteObfuscation

31        Application.Calculation = xlCalculationManual
32        Application.ScreenUpdating = False
33        Set Form = New AddStatistic
34        With Form
35            .Caption = "Komplett Code-Verschleierung"
36            .lbOK.Caption = "Start"
37            .chQuestion.visible = True
38            .chQuestion2.visible = True
39            .chQuestion.Value = True
40            .chQuestion2.Value = True
41            .chQuestion.Caption = "String-Werte erfassen"
42            .chQuestion2.Caption = "Sicheren Modus verwenden"
43            .chQuestion2.ControlTipText = "API / Moduletyp 100 werden ausgeschlossen "
44            .lbWord.Caption = 1
45            .Show
46            sNameWB = .cmbMain.Value
47        End With
48        If sNameWB = vbNullString Then Exit Sub
49        If sNameWB Like "*.docm" Or sNameWB Like "*.DOCM" Then
              Dim objWrdApp As Object
50            Set objWrdApp = GetObject(, "Word.Application")
51            Set objWB = objWrdApp.Documents(sNameWB)
52        Else
53            Set objWB = Workbooks(sNameWB)
54        End If

          Dim vbProj  As Object
55        Set vbProj = objWB.VBProject
56        If vbProj.Protection = 1 Then
57            Call MsgBox("The VBA project is password protected, remove the password from the project!", vbCritical, "Removing Formatting:")
58            Exit Sub
59        End If
          
60        Call DelegateMainObfParser(objWB, Form.chQuestion.Value, Form.chQuestion2.Value, True)
61        Call MainObfuscation(objWB, Form.chQuestion.Value, True)
       
          Dim iFile       As Integer
          Dim Arr()       As Variant
          Dim ListCode() As Variant
          
62        ReDim ListCode(0 To vbProj.VBComponents.Count - 1, 2) '.Clear
63        Set vbProj = objWB.VBProject
          
64        For iFile = 1 To vbProj.VBComponents.Count
65            ListCode(iFile - 1, 0) = iFile
66            ListCode(iFile - 1, 1) = ComponentTypeToString(vbProj.VBComponents(iFile).Type)
67            ListCode(iFile - 1, 2) = vbProj.VBComponents(iFile).Name
68        Next iFile
69        Arr = ListCode
70        Call Sort2_asc(Arr, 1)
71        ListCode = Arr
72        For iFile = 0 To vbProj.VBComponents.Count - 1
73            ListCode(iFile, 0) = iFile + 1
74        Next iFile

75        If Not objWB.Name Like "*_obf_*" Then
76            sPath = Left(objWB.FullName, Len(objWB.FullName) - Len(objWB.Name))
77            If sPath = vbNullString Then
78                Call MsgBox("The file is not saved, you need to save the file to continue: [" & objWB.Name & " ]", vbInformation, "Mistake:")
79                Exit Sub
80            End If
81            arrNameFile = Split(objWB.Name, ".")
82            objWB.SaveAs Filename:=sPath & arrNameFile(0) & "_obf_" & Replace(Now(), ":", ".") & "." & arrNameFile(1)    ', FileFormat:=wb.FileFormat
83        End If
84        j = -1

85        Set vbProj = objWB.VBProject
86        For i = 0 To vbProj.VBComponents.Count - 1
              
87            Set vbComp = vbProj.VBComponents(ListCode(i, 2))


88            Call K_AddNumbersLine.RemoveLineNumbers(vbComp, vbLineNumbers_LabelTypes.vbLabelColon)
89            Call K_AddNumbersLine.RemoveLineNumbers(vbComp, vbLineNumbers_LabelTypes.vbLabelTab)


90            Call N_Obfuscation.Remove_Comments(vbComp.CodeModule)


91            Call N_Obfuscation.TrimLinesTabAndSpase(vbComp.CodeModule)


92            Call N_Obfuscation.Remove_OptionExplicit(vbComp.CodeModule)


93            Call N_Obfuscation.Remove_EmptyLines(vbComp.CodeModule)

        
94            Call N_Obfuscation.RemoveBreaksLineInCode(vbComp.CodeModule)
          

95        Next i

96        Application.DisplayAlerts = False
97        Set objWkSh = objWB.Worksheets(NAME_SH)
98        objWkSh.Delete
99        If Form.chQuestion.Value = True Then
100           Set objWkSh = objWB.Worksheets(NAME_SH_STR)
101           objWkSh.Delete
102       End If
103       Application.DisplayAlerts = True
          
104       objWB.Save
          

105       Call MsgBox(objWB.Name & " encrypted!", vbInformation, "Code encryption:")
       
106       Set Form = Nothing
107       Application.Calculation = xlCalculationAutomatic
108       Application.ScreenUpdating = True
109       Exit Sub
ErrCompleteObfuscation:
110       Application.Calculation = xlCalculationAutomatic
111       Application.ScreenUpdating = True
112       Call MsgBox("Error in N_ObfMainNew.StartCompleteObfuscation" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
113       Call WriteErrorLog("StartCompleteObfuscation")
End Sub

Public Sub StartObfuscation()
          Dim Form        As AddStatistic
          Dim sNameWB     As String
          Dim objWB       As Object

          'On Error GoTo ErrStartParser
114       Set Form = New AddStatistic
115       With Form
116           .Caption = "Code obfuscation:"
117           .lbOK.Caption = "OBFUSCATE"
118           .chQuestion.visible = True
119           .chQuestion.Value = True
120           .lbWord.Caption = 1
121           .Show
122           sNameWB = .cmbMain.Value
123       End With
124       If sNameWB = vbNullString Then Exit Sub

125       If sNameWB Like "*.docm" Or sNameWB Like "*.DOCM" Then
              Dim objWrdApp As Object
126           Set objWrdApp = GetObject(, "Word.Application")
127           Set objWB = objWrdApp.Documents(sNameWB)
128       Else
129           Set objWB = Workbooks(sNameWB)
130       End If

131       Call MainObfuscation(objWB, Form.chQuestion.Value)
132       Set Form = Nothing
133       Exit Sub
ErrStartParser:
134       Application.Calculation = xlCalculationAutomatic
135       Application.ScreenUpdating = True
136       Call MsgBox("Error in N_ObfParserVBA.StartParser" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
137       Call WriteErrorLog("AddShapeStatistic")
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 11:30    CalDymos
Private Sub MainObfuscation(ByRef objWB As Object, Optional bEncodeStr As Boolean = False, Optional bNoFinishMessage As Boolean = False)
138       On Error GoTo ErrStartParser
139       If objWB.VBProject.Protection = vbext_pp_locked Then
140           Call MsgBox("The project is protected, remove the password!", vbCritical, "Project:")
141       Else
142           If ActiveSheet.Name = NAME_SH Then
143               Application.ScreenUpdating = False
144               Application.Calculation = xlCalculationManual
145               Application.EnableEvents = False

146               Call Obfuscation(objWB, bEncodeStr)

147               With ActiveWorkbook.Worksheets(NAME_SH)
148                   Call SortTabel(ActiveWorkbook.Worksheets(NAME_SH), .Range(.Cells(1, 1), .Cells(1, 13)).Address, "M1", 1)
149               End With

150               Application.EnableEvents = True
151               Application.Calculation = xlCalculationAutomatic
152               Application.ScreenUpdating = True
153               If Not bNoFinishMessage Then
154                   Call MsgBox("Book code [" & objWB.Name & "] encrypted!", vbInformation, "Code encryption:")
155               End If
156           Else
157               Call MsgBox("Create or navigate to the sheet: [" & NAME_SH & "]", vbCritical, "Activating the sheet:")
158           End If
159       End If
160       Exit Sub
ErrStartParser:
161       Application.EnableEvents = True
162       Application.Calculation = xlCalculationAutomatic
163       Application.ScreenUpdating = True
164       Call MsgBox("Error in MainObfuscation" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : FelterAdd - filtration in the desired order before before passing through the loops
'* Created    : 29-07-2020 09:58
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub FelterAdd()
          Dim LastRow     As Long
165       With ActiveWorkbook.Worksheets(NAME_SH)
166           LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
167           If LastRow > 1 Then
168               .Range(.Cells(2, 12), .Cells(LastRow, 12)).FormulaR1C1 = "=LEN(RC[-4])"
169               .Range(.Cells(2, 13), .Cells(LastRow, 13)).FormulaR1C1 = "=R[-1]C+1"
170               .Range(.Cells(2, 13), .Cells(LastRow, 13)).Value = .Range(.Cells(2, 13), .Cells(LastRow, 13)).Value
171               Call SortTabel(ActiveWorkbook.Worksheets(NAME_SH), .Range(.Cells(1, 1), .Cells(1, 13)).Address, "L1", 2)
172           End If
173       End With
End Sub


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : StringCrypt - Function to encrypt the strings
'* Created    : 07-09-2023 11:32
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):             Description
'*
'* ByVal Inp As String   :
'* Key As String         :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'This function must be customized by the user and represents the counter function to the encoding function stored in varStrCryptFunc
Private Function StringCrypt(ByVal Inp As String, Key As String) As String
          Dim strEn As String

174       strEn = Inp
175       'code line
176       '....
177       '....
178       '....

191      StringCrypt = strEn
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : Obfuscation - ãëàâíàÿ ïðîöåäóðà øèôðîâàíèÿ
'* Created    : 20-04-2020 18:26
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objWB As Workbook               : êíèãà
'* Optional bEncodeStr As Boolean = True : øèôðîâàòü ñòðîêîâûå çíà÷åíèÿ
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 11:42    CalDymos

Private Sub Obfuscation(ByRef objWB As Workbook, Optional bEncodeStr As Boolean = True)
          Dim arrData     As Variant
          Dim i           As Long
          Dim j           As Long
          Dim skey        As String
          Dim sCode       As String
          Dim sFinde      As String
          Dim sReplace    As String
          Dim sPattern    As String
          
          Dim k As Long
          Dim objWkSh As Worksheet
          Dim lSize As Long
          Dim bModulExist As Boolean

          Dim objDictName As Scripting.Dictionary
          Dim objDictFuncAndsub As Scripting.Dictionary
          Dim objDictModule As Scripting.Dictionary
          Dim objDictModuleOld As Scripting.Dictionary
          Dim objVBCitem  As VBIDE.VBComponent
          Dim dTime       As Date

192       dTime = Now()
193       Debug.Print "Start:" & VBA.Format$(Now() - dTime, "Long Time")

194       Set objDictName = New Scripting.Dictionary
195       Set objDictFuncAndsub = New Scripting.Dictionary
196       Set objDictModule = New Scripting.Dictionary
197       Set objDictModuleOld = New Scripting.Dictionary

          'save and Load
198       If Not sFolderHave(objWB.Path & Application.PathSeparator & OBF_RELEASE_PATH) Then MkDir (objWB.Path & Application.PathSeparator & OBF_RELEASE_PATH)
199       objWB.SaveAs Filename:=objWB.Path & Application.PathSeparator & OBF_RELEASE_PATH & Application.PathSeparator & C_PublicFunctions.sGetBaseName(objWB.FullName) & "_obf_" & Replace(Now(), ":", ".") & "." & C_PublicFunctions.sGetExtensionName(objWB.FullName)    ', FileFormat:=objWB.FileFormat

200       Debug.Print "File saving - completed:" & VBA.Format$(Now() - dTime, "Long Time")
          
          ' Insert string encryption function
201       If bEncodeStr Then
202           Set objWkSh = ActiveWorkbook.Worksheets(NAME_SH_STR)
203           If Not objWkSh Is Nothing And IsArray(varStrCryptFunc) Then
204               For Each objVBCitem In objWB.VBProject.VBComponents
205                   If objVBCitem.Name = objWkSh.Cells(2, 17).Value Then
206                       bModulExist = True
207                   End If
208               Next
209               If Not bModulExist Then
210                   Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
211                   objVBCitem.Name = objWkSh.Cells(2, 17).Value
212               End If
                  
                  
213               With objWB.VBProject.VBComponents(objWkSh.Cells(2, 17).Value).CodeModule
214                   lSize = .CountOfLines + 1
215                   For i = 0 To UBound(varStrCryptFunc)
216                       .InsertLines i + lSize, varStrCryptFunc(i)
217                   Next
218               End With
219           End If

          
              ' Insert constant with the key
220           If Not objWkSh Is Nothing Then
221               For Each objVBCitem In objWB.VBProject.VBComponents
222                   If objVBCitem.Name = objWkSh.Cells(2, 14).Value Then
223                       bModulExist = True
224                   End If
225               Next
226               If Not bModulExist Then
227                   Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
228                   objVBCitem.Name = objWkSh.Cells(2, 14).Value
229               End If
                  
                  
230               With objWB.VBProject.VBComponents(objWkSh.Cells(2, 14).Value).CodeModule
231                   lSize = .CountOfLines
232                   For i = 1 To lSize
233                       If Trim$(.Lines(i, 1)) = "Option Explicit" Then
234                       ElseIf Trim$(.Lines(i, 1)) = "Option Base 1" Then
235                       ElseIf Trim$(.Lines(i, 1)) = "Option Private Module" Then
236                       ElseIf Trim$(.Lines(i, 1)) = "Option Compare Text" Then
237                       Else
238                           .InsertLines i, "Public Const " & asCryptKey(0) & " As String = " & Chr$(34) & asCryptKey(1) & Chr$(34)
239                           Exit For
240                       End If
241                   Next
242               End With
243           End If
244       End If
          'Filtern
245       Call FelterAdd

          'Einlesen der Daten
246       With ActiveWorkbook.Worksheets(NAME_SH)
247           .Activate
248           i = .Cells(Rows.Count, 1).End(xlUp).Row
249           arrData = .Range(Cells(2, 1), Cells(i, 10)).Value2
250       End With

          'Sammlung verschlüsselter Namen und Subs / Functions
251       For i = LBound(arrData) To UBound(arrData)
252           If arrData(i, 9) = "yes" Then
                  'Sammlung verschlüsselter Namen
253               If objDictName.Exists(arrData(i, 8)) = False Then objDictName.Add arrData(i, 8), arrData(i, 10)
                  'Sammlung der Subs und Functions
254               If objDictFuncAndsub.Exists(arrData(i, 6)) = False Then objDictFuncAndsub.Add arrData(i, 6), arrData(i, 5)
255           End If
256       Next i

          'Codesammlung aus Modulen
257       For Each objVBCitem In objWB.VBProject.VBComponents
258           If objDictModule.Exists(objVBCitem.Name) = False Then
259               sCode = GetCodeFromModule(objVBCitem)
                  'Beseitigung von Zeilenumbrüchen
260               sCode = VBA.Replace(sCode, " _" & vbNewLine, " XXXXX") 'changed : am 24.04 CalDymos
261               objDictModule.Add objVBCitem.Name, sCode
262               objDictModuleOld.Add objVBCitem.Name, sCode
263               sCode = vbNullString
264           End If
265       Next objVBCitem
          'Ende der Sammlung

266       Debug.Print "Data collection - completed:" & VBA.Format$(Now() - dTime, "Long Time")

          'Schleifen
267       sCode = vbNullString
268       With objDictName
269           For i = 0 To .Count - 1
270               For j = 0 To objDictModule.Count - 1
271                   sFinde = .Keys(i)
272                   sReplace = .Items(i)
273                   If InStr(sFinde, "WName") <> 0 Then
274                       Debug.Print "Found"
275                   End If
276                   skey = objDictModule.Keys(j)
277                   sCode = objDictModule.Item(skey)
278                   If sCode Like "*" & sFinde & "*" And VBA.Len(sFinde) > 1 Then
                          '------------------------------------------------ changed: 31.08 CalDymos
279                       sPattern = "([\*\.\^\*\+\#\(\)\-\=\/\,\:\;\s])" & sFinde & "([\*\.\^\*\+\!\@\#\$\%\&\(\)\-\=\/\,\:\;\s]|$)"
280                       sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
281                       If InStr(sCode, "Application.OnTime") <> 0 And InStr(sCode, sFinde) <> 0 Then
282                           If objDictFuncAndsub.Exists(sFinde) Then
283                               sPattern = "([\" & VBA.Chr$(34) & "])" & sFinde & "([\" & VBA.Chr$(34) & "]|$)"
284                               sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
285                               If WorksheetExist(NAME_SH_STR, objWB) Then
                                      'Replace the string in the Excel sheet with coded
286                                   With ActiveWorkbook.Worksheets(NAME_SH_STR)
287                                       .Activate
288                                       For k = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
289                                           If .Cells(k, 2).Value2 = skey And .Cells(k, 5).Value = Chr$(34) & sFinde & Chr$(34) Then
290                                               .Cells(k, 5).Value = Chr$(34) & sReplace & Chr$(34)
291                                           End If
292                                       Next
293                                   End With
294                               End If
295                           End If
296                       End If
                          '------------------------------------------------
297                       If sCode <> vbNullString Then objDictModule.Item(skey) = sCode
298                   End If
                      'Regulierungsrahmen für Events, vor allem für Formulare
299                   If sCode Like "* " & Chr$(83) & "ub *" & sFinde & "_*(*)*" Then
300                       sPattern = "([\s])(Sub)([\s])" & sFinde & "(\_{1}[A-Za-zÀ-ßà-ÿ¨¸]{4,40}\([A-Za-zÀ-ßà-ÿ¨¸\s\.\,]{0,100}\))"
301                       sCode = RegExpFindReplace(sCode, sPattern, "$1$2$3" & sReplace & "$4", True, False, False)
302                       If sCode <> vbNullString Then objDictModule.Item(skey) = sCode
303                       sPattern = "([\s])" & sFinde & "(\_{1}[A-Za-zÀ-ßà-ÿ¨¸]{4,40}(?:\:\s|\n|\r))"
304                       sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
305                       If sCode <> vbNullString Then objDictModule.Item(skey) = sCode
306                   End If
307                   sCode = vbNullString
308               Next j
309               DoEvents
310               If .Count > 1 Then
311                   Application.StatusBar = "Data encryption - completed:" & Format(i / (.Count - 1), "Percent") & ", " & i & "from" & .Count - 1
312               Else
313                   Application.StatusBar = "Data encryption - completed:" & Format(i / .Count, "Percent") & ", " & i & "from" & .Count
314               End If
315           Next i
316       End With
317       Application.StatusBar = False
          'Ende

          'Übertragung
318       sCode = vbNullString

319       For j = 0 To objDictModule.Count - 1
              Dim arrNew  As Variant
              Dim arrOld  As Variant
              Dim sTemp   As String
320           arrNew = VBA.Split(objDictModule.Items(j), vbNewLine)
321           arrOld = VBA.Split(objDictModuleOld.Items(j), vbNewLine)
322           For i = LBound(arrNew) To UBound(arrNew)
323               If arrNew(i) = vbNullString Or VBA.Left$(VBA.Trim$(arrNew(i)), 1) = "'" Then
324                   sTemp = vbNullString
325               Else
326                   sTemp = "'" & arrOld(i) & vbNewLine
327               End If
328               sCode = sCode & sTemp & arrNew(i) & vbNewLine
329               sTemp = vbNullString
330           Next i
331           skey = objDictModule.Keys(j)
332           objDictModule.Item(skey) = sCode
333           sCode = vbNullString
334       Next j
335       Debug.Print "String alternation - completed:" & VBA.Format$(Now() - dTime, "Long Time")


336       Debug.Print "Data encryption - completed:" & VBA.Format$(Now() - dTime, "Long Time")
          'code laden
337       For j = 0 To objDictModule.Count - 1
338           Set objVBCitem = objWB.VBProject.VBComponents(objDictModule.Keys(j))
339           sCode = objDictModule.Items(j)
              'âîçâðàò ïåðåíîñ ñòðîê
340           sCode = VBA.Replace(sCode, " XXXXX", " _" & vbNewLine) 'changed : am 24.04 CalDymos
341           Call SetCodeInModule(objVBCitem, sCode)
342       Next j

343       Debug.Print "Code loading- completed:" & VBA.Format$(Now() - dTime, "Long Time")

          'Controls umbenennen
344       For i = LBound(arrData) To UBound(arrData)
345           If arrData(i, 9) = "yes" And objDictName.Exists(arrData(i, 8)) Then
346               If arrData(i, 1) = "Control" Then
347                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
348                   objVBCitem.Designer.Controls(arrData(i, 8)).Name = arrData(i, 10)
349               End If
350           End If
351       Next i

352       Debug.Print "Renaming of controls - completed:" & VBA.Format$(Now() - dTime, "Long Time")

          'Ändern von Modulen
353       For i = LBound(arrData) To UBound(arrData)
354           If arrData(i, 9) = "yes" And objDictName.Exists(arrData(i, 8)) Then
355               If arrData(i, 1) = "Module" And VBA.CByte(arrData(i, 2)) <> 100 Then
356                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
357                   objVBCitem.Name = arrData(i, 10)
358               ElseIf arrData(i, 1) = "Module" And VBA.CByte(arrData(i, 2)) = 100 Then
359                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
360                   objVBCitem.Name = arrData(i, 10)
361               End If
362           End If
363       Next i

          'øèôðîâàíèå ñòðîê
364       If bEncodeStr Then Call EncodedStringCode(objWB)

365       Debug.Print "Renaming modules- completed:" & VBA.Format$(Now() - dTime, "Long Time")
366       objWB.Save

End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : EncodedStringCode - øèôðîâàíèå ñòðîêîâûé çíà÷åíèé êîäà
'* Created    : 29-07-2020 10:00
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef objWB As Workbook :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 11:51    CalDymos

Private Sub EncodedStringCode(ByRef objWB As Workbook)
          Dim arrData     As Variant
          Dim i           As Long
          Dim k As Long
          Dim sCodeString As String
          Dim objVBCitem  As VBIDE.VBComponent
          Dim sCode       As String
          Dim varStr() As String
          Dim strCryptFuncCipher As String
          Dim strCryptKeyCipher As String
          Dim strKey As String

          'Datenverarbeitung
367       With ActiveWorkbook.Worksheets(NAME_SH_STR)
368           .Activate
369           arrData = .Range(Cells(2, 1), Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 9)).Value2
370           strCryptFuncCipher = .Cells(2, 18).Value
371           strCryptKeyCipher = .Cells(2, 15).Value
372           strKey = .Cells(2, 13).Value
373       End With
          'Zeilenzusammenstellung
374       sCodeString = "Option Explicit" & VBA.Chr$(13)
375       For i = LBound(arrData) To UBound(arrData)
376           If arrData(i, 7) = "yes" Then
377               sCodeString = sCodeString & "Public Const " & arrData(i, 8) & " as String = " & Chr$(34) & StringCrypt(TrimA(arrData(i, 5), Chr$(34)), strKey, True) & Chr$(34) & VBA.Chr$(13)
378           End If
379       Next i
          Dim NameOldMOdule As String
380       For i = LBound(arrData) To UBound(arrData)
381           If arrData(i, 7) = "yes" Then
382               If NameOldMOdule <> arrData(i, 9) Then
383                   sCode = vbNullString
384                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 9))
385                   sCode = GetCodeFromModule(objVBCitem)
386                   If sCode <> vbNullString Then
387                       varStr = VBA.Split(sCode, vbNewLine)
388                       For k = 0 To UBound(varStr())
389                           If CStr(varStr(k)) = "" Then
                                  'Do Nothing
390                           ElseIf Left$(CStr(varStr(k)), 1) = "'" Then
                                  'Do Nothing
391                           ElseIf IsSubOrFunc(varStr(k)) Then
                                  'Do Nothing
392                           ElseIf VBA.InStr(1, CStr(varStr(k)), arrData(i, 5)) <> 0 Then
393                               varStr(k) = VBA.Trim$(VBA.Replace(CStr(varStr(k)), arrData(i, 5), strCryptFuncCipher & "(" & arrData(i, 8) & ", " & strCryptKeyCipher & ")"))
394                           End If
395                       Next
396                       sCode = Join(varStr, vbNewLine)
397                   End If
398                   NameOldMOdule = arrData(i, 9)
399               Else
400                   If sCode <> vbNullString Then
401                       varStr = VBA.Split(sCode, vbNewLine)
402                       For k = 0 To UBound(varStr)
403                           If CStr(varStr(k)) = "" Then
                                  'Do Nothing
404                           ElseIf Left$(CStr(varStr(k)), 1) = "'" Then
                                  'Do Nothing
405                           ElseIf IsSubOrFunc(varStr(k)) Then
                                  'Do Nothing
406                           ElseIf VBA.InStr(1, CStr(varStr(k)), arrData(i, 5)) <> 0 Then
407                               varStr(k) = VBA.Trim$(VBA.Replace(CStr(varStr(k)), arrData(i, 5), strCryptFuncCipher & "(" & arrData(i, 8) & ", " & strCryptKeyCipher & ")"))
408                           End If
409                       Next
410                       sCode = Join(varStr, vbNewLine)
411                   End If
412               End If
413               If i = UBound(arrData) Then
414                   Call SetCodeInModule(objVBCitem, sCode)
415                   Set objVBCitem = Nothing
416               Else
417                   If arrData(i + 1, 9) <> arrData(i, 9) Then
418                       Call SetCodeInModule(objVBCitem, sCode)
419                       Set objVBCitem = Nothing
420                   End If
421               End If

422           End If
423           DoEvents
424           If i Mod 100 = 0 Then Application.StatusBar = "String encryption - completed:" & Format(i / UBound(arrData), "Percent") & ", " & i & "from" & UBound(arrData)
425       Next i
426       Application.StatusBar = False
          Dim sName       As String
427       sName = ActiveWorkbook.Worksheets(NAME_SH_STR).Cells(2, 11).Value
428       If sName <> vbNullString Then
429           Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
430           objVBCitem.Name = sName
431           Call SetCodeInModule(objVBCitem, sCodeString)
432       End If
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetCodeFromModule - ïîëó÷èòü êîä èç ìîäóëÿ â ñòðîêîâóþ ïåðåìåííóþ
'* Created    : 20-04-2020 18:20
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objVBComp As VBIDE.VBComponent : ìîäóëü VBA
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function GetCodeFromModule(ByRef objVBComp As VBIDE.VBComponent) As String
433       GetCodeFromModule = vbNullString
434       With objVBComp.CodeModule
435           If .CountOfLines > 0 Then
436               GetCodeFromModule = .Lines(1, .CountOfLines)
437           End If
438       End With
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SetCodeInModule çàãðóçèòü êîä èç ñòðîêîâîé ïåðåìåíîé â ìîäóëü
'* Created    : 20-04-2020 18:21
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objVBComp As VBIDE.VBComponent : ìîäóëü VBA
'* ByVal SCode As String                : ñòðîêîâàÿ ïåðåìåííàÿ
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub SetCodeInModule(ByRef objVBComp As VBIDE.VBComponent, ByVal sCode As String)
439       With objVBComp.CodeModule
440           If .CountOfLines > 0 Then
                  'Debug.Print .CountOfLines
441               Call .DeleteLines(1, .CountOfLines)
                  'Debug.Print sCode
442               Call .InsertLines(1, VBA.Trim$(sCode))
443           End If
444       End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SortTabel - ñîðòèðîâêà äèàïàçîíà äàííûõ
'* Created    : 29-07-2020 10:03
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByRef WS As Worksheet       :
'* ByVal sRng As String        :
'* sKey1 As String             :
'* Optional bOrder As Byte = 2 :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub SortTabel(ByRef ws As Worksheet, ByVal sRng As String, sKey1 As String, Optional bOrder As Byte = 2)
445       With ws
446           On Error GoTo errMsg
447           .Activate
448           .Range(sRng).AutoFilter
Repeatnext:
449           .AutoFilter.Sort.SortFields.Clear
450           .AutoFilter.Sort.SortFields.Add Key:=Range(sKey1), SortOn:=xlSortOnValues, Order:=bOrder, DataOption:=xlSortNormal
451           With .AutoFilter.Sort
452               .Header = xlYes
453               .MatchCase = False
454               .Orientation = xlTopToBottom
455               .SortMethod = xlPinYin
456               .Apply
457           End With
458       End With
459       Exit Sub
errMsg:
460       If Err.Number = 91 Then
461           ws.Range(sRng).AutoFilter
462           Err.Clear
463           GoTo Repeatnext
464       Else
465           Call MsgBox(Err.Description, vbCritical, "Mistake:")
466       End If
End Sub
