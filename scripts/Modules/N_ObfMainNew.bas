Attribute VB_Name = "N_ObfMainNew"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : N_ObfMainNew - Modul zur Code-Verschleierung
'* Created    : 08-10-2020 14:11
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 11:26    CalDymos
'* Updated    : 12-09-2023 13:29    CalDymos
'* Updated    : 13-09-2023 13:29    CalDymos

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

Private Sub Sort2_asc(arr(), col As Long)
          Dim temp()      As Variant
          Dim lb2 As Long, ub2 As Long, lTop As Long, lBot As Long

8         lTop = LBound(arr, 1)
9         lBot = UBound(arr, 1)
10        lb2 = LBound(arr, 2)
11        ub2 = UBound(arr, 2)
12        ReDim temp(lb2 To ub2)

13        Call QSort2_asc(arr(), col, lTop, lBot, temp(), lb2, ub2)
End Sub
Private Sub QSort2_asc(arr(), C As Long, ByVal top As Long, ByVal bot As Long, temp(), lb2 As Long, ub2 As Long)
          Dim t As Long, LB As Long, MidItem, j As Long

14        MidItem = arr((top + bot) \ 2, C)
15        t = top: LB = bot

16        Do
17            Do While arr(t, C) < MidItem: t = t + 1: Loop
18            Do While arr(LB, C) > MidItem: LB = LB - 1: Loop
19            If t < LB Then
20                For j = lb2 To ub2: temp(j) = arr(t, j): Next j
21                For j = lb2 To ub2: arr(t, j) = arr(LB, j): Next j
22                For j = lb2 To ub2: arr(LB, j) = temp(j): Next j
23                t = t + 1: LB = LB - 1
24            ElseIf t = LB Then
25                t = t + 1: LB = LB - 1
26            End If
27        Loop While t <= LB

28        If t < bot Then QSort2_asc arr(), C, t, bot, temp(), lb2, ub2
29        If top < LB Then QSort2_asc arr(), C, top, LB, temp(), lb2, ub2

End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : StartCompleteObfuscation - Performs a complete obfuscation, including removing all comments and formatting.
'* Created    : 07-09-2023 07:02
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Modified   : Date and Time       Author              Description
'* Updated    : 13-09-2023 14:48    CalDymos
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
          Dim arr()       As Variant
          Dim ListCode() As Variant
          
62        ReDim ListCode(0 To vbProj.VBComponents.Count - 1, 2) '.Clear
63        Set vbProj = objWB.VBProject
          
64        For iFile = 1 To vbProj.VBComponents.Count
65            ListCode(iFile - 1, 0) = iFile
66            ListCode(iFile - 1, 1) = ComponentTypeToString(vbProj.VBComponents(iFile).Type)
67            ListCode(iFile - 1, 2) = vbProj.VBComponents(iFile).Name
68        Next iFile
69        arr = ListCode
70        Call Sort2_asc(arr, 1)
71        ListCode = arr
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
99        Set objWkSh = objWB.Worksheets(NAME_SH_CTL)
100       objWkSh.Delete
101       If Form.chQuestion.Value = True Then
102           Set objWkSh = objWB.Worksheets(NAME_SH_STR)
103           objWkSh.Delete
104       End If
105       Application.DisplayAlerts = True
          
106       objWB.Save
          

107       Call MsgBox(objWB.Name & " encrypted!", vbInformation, "Code encryption:")
       
108       Set Form = Nothing
109       Application.Calculation = xlCalculationAutomatic
110       Application.ScreenUpdating = True
111       Exit Sub
ErrCompleteObfuscation:
112       Application.Calculation = xlCalculationAutomatic
113       Application.ScreenUpdating = True
114       Call MsgBox("Error in N_ObfMainNew.StartCompleteObfuscation" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
115       Call WriteErrorLog("StartCompleteObfuscation")
End Sub

Public Sub StartObfuscation()
          Dim Form        As AddStatistic
          Dim sNameWB     As String
          Dim objWB       As Object

          'On Error GoTo ErrStartParser
116       Set Form = New AddStatistic
117       With Form
118           .Caption = "Code obfuscation:"
119           .lbOK.Caption = "OBFUSCATE"
120           .chQuestion.visible = True
121           .chQuestion.Value = True
122           .lbWord.Caption = 1
123           .Show
124           sNameWB = .cmbMain.Value
125       End With
126       If sNameWB = vbNullString Then Exit Sub

127       If sNameWB Like "*.docm" Or sNameWB Like "*.DOCM" Then
              Dim objWrdApp As Object
128           Set objWrdApp = GetObject(, "Word.Application")
129           Set objWB = objWrdApp.Documents(sNameWB)
130       Else
131           Set objWB = Workbooks(sNameWB)
132       End If

133       Call MainObfuscation(objWB, Form.chQuestion.Value)
134       Set Form = Nothing
135       Exit Sub
ErrStartParser:
136       Application.Calculation = xlCalculationAutomatic
137       Application.ScreenUpdating = True
138       Call MsgBox("Error in N_ObfParserVBA.StartParser" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
139       Call WriteErrorLog("AddShapeStatistic")
End Sub

'* Modified   : Date and Time       Author              Description
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 11:30    CalDymos
'* Updated    : 12-09-2023 13:31    CalDymos
Private Sub MainObfuscation(ByRef objWB As Workbook, Optional bEncodeStr As Boolean = False, Optional bNoFinishMessage As Boolean = False)
140       On Error GoTo ErrStartParser
141       If objWB.VBProject.Protection = vbext_pp_locked Then
142           Call MsgBox("The project is protected, remove the password!", vbCritical, "Project:")
143       Else
144           If objWB.ActiveSheet.Name = NAME_SH Then
145               Application.ScreenUpdating = False
146               Application.Calculation = xlCalculationManual
147               Application.EnableEvents = False

148               If Obfuscation(objWB, bEncodeStr) Then
                      
                      
149                   With objWB.Worksheets(NAME_SH)
150                       Call SortTabel(objWB.Worksheets(NAME_SH), .Range(.Cells(1, 1), .Cells(1, 13)).Address, "M1", 1)
151                   End With

152                   Application.EnableEvents = True
153                   Application.Calculation = xlCalculationAutomatic
154                   Application.ScreenUpdating = True
155                   If Not bNoFinishMessage Then
156                       Call MsgBox("Book code [" & objWB.Name & "] encrypted!", vbInformation, "Code encryption:")
157                   End If
158               Else
159                   Application.EnableEvents = True
160                   Application.Calculation = xlCalculationAutomatic
161                   Application.ScreenUpdating = True
162               End If
163           Else
164               Call MsgBox("Create or navigate to the sheet: [" & NAME_SH & "]", vbCritical, "Activating the sheet:")
165           End If
166       End If
167       Exit Sub
ErrStartParser:
168       Application.EnableEvents = True
169       Application.Calculation = xlCalculationAutomatic
170       Application.ScreenUpdating = True
171       Call MsgBox("Error in MainObfuscation" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
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
172       With ActiveWorkbook.Worksheets(NAME_SH)
173           LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
174           If LastRow > 1 Then
175               .Range(.Cells(2, 12), .Cells(LastRow, 12)).FormulaR1C1 = "=LEN(RC[-4])"
176               .Range(.Cells(2, 13), .Cells(LastRow, 13)).FormulaR1C1 = "=R[-1]C+1"
177               .Range(.Cells(2, 13), .Cells(LastRow, 13)).Value = .Range(.Cells(2, 13), .Cells(LastRow, 13)).Value
178               Call SortTabel(ActiveWorkbook.Worksheets(NAME_SH), .Range(.Cells(1, 1), .Cells(1, 13)).Address, "L1", 2)
179           End If
180       End With
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
'* ByVal Mode As Boolean :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'This function must be customized by the user and represents the counter function to the encoding function stored in varCryptFunc

Private Function StringCrypt(ByVal Inp As String, Key As String, ByVal Mode As Boolean) As String
          Dim z As String
          Dim i As Integer, Position As Integer
          Dim cptZahl As Long, orgZahl As Long
          Dim keyZahl As Long, cptString As String
          
181       For i = 1 To Len(Inp)
182           Position = Position + 1
183           If Position > Len(Key) Then Position = 1
184           keyZahl = Asc(Mid(Key, Position, 1))
                  
185           If Mode Then
                  
                  'Verschlьsseln
186               orgZahl = Asc(Mid(Inp, i, 1))
187               cptZahl = orgZahl Xor keyZahl
188               cptString = Hex(cptZahl)
189               If Len(cptString) < 2 Then cptString = "0" & cptString
190               z = z & cptString
                  
191           Else
                  
                  'Entschlьsseln
192               If i > Len(Inp) \ 2 Then Exit For
193               cptZahl = CByte("&H" & Mid$(Inp, i * 2 - 1, 2))
194               orgZahl = cptZahl Xor keyZahl
195               z = z & Chr$(orgZahl)
                  
196           End If
197       Next i
           
198       StringCrypt = z
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddProcsToVBProject
'* Created    : 13-09-2023 13:32
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):             Description
'*
'* ByRef objWB As Workbook :
'* AddProcs(               :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub AddProcsToVBProject(ByRef objWB As Workbook, AddProcs() As CAddProc)
          Dim k As Long
          Dim i As Long
          Dim objVBCitem  As VBIDE.VBComponent
          Dim objProc As CAddProc
          Dim sCodeLines As String
          Dim startLine As Long
          Dim numOfLines As Long

199       If Not IsArrayEmpty(AddProcs()) Then
200           For i = 0 To UBound(AddProcs())
201               Set objProc = AddProcs(i)
202               If Not CodeModuleExist(objWB, objProc.ModuleName) Then
203                   Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
204                   objVBCitem.Name = objProc.ModuleName
205               End If
                  
206               With objWB.VBProject.VBComponents(objProc.ModuleName).CodeModule
207                   If Not ProcInCodeModuleExist(objWB.VBProject.VBComponents(objProc.ModuleName).CodeModule, objProc.Name) Then
208                       sCodeLines = Join(objProc.CodeLines, vbCrLf)
209                       .AddFromString sCodeLines
210                   ElseIf objProc.BehavProcExists = enumBehavProcExistInsCodeAtBegin Then
211                       startLine = .ProcBodyLine(objProc.Name, vbext_pk_Proc) + 1
          
212                       For k = 1 To UBound(objProc.CodeLines) - 1
213                           .InsertLines startLine, objProc.GetCodeLine(k)
214                           startLine = startLine + 1
215                       Next
216                   ElseIf objProc.BehavProcExists = enumBehavProcExistInsCodeAtEnd Then
217                       startLine = .ProcStartLine(objProc.Name, vbext_pk_Proc) + .ProcCountLines(objProc.Name, vbext_pk_Proc) - 1
          
218                       For k = 1 To UBound(objProc.CodeLines) - 1
219                           .InsertLines startLine, objProc.GetCodeLine(k)
220                           startLine = startLine + 1
221                       Next
222                   End If
223               End With
224           Next
225       End If
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddGlobalVarsToVBProject
'* Created    : 13-09-2023 13:32
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):             Description
'*
'* ByRef objWB As Workbook :
'* AddGlobalVars(          :
'*
'* Modified   : Date and Time       Author              Description
'* Updated    : 13-09-2023 14:48    CalDymos
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub AddGlobalVarsToVBProject(ByRef objWB As Workbook, AddGlobalVars() As CAddGlobalVar)

          Dim k As Long
          Dim i As Long
          Dim objVBCitem  As VBIDE.VBComponent
          Dim objVar As CAddGlobalVar
          Dim sCodeLines As String
          Dim lSize As Long
              
226       If Not IsArrayEmpty(AddGlobalVars()) Then
227           For i = 0 To UBound(AddGlobalVars())
228               Set objVar = AddGlobalVars(i)
229               If Not CodeModuleExist(objWB, objVar.ModuleName) Then
230                   Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
231                   objVBCitem.Name = objVar.ModuleName
232               End If
                  
                  
233               With objWB.VBProject.VBComponents(objVar.ModuleName).CodeModule
234                   sCodeLines = vbNullString
235                   Select Case objVar.Visibility
                          Case enumVisibility.enumVisibilityPublic
236                           sCodeLines = sCodeLines & "Public "
237                       Case enumVisibility.enumVisibilityPrivate
238                           sCodeLines = sCodeLines & "Private "
239                   End Select
240                   If objVar.IsConstant Then sCodeLines = sCodeLines & "Const "
241                   sCodeLines = sCodeLines & objVar.Name & " As "
242                   Select Case objVar.DateType
                          Case enumDataType.enumDataTypeString
243                           sCodeLines = sCodeLines & "String "
244                       Case enumDataType.enumDataTypeInt
245                           sCodeLines = sCodeLines & "Integer "
246                       Case enumDataType.enumDataTypeLong
247                           sCodeLines = sCodeLines & "Long "
248                       Case enumDataType.enumDataTypeBool
249                           sCodeLines = sCodeLines & "Boolean "
250                       Case enumDataType.enumDataTypeByte
251                           sCodeLines = sCodeLines & "Byte "
252                   End Select
253                   If objVar.IsConstant Then
254                       sCodeLines = sCodeLines & "= "
255                       Select Case objVar.DateType
                              Case enumDataType.enumDataTypeString
256                               sCodeLines = sCodeLines & Chr$(34) & objVar.Value & Chr$(34)
257                           Case enumDataType.enumDataTypeInt
258                               sCodeLines = sCodeLines & CStr(CInt(objVar.Value))
259                           Case enumDataType.enumDataTypeLong
260                               sCodeLines = sCodeLines & CStr(CLng(objVar.Value))
261                           Case enumDataType.enumDataTypeBool
262                               sCodeLines = sCodeLines & CStr(CBool(objVar.Value))
263                           Case enumDataType.enumDataTypeByte
264                               sCodeLines = sCodeLines & CStr(CByte(objVar.Value))
265                       End Select
266                   End If
                                    
267                   If .Find("#IF", 1, 1, .CountOfLines - 1, 500) Then 'Check for compiler #If, because with compiler #If,  .AddFromString does not work correctly
268                       lSize = .CountOfLines
269                       For k = 1 To lSize
270                           If Trim$(.Lines(k, 1)) = "Option Explicit" Then
271                           ElseIf Trim$(.Lines(k, 1)) = "Option Base 1" Then
272                           ElseIf Trim$(.Lines(k, 1)) = "Option Private Module" Then
273                           ElseIf Trim$(.Lines(k, 1)) = "Option Compare Text" Then
274                           Else
275                               .InsertLines k, sCodeLines
                                  Exit For
276                           End If
277                       Next
278                   Else
279                       .AddFromString sCodeLines
280                   End If
281                   sCodeLines = vbNullString
282               End With
283           Next
284       End If
          
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ObfuscateCtlProperties
'* Created    : 13-09-2023 13:32
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):             Description
'*
'* ByRef objWB As Workbook :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ObfuscateCtlProperties(ByRef objWB As Workbook)
          Dim arrData As Variant
          Dim i As Long
          Dim vbc As VBIDE.VBComponent
          Dim objCtl     As MSForms.control
          Dim objTxtBox As MSForms.TextBox
          Dim objLbl As MSForms.Label
          Dim objCmdBtn As MSForms.CommandButton
          Dim objFrame As MSForms.Frame
          Dim objChkBox As MSForms.CheckBox
          Dim objOptBtn As MSForms.OptionButton
          Dim objTglBtn As MSForms.ToggleButton
          
          'Einlesen der Daten
285       With objWB.Worksheets(NAME_SH_CTL)
286           .Activate
287           i = .Cells(Rows.Count, 1).End(xlUp).Row
288           arrData = .Range(Cells(2, 1), Cells(i, 5)).Value2
289       End With
          
290       For i = LBound(arrData) To UBound(arrData)
          

291           If arrData(i, 1) = "UserForm" Then
292               For Each vbc In objWB.VBProject.VBComponents
293                   If Not vbc.Designer Is Nothing And vbc.Type = vbext_ct_MSForm Then
294                       If vbc.Name = arrData(i, 2) Then
295                           vbc.Properties(arrData(i, 4)) = ""
296                           Exit For
297                       End If
298                   End If
299               Next
300           Else
301               For Each vbc In objWB.VBProject.VBComponents
302                   If Not vbc.Designer Is Nothing And vbc.Type = vbext_ct_MSForm Then
303                       If vbc.Name = arrData(i, 2) Then
304                           For Each objCtl In vbc.Designer.Controls
305                               If objCtl.Name = arrData(i, 3) Then
306                                   Select Case arrData(i, 4)
                                          Case "ControlTipText"
307                                           objCtl.ControlTipText = ""
308                                       Case "Tag"
309                                           objCtl.Tag = ""
310                                       Case "Text"
311                                           Set objTxtBox = objCtl
312                                           objTxtBox.Text = ""
313                                       Case "Caption"
314                                           Select Case arrData(i, 1)
                                                  Case "CommandButton"
315                                                   Set objCmdBtn = objCtl
316                                                   objCmdBtn.Caption = ""
317                                               Case "Label"
318                                                   Set objLbl = objCtl
319                                                   objLbl.Caption = ""
320                                               Case "Frame"
321                                                   Set objFrame = objCtl
322                                                   objFrame.Caption = ""
323                                               Case "CheckBox"
324                                                   Set objChkBox = objCtl
325                                                   objChkBox.Caption = ""
326                                               Case "OptionButton"
327                                                   Set objOptBtn = objCtl
328                                                   objOptBtn.Caption = ""
329                                               Case "ToggleButton"
330                                                   Set objTglBtn = objCtl
331                                                   objTglBtn.Caption = ""
332                                           End Select
333                                       Case "Width"
334                                           objCtl.Width = 0
335                                       Case "Height"
336                                           objCtl.Height = 0
337                                       Case "Left"
338                                           objCtl.Left = -32768
339                                       Case "Top"
340                                           objCtl.top = -32768
341                                   End Select
342                                   Exit For
343                               End If
344                           Next
345                           Exit For
346                       End If
347                   End If
348               Next
349           End If
             
        
350       Next i
          
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : Obfuscation - главная процедура шифрования
'* Created    : 20-04-2020 18:26
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objWB As Workbook               : книга
'* Optional bEncodeStr As Boolean = True : шифровать строковые значения
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 11:42    CalDymos
'* Updated    : 13-09-2023 13:32    CalDymos

Private Function Obfuscation(ByRef objWB As Workbook, Optional bEncodeStr As Boolean = True) As Boolean
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

          Dim objDictName As Scripting.Dictionary
          Dim objDictFuncAndsub As Scripting.Dictionary
          Dim objDictModule As Scripting.Dictionary
          Dim objDictModuleOld As Scripting.Dictionary
          Dim objVBCitem  As VBIDE.VBComponent
          Dim dTime       As Date
          
          Dim sCodeLines As String
                    
351       dTime = Now()
352       Debug.Print "Start:" & VBA.Format$(Now() - dTime, "Long Time")

353       Set objDictName = New Scripting.Dictionary
354       Set objDictFuncAndsub = New Scripting.Dictionary
355       Set objDictModule = New Scripting.Dictionary
356       Set objDictModuleOld = New Scripting.Dictionary
          
357       If bEncodeStr Then
358           If Not IsArrayEmpty(CryptFunc()) And Not IsArrayEmpty(CryptKey()) Then
359               If UBound(CryptFunc()) = -1 Or UBound(CryptKey()) = -1 Then
360                   Call MsgBox("Please start the VBA code parser before !", vbCritical, "Run VBA code parser")
361                   Exit Function
362               End If
363           Else
364               Call MsgBox("Please start the VBA code parser before !", vbCritical, "Run VBA code parser")
365               Exit Function
366           End If
367       End If
          
368       If Not IsArrayEmpty(AddCtrlProp()) Then
369           If UBound(AddCtrlProp()) = -1 Then
370               Call MsgBox("Please start the VBA code parser before !", vbCritical, "Run VBA code parser")
371               Exit Function
372           Else
373           End If
374       Else
375           Call MsgBox("Please start the VBA code parser before !", vbCritical, "Run VBA code parser")
376           Exit Function
377       End If
          
          'save and Load
378       If Not sFolderHave(objWB.Path & Application.PathSeparator & OBF_RELEASE_PATH) Then MkDir (objWB.Path & Application.PathSeparator & OBF_RELEASE_PATH)
379       objWB.SaveAs Filename:=objWB.Path & Application.PathSeparator & OBF_RELEASE_PATH & Application.PathSeparator & C_PublicFunctions.sGetBaseName(objWB.FullName) & "_obf_" & Replace(Now(), ":", ".") & "." & C_PublicFunctions.sGetExtensionName(objWB.FullName)    ', FileFormat:=objWB.FileFormat

380       Debug.Print "File saving - completed:" & VBA.Format$(Now() - dTime, "Long Time")
              
381       If bEncodeStr Then
              ' Insert constant with the key
382           AddGlobalVarsToVBProject objWB, CryptKey()
              
              ' Insert string encryption function
383           AddProcsToVBProject objWB, CryptFunc()
384       End If
          
          'Insert Control Properties
385       AddProcsToVBProject objWB, AddCtrlProp()
          
          'Obfuscate Control Properties
386       ObfuscateCtlProperties objWB
          
          'Filtern
387       Call FelterAdd

          'Einlesen der Daten
388       With objWB.Worksheets(NAME_SH)
389           .Activate
390           i = .Cells(Rows.Count, 1).End(xlUp).Row
391           arrData = .Range(Cells(2, 1), Cells(i, 10)).Value2
392       End With

          'Sammlung verschlьsselter Namen und Subs / Functions
393       For i = LBound(arrData) To UBound(arrData)
394           If arrData(i, 9) = "yes" Then
                  'Sammlung verschlьsselter Namen
395               If objDictName.Exists(arrData(i, 8)) = False Then objDictName.Add arrData(i, 8), arrData(i, 10)
                  'Sammlung der Subs und Functions
396               If objDictFuncAndsub.Exists(arrData(i, 6)) = False Then objDictFuncAndsub.Add arrData(i, 6), arrData(i, 5)
397           End If
398       Next i

          'Codesammlung aus Modulen
399       For Each objVBCitem In objWB.VBProject.VBComponents
400           If objDictModule.Exists(objVBCitem.Name) = False Then
401               sCode = GetCodeFromModule(objVBCitem)
                  'Beseitigung von Zeilenumbrьchen
402               sCode = VBA.Replace(sCode, " _" & vbNewLine, " XXXXX") 'changed : am 24.04 CalDymos
403               objDictModule.Add objVBCitem.Name, sCode
404               objDictModuleOld.Add objVBCitem.Name, sCode
405               sCode = vbNullString
406           End If
407       Next objVBCitem
          'Ende der Sammlung

408       Debug.Print "Data collection - completed:" & VBA.Format$(Now() - dTime, "Long Time")

          'Schleifen
409       sCode = vbNullString
410       With objDictName
411           For i = 0 To .Count - 1
412               For j = 0 To objDictModule.Count - 1
413                   sFinde = .Keys(i)
414                   sReplace = .Items(i)
415                   skey = objDictModule.Keys(j)
416                   sCode = objDictModule.Item(skey)
417                   If sCode Like "*" & sFinde & "*" And VBA.Len(sFinde) > 1 Then
                          '------------------------------------------------ changed: 31.08 CalDymos
418                       sPattern = "([\*\.\^\*\+\#\(\)\-\=\/\,\:\;\s])" & sFinde & "([\*\.\^\*\+\!\@\#\$\%\&\(\)\-\=\/\,\:\;\s]|$)"
419                       sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
420                       If InStr(sCode, "Application.OnTime") <> 0 And InStr(sCode, sFinde) <> 0 Then
421                           If objDictFuncAndsub.Exists(sFinde) Then
422                               sPattern = "([\" & VBA.Chr$(34) & "])" & sFinde & "([\" & VBA.Chr$(34) & "]|$)"
423                               sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
424                               If WorksheetExist(NAME_SH_STR, objWB) Then
                                      'Replace the string in the Excel sheet with coded
425                                   With objWB.Worksheets(NAME_SH_STR)
426                                       .Activate
427                                       For k = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
428                                           If .Cells(k, 2).Value2 = skey And .Cells(k, 5).Value = Chr$(34) & sFinde & Chr$(34) Then
429                                               .Cells(k, 5).Value = Chr$(34) & sReplace & Chr$(34)
430                                           End If
431                                       Next
432                                   End With
433                               End If
434                           End If
435                       End If
                          '------------------------------------------------
436                       If sCode <> vbNullString Then objDictModule.Item(skey) = sCode
437                   End If
                      'Regulierungsrahmen fьr Events, vor allem fьr Formulare
438                   If sCode Like "* " & Chr$(83) & "ub *" & sFinde & "_*(*)*" Then
439                       sPattern = "([\s])(Sub)([\s])" & sFinde & "(\_{1}[A-Za-zА-Яа-яЁё]{4,40}\([A-Za-zА-Яа-яЁё\s\.\,]{0,100}\))"
440                       sCode = RegExpFindReplace(sCode, sPattern, "$1$2$3" & sReplace & "$4", True, False, False)
441                       If sCode <> vbNullString Then objDictModule.Item(skey) = sCode
442                       sPattern = "([\s])" & sFinde & "(\_{1}[A-Za-zА-Яа-яЁё]{4,40}(?:\:\s|\n|\r))"
443                       sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
444                       If sCode <> vbNullString Then objDictModule.Item(skey) = sCode
445                   End If
446                   sCode = vbNullString
447               Next j
448               DoEvents
449               If .Count > 1 Then
450                   Application.StatusBar = "Data encryption - completed:" & Format(i / (.Count - 1), "Percent") & ", " & i & "from" & .Count - 1
451               Else
452                   Application.StatusBar = "Data encryption - completed:" & Format(i / .Count, "Percent") & ", " & i & "from" & .Count
453               End If
454           Next i
455       End With
456       Application.StatusBar = False
          'Ende

          'Ьbertragung
457       sCode = vbNullString

458       For j = 0 To objDictModule.Count - 1
              Dim arrNew  As Variant
              Dim arrOld  As Variant
              Dim sTemp   As String
459           arrNew = VBA.Split(objDictModule.Items(j), vbNewLine)
460           arrOld = VBA.Split(objDictModuleOld.Items(j), vbNewLine)
461           For i = LBound(arrNew) To UBound(arrNew)
462               If arrNew(i) = vbNullString Or VBA.Left$(VBA.Trim$(arrNew(i)), 1) = "'" Then
463                   sTemp = vbNullString
464               Else
465                   sTemp = "'" & arrOld(i) & vbNewLine
466               End If
467               sCode = sCode & sTemp & arrNew(i) & vbNewLine
468               sTemp = vbNullString
469           Next i
470           skey = objDictModule.Keys(j)
471           objDictModule.Item(skey) = sCode
472           sCode = vbNullString
473       Next j
474       Debug.Print "String alternation - completed:" & VBA.Format$(Now() - dTime, "Long Time")


475       Debug.Print "Data encryption - completed:" & VBA.Format$(Now() - dTime, "Long Time")
          'code laden
476       For j = 0 To objDictModule.Count - 1
477           Set objVBCitem = objWB.VBProject.VBComponents(objDictModule.Keys(j))
478           sCode = objDictModule.Items(j)
              'возврат перенос строк
479           sCode = VBA.Replace(sCode, " XXXXX", " _" & vbNewLine) 'changed : am 24.04 CalDymos
480           Call SetCodeInModule(objVBCitem, sCode)
481       Next j

482       Debug.Print "Code loading- completed:" & VBA.Format$(Now() - dTime, "Long Time")

          'Controls umbenennen
483       For i = LBound(arrData) To UBound(arrData)
484           If arrData(i, 9) = "yes" And objDictName.Exists(arrData(i, 8)) Then
485               If arrData(i, 1) = "Control" Then
486                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
487                   objVBCitem.Designer.Controls(arrData(i, 8)).Name = arrData(i, 10)
488               End If
489           End If
490       Next i

491       Debug.Print "Renaming of controls - completed:" & VBA.Format$(Now() - dTime, "Long Time")

          'Дndern von Modulen
492       For i = LBound(arrData) To UBound(arrData)
493           If arrData(i, 9) = "yes" And objDictName.Exists(arrData(i, 8)) Then
494               If arrData(i, 1) = "Module" And VBA.CByte(arrData(i, 2)) <> 100 Then
495                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
496                   objVBCitem.Name = arrData(i, 10)
497               ElseIf arrData(i, 1) = "Module" And VBA.CByte(arrData(i, 2)) = 100 Then
498                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
499                   objVBCitem.Name = arrData(i, 10)
500               End If
501           End If
502       Next i

          'шифрование строк
503       If bEncodeStr Then Call EncodedStringCode(objWB)

504       Debug.Print "Renaming modules- completed:" & VBA.Format$(Now() - dTime, "Long Time")
505       objWB.Save
          
506       Obfuscation = True
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : EncodedStringCode - шифрование строковый значений кода
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
507       With objWB.Worksheets(NAME_SH_STR)
508           .Activate
509           arrData = .Range(Cells(2, 1), Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 9)).Value2
510           strCryptFuncCipher = .Cells(2, 18).Value
511           strCryptKeyCipher = .Cells(2, 15).Value
512           strKey = .Cells(2, 13).Value
513       End With
          'Zeilenzusammenstellung
514       sCodeString = "Option Explicit" & VBA.Chr$(13)
515       For i = LBound(arrData) To UBound(arrData)
516           If arrData(i, 7) = "yes" Then
517               sCodeString = sCodeString & "Public Const " & arrData(i, 8) & " as String = " & Chr$(34) & StringCrypt(TrimA(arrData(i, 5), Chr$(34)), strKey, True) & Chr$(34) & VBA.Chr$(13)
518           End If
519       Next i
          Dim NameOldMOdule As String
520       For i = LBound(arrData) To UBound(arrData)
521           If arrData(i, 7) = "yes" Then
522               If NameOldMOdule <> arrData(i, 9) Then
523                   sCode = vbNullString
524                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 9))
525                   sCode = GetCodeFromModule(objVBCitem)
526                   If sCode <> vbNullString Then
527                       varStr = VBA.Split(sCode, vbNewLine)
528                       For k = 0 To UBound(varStr())
529                           If CStr(varStr(k)) = "" Then
                                  'Do Nothing
530                           ElseIf Left$(CStr(varStr(k)), 1) = "'" Then
                                  'Do Nothing
531                           ElseIf IsSubOrFunc(varStr(k)) Then
                                  'Do Nothing
532                           ElseIf VBA.InStr(1, CStr(varStr(k)), arrData(i, 5)) <> 0 Then
533                               varStr(k) = VBA.Trim$(VBA.Replace(CStr(varStr(k)), arrData(i, 5), strCryptFuncCipher & "(" & arrData(i, 8) & ", " & strCryptKeyCipher & ")"))
534                           End If
535                       Next
536                       sCode = Join(varStr, vbNewLine)
537                   End If
538                   NameOldMOdule = arrData(i, 9)
539               Else
540                   If sCode <> vbNullString Then
541                       varStr = VBA.Split(sCode, vbNewLine)
542                       For k = 0 To UBound(varStr)
543                           If CStr(varStr(k)) = "" Then
                                  'Do Nothing
544                           ElseIf Left$(CStr(varStr(k)), 1) = "'" Then
                                  'Do Nothing
545                           ElseIf IsSubOrFunc(varStr(k)) Then
                                  'Do Nothing
546                           ElseIf VBA.InStr(1, CStr(varStr(k)), arrData(i, 5)) <> 0 Then
547                               varStr(k) = VBA.Trim$(VBA.Replace(CStr(varStr(k)), arrData(i, 5), strCryptFuncCipher & "(" & arrData(i, 8) & ", " & strCryptKeyCipher & ")"))
548                           End If
549                       Next
550                       sCode = Join(varStr, vbNewLine)
551                   End If
552               End If
553               If i = UBound(arrData) Then
554                   Call SetCodeInModule(objVBCitem, sCode)
555                   Set objVBCitem = Nothing
556               Else
557                   If arrData(i + 1, 9) <> arrData(i, 9) Then
558                       Call SetCodeInModule(objVBCitem, sCode)
559                       Set objVBCitem = Nothing
560                   End If
561               End If

562           End If
563           DoEvents
564           If i Mod 100 = 0 Then Application.StatusBar = "String encryption - completed:" & Format(i / UBound(arrData), "Percent") & ", " & i & "from" & UBound(arrData)
565       Next i
566       Application.StatusBar = False
          Dim sName       As String
567       sName = objWB.Worksheets(NAME_SH_STR).Cells(2, 11).Value
568       If sName <> vbNullString Then
569           Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
570           objVBCitem.Name = sName
571           Call SetCodeInModule(objVBCitem, sCodeString)
572       End If
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetCodeFromModule - получить код из модуля в строковую переменную
'* Created    : 20-04-2020 18:20
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objVBComp As VBIDE.VBComponent : модуль VBA
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function GetCodeFromModule(ByRef objVBComp As VBIDE.VBComponent) As String
573       GetCodeFromModule = vbNullString
574       With objVBComp.CodeModule
575           If .CountOfLines > 0 Then
576               GetCodeFromModule = .Lines(1, .CountOfLines)
577           End If
578       End With
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SetCodeInModule загрузить код из строковой переменой в модуль
'* Created    : 20-04-2020 18:21
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef objVBComp As VBIDE.VBComponent : модуль VBA
'* ByVal SCode As String                : строковая переменная
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub SetCodeInModule(ByRef objVBComp As VBIDE.VBComponent, ByVal sCode As String)
579       With objVBComp.CodeModule
580           If .CountOfLines > 0 Then
                  'Debug.Print .CountOfLines
581               Call .DeleteLines(1, .CountOfLines)
                  'Debug.Print sCode
582               Call .InsertLines(1, VBA.Trim$(sCode))
583           End If
584       End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SortTabel - сортировка диапазона данных
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
585       With ws
586           On Error GoTo errMsg
587           .Activate
588           .Range(sRng).AutoFilter
Repeatnext:
589           .AutoFilter.Sort.SortFields.Clear
590           .AutoFilter.Sort.SortFields.Add Key:=Range(sKey1), SortOn:=xlSortOnValues, Order:=bOrder, DataOption:=xlSortNormal
591           With .AutoFilter.Sort
592               .Header = xlYes
593               .MatchCase = False
594               .Orientation = xlTopToBottom
595               .SortMethod = xlPinYin
596               .Apply
597           End With
598       End With
599       Exit Sub
errMsg:
600       If Err.Number = 91 Then
601           ws.Range(sRng).AutoFilter
602           Err.Clear
603           GoTo Repeatnext
604       Else
605           Call MsgBox(Err.Description, vbCritical, "Mistake:")
606       End If
End Sub
