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
'* Updated    : 14-09-2023 06:45    CalDymos

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
'* Modified   : Date and Time       Author              Description
'* Updated    : 14-09-2023 06:45    CalDymos
Private Sub AddProcsToVBProject(ByRef objWB As Workbook, AddProcs() As CAddProc)
          Dim k As Long
          Dim i As Long
          Dim j As Long
          Dim objVBCitem  As VBIDE.VBComponent
          Dim objProc As CAddProc
          Dim sCodeLines As String
          Dim startLine As Long
          Dim CompIfLine As Long

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
209                       CompIfLine = -1
210                       For j = 1 To .CountOfDeclarationLines + 1
211                           If InStr(.Lines(j, 1), "#If") <> 0 Then
                                  'Debug.Print j
212                               CompIfLine = j
213                               Exit For
214                           End If
215                       Next j

216                       If CompIfLine <> -1 Then
217                           .InsertLines .CountOfLines, sCodeLines
218                       Else
219                           .AddFromString sCodeLines
220                       End If
221                   ElseIf objProc.BehavProcExists = enumBehavProcExistInsCodeAtBegin Then
222                       startLine = .ProcBodyLine(objProc.Name, vbext_pk_Proc) + 1
          
223                       For k = 1 To UBound(objProc.CodeLines) - 1
224                           .InsertLines startLine, objProc.GetCodeLine(k)
225                           startLine = startLine + 1
226                       Next
227                   ElseIf objProc.BehavProcExists = enumBehavProcExistInsCodeAtEnd Then
228                       startLine = .ProcStartLine(objProc.Name, vbext_pk_Proc) + .ProcCountLines(objProc.Name, vbext_pk_Proc) - 1
          
229                       For k = 1 To UBound(objProc.CodeLines) - 1
230                           .InsertLines startLine, objProc.GetCodeLine(k)
231                           startLine = startLine + 1
232                       Next
233                   End If
234               End With
235           Next
236       End If
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
              
237       If Not IsArrayEmpty(AddGlobalVars()) Then
238           For i = 0 To UBound(AddGlobalVars())
239               Set objVar = AddGlobalVars(i)
240               If Not CodeModuleExist(objWB, objVar.ModuleName) Then
241                   Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
242                   objVBCitem.Name = objVar.ModuleName
243               End If
                  
                  
244               With objWB.VBProject.VBComponents(objVar.ModuleName).CodeModule
245                   sCodeLines = vbNullString
246                   Select Case objVar.Visibility
                          Case enumVisibility.enumVisibilityPublic
247                           sCodeLines = sCodeLines & "Public "
248                       Case enumVisibility.enumVisibilityPrivate
249                           sCodeLines = sCodeLines & "Private "
250                   End Select
251                   If objVar.IsConstant Then sCodeLines = sCodeLines & "Const "
252                   sCodeLines = sCodeLines & objVar.Name & " As "
253                   Select Case objVar.DateType
                          Case enumDataType.enumDataTypeString
254                           sCodeLines = sCodeLines & "String "
255                       Case enumDataType.enumDataTypeInt
256                           sCodeLines = sCodeLines & "Integer "
257                       Case enumDataType.enumDataTypeLong
258                           sCodeLines = sCodeLines & "Long "
259                       Case enumDataType.enumDataTypeBool
260                           sCodeLines = sCodeLines & "Boolean "
261                       Case enumDataType.enumDataTypeByte
262                           sCodeLines = sCodeLines & "Byte "
263                   End Select
264                   If objVar.IsConstant Then
265                       sCodeLines = sCodeLines & "= "
266                       Select Case objVar.DateType
                              Case enumDataType.enumDataTypeString
267                               sCodeLines = sCodeLines & Chr$(34) & objVar.Value & Chr$(34)
268                           Case enumDataType.enumDataTypeInt
269                               sCodeLines = sCodeLines & CStr(CInt(objVar.Value))
270                           Case enumDataType.enumDataTypeLong
271                               sCodeLines = sCodeLines & CStr(CLng(objVar.Value))
272                           Case enumDataType.enumDataTypeBool
273                               sCodeLines = sCodeLines & CStr(CBool(objVar.Value))
274                           Case enumDataType.enumDataTypeByte
275                               sCodeLines = sCodeLines & CStr(CByte(objVar.Value))
276                       End Select
277                   End If
                                    
278                   If .Find("#IF", 1, 1, .CountOfLines - 1, 500) Then 'Check for compiler #If, because with compiler #If,  .AddFromString does not work correctly
279                       lSize = .CountOfLines
280                       For k = 1 To lSize
281                           If Trim$(.Lines(k, 1)) = "Option Explicit" Then
282                           ElseIf Trim$(.Lines(k, 1)) = "Option Base 1" Then
283                           ElseIf Trim$(.Lines(k, 1)) = "Option Private Module" Then
284                           ElseIf Trim$(.Lines(k, 1)) = "Option Compare Text" Then
285                           Else
286                               .InsertLines k, sCodeLines
287                               Exit For
288                           End If
289                       Next
290                   Else
291                       .AddFromString sCodeLines
292                   End If
293                   sCodeLines = vbNullString
294               End With
295           Next
296       End If
          
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
297       With objWB.Worksheets(NAME_SH_CTL)
298           .Activate
299           i = .Cells(Rows.Count, 1).End(xlUp).Row
300           arrData = .Range(Cells(2, 1), Cells(i, 5)).Value2
301       End With
          
302       For i = LBound(arrData) To UBound(arrData)
          

303           If arrData(i, 1) = "UserForm" Then
304               For Each vbc In objWB.VBProject.VBComponents
305                   If Not vbc.Designer Is Nothing And vbc.Type = vbext_ct_MSForm Then
306                       If vbc.Name = arrData(i, 2) Then
307                           vbc.Properties(arrData(i, 4)) = ""
308                           Exit For
309                       End If
310                   End If
311               Next
312           Else
313               For Each vbc In objWB.VBProject.VBComponents
314                   If Not vbc.Designer Is Nothing And vbc.Type = vbext_ct_MSForm Then
315                       If vbc.Name = arrData(i, 2) Then
316                           For Each objCtl In vbc.Designer.Controls
317                               If objCtl.Name = arrData(i, 3) Then
318                                   Select Case arrData(i, 4)
                                          Case "ControlTipText"
319                                           objCtl.ControlTipText = ""
320                                       Case "Tag"
321                                           objCtl.Tag = ""
322                                       Case "Text"
323                                           Set objTxtBox = objCtl
324                                           objTxtBox.Text = ""
325                                       Case "Caption"
326                                           Select Case arrData(i, 1)
                                                  Case "CommandButton"
327                                                   Set objCmdBtn = objCtl
328                                                   objCmdBtn.Caption = ""
329                                               Case "Label"
330                                                   Set objLbl = objCtl
331                                                   objLbl.Caption = ""
332                                               Case "Frame"
333                                                   Set objFrame = objCtl
334                                                   objFrame.Caption = ""
335                                               Case "CheckBox"
336                                                   Set objChkBox = objCtl
337                                                   objChkBox.Caption = ""
338                                               Case "OptionButton"
339                                                   Set objOptBtn = objCtl
340                                                   objOptBtn.Caption = ""
341                                               Case "ToggleButton"
342                                                   Set objTglBtn = objCtl
343                                                   objTglBtn.Caption = ""
344                                           End Select
345                                       Case "Width"
346                                           objCtl.Width = 0
347                                       Case "Height"
348                                           objCtl.Height = 0
349                                       Case "Left"
350                                           objCtl.Left = -32768
351                                       Case "Top"
352                                           objCtl.top = -32768
353                                   End Select
354                                   Exit For
355                               End If
356                           Next
357                           Exit For
358                       End If
359                   End If
360               Next
361           End If
             
        
362       Next i
          
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
                    
363       dTime = Now()
364       Debug.Print "Start:" & VBA.Format$(Now() - dTime, "Long Time")

365       Set objDictName = New Scripting.Dictionary
366       Set objDictFuncAndsub = New Scripting.Dictionary
367       Set objDictModule = New Scripting.Dictionary
368       Set objDictModuleOld = New Scripting.Dictionary
          
369       If bEncodeStr Then
370           If Not IsArrayEmpty(CryptFunc()) And Not IsArrayEmpty(CryptKey()) Then
371               If UBound(CryptFunc()) = -1 Or UBound(CryptKey()) = -1 Then
372                   Call MsgBox("Please start Step 1 : VBA code parser before !", vbCritical, "Run VBA code parser")
373                   Exit Function
374               End If
375           Else
376               Call MsgBox("Please start Step 1 : VBA code parser before !", vbCritical, "Run VBA code parser")
377               Exit Function
378           End If
379       End If
          
380       If Not IsArrayEmpty(AddCtrlProp()) Then
381           If UBound(AddCtrlProp()) = -1 Then
382               Call MsgBox("Please start Step 1 : VBA code parser before !", vbCritical, "Run VBA code parser")
383               Exit Function
384           Else
385           End If
386       Else
387           Call MsgBox("Please start Step 1 : VBA code parser before !", vbCritical, "Run VBA code parser")
388           Exit Function
389       End If
          
          'save and Load
390       If Not sFolderHave(objWB.Path & Application.PathSeparator & OBF_RELEASE_PATH) Then MkDir (objWB.Path & Application.PathSeparator & OBF_RELEASE_PATH)
391       objWB.SaveAs Filename:=objWB.Path & Application.PathSeparator & OBF_RELEASE_PATH & Application.PathSeparator & C_PublicFunctions.sGetBaseName(objWB.FullName) & "_obf_" & Replace(Now(), ":", ".") & "." & C_PublicFunctions.sGetExtensionName(objWB.FullName)    ', FileFormat:=objWB.FileFormat

392       Debug.Print "File saving - completed:" & VBA.Format$(Now() - dTime, "Long Time")
              
393       If bEncodeStr Then
              ' Insert constant with the key
394           AddGlobalVarsToVBProject objWB, CryptKey()
              
              ' Insert string encryption function
395           AddProcsToVBProject objWB, CryptFunc()
396       End If
          
          'Insert Control Properties
397       AddProcsToVBProject objWB, AddCtrlProp()
          
          'Obfuscate Control Properties
398       ObfuscateCtlProperties objWB
          
          'Filtern
399       Call FelterAdd

          'Einlesen der Daten
400       With objWB.Worksheets(NAME_SH)
401           .Activate
402           i = .Cells(Rows.Count, 1).End(xlUp).Row
403           arrData = .Range(Cells(2, 1), Cells(i, 10)).Value2
404       End With

          'Sammlung verschlьsselter Namen und Subs / Functions
405       For i = LBound(arrData) To UBound(arrData)
406           If arrData(i, 9) = "yes" Then
                  'Sammlung verschlьsselter Namen
407               If objDictName.Exists(arrData(i, 8)) = False Then objDictName.Add arrData(i, 8), arrData(i, 10)
                  'Sammlung der Subs und Functions
408               If objDictFuncAndsub.Exists(arrData(i, 6)) = False Then objDictFuncAndsub.Add arrData(i, 6), arrData(i, 5)
409           End If
410       Next i

          'Codesammlung aus Modulen
411       For Each objVBCitem In objWB.VBProject.VBComponents
412           If objDictModule.Exists(objVBCitem.Name) = False Then
413               sCode = GetCodeFromModule(objVBCitem)
                  'Beseitigung von Zeilenumbrьchen
414               sCode = VBA.Replace(sCode, " _" & vbNewLine, " XXXXX") 'changed : am 24.04 CalDymos
415               objDictModule.Add objVBCitem.Name, sCode
416               objDictModuleOld.Add objVBCitem.Name, sCode
417               sCode = vbNullString
418           End If
419       Next objVBCitem
          'Ende der Sammlung

420       Debug.Print "Data collection - completed:" & VBA.Format$(Now() - dTime, "Long Time")

          'Schleifen
421       sCode = vbNullString
422       With objDictName
423           For i = 0 To .Count - 1
424               For j = 0 To objDictModule.Count - 1
425                   sFinde = .Keys(i)
426                   sReplace = .Items(i)
427                   skey = objDictModule.Keys(j)
428                   sCode = objDictModule.Item(skey)
429                   If sCode Like "*" & sFinde & "*" And VBA.Len(sFinde) > 1 Then
                          '------------------------------------------------ changed: 31.08 CalDymos
430                       sPattern = "([\*\.\^\*\+\#\(\)\-\=\/\,\:\;\s])" & sFinde & "([\*\.\^\*\+\!\@\#\$\%\&\(\)\-\=\/\,\:\;\s]|$)"
431                       sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
432                       If InStr(sCode, "Application.OnTime") <> 0 And InStr(sCode, sFinde) <> 0 Then
433                           If objDictFuncAndsub.Exists(sFinde) Then
434                               sPattern = "([\" & VBA.Chr$(34) & "])" & sFinde & "([\" & VBA.Chr$(34) & "]|$)"
435                               sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
436                               If WorksheetExist(NAME_SH_STR, objWB) Then
                                      'Replace the string in the Excel sheet with coded
437                                   With objWB.Worksheets(NAME_SH_STR)
438                                       .Activate
439                                       For k = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
440                                           If .Cells(k, 2).Value2 = skey And .Cells(k, 5).Value = Chr$(34) & sFinde & Chr$(34) Then
441                                               .Cells(k, 5).Value = Chr$(34) & sReplace & Chr$(34)
442                                           End If
443                                       Next
444                                   End With
445                               End If
446                           End If
447                       End If
                          '------------------------------------------------
448                       If sCode <> vbNullString Then objDictModule.Item(skey) = sCode
449                   End If
                      'Regulierungsrahmen fьr Events, vor allem fьr Formulare
450                   If sCode Like "* " & Chr$(83) & "ub *" & sFinde & "_*(*)*" Then
451                       sPattern = "([\s])(Sub)([\s])" & sFinde & "(\_{1}[A-Za-zА-Яа-яЁё]{4,40}\([A-Za-zА-Яа-яЁё\s\.\,]{0,100}\))"
452                       sCode = RegExpFindReplace(sCode, sPattern, "$1$2$3" & sReplace & "$4", True, False, False)
453                       If sCode <> vbNullString Then objDictModule.Item(skey) = sCode
454                       sPattern = "([\s])" & sFinde & "(\_{1}[A-Za-zА-Яа-яЁё]{4,40}(?:\:\s|\n|\r))"
455                       sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
456                       If sCode <> vbNullString Then objDictModule.Item(skey) = sCode
457                   End If
458                   sCode = vbNullString
459               Next j
460               DoEvents
461               If .Count > 1 Then
462                   Application.StatusBar = "Data encryption - completed:" & Format(i / (.Count - 1), "Percent") & ", " & i & "from" & .Count - 1
463               Else
464                   Application.StatusBar = "Data encryption - completed:" & Format(i / .Count, "Percent") & ", " & i & "from" & .Count
465               End If
466           Next i
467       End With
468       Application.StatusBar = False
          'Ende

          'Ьbertragung
469       sCode = vbNullString

470       For j = 0 To objDictModule.Count - 1
              Dim arrNew  As Variant
              Dim arrOld  As Variant
              Dim sTemp   As String
471           arrNew = VBA.Split(objDictModule.Items(j), vbNewLine)
472           arrOld = VBA.Split(objDictModuleOld.Items(j), vbNewLine)
473           For i = LBound(arrNew) To UBound(arrNew)
474               If arrNew(i) = vbNullString Or VBA.Left$(VBA.Trim$(arrNew(i)), 1) = "'" Then
475                   sTemp = vbNullString
476               Else
477                   sTemp = "'" & arrOld(i) & vbNewLine
478               End If
479               sCode = sCode & sTemp & arrNew(i) & vbNewLine
480               sTemp = vbNullString
481           Next i
482           skey = objDictModule.Keys(j)
483           objDictModule.Item(skey) = sCode
484           sCode = vbNullString
485       Next j
486       Debug.Print "String alternation - completed:" & VBA.Format$(Now() - dTime, "Long Time")


487       Debug.Print "Data encryption - completed:" & VBA.Format$(Now() - dTime, "Long Time")
          'code laden
488       For j = 0 To objDictModule.Count - 1
489           Set objVBCitem = objWB.VBProject.VBComponents(objDictModule.Keys(j))
490           sCode = objDictModule.Items(j)
              'возврат перенос строк
491           sCode = VBA.Replace(sCode, " XXXXX", " _" & vbNewLine) 'changed : am 24.04 CalDymos
492           Call SetCodeInModule(objVBCitem, sCode)
493       Next j

494       Debug.Print "Code loading- completed:" & VBA.Format$(Now() - dTime, "Long Time")

          'Controls umbenennen
495       For i = LBound(arrData) To UBound(arrData)
496           If arrData(i, 9) = "yes" And objDictName.Exists(arrData(i, 8)) Then
497               If arrData(i, 1) = "Control" Then
498                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
499                   objVBCitem.Designer.Controls(arrData(i, 8)).Name = arrData(i, 10)
500               End If
501           End If
502       Next i

503       Debug.Print "Renaming of controls - completed:" & VBA.Format$(Now() - dTime, "Long Time")

          'Дndern von Modulen
504       For i = LBound(arrData) To UBound(arrData)
505           If arrData(i, 9) = "yes" And objDictName.Exists(arrData(i, 8)) Then
506               If arrData(i, 1) = "Module" And VBA.CByte(arrData(i, 2)) <> 100 Then
507                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
508                   objVBCitem.Name = arrData(i, 10)
509               ElseIf arrData(i, 1) = "Module" And VBA.CByte(arrData(i, 2)) = 100 Then
510                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
511                   objVBCitem.Name = arrData(i, 10)
512               End If
513           End If
514       Next i

          'шифрование строк
515       If bEncodeStr Then Call EncodedStringCode(objWB)

516       Debug.Print "Renaming modules- completed:" & VBA.Format$(Now() - dTime, "Long Time")
517       objWB.Save
          
518       Obfuscation = True
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
519       With objWB.Worksheets(NAME_SH_STR)
520           .Activate
521           arrData = .Range(Cells(2, 1), Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 9)).Value2
522           strCryptFuncCipher = .Cells(2, 18).Value
523           strCryptKeyCipher = .Cells(2, 15).Value
524           strKey = .Cells(2, 13).Value
525       End With
          'Zeilenzusammenstellung
526       sCodeString = "Option Explicit" & VBA.Chr$(13)
527       For i = LBound(arrData) To UBound(arrData)
528           If arrData(i, 7) = "yes" Then
529               sCodeString = sCodeString & "Public Const " & arrData(i, 8) & " as String = " & Chr$(34) & StringCrypt(TrimA(arrData(i, 5), Chr$(34)), strKey, True) & Chr$(34) & VBA.Chr$(13)
530           End If
531       Next i
          Dim NameOldMOdule As String
532       For i = LBound(arrData) To UBound(arrData)
533           If arrData(i, 7) = "yes" Then
534               If NameOldMOdule <> arrData(i, 9) Then
535                   sCode = vbNullString
536                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 9))
537                   sCode = GetCodeFromModule(objVBCitem)
538                   If sCode <> vbNullString Then
539                       varStr = VBA.Split(sCode, vbNewLine)
540                       For k = 0 To UBound(varStr())
541                           If CStr(varStr(k)) = "" Then
                                  'Do Nothing
542                           ElseIf Left$(CStr(varStr(k)), 1) = "'" Then
                                  'Do Nothing
543                           ElseIf IsSubOrFunc(varStr(k)) Then
                                  'Do Nothing
544                           ElseIf VBA.InStr(1, CStr(varStr(k)), arrData(i, 5)) <> 0 Then
545                               varStr(k) = VBA.Trim$(VBA.Replace(CStr(varStr(k)), arrData(i, 5), strCryptFuncCipher & "(" & arrData(i, 8) & ", " & strCryptKeyCipher & ")"))
546                           End If
547                       Next
548                       sCode = Join(varStr, vbNewLine)
549                   End If
550                   NameOldMOdule = arrData(i, 9)
551               Else
552                   If sCode <> vbNullString Then
553                       varStr = VBA.Split(sCode, vbNewLine)
554                       For k = 0 To UBound(varStr)
555                           If CStr(varStr(k)) = "" Then
                                  'Do Nothing
556                           ElseIf Left$(CStr(varStr(k)), 1) = "'" Then
                                  'Do Nothing
557                           ElseIf IsSubOrFunc(varStr(k)) Then
                                  'Do Nothing
558                           ElseIf VBA.InStr(1, CStr(varStr(k)), arrData(i, 5)) <> 0 Then
559                               varStr(k) = VBA.Trim$(VBA.Replace(CStr(varStr(k)), arrData(i, 5), strCryptFuncCipher & "(" & arrData(i, 8) & ", " & strCryptKeyCipher & ")"))
560                           End If
561                       Next
562                       sCode = Join(varStr, vbNewLine)
563                   End If
564               End If
565               If i = UBound(arrData) Then
566                   Call SetCodeInModule(objVBCitem, sCode)
567                   Set objVBCitem = Nothing
568               Else
569                   If arrData(i + 1, 9) <> arrData(i, 9) Then
570                       Call SetCodeInModule(objVBCitem, sCode)
571                       Set objVBCitem = Nothing
572                   End If
573               End If

574           End If
575           DoEvents
576           If i Mod 100 = 0 Then Application.StatusBar = "String encryption - completed:" & Format(i / UBound(arrData), "Percent") & ", " & i & "from" & UBound(arrData)
577       Next i
578       Application.StatusBar = False
          Dim sName       As String
579       sName = objWB.Worksheets(NAME_SH_STR).Cells(2, 11).Value
580       If sName <> vbNullString Then
581           Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
582           objVBCitem.Name = sName
583           Call SetCodeInModule(objVBCitem, sCodeString)
584       End If
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
585       GetCodeFromModule = vbNullString
586       With objVBComp.CodeModule
587           If .CountOfLines > 0 Then
588               GetCodeFromModule = .Lines(1, .CountOfLines)
589           End If
590       End With
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
591       With objVBComp.CodeModule
592           If .CountOfLines > 0 Then
                  'Debug.Print .CountOfLines
593               Call .DeleteLines(1, .CountOfLines)
                  'Debug.Print sCode
594               Call .InsertLines(1, VBA.Trim$(sCode))
595           End If
596       End With
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
597       With ws
598           On Error GoTo errMsg
599           .Activate
600           .Range(sRng).AutoFilter
Repeatnext:
601           .AutoFilter.Sort.SortFields.Clear
602           .AutoFilter.Sort.SortFields.Add Key:=Range(sKey1), SortOn:=xlSortOnValues, Order:=bOrder, DataOption:=xlSortNormal
603           With .AutoFilter.Sort
604               .Header = xlYes
605               .MatchCase = False
606               .Orientation = xlTopToBottom
607               .SortMethod = xlPinYin
608               .Apply
609           End With
610       End With
611       Exit Sub
errMsg:
612       If Err.Number = 91 Then
613           ws.Range(sRng).AutoFilter
614           Err.Clear
615           GoTo Repeatnext
616       Else
617           Call MsgBox(Err.Description, vbCritical, "Mistake:")
618       End If
End Sub
