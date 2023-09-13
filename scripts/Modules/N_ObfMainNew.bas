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
'* Updated    : 13-09-2023 13:29    CalDymos            Parser functions changed / added

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
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 11:30    CalDymos
'* Updated    : 12-09-2023 13:31    CalDymos
Private Sub MainObfuscation(ByRef objWB As Workbook, Optional bEncodeStr As Boolean = False, Optional bNoFinishMessage As Boolean = False)
138       On Error GoTo ErrStartParser
139       If objWB.VBProject.Protection = vbext_pp_locked Then
140           Call MsgBox("The project is protected, remove the password!", vbCritical, "Project:")
141       Else
142           If objWB.ActiveSheet.Name = NAME_SH Then
143               Application.ScreenUpdating = False
144               Application.Calculation = xlCalculationManual
145               Application.EnableEvents = False

146               If Obfuscation(objWB, bEncodeStr) Then
                      
                      
147                   With objWB.Worksheets(NAME_SH)
148                       Call SortTabel(objWB.Worksheets(NAME_SH), .Range(.Cells(1, 1), .Cells(1, 13)).Address, "M1", 1)
149                   End With

150                   Application.EnableEvents = True
151                   Application.Calculation = xlCalculationAutomatic
152                   Application.ScreenUpdating = True
153                   If Not bNoFinishMessage Then
154                       Call MsgBox("Book code [" & objWB.Name & "] encrypted!", vbInformation, "Code encryption:")
155                   End If
156               Else
157                   Application.EnableEvents = True
158                   Application.Calculation = xlCalculationAutomatic
159                   Application.ScreenUpdating = True
160               End If
161           Else
162               Call MsgBox("Create or navigate to the sheet: [" & NAME_SH & "]", vbCritical, "Activating the sheet:")
163           End If
164       End If
165       Exit Sub
ErrStartParser:
166       Application.EnableEvents = True
167       Application.Calculation = xlCalculationAutomatic
168       Application.ScreenUpdating = True
169       Call MsgBox("Error in MainObfuscation" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbCritical, "Mistake:")
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
170       With ActiveWorkbook.Worksheets(NAME_SH)
171           LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
172           If LastRow > 1 Then
173               .Range(.Cells(2, 12), .Cells(LastRow, 12)).FormulaR1C1 = "=LEN(RC[-4])"
174               .Range(.Cells(2, 13), .Cells(LastRow, 13)).FormulaR1C1 = "=R[-1]C+1"
175               .Range(.Cells(2, 13), .Cells(LastRow, 13)).Value = .Range(.Cells(2, 13), .Cells(LastRow, 13)).Value
176               Call SortTabel(ActiveWorkbook.Worksheets(NAME_SH), .Range(.Cells(1, 1), .Cells(1, 13)).Address, "L1", 2)
177           End If
178       End With
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
          
179       For i = 1 To Len(Inp)
180           Position = Position + 1
181           If Position > Len(Key) Then Position = 1
182           keyZahl = Asc(Mid(Key, Position, 1))
                  
183           If Mode Then
                  
                  'Verschlьsseln
184               orgZahl = Asc(Mid(Inp, i, 1))
185               cptZahl = orgZahl Xor keyZahl
186               cptString = Hex(cptZahl)
187               If Len(cptString) < 2 Then cptString = "0" & cptString
188               z = z & cptString
                  
189           Else
                  
                  'Entschlьsseln
190               If i > Len(Inp) \ 2 Then Exit For
191               cptZahl = CByte("&H" & Mid$(Inp, i * 2 - 1, 2))
192               orgZahl = cptZahl Xor keyZahl
193               z = z & Chr$(orgZahl)
                  
194           End If
195       Next i
           
196       StringCrypt = z
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

197       If Not IsArrayEmpty(AddProcs()) Then
198           For i = 0 To UBound(AddProcs())
199               Set objProc = AddProcs(i)
200               If Not CodeModuleExist(objWB, objProc.ModuleName) Then
201                   Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
202                   objVBCitem.Name = objProc.ModuleName
203               End If
                  
204               With objWB.VBProject.VBComponents(objProc.ModuleName).CodeModule
205                   If Not ProcInCodeModuleExist(objWB.VBProject.VBComponents(objProc.ModuleName).CodeModule, objProc.Name) Then
206                       sCodeLines = Join(objProc.CodeLines, vbCrLf)
207                       .AddFromString sCodeLines
208                   ElseIf objProc.BehavProcExists = enumBehavProcExistInsCodeAtBegin Then
209                       startLine = .ProcBodyLine(objProc.Name, vbext_pk_Proc) + 1
          
210                       For k = 1 To UBound(objProc.CodeLines) - 1
211                           .InsertLines startLine, objProc.GetCodeLine(k)
212                           startLine = startLine + 1
213                       Next
214                   ElseIf objProc.BehavProcExists = enumBehavProcExistInsCodeAtEnd Then
215                       startLine = .ProcStartLine(objProc.Name, vbext_pk_Proc) + .ProcCountLines(objProc.Name, vbext_pk_Proc) - 1
          
216                       For k = 1 To UBound(objProc.CodeLines) - 1
217                           .InsertLines startLine, objProc.GetCodeLine(k)
218                           startLine = startLine + 1
219                       Next
220                   End If
221               End With
222           Next
223       End If
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
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub AddGlobalVarsToVBProject(ByRef objWB As Workbook, AddGlobalVars() As CAddGlobalVar)

          Dim k As Long
          Dim i As Long
          Dim objVBCitem  As VBIDE.VBComponent
          Dim objVar As CAddGlobalVar
          Dim sCodeLines As String
              
224       If Not IsArrayEmpty(AddGlobalVars()) Then
225           For i = 0 To UBound(AddGlobalVars())
226               Set objVar = AddGlobalVars(i)
227               If Not CodeModuleExist(objWB, objVar.ModuleName) Then
228                   Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
229                   objVBCitem.Name = objVar.ModuleName
230               End If
                  
                  
231               With objWB.VBProject.VBComponents(objVar.ModuleName).CodeModule
232                   sCodeLines = vbNullString
233                   Select Case objVar.Visibility
                          Case enumVisibility.enumVisibilityPublic
234                           sCodeLines = sCodeLines & "Public "
235                       Case enumVisibility.enumVisibilityPrivate
236                           sCodeLines = sCodeLines & "Private "
237                   End Select
238                   If objVar.IsConstant Then sCodeLines = sCodeLines & "Const "
239                   sCodeLines = sCodeLines & objVar.Name & " As "
240                   Select Case objVar.DateType
                          Case enumDataType.enumDataTypeString
241                           sCodeLines = sCodeLines & "String "
242                       Case enumDataType.enumDataTypeInt
243                           sCodeLines = sCodeLines & "Integer "
244                       Case enumDataType.enumDataTypeLong
245                           sCodeLines = sCodeLines & "Long "
246                       Case enumDataType.enumDataTypeBool
247                           sCodeLines = sCodeLines & "Boolean "
248                       Case enumDataType.enumDataTypeByte
249                           sCodeLines = sCodeLines & "Byte "
250                   End Select
251                   If objVar.IsConstant Then
252                       sCodeLines = sCodeLines & "= "
253                       Select Case objVar.DateType
                              Case enumDataType.enumDataTypeString
254                               sCodeLines = sCodeLines & Chr$(34) & objVar.Value & Chr$(34)
255                           Case enumDataType.enumDataTypeInt
256                               sCodeLines = sCodeLines & CStr(CInt(objVar.Value))
257                           Case enumDataType.enumDataTypeLong
258                               sCodeLines = sCodeLines & CStr(CLng(objVar.Value))
259                           Case enumDataType.enumDataTypeBool
260                               sCodeLines = sCodeLines & CStr(CBool(objVar.Value))
261                           Case enumDataType.enumDataTypeByte
262                               sCodeLines = sCodeLines & CStr(CByte(objVar.Value))
263                       End Select
264                   End If
                                    
265                   .AddFromString sCodeLines
266                   sCodeLines = vbNullString
267               End With
268           Next
269       End If
          
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
270       With objWB.Worksheets(NAME_SH_CTL)
271           .Activate
272           i = .Cells(Rows.Count, 1).End(xlUp).Row
273           arrData = .Range(Cells(2, 1), Cells(i, 5)).Value2
274       End With
          
275       For i = LBound(arrData) To UBound(arrData)
          

276           If arrData(i, 1) = "UserForm" Then
277               For Each vbc In objWB.VBProject.VBComponents
278                   If Not vbc.Designer Is Nothing And vbc.Type = vbext_ct_MSForm Then
279                       If vbc.Name = arrData(i, 2) Then
280                           vbc.Properties(arrData(i, 4)) = ""
281                           Exit For
282                       End If
283                   End If
284               Next
285           Else
286               For Each vbc In objWB.VBProject.VBComponents
287                   If Not vbc.Designer Is Nothing And vbc.Type = vbext_ct_MSForm Then
288                       If vbc.Name = arrData(i, 2) Then
289                           For Each objCtl In vbc.Designer.Controls
290                               If objCtl.Name = arrData(i, 3) Then
291                                   Select Case arrData(i, 4)
                                          Case "ControlTipText"
292                                           objCtl.ControlTipText = ""
293                                       Case "Tag"
294                                           objCtl.Tag = ""
295                                       Case "Text"
296                                           Set objTxtBox = objCtl
297                                           objTxtBox.Text = ""
298                                       Case "Caption"
299                                           Select Case arrData(i, 1)
                                                  Case "CommandButton"
300                                                   Set objCmdBtn = objCtl
301                                                   objCmdBtn.Caption = ""
302                                               Case "Label"
303                                                   Set objLbl = objCtl
304                                                   objLbl.Caption = ""
305                                               Case "Frame"
306                                                   Set objFrame = objCtl
307                                                   objFrame.Caption = ""
308                                               Case "CheckBox"
309                                                   Set objChkBox = objCtl
310                                                   objChkBox.Caption = ""
311                                               Case "OptionButton"
312                                                   Set objOptBtn = objCtl
313                                                   objOptBtn.Caption = ""
314                                               Case "ToggleButton"
315                                                   Set objTglBtn = objCtl
316                                                   objTglBtn.Caption = ""
317                                           End Select
318                                       Case "Width"
319                                           objCtl.Width = 0
320                                       Case "Height"
321                                           objCtl.Height = 0
322                                       Case "Left"
323                                           objCtl.Left = -32768
324                                       Case "Top"
325                                           objCtl.top = -32768
326                                   End Select
327                                   Exit For
328                               End If
329                           Next
330                           Exit For
331                       End If
332                   End If
333               Next
334           End If
             
        
335       Next i
          
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
                    
336       dTime = Now()
337       Debug.Print "Start:" & VBA.Format$(Now() - dTime, "Long Time")

338       Set objDictName = New Scripting.Dictionary
339       Set objDictFuncAndsub = New Scripting.Dictionary
340       Set objDictModule = New Scripting.Dictionary
341       Set objDictModuleOld = New Scripting.Dictionary
          
342       If bEncodeStr Then
343           If Not IsArrayEmpty(CryptFunc()) And Not IsArrayEmpty(CryptKey()) Then
344               If UBound(CryptFunc()) = -1 Or UBound(CryptKey()) = -1 Then
345                   Call MsgBox("Please start the VBA code parser before !", vbCritical, "Run VBA code parser")
346                   Exit Function
347               End If
348           Else
349               Call MsgBox("Please start the VBA code parser before !", vbCritical, "Run VBA code parser")
350               Exit Function
351           End If
352       End If
          
353       If Not IsArrayEmpty(AddCtrlProp()) Then
354           If UBound(AddCtrlProp()) = -1 Then
355               Call MsgBox("Please start the VBA code parser before !", vbCritical, "Run VBA code parser")
356               Exit Function
357           Else
358           End If
359       Else
360           Call MsgBox("Please start the VBA code parser before !", vbCritical, "Run VBA code parser")
361           Exit Function
362       End If
          
          'save and Load
363       If Not sFolderHave(objWB.Path & Application.PathSeparator & OBF_RELEASE_PATH) Then MkDir (objWB.Path & Application.PathSeparator & OBF_RELEASE_PATH)
364       objWB.SaveAs Filename:=objWB.Path & Application.PathSeparator & OBF_RELEASE_PATH & Application.PathSeparator & C_PublicFunctions.sGetBaseName(objWB.FullName) & "_obf_" & Replace(Now(), ":", ".") & "." & C_PublicFunctions.sGetExtensionName(objWB.FullName)    ', FileFormat:=objWB.FileFormat

365       Debug.Print "File saving - completed:" & VBA.Format$(Now() - dTime, "Long Time")
              
366       If bEncodeStr Then
              ' Insert constant with the key
367           AddGlobalVarsToVBProject objWB, CryptKey()
              
              ' Insert string encryption function
368           AddProcsToVBProject objWB, CryptFunc()
369       End If
          
          'Insert Control Properties
370       AddProcsToVBProject objWB, AddCtrlProp()
          
          'Obfuscate Control Properties
371       ObfuscateCtlProperties objWB
          
          'Filtern
372       Call FelterAdd

          'Einlesen der Daten
373       With objWB.Worksheets(NAME_SH)
374           .Activate
375           i = .Cells(Rows.Count, 1).End(xlUp).Row
376           arrData = .Range(Cells(2, 1), Cells(i, 10)).Value2
377       End With

          'Sammlung verschlьsselter Namen und Subs / Functions
378       For i = LBound(arrData) To UBound(arrData)
379           If arrData(i, 9) = "yes" Then
                  'Sammlung verschlьsselter Namen
380               If objDictName.Exists(arrData(i, 8)) = False Then objDictName.Add arrData(i, 8), arrData(i, 10)
                  'Sammlung der Subs und Functions
381               If objDictFuncAndsub.Exists(arrData(i, 6)) = False Then objDictFuncAndsub.Add arrData(i, 6), arrData(i, 5)
382           End If
383       Next i

          'Codesammlung aus Modulen
384       For Each objVBCitem In objWB.VBProject.VBComponents
385           If objDictModule.Exists(objVBCitem.Name) = False Then
386               sCode = GetCodeFromModule(objVBCitem)
                  'Beseitigung von Zeilenumbrьchen
387               sCode = VBA.Replace(sCode, " _" & vbNewLine, " XXXXX") 'changed : am 24.04 CalDymos
388               objDictModule.Add objVBCitem.Name, sCode
389               objDictModuleOld.Add objVBCitem.Name, sCode
390               sCode = vbNullString
391           End If
392       Next objVBCitem
          'Ende der Sammlung

393       Debug.Print "Data collection - completed:" & VBA.Format$(Now() - dTime, "Long Time")

          'Schleifen
394       sCode = vbNullString
395       With objDictName
396           For i = 0 To .Count - 1
397               For j = 0 To objDictModule.Count - 1
398                   sFinde = .Keys(i)
399                   sReplace = .Items(i)
400                   skey = objDictModule.Keys(j)
401                   sCode = objDictModule.Item(skey)
402                   If sCode Like "*" & sFinde & "*" And VBA.Len(sFinde) > 1 Then
                          '------------------------------------------------ changed: 31.08 CalDymos
403                       sPattern = "([\*\.\^\*\+\#\(\)\-\=\/\,\:\;\s])" & sFinde & "([\*\.\^\*\+\!\@\#\$\%\&\(\)\-\=\/\,\:\;\s]|$)"
404                       sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
405                       If InStr(sCode, "Application.OnTime") <> 0 And InStr(sCode, sFinde) <> 0 Then
406                           If objDictFuncAndsub.Exists(sFinde) Then
407                               sPattern = "([\" & VBA.Chr$(34) & "])" & sFinde & "([\" & VBA.Chr$(34) & "]|$)"
408                               sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
409                               If WorksheetExist(NAME_SH_STR, objWB) Then
                                      'Replace the string in the Excel sheet with coded
410                                   With objWB.Worksheets(NAME_SH_STR)
411                                       .Activate
412                                       For k = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
413                                           If .Cells(k, 2).Value2 = skey And .Cells(k, 5).Value = Chr$(34) & sFinde & Chr$(34) Then
414                                               .Cells(k, 5).Value = Chr$(34) & sReplace & Chr$(34)
415                                           End If
416                                       Next
417                                   End With
418                               End If
419                           End If
420                       End If
                          '------------------------------------------------
421                       If sCode <> vbNullString Then objDictModule.Item(skey) = sCode
422                   End If
                      'Regulierungsrahmen fьr Events, vor allem fьr Formulare
423                   If sCode Like "* " & Chr$(83) & "ub *" & sFinde & "_*(*)*" Then
424                       sPattern = "([\s])(Sub)([\s])" & sFinde & "(\_{1}[A-Za-zА-Яа-яЁё]{4,40}\([A-Za-zА-Яа-яЁё\s\.\,]{0,100}\))"
425                       sCode = RegExpFindReplace(sCode, sPattern, "$1$2$3" & sReplace & "$4", True, False, False)
426                       If sCode <> vbNullString Then objDictModule.Item(skey) = sCode
427                       sPattern = "([\s])" & sFinde & "(\_{1}[A-Za-zА-Яа-яЁё]{4,40}(?:\:\s|\n|\r))"
428                       sCode = RegExpFindReplace(sCode, sPattern, "$1" & sReplace & "$2", True, False, False)
429                       If sCode <> vbNullString Then objDictModule.Item(skey) = sCode
430                   End If
431                   sCode = vbNullString
432               Next j
433               DoEvents
434               If .Count > 1 Then
435                   Application.StatusBar = "Data encryption - completed:" & Format(i / (.Count - 1), "Percent") & ", " & i & "from" & .Count - 1
436               Else
437                   Application.StatusBar = "Data encryption - completed:" & Format(i / .Count, "Percent") & ", " & i & "from" & .Count
438               End If
439           Next i
440       End With
441       Application.StatusBar = False
          'Ende

          'Ьbertragung
442       sCode = vbNullString

443       For j = 0 To objDictModule.Count - 1
              Dim arrNew  As Variant
              Dim arrOld  As Variant
              Dim sTemp   As String
444           arrNew = VBA.Split(objDictModule.Items(j), vbNewLine)
445           arrOld = VBA.Split(objDictModuleOld.Items(j), vbNewLine)
446           For i = LBound(arrNew) To UBound(arrNew)
447               If arrNew(i) = vbNullString Or VBA.Left$(VBA.Trim$(arrNew(i)), 1) = "'" Then
448                   sTemp = vbNullString
449               Else
450                   sTemp = "'" & arrOld(i) & vbNewLine
451               End If
452               sCode = sCode & sTemp & arrNew(i) & vbNewLine
453               sTemp = vbNullString
454           Next i
455           skey = objDictModule.Keys(j)
456           objDictModule.Item(skey) = sCode
457           sCode = vbNullString
458       Next j
459       Debug.Print "String alternation - completed:" & VBA.Format$(Now() - dTime, "Long Time")


460       Debug.Print "Data encryption - completed:" & VBA.Format$(Now() - dTime, "Long Time")
          'code laden
461       For j = 0 To objDictModule.Count - 1
462           Set objVBCitem = objWB.VBProject.VBComponents(objDictModule.Keys(j))
463           sCode = objDictModule.Items(j)
              'возврат перенос строк
464           sCode = VBA.Replace(sCode, " XXXXX", " _" & vbNewLine) 'changed : am 24.04 CalDymos
465           Call SetCodeInModule(objVBCitem, sCode)
466       Next j

467       Debug.Print "Code loading- completed:" & VBA.Format$(Now() - dTime, "Long Time")

          'Controls umbenennen
468       For i = LBound(arrData) To UBound(arrData)
469           If arrData(i, 9) = "yes" And objDictName.Exists(arrData(i, 8)) Then
470               If arrData(i, 1) = "Control" Then
471                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
472                   objVBCitem.Designer.Controls(arrData(i, 8)).Name = arrData(i, 10)
473               End If
474           End If
475       Next i

476       Debug.Print "Renaming of controls - completed:" & VBA.Format$(Now() - dTime, "Long Time")

          'Дndern von Modulen
477       For i = LBound(arrData) To UBound(arrData)
478           If arrData(i, 9) = "yes" And objDictName.Exists(arrData(i, 8)) Then
479               If arrData(i, 1) = "Module" And VBA.CByte(arrData(i, 2)) <> 100 Then
480                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
481                   objVBCitem.Name = arrData(i, 10)
482               ElseIf arrData(i, 1) = "Module" And VBA.CByte(arrData(i, 2)) = 100 Then
483                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 3))
484                   objVBCitem.Name = arrData(i, 10)
485               End If
486           End If
487       Next i

          'шифрование строк
488       If bEncodeStr Then Call EncodedStringCode(objWB)

489       Debug.Print "Renaming modules- completed:" & VBA.Format$(Now() - dTime, "Long Time")
490       objWB.Save
          
491       Obfuscation = True
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
492       With objWB.Worksheets(NAME_SH_STR)
493           .Activate
494           arrData = .Range(Cells(2, 1), Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 9)).Value2
495           strCryptFuncCipher = .Cells(2, 18).Value
496           strCryptKeyCipher = .Cells(2, 15).Value
497           strKey = .Cells(2, 13).Value
498       End With
          'Zeilenzusammenstellung
499       sCodeString = "Option Explicit" & VBA.Chr$(13)
500       For i = LBound(arrData) To UBound(arrData)
501           If arrData(i, 7) = "yes" Then
502               sCodeString = sCodeString & "Public Const " & arrData(i, 8) & " as String = " & Chr$(34) & StringCrypt(TrimA(arrData(i, 5), Chr$(34)), strKey, True) & Chr$(34) & VBA.Chr$(13)
503           End If
504       Next i
          Dim NameOldMOdule As String
505       For i = LBound(arrData) To UBound(arrData)
506           If arrData(i, 7) = "yes" Then
507               If NameOldMOdule <> arrData(i, 9) Then
508                   sCode = vbNullString
509                   Set objVBCitem = objWB.VBProject.VBComponents(arrData(i, 9))
510                   sCode = GetCodeFromModule(objVBCitem)
511                   If sCode <> vbNullString Then
512                       varStr = VBA.Split(sCode, vbNewLine)
513                       For k = 0 To UBound(varStr())
514                           If CStr(varStr(k)) = "" Then
                                  'Do Nothing
515                           ElseIf Left$(CStr(varStr(k)), 1) = "'" Then
                                  'Do Nothing
516                           ElseIf IsSubOrFunc(varStr(k)) Then
                                  'Do Nothing
517                           ElseIf VBA.InStr(1, CStr(varStr(k)), arrData(i, 5)) <> 0 Then
518                               varStr(k) = VBA.Trim$(VBA.Replace(CStr(varStr(k)), arrData(i, 5), strCryptFuncCipher & "(" & arrData(i, 8) & ", " & strCryptKeyCipher & ")"))
519                           End If
520                       Next
521                       sCode = Join(varStr, vbNewLine)
522                   End If
523                   NameOldMOdule = arrData(i, 9)
524               Else
525                   If sCode <> vbNullString Then
526                       varStr = VBA.Split(sCode, vbNewLine)
527                       For k = 0 To UBound(varStr)
528                           If CStr(varStr(k)) = "" Then
                                  'Do Nothing
529                           ElseIf Left$(CStr(varStr(k)), 1) = "'" Then
                                  'Do Nothing
530                           ElseIf IsSubOrFunc(varStr(k)) Then
                                  'Do Nothing
531                           ElseIf VBA.InStr(1, CStr(varStr(k)), arrData(i, 5)) <> 0 Then
532                               varStr(k) = VBA.Trim$(VBA.Replace(CStr(varStr(k)), arrData(i, 5), strCryptFuncCipher & "(" & arrData(i, 8) & ", " & strCryptKeyCipher & ")"))
533                           End If
534                       Next
535                       sCode = Join(varStr, vbNewLine)
536                   End If
537               End If
538               If i = UBound(arrData) Then
539                   Call SetCodeInModule(objVBCitem, sCode)
540                   Set objVBCitem = Nothing
541               Else
542                   If arrData(i + 1, 9) <> arrData(i, 9) Then
543                       Call SetCodeInModule(objVBCitem, sCode)
544                       Set objVBCitem = Nothing
545                   End If
546               End If

547           End If
548           DoEvents
549           If i Mod 100 = 0 Then Application.StatusBar = "String encryption - completed:" & Format(i / UBound(arrData), "Percent") & ", " & i & "from" & UBound(arrData)
550       Next i
551       Application.StatusBar = False
          Dim sName       As String
552       sName = objWB.Worksheets(NAME_SH_STR).Cells(2, 11).Value
553       If sName <> vbNullString Then
554           Set objVBCitem = objWB.VBProject.VBComponents.Add(vbext_ct_StdModule)
555           objVBCitem.Name = sName
556           Call SetCodeInModule(objVBCitem, sCodeString)
557       End If
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
558       GetCodeFromModule = vbNullString
559       With objVBComp.CodeModule
560           If .CountOfLines > 0 Then
561               GetCodeFromModule = .Lines(1, .CountOfLines)
562           End If
563       End With
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
564       With objVBComp.CodeModule
565           If .CountOfLines > 0 Then
                  'Debug.Print .CountOfLines
566               Call .DeleteLines(1, .CountOfLines)
                  'Debug.Print sCode
567               Call .InsertLines(1, VBA.Trim$(sCode))
568           End If
569       End With
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
570       With ws
571           On Error GoTo errMsg
572           .Activate
573           .Range(sRng).AutoFilter
Repeatnext:
574           .AutoFilter.Sort.SortFields.Clear
575           .AutoFilter.Sort.SortFields.Add Key:=Range(sKey1), SortOn:=xlSortOnValues, Order:=bOrder, DataOption:=xlSortNormal
576           With .AutoFilter.Sort
577               .Header = xlYes
578               .MatchCase = False
579               .Orientation = xlTopToBottom
580               .SortMethod = xlPinYin
581               .Apply
582           End With
583       End With
584       Exit Sub
errMsg:
585       If Err.Number = 91 Then
586           ws.Range(sRng).AutoFilter
587           Err.Clear
588           GoTo Repeatnext
589       Else
590           Call MsgBox(Err.Description, vbCritical, "Mistake:")
591       End If
End Sub
