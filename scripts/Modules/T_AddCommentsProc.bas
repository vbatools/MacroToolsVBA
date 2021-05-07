Attribute VB_Name = "T_AddCommentsProc"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : T_AddCommentsProc - ћодуль авто документировани€ кода проекта VBA
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

'табы сокращенно
Private Const vbTab2 = vbTab & vbTab
Private Const vbTab4 = vbTab2 & vbTab2
'формат даты
Private Const ctFormat = "dd-mm-yyyy hh:nn"

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : sysAddHeaderTop - создание основного коментари€
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub sysAddHeaderTop()
25:    Call sysAddHeader(Application.VBE.ActiveCodePane)
26: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : sysAddModifiedTop - создание строки обновлени€
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub sysAddModifiedTop()
36:    Call sysAddModified(Application.VBE.ActiveCodePane)
37: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : sysAddTODOTop - создание коментари€ TODO
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub sysAddTODOTop()
47:    Call sysAddTODO(Application.VBE.ActiveCodePane)
48: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ShowTODOList - вызов формы со списком задач TODO
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub ShowTODOList()
58:    Call ModuleTODO.Show(0)
59: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : sysAddHeader - создание основного коментари€
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                         Description
'*
'* ByRef CurentCodePane As CodePane : активна€ код нанель VBE
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub sysAddHeader(ByRef CurentCodePane As CodePane)
73:    Dim nLine  As Long
74:    Dim i      As Byte
75:    Dim ProcKind As VBIDE.vbext_ProcKind
76:    Dim sProc  As String
77:    Dim sTemp  As String
78:    Dim sTime  As String
79:    Dim sType  As String
80:    Dim txtName As String
81:    Dim txtContacts As String
82:    Dim txtCopyright As String
83:    Dim txtOther As String
84:    Dim sProcDeclartion As String
85:    Dim sProcArguments As String
86:    Dim TBComment As ListObject
87:    Set TBComment = SHSNIPPETS.ListObjects(C_Const.TB_COMMENT)
88:    With TBComment.ListColumns(2)
89:        txtName = .Range(2, 1).Value
90:        If txtName = vbNullString Then txtName = Environ("UserName")
91:        txtContacts = .Range(3, 1).Value
92:        If txtContacts <> vbNullString Then txtContacts = "'* Contacts   :" & vbTab & txtContacts & vbCrLf
93:        txtCopyright = .Range(4, 1).Value
94:        If txtCopyright <> vbNullString Then txtCopyright = "'* Copyright  :" & vbTab & txtCopyright & vbCrLf
95:        txtOther = .Range(5, 1).Value
96:        If txtOther <> vbNullString Then txtOther = "'* Other     :" & vbTab & txtOther & vbCrLf
97:    End With
98:
99:    On Error Resume Next
100:    With CurentCodePane
101:        'получить начальную строку и им€ текущей процедуры
102:        GetCurrentProcInfo nLine, sProc, CurentCodePane
103:
104:        'создание '* * *' строки блока
105:        sTemp = Replace(String(90, "*"), "**", "* ")
106:
107:        'формат даты
108:        sTime = Format(Now, ctFormat)
109:
110:        'setup a type label
111:        If sProc = "" Then
112:            'верхней части модул€
113:            sType = "* Module     :"
114:            sProc = .CodeModule.Name
115:            nLine = 1
116:        Else
117:            For i = 0 To 4
118:                ProcKind = i
119:                sProcDeclartion = GetProcedureDeclaration(.CodeModule, sProc, ProcKind, LineSplitRemove)
120:                If sProcDeclartion <> vbNullString Then Exit For
121:            Next
122:            sProcArguments = AddStringParamertFromProcedureDeclaration(sProcDeclartion)
123:            sType = TypeProcedyreComment(sProcDeclartion)
124:        End If
125:
126:        'создание текстового блока дл€ вставки
127:        sTemp = "'" & sTemp & vbCrLf & _
                           "'" & sType & vbTab & sProc & vbCrLf & _
                           "'* Created    :" & vbTab & sTime & vbTab & vbCrLf & _
                           "'* Author     :" & vbTab & txtName & vbCrLf & _
                           txtContacts & _
                           txtCopyright & _
                           txtOther & _
                           sProcArguments & _
                           "'" & sTemp
136:        'вставка
137:        .CodeModule.InsertLines nLine, sTemp
138:    End With
139: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : sysAddModified - создание строки обновлени€
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                         Description
'*
'* ByRef CurentCodePane As CodePane : активна€ код нанель VBE
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub sysAddModified(ByRef CurentCodePane As CodePane)
153:    Dim nLine  As Long
154:    Dim sProc  As String
155:    Dim sTime  As String
156:    Dim sSecondLine As String
157:    Dim sUser As String
158:
159:    Const sUPDATE As String = "'* Updated    :"
160:    Const sFersLine As String = "'* Modified   :" & vbTab & "Date and Time" & vbTab2 & "Author" & vbTab4 & "Description" & vbCrLf
161:
162:    On Error Resume Next
163:    With CurentCodePane
164:        'получить начальную строку и им€ текущей процедуры
165:        GetCurrentProcInfo nLine, sProc, CurentCodePane
166:
167:        'формат даты
168:        sTime = Format(Now, ctFormat)
169:        sUser = SHSNIPPETS.ListObjects(C_Const.TB_COMMENT).Range(2, 2).Value
170:        If sUser = vbNullString Then sUser = Environ("UserName")
171:
172:        sSecondLine = sUPDATE & vbTab & sTime & vbTab & sUser & vbTab2
173:        If Not .CodeModule.Lines(nLine - 2, 1) Like sUPDATE & "*" Then
174:            sSecondLine = sFersLine & sSecondLine
175:        End If
176:        'вставка
177:        .CodeModule.InsertLines nLine - 1, sSecondLine
178:    End With
179: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : sysAddTODO - создание строки TODO
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                         Description
'*
'* ByRef CurentCodePane As CodePane : активна€ код нанель VBE
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub sysAddTODO(ByRef CurentCodePane As CodePane)
193:    Dim nLine  As Long
194:    Dim lStartLine As Long
195:    Dim lStartColumn As Long
196:    Dim lEndLine As Long
197:    Dim lEndColumn As Long
198:    Dim txtName As String
199:    Dim sFersLine As String
200:    Dim sSpec As String
201:
202:
203:    txtName = SHSNIPPETS.ListObjects(C_Const.TB_COMMENT).ListColumns(2).Range(2, 1).Value
204:    If txtName = vbNullString Then txtName = Environ("UserName")
205:
206:    On Error Resume Next
207:    With CurentCodePane
208:        'вставка
209:        .GetSelection lStartLine, lStartColumn, lEndLine, lEndColumn
210:        sSpec = VBA.String$(lStartColumn - 1, " ")
211:         sFersLine = sSpec & "'* TODO Created: " & VBA.Format$(Now, ctFormat) & " Author: " & txtName & vbCrLf & sSpec & "'*"
212:        .CodeModule.InsertLines lStartLine, sFersLine
213:    End With
214: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : GetCurrentProcInfo - получение строки и название процедуры
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                         Description
'*
'* ByRef nLine As Long              : номер строки
'* ByRef sProc As String            : название процедуры
'* ByRef CurentCodePane As CodePane : активна€ код нанель VBE
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Sub GetCurrentProcInfo(ByRef nLine As Long, ByRef sProc As String, ByRef CurentCodePane As CodePane)
230:    Dim t      As Long
231:
232:    With CurentCodePane
233:        'получаем им€ процедуры из положени€ курсора
234:        .GetSelection nLine, t, t, t
235:        sProc = .CodeModule.ProcOfLine(nLine, vbext_pk_Proc)
236:
237:        If sProc = "" Then
238:            'мы находимс€ в разделе объ€влени€; пропустите существующие пользовательские строки комментариев
239:            Do While .CodeModule.Find("'*", nLine, 1, .CodeModule.CountOfDeclarationLines, 2)
240:                nLine = nLine + 1
241:                If nLine > .CodeModule.CountOfDeclarationLines Then Exit Do
242:            Loop
243:        Else
244:            'в нутри поцедуры -> получаем номер первой строки
245:            nLine = .CodeModule.ProcBodyLine(sProc, vbext_pk_Proc)
246:        End If
247:    End With
248:
249: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddStringParamertFromProcedureDeclaration - возращает строку коментари€ с параметрами функции или процедуры
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                     Description
'*
'* ByVal sPocDeclartion As String : строка декларировани€ функции или процедуры
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Private Function AddStringParamertFromProcedureDeclaration(ByVal sPocDeclartion As String) As String
263:    Dim sDeclaration As String
264:    sDeclaration = Right$(sPocDeclartion, Len(sPocDeclartion) - InStr(1, sPocDeclartion, "("))
265:    sDeclaration = Left$(sDeclaration, InStr(1, sDeclaration, ")") - 1)
266:    'если нет параметров то возращем пусто
267:    If sDeclaration = vbNullString Then Exit Function
268:
269:    Dim arStr() As String
270:    Dim sTemp  As String
271:    Dim i      As Byte
272:    Dim iMaxLen As Byte
273:    Dim iTempLen As Byte
274:
275:    arStr = Split(sDeclaration, ",")
276:    iMaxLen = 0
277:    For i = 0 To UBound(arStr)
278:        iTempLen = Len(Trim$(arStr(i)))
279:        If iMaxLen < iTempLen Then iMaxLen = iTempLen
280:    Next i
281:
282:    sDeclaration = "'* Argument(s):" & String$(iMaxLen - Len(Trim$("'* Argument(s):")), " ") & vbTab2 & "Description" & vbCrLf & "'*" & vbCrLf
283:    For i = 0 To UBound(arStr)
284:        sTemp = "'* " & Trim$(arStr(i)) & String$(iMaxLen - Len(Trim$(arStr(i))), " ") & " :"
285:        sDeclaration = sDeclaration & sTemp & vbCrLf
286:    Next i
287:    AddStringParamertFromProcedureDeclaration = sDeclaration & "'* " & vbCrLf
288: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : TypeProcedyreComment - функци€ возвращает тип функции или процедуры
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef StrDeclarationProcedure As String : строка декларировани€ функции или процедуры
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function TypeProcedyreComment(ByRef StrDeclarationProcedure As String) As String
302:    If StrDeclarationProcedure Like "*Sub*" Then
303: TypeProcedyreComment = "* Sub        :"
304:    ElseIf StrDeclarationProcedure Like "*Function*" Then
305: TypeProcedyreComment = "* Function   :"
306:    ElseIf StrDeclarationProcedure Like "*Property Set*" Then
307:        TypeProcedyreComment = "* Prop Set   :"
308:    ElseIf StrDeclarationProcedure Like "*Property Get*" Then
309:        TypeProcedyreComment = "* Prop Get   :"
310:    ElseIf StrDeclarationProcedure Like "*Property Let*" Then
311:        TypeProcedyreComment = "* Prop Let   :"
312:    Else
313:        TypeProcedyreComment = "* Un Type    :"
314:    End If
End Function
