Attribute VB_Name = "X_InfoFile"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : X_InfoFile - модуль редактирования свойств книги excel
'* Created    : 20-07-2020 12:31
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ShowProp функция создания списка свойств файла
'* Created    : 20-07-2020 12:32
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook : ссылка на книгу
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function ShowProp(ByRef wb As Workbook) As String
22:    Dim DP          As DocumentProperty
23:    Dim Txt         As String
24:
25:    With wb
26:        On Error Resume Next
27:        For Each DP In .BuiltinDocumentProperties
28:            Txt = Txt & DP.Name & "||" & " " & DP.Value & vbNewLine
29:        Next
30:        On Error GoTo 0
31:    End With
32:    ShowProp = VBA.Left$(Txt, VBA.Len(Txt) - 2)
33: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DelAllProp - функция удаления всех свойств книги, возращает количество удаленных свойств
'* Created    : 20-07-2020 12:33
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook : ссылка на книгу
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function DelAllProp(ByRef wb As Workbook) As Byte
47:    Dim DP          As DocumentProperty
48:    Dim byItem      As Byte
49:    With wb
50:        On Error Resume Next
51:        For Each DP In .BuiltinDocumentProperties
52:            If DP.Value <> vbNullString Then
53:                DP.Value = vbNullString
54:                byItem = byItem + 1
55:            End If
56:        Next
57:        On Error GoTo 0
58:    End With
59:    DelAllProp = byItem
60: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WriteOneProp - редактирование значения одного выбраного свойства
'* Created    : 20-07-2020 12:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByRef WB As Workbook     : ссылка на книгу
'* ByVal NameProp As String : название совйства
'* ByVal Val As String      : новое значение свойства
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function WriteOneProp(ByRef wb As Workbook, ByVal NameProp As String, ByVal Val As String) As String
76:    Dim DP          As DocumentProperty
77:    On Error GoTo errMsg
78:    Set DP = wb.BuiltinDocumentProperties(NameProp)
79:    DP.Value = Val
80:    WriteOneProp = True
81:    Exit Function
errMsg:
83:    WriteOneProp = Err.Description
84:    On Error GoTo 0
85: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ShowCustomProp - создания списка пользовательских свойств файла
'* Created    : 20-07-2020 12:35
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook : ссылка на книгу
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function ShowCustomProp(ByRef wb As Workbook) As String
99:    Dim i           As Integer
100:    Dim Txt         As String
101:    With wb
102:        If .CustomDocumentProperties.Count > 0 Then
103:            For i = 1 To .CustomDocumentProperties.Count
104:                Txt = Txt & .CustomDocumentProperties(i).Name & "||" & " " & .CustomDocumentProperties(i).Value & vbNewLine
105:            Next i
106:        End If
107:    End With
108:    ShowCustomProp = Txt
109: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DelAllCustomProp - функция удаления всех пользовательских свойств книги, возращает количество удаленных свойств
'* Created    : 20-07-2020 12:35
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook : ссылка на книгу
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function DelAllCustomProp(ByRef wb As Workbook) As Byte
123:    Dim i           As Integer
124:    Dim byItem      As Byte
125:    With wb
126:        If .CustomDocumentProperties.Count > 0 Then
127:            For i = .CustomDocumentProperties.Count To 1 Step -1
128:                .CustomDocumentProperties(i).Delete
129:                byItem = byItem + 1
130:            Next i
131:        End If
132:    End With
133:    DelAllCustomProp = byItem
134: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DelOneCustomProp - удаление выпраного пользовательского свойства
'* Created    : 20-07-2020 12:36
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByRef WB As Workbook     : ссылка на книгу
'* ByVal NameProp As String : название совйства
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function DelOneCustomProp(ByRef wb As Workbook, ByVal NameProp As String) As Boolean
149:    Dim i           As Integer
150:    With wb
151:        If .CustomDocumentProperties.Count > 0 Then
152:            For i = 1 To .CustomDocumentProperties.Count
153:                If .CustomDocumentProperties(i).Name = NameProp Then
154:                    .CustomDocumentProperties(i).Delete
155:                    DelOneCustomProp = True
156:                    Exit Function
157:                End If
158:            Next i
159:        End If
160:    End With
161: End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddOneCustomProp - процедура создания пользовательского свойства
'* Created    : 20-07-2020 12:37
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByRef WB As Workbook     : ссылка на книгу
'* ByVal NameProp As String : название совйства
'* ByVal Val As Variant     : значение свойства
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub AddOneCustomProp(ByRef wb As Workbook, ByVal NameProp As String, ByVal Val As Variant)
177:    Call DelOneCustomProp(wb, NameProp)
178:    Call wb.CustomDocumentProperties.Add(NameProp, False, msoPropertyTypeString, VBA.CStr(Val))
179: End Sub
Public Function GetOneProp(ByRef wb As Workbook, ByVal NameProp As String) As String
181:    GetOneProp = wb.BuiltinDocumentProperties(NameProp).Value
End Function
