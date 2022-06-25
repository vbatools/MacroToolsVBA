Attribute VB_Name = "ZB_DeleteLinksFile"
Option Explicit
Option Private Module

Private Const sFULL_PATH As String = "Full path to the file:"
Private Const sFILE_LINKS As String = "File with a link"
Private Const sDELETE As String = "DELETE"

    Public Sub deleteAllLinksInFile()
9:    Dim sFileNameFull As Variant
10:    Dim bFlag       As Boolean
11:
12:    sFileNameFull = SelectedFile(vbNullString, False, "*.xls;*.xlsm;*.xlsx")
13:    If TypeName(sFileNameFull) = "Empty" Then Exit Sub
14:
15:    If MsgBox("Create backup files ?", vbYesNo + vbQuestion, "Removing passwords:") = vbYes Then
16:        bFlag = True
17:    End If
18:
19:    On Error GoTo errMsg
20:
21:    Dim sFullNameFile As String
22:    Dim cEditOpenXML As clsEditOpenXML
23:    Dim sPathLinks  As String
24:    Dim bMsg        As Boolean
25:
26:    sFullNameFile = sFileNameFull(1)
27:    Set cEditOpenXML = New clsEditOpenXML
28:    With cEditOpenXML
29:        .CreateBackupXML = bFlag
30:        .SourceFile = sFullNameFile
31:        .UnzipFile
32:        sPathLinks = .XLFolder & "externalLinks"
33:        If FileHave(sPathLinks, Directory) Then
34:            Dim objFso As Object
35:            Set objFso = CreateObject("Scripting.FileSystemObject")
36:            objFso.DeleteFolder (sPathLinks)
37:            Set objFso = Nothing
38:            bMsg = True
39:        End If
40:        .ZipAllFilesInFolder
41:    End With
42:    Set cEditOpenXML = Nothing
43:    If bMsg Then
44:        Call MsgBox("The links in the file were completely deleted: [" & sGetBaseName(sFullNameFile) & "]", vbInformation, "Deleting links:")
45:    Else
46:        Call MsgBox("In the file: [" & sGetBaseName(sFullNameFile) & "] no links to other files!", vbInformation, "Deleting links:")
47:    End If
48:
49:    Exit Sub
errMsg:
51:    Select Case Err.Number
        Case Else
53:            Call MsgBox("Error in deleteAllLinksInFile" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbCritical, "Mistake:")
54:            Call WriteErrorLog("deleteAllLinksInFile")
55:    End Select
56:    Set cEditOpenXML = Nothing
57: End Sub

     Public Sub getListAllLinksInFile()
60:    Dim sFileNameFull As Variant
61:
62:    sFileNameFull = SelectedFile(vbNullString, False, "*.xls;*.xlsm;*.xlsx")
63:    If TypeName(sFileNameFull) = "Empty" Then Exit Sub
64:
65:    On Error GoTo errMsg
66:
67:    Dim sFullNameFile As String
68:    Dim cEditOpenXML As clsEditOpenXML
69:    Dim sPathLinks  As String
70:    Dim bMsg        As Boolean
71:
72:    sFullNameFile = sFileNameFull(1)
73:    Set cEditOpenXML = New clsEditOpenXML
74:    With cEditOpenXML
75:        .CreateBackupXML = False
76:        .SourceFile = sFullNameFile
77:        .UnzipFile
78:        sPathLinks = .XLFolder & "externalLinks\_rels"
79:        If FileHave(sPathLinks, Directory) Then
80:            Dim objFso As Object
81:            Dim objFolder As Object
82:            Dim objFile As Object
83:            Dim i   As Integer
84:            Dim j   As Integer
85:            Dim arrFile() As String
86:            Dim sXML As String
87:            Const sTARGET As String = " Target="
88:
89:            Set objFso = CreateObject("Scripting.FileSystemObject")
90:            Set objFolder = objFso.GetFolder(sPathLinks)
91:
92:            For Each objFile In objFolder.Files
93:                If objFile.Name Like "*.rels" Then
94:                    j = j + 1
95:                    ReDim Preserve arrFile(1 To 2, 1 To j)
96:                    arrFile(1, j) = objFile.Name
97:                    sXML = .GetXMLFromFile(arrFile(1, j), sPathLinks & Application.PathSeparator)
98:                    If sXML Like "*" & sTARGET & VBA.Chr$(34) & "*" Then
99:                        sXML = VBA.Right$(sXML, VBA.Len(sXML) - VBA.InStr(1, sXML, sTARGET) - VBA.Len(sTARGET))
100:                        sXML = VBA.Left$(sXML, VBA.InStr(1, sXML, VBA.Chr$(34)) - 1)
101:                        arrFile(2, j) = sXML
102:                    End If
103:                End If
104:            Next
105:            Set objFolder = Nothing
106:            Set objFso = Nothing
107:            bMsg = True
108:        End If
109:        .ZipAllFilesInFolder
110:    End With
111:    Set cEditOpenXML = Nothing
112:
113:    If bMsg Then
114:        ActiveWorkbook.Worksheets.Add
115:        With ActiveCell
116:            .Value = sFULL_PATH
117:            .Offset(0, 1).Value = sFullNameFile
118:            .Offset(1, 0).Value = sFILE_LINKS
119:            .Offset(1, 1).Value = "The file to which the link goes"
120:            .Offset(1, 2).Value = "Action (put down)"
121:            .Offset(2, 0).Resize(UBound(arrFile, 2), UBound(arrFile, 1)).Value2 = WorksheetFunction.Transpose(arrFile)
122:            .Offset(2, 2).Resize(UBound(arrFile, 2), 1).Value2 = sDELETE
123:        End With
124:        Call MsgBox("Creating a list of links from a file:[" & sGetBaseName(sFullNameFile) & "]", vbInformation, "Creating a list:")
125:    Else
126:        Call MsgBox("In the file: [" & sGetBaseName(sFullNameFile) & "] no links to other files!", vbInformation, "Creating a list:")
127:    End If
128:
129:    Exit Sub
errMsg:
131:    Select Case Err.Number
        Case Else
133:            Call MsgBox("Error in getListAllLinksInFile" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbCritical, "Mistake:")
134:            Call WriteErrorLog("getListAllLinksInFile")
135:    End Select
136:    Set cEditOpenXML = Nothing
137:
138: End Sub

Public Sub deleteLinksOnList()
141:    Dim bFlag       As Boolean
142:    Dim arrVal      As Variant
143:    Dim errMsg      As String
144:    Dim sFullNameFile As String
145:
146:    On Error GoTo errMsg
147:
148:    With ActiveSheet
149:        Dim lLastRow As Long
150:        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
151:        If lLastRow < 3 Then
152:            Call MsgBox("There is no data table!", vbCritical, "Mistake:")
153:            Exit Sub
154:        End If
155:        If .Cells(1, 1).Value <> sFULL_PATH Then
156:            errMsg = "Field not found [" & sFULL_PATH & "]" & vbNewLine
157:        End If
158:        If .Cells(2, 1).Value <> sFILE_LINKS Then
159:            errMsg = errMsg & "Field not found [" & sFILE_LINKS & "]" & vbNewLine
160:        End If
161:
162:        sFullNameFile = .Cells(1, 2).Value
163:
164:        If sFullNameFile = vbNullString Then
165:            errMsg = errMsg & "The file path is not set:" & vbNewLine
166:        ElseIf Not FileHave(sFullNameFile) Then
167:            errMsg = errMsg & "The path to the file does not exist" & vbNewLine
168:        End If
169:
170:        If errMsg <> vbNullString Then
171:            Call MsgBox("The data table is not recognized:" & vbNewLine & errMsg, vbCritical, "Mistake:")
172:            Exit Sub
173:        End If
174:
175:        arrVal = .Range(.Cells(3, 1), .Cells(lLastRow, 3)).Value2
176:    End With
177:
178:    If MsgBox("Create backup files ?", vbYesNo + vbQuestion, "Removing passwords:") = vbYes Then
179:        bFlag = True
180:    End If
181:
182:    Dim cEditOpenXML As clsEditOpenXML
183:    Dim sPathLinks  As String
184:    Dim sPathLinksRels As String
185:    Dim bMsg        As Boolean
186:    Dim i           As Integer
187:    Dim iCount      As Integer
188:    Dim sfileName   As String
189:
190:    Set cEditOpenXML = New clsEditOpenXML
191:    With cEditOpenXML
192:        .CreateBackupXML = bFlag
193:        .SourceFile = sFullNameFile
194:        .UnzipFile
195:        sPathLinks = .XLFolder & "externalLinks" & Application.PathSeparator
196:        sPathLinksRels = sPathLinks & "_rels" & Application.PathSeparator
197:
198:        For i = 1 To UBound(arrVal)
199:            sfileName = arrVal(i, 1)
200:            If arrVal(i, 3) = sDELETE And FileHave(sPathLinksRels & sfileName) Then
201:                Call Kill(sPathLinks & VBA.Replace(sfileName, ".rels", vbNullString))
202:                Call Kill(sPathLinksRels & sfileName)
203:                bMsg = True
204:                iCount = iCount + 1
205:            End If
206:        Next i
207:        .ZipAllFilesInFolder
208:    End With
209:    Set cEditOpenXML = Nothing
210:
211:    If bMsg Then
212:        Call MsgBox("The links in the file were deleted: [" & sGetBaseName(sFullNameFile) & "]" & vbNewLine & "Deleted: [" & iCount & "] connections!", vbInformation, "Deleting links:")
213:    Else
214:        Call MsgBox("In the file: [" & sGetBaseName(sFullNameFile) & "] no links to other files!", vbInformation, "Deleting links:")
215:    End If
216:
217:    Exit Sub
errMsg:
219:    Select Case Err.Number
        Case Else
221:            Call MsgBox("Error in deleteLinksOnList" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl, vbOKOnly + vbCritical, "Mistake:")
222:            Call WriteErrorLog("deleteLinksOnList")
223:    End Select
224:    Set cEditOpenXML = Nothing
End Sub
