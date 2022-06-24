VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEditOpenXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : clsEditOpenXML - работа с XML
'* Created    : 15-09-2019 15:53
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-10-2019 12:43    VBATools  add module delete XML Node
'* Updated    : 30-10-2019 13:13    VBATools  add project is uniwable for Word
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private mbCreateBackup As Boolean
Private mvSourceFile As Variant
Private mvXLFolder As Variant
Private msSheet2Change As String
Private msSheetId As String
Private msSheetFileName As String
Private mbAddedZip As Boolean
Private mvUnzipFolder As Variant
Private mvXMLFolderRoot As Variant
Private mvTipeFile As String
Private mvTipeFileRels As String

Private Const sUNZIPPED As String = "UnZipped_"
Public Enum enBaseName
    dimension = 1
    SheetViews = 2
    sheetFormatPr = 3
    sheetData = 4
    sheetProtection = 5
    pageMargins = 6
    sDrawing = 7
    workbookProtection = 8
End Enum
Public Enum XMLFolder
    XMLFolder_root = 1
    XMLFolder_rels = 2
    XMLFolder_xl = 3
    XMLFolder_customUI = 4
    XMLFolder_docProps = 5
End Enum

    Public Sub CopyFiles2(ByVal sFileSpec As String, ByVal sSourcePath As String, ByVal sTargetPath As String)
43:    Dim sFile  As String
44:    sFile = Dir(XLFolder & sSourcePath & sFileSpec)
45:    Do While Len(sFile) > 0
46:        FileCopy XLFolder & sSourcePath & sFile, sTargetPath    '& sFile
47:        sFile = Dir()
48:    Loop
49: End Sub

    Public Function GetXMLFromFile(sfileName As String, Optional sFolder As String = vbNullString) As String
52:    '-------------------------------------------------------------------------
53:    ' Procedure : GetXMLFromFile
54:    ' Purpose   : Gets the XML code from the foldername\filename
55:    '-------------------------------------------------------------------------
       If sFolder = vbNullString Then sFolder = XLFolder
56:    Dim oXMLDoc As MSXML2.DOMDocument
57:    If Len(XLFolder) = 0 Then
58:        GetXMLFromFile = vbNullString
59:    Else
60:        Set oXMLDoc = New MSXML2.DOMDocument
61:        oXMLDoc.Load sFolder & sfileName
62:        GetXMLFromFile = oXMLDoc.XML
63:        Set oXMLDoc = Nothing
64:    End If
65: End Function

    Public Function DelPartXMLFromFile(sfileName As String, TupeXML As enBaseName) As Boolean
68:    '-------------------------------------------------------------------------
69:    ' Procedure : DelPartXMLFromFile
70:    ' Purpose   : Del Part XML in XML code from the foldername\filename
71:    '-------------------------------------------------------------------------
72:    DelPartXMLFromFile = False
73:    If Len(XLFolder) <> 0 Then
74:        Dim oXMLDoc As MSXML2.DOMDocument
75:        Dim i  As Integer
76:        Dim XMLChild As String, XMLParent As String, sURI As String
77:        Set oXMLDoc = New MSXML2.DOMDocument
78:        oXMLDoc.Load XLFolder & sfileName
79:        XMLParent = oXMLDoc.XML
80:        For i = oXMLDoc.ChildNodes(1).ChildNodes.Length - 1 To 0 Step -1
81:            If BaseName(TupeXML) = oXMLDoc.ChildNodes(1).ChildNodes(i).BaseName Then
82:                XMLChild = oXMLDoc.ChildNodes(1).ChildNodes(i).XML
83:                sURI = "xmlns=" & Chr(34) & oXMLDoc.ChildNodes(1).ChildNodes(i).NamespaceURI & Chr(34) & " "
84:                XMLChild = Replace(XMLChild, sURI, "")
85:                XMLParent = Replace(XMLParent, XMLChild, vbNullString)
86:                oXMLDoc.loadXML XMLParent
87:                oXMLDoc.Save XLFolder & sfileName
88:                DelPartXMLFromFile = True
89:                Exit For
90:            End If
91:        Next i
92:        Set oXMLDoc = Nothing
93:    End If
94: End Function

     Public Sub WriteXML2File(sXML As String, sfileName As String, sXMLFolder As XMLFolder)
97:    '-------------------------------------------------------------------------
98:    ' Procedure : WriteXML2File
99:    ' Purpose   : Writes sXML to sFileName
100:    '             Modified by Ken Puls 2009-05-12
101:    '             Adjusted to add ability to write to customUI container
102:    '-------------------------------------------------------------------------
103:    Dim oXMLDoc As MSXML2.DOMDocument
104:    Set oXMLDoc = New MSXML2.DOMDocument
105:
106:    'If attempting to write a customUI component, test to see if one exists
107:
108:    'Should probably test the .rels file to see if the CustomUI relationship exists...
109:    If sXMLFolder = XMLFolder_customUI Then
110:        If Not FolderExists(XMLFolder(XMLFolder_customUI)) Then
111:
112:            MkDir XMLFolder(XMLFolder_customUI)
113:            'Write the XML to the file
114:            oXMLDoc.loadXML sXML
115:            oXMLDoc.Save XMLFolder(sXMLFolder) & sfileName
116:            'CustomUI has not been created yet.  Rels file needs to be adjusted
117:            AddCustomUIToRels
118:        End If
119:    End If
120:
121:    'Write the XML to the file
122:    oXMLDoc.loadXML sXML
123:    oXMLDoc.Save XMLFolder(sXMLFolder) & sfileName
124: End Sub

     Public Sub AddCustomUIToRels()
127:    'Macro Purpose: Add the customUI relationship to the rels file
128:
129:    Dim oXMLDoc As MSXML2.DOMDocument
130:    '    Dim oXMLElement As MSXML2.IXMLDOMElement
131:    Dim oXMLElement As MSXML2.IXMLDOMNode
132:    Dim oXMLAttrib As MSXML2.IXMLDOMAttribute
133:    Dim oNamedNodeMap As MSXML2.IXMLDOMNamedNodeMap
134:    Dim oXMLRelsList As MSXML2.IXMLDOMNodeList
135:    'Create a new XML document
136:    Set oXMLDoc = New MSXML2.DOMDocument
137:    'Attach to the root element of the .rels file
138:    oXMLDoc.Load XMLFolder(XMLFolder_rels) & ".rels"
139:
140:    'Create a new relationship element in the .rels file
141:    Set oXMLElement = oXMLDoc.createNode(1, "Relationship", _
                    "http://schemas.openxmlformats.org/package/2006/relationships")
143:    Set oNamedNodeMap = oXMLElement.Attributes
144:
145:    'Create ID attribute for the element
146:    Set oXMLAttrib = oXMLDoc.createAttribute("Id")
147:    oXMLAttrib.NodeValue = "cuID"
148:    oNamedNodeMap.setNamedItem oXMLAttrib
149:
150:    'Create Type attribute for the element
151:    Set oXMLAttrib = oXMLDoc.createAttribute("Type")
152:    oXMLAttrib.NodeValue = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
153:    oNamedNodeMap.setNamedItem oXMLAttrib
154:
155:    'Create Target element for the attribute
156:    Set oXMLAttrib = oXMLDoc.createAttribute("Target")
157:    oXMLAttrib.NodeValue = "customUI/customUI.xml"
158:    oNamedNodeMap.setNamedItem oXMLAttrib
159:
160:    'Now insert the new node at the proper location
161:    Set oXMLRelsList = oXMLDoc.SelectNodes("/Relationships")
162:    oXMLRelsList.Item(0).appendChild oXMLElement
163:    'Save the .rels file
164:    oXMLDoc.Save XMLFolder(XMLFolder_rels) & ".rels"
165:
166:    Set oXMLAttrib = Nothing
167:    Set oXMLElement = Nothing
168:    Set oXMLDoc = Nothing
169: End Sub

     Private Function GetSheetIdFromSheetName(sSheetName) As String
172:    '-------------------------------------------------------------------------
173:    ' Procedure : GetSheetIdFromSheetName
174:    ' Purpose   : Finds out what the SheetId of sSheetname is
175:    '             by reading Workbook.xml
176:    '-------------------------------------------------------------------------
177:    Dim oXMLDoc As MSXML2.DOMDocument
178:    Dim oXMLNode As MSXML2.IXMLDOMNode
179:    Dim oXMLNodeList As MSXML2.IXMLDOMNodeList
180:    If mvXLFolder <> "" And Sheet2Change <> "" Then
181:        Set oXMLDoc = New MSXML2.DOMDocument
182:        oXMLDoc.Load XLFolder & "workbook.xml"
183:        Set oXMLNodeList = oXMLDoc.SelectNodes("/workbook/sheets/sheet")
184:        For Each oXMLNode In oXMLNodeList
185:            If oXMLNode.Attributes.getNamedItem("name").NodeValue = sSheetName Then
186:                GetSheetIdFromSheetName = oXMLNode.Attributes.getNamedItem("r:id").NodeValue
187:                Exit Function
188:            End If
189:        Next
190:    End If
191: End Function

     Public Function GetSheetFileNameFromId(sSheetId As String, Optional sType As String = "Id") As String
194:    '-------------------------------------------------------------------------
195:    ' Procedure : GetSheetFileNameFromId
196:    ' Purpose   : Fetches the name of the xml file belonging to the sheet with id SheetId.
197:    '-------------------------------------------------------------------------
198:    Dim oXMLDoc As MSXML2.DOMDocument
199:    Dim oXMLNode As MSXML2.IXMLDOMNode
200:    Dim oXMLNodeList As MSXML2.IXMLDOMNodeList
201:    If mvXLFolder <> "" And Sheet2Change <> "" Then
202:        Set oXMLDoc = New MSXML2.DOMDocument
203:        oXMLDoc.Load XLFolder & mvTipeFileRels
204:
205:        Set oXMLNodeList = oXMLDoc.SelectNodes("/Relationships/Relationship")
206:        For Each oXMLNode In oXMLNodeList
207:            If sType = "Id" Then
208:                If oXMLNode.Attributes.getNamedItem("Id").NodeValue = sSheetId Then
209:                    GetSheetFileNameFromId = oXMLNode.Attributes.getNamedItem("Target").NodeValue
210:                    Exit Function
211:                End If
212:            Else
213:                If oXMLNode.Attributes.getNamedItem("Type").NodeValue Like sSheetId Then
214:                    GetSheetFileNameFromId = oXMLNode.Attributes.getNamedItem("Target").NodeValue
215:                    Exit Function
216:                End If
217:            End If
218:        Next
219:    End If
220: End Function

     Public Function GetSheetNameFromId(sId As String) As String
223:    '-------------------------------------------------------------------------
224:    ' Procedure : GetSheetNameFromId
225:    ' Purpose   : Returns the sheetname belonging to a sheetId
226:    '-------------------------------------------------------------------------
227:    Dim oXMLDoc As MSXML2.DOMDocument
228:    Dim oXMLNode As MSXML2.IXMLDOMNode
229:    Dim oXMLNodeList As MSXML2.IXMLDOMNodeList
230:    If mvXLFolder <> "" Then
231:        Set oXMLDoc = New MSXML2.DOMDocument
232:        oXMLDoc.Load XLFolder & "workbook.xml"
233:        Set oXMLNodeList = oXMLDoc.SelectNodes("/workbook/sheets/sheet")
234:        For Each oXMLNode In oXMLNodeList
235:            If oXMLNode.Attributes.getNamedItem("r:id").NodeValue = "rId" & Val(sId) + 1 Then
236:                GetSheetNameFromId = oXMLNode.Attributes.getNamedItem("name").NodeValue
237:                'Got it, get out
238:                Exit Function
239:            End If
240:        Next
241:    End If
242: End Function

     Public Sub ZipAllFilesInFolder()
245:    '-------------------------------------------------------------------------
246:    ' Procedure : ZipAllFilesInFolder
247:    ' Purpose   : Zips all files in a folder (including subfolders) whilst retaining the folder structure
248:    '-------------------------------------------------------------------------
249:    Dim oShellApp As Object
250:    Dim sDate  As String
251:    Dim sDefPath As String
252:    Dim vFileNameZip As Variant
253:    Dim FSO    As Object
254:    Dim lFileCt As Long
255:    Set FSO = CreateObject("scripting.filesystemobject")
256:
257:    'To ensure a unique filename,
258:    'append date and time to the name of the current file
259:
260:    sDate = Format(Now, " dd-mmm-yy h-mm-ss")
261:    vFileNameZip = SourceFile & sDate & ".zip"
262:
263:    'Create empty Zip File
264:    NewZip vFileNameZip
265:
266:    Set oShellApp = CreateObject("Shell.Application")
267:
268:    'Count how many items are in the "old" folder
269:    lFileCt = oShellApp.Namespace(FolderName & sUNZIPPED & FileNameXML & Application.PathSeparator).Items.Count
270:
271:    'Copy the files to the compressed folder
272:    oShellApp.Namespace(vFileNameZip).CopyHere oShellApp.Namespace(FolderName & sUNZIPPED & FileNameXML & Application.PathSeparator).Items
273:
274:    'Keep script waiting until we have same # of files in the new folder
275:    On Error Resume Next
276:    Do Until oShellApp.Namespace(vFileNameZip).Items.Count = lFileCt
277:        Application.Wait (Now + TimeValue("0:00:01"))
278:    Loop
279:    DoEvents
280:
281:    'Remove original file
282:    Kill SourceFile
283:
284:    'Rename new zipped file to same name as original file (with .zip appended)
285:    Name vFileNameZip As SourceFile
286:    On Error Resume Next
287:
288:    'Now remove old folder, just in case something went haywire
289:    FSO.DeleteFolder FolderName & sUNZIPPED & FileNameXML, True
290:    On Error GoTo 0
291:
292:    Set oShellApp = Nothing
293: End Sub

     Public Sub UnzipFile()
296:    '-------------------------------------------------------------------------
297:    ' Procedure : UnzipFile
298:    ' Purpose   : Unzips all files in a zip file to a designated folder
299:    '-------------------------------------------------------------------------
300:    'Courtesy www.rondebruin.nl
301:    Dim FSO    As Object
302:    Dim oShellApp As Object
303:    Set FSO = CreateObject("scripting.filesystemobject")
304:
305:    'Derive the folder to unzip to from the location of the sourcefile
306:    XMLFolderRoot = FolderName
307:
308:    'A dedicated unzip folder will be created in the same folder as the sourcefile,
309:    'called ..\Unzipped Filename\
310:    If Right(XMLFolderRoot, 1) <> Application.PathSeparator Then
311:        XMLFolderRoot = XMLFolderRoot & Application.PathSeparator & sUNZIPPED & FileNameXML & " \ """
312:    Else
313:        XMLFolderRoot = XMLFolderRoot & sUNZIPPED & FileNameXML & Application.PathSeparator
314:    End If
315:    On Error Resume Next
316:    'Remove all previous existing folders
317:    FSO.DeleteFolder XMLFolderRoot & "*", True
318:    Kill XMLFolderRoot & "*.*"
319:    On Error GoTo 0
320:
321:    'Create normal folder
322:    If FolderExists(XMLFolderRoot) = False Then
323:        MkDir XMLFolderRoot
324:    End If
325:
326:    XLFolder = XMLFolderRoot & mvTipeFile & Application.PathSeparator
327:
328:    Set oShellApp = CreateObject("Shell.Application")
329:    'Copy the files in the newly created folder
330:    oShellApp.Namespace(XMLFolderRoot).CopyHere oShellApp.Namespace(SourceFile).Items
331:
332:    On Error Resume Next
333:    'Clean up temp folder
334:    FSO.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
335:
336:    Set oShellApp = Nothing
337:    Set FSO = Nothing
338:    Exit Sub
339: End Sub

     Sub NewZip(sPath)
342:    If Len(Dir(sPath)) > 0 Then Kill sPath
343:    Open sPath For Output As #1
344:    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
345:    Close #1
346: End Sub

     Public Property Get CreateBackupXML() As Boolean
349:    CreateBackupXML = mbCreateBackup
350: End Property

     Public Property Let CreateBackupXML(ByVal bCreateBackup As Boolean)
353:    mbCreateBackup = bCreateBackup
354: End Property

     Private Sub Class_Initialize()
357:    'Set defaults
358:    CreateBackupXML = True
359: End Sub

     Public Property Get SourceFile() As Variant
362:    SourceFile = mvSourceFile
363: End Property

     Public Property Let SourceFile(ByVal vSourceFile As Variant)    '''
366:    mvSourceFile = vSourceFile
367:    Call WathIsFileType(vSourceFile)
368:    If CreateBackupXML Then
369:        If Len(Dir(vSourceFile & "(backup)")) > 0 Then
370:            Kill vSourceFile & "(backup)"
371:        End If
372:        FileCopy vSourceFile, vSourceFile & "(backup)"
373:    End If
374:    If Not vSourceFile Like "*.zip" Then
375:        Name vSourceFile As vSourceFile & ".zip"
376:        mvSourceFile = mvSourceFile & ".zip"
377:        AddedZip = True
378:    End If
379: End Property

     Public Property Get FolderName() As Variant
382:    FolderName = Mid(SourceFile, 1, InStrRev(SourceFile, Application.PathSeparator))
383: End Property

     Public Property Get FileNameXML() As Variant
386:    If SourceFile <> "" Then
387:        FileNameXML = Mid(SourceFile, InStrRev(SourceFile, Application.PathSeparator) + 1, Len(SourceFile))
388:    End If
389: End Property

     Public Property Get XLFolder() As Variant
392:    XLFolder = mvXLFolder
393: End Property

     Public Property Let XLFolder(ByVal vXLFolder As Variant)
396:    mvXLFolder = vXLFolder
397: End Property

     Public Property Get XMLFolder(sXMLFolder As XMLFolder) As String
400:    Select Case sXMLFolder
        Case Is = XMLFolder_root
402:            XMLFolder = mvXMLFolderRoot
403:        Case Is = XMLFolder_customUI
404:            XMLFolder = mvXMLFolderRoot & "customUI" & Application.PathSeparator
405:        Case Is = XMLFolder_docProps
406:            XMLFolder = mvXMLFolderRoot & "docProps" & Application.PathSeparator
407:        Case Is = XMLFolder_rels
408:            XMLFolder = mvXMLFolderRoot & "_rels" & Application.PathSeparator
409:        Case Is = XMLFolder_xl
410:            XMLFolder = mvXMLFolderRoot & mvTipeFile & Application.PathSeparator
411:    End Select
412: End Property

     Public Property Get BaseName(sBaseName As enBaseName) As String
415:    Select Case sBaseName
        Case Is = dimension
417:            BaseName = "dimension"
418:        Case Is = SheetViews
419:            BaseName = "SheetViews"
420:        Case Is = sheetFormatPr
421:            BaseName = "sheetFormatPr"
422:        Case Is = sheetData
423:            BaseName = "sheetData"
424:        Case Is = sheetProtection
425:            BaseName = "sheetProtection"
426:        Case Is = pageMargins
427:            BaseName = "pageMargins"
428:        Case Is = sDrawing
429:            BaseName = "Drawing"
430:        Case Is = workbookProtection
431:            BaseName = "workbookProtection"
432:    End Select
433: End Property

     Public Property Get Sheet2Change() As String
436:    Sheet2Change = msSheet2Change
437: End Property

     Public Property Let Sheet2Change(ByVal sSheet2Change As String)
440:    msSheet2Change = sSheet2Change
441:    SheetId = GetSheetIdFromSheetName(sSheet2Change)
442:    If SheetId <> "" Then
443:        SheetFileName = GetSheetFileNameFromId(SheetId)
444:    End If
445: End Property

     Public Property Get SheetId() As String
448:    SheetId = msSheetId
449: End Property

     Public Property Let SheetId(ByVal sSheetId As String)
452:    msSheetId = sSheetId
453: End Property

     Public Property Get SheetFileName() As String
456:    SheetFileName = msSheetFileName
457: End Property

     Public Property Let SheetFileName(ByVal sSheetFileName As String)
460:    msSheetFileName = sSheetFileName
461: End Property

     Private Property Get AddedZip() As Boolean
464:    AddedZip = mbAddedZip
465: End Property

     Private Property Let AddedZip(ByVal bAddedZip As Boolean)
468:    mbAddedZip = bAddedZip
469: End Property

     Private Sub Class_Terminate()
472:    Dim FSO    As Object
473:    If AddedZip Then
474:        'Remove .zip from file again
475:        Name SourceFile As Left(SourceFile, Len(SourceFile) - 4)
476:    End If
477:    'Remove zip folder
478:    On Error Resume Next
479:    Set FSO = CreateObject("scripting.filesystemobject")
480:    'Omit trailing backslash
481:    FSO.DeleteFolder Left(UnzipFolder, Len(UnzipFolder) - 1), True
482:    Set FSO = Nothing
483: End Sub

     Private Property Get UnzipFolder() As Variant
486:    UnzipFolder = mvUnzipFolder
487: End Property

     Private Property Let UnzipFolder(ByVal vUnzipFolder As Variant)
490:    mvUnzipFolder = vUnzipFolder
491: End Property

     Private Property Get XMLFolderRoot() As Variant
494:    XMLFolderRoot = mvXMLFolderRoot
495: End Property

     Private Property Let XMLFolderRoot(ByVal vXMLFolderRoot As Variant)
498:    mvXMLFolderRoot = vXMLFolderRoot
499: End Property

     Private Sub WathIsFileType(ByVal sPath As String)
502:    Select Case sGetExtensionName(sPath)
        Case "docm", "dotm", "dotx", "docx", "doc", "dot":
504:            mvTipeFile = "word"
505:            mvTipeFileRels = "_rels\document.xml.rels"
506:        Case "xlsm", "xlsb", "xlam", "xla", "xlsx", "xls":
507:            mvTipeFile = "xl"
508:            mvTipeFileRels = "_rels\workbook.xml.rels"
509:        Case "pptx", "pptm", "ppt", "potx", "potm", "pot":
510:            mvTipeFile = "ppt"
511:            mvTipeFileRels = "_rels\presentation.xml.rels"
512:    End Select
513: End Sub
     Public Property Get TypeFile() As String
515:    TypeFile = mvTipeFile
516: End Property
Public Property Get TipeFileRels() As String
518:    TipeFileRels = mvTipeFileRels
End Property
