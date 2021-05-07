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
46:    Dim sFile  As String
47:    sFile = Dir(XLFolder & sSourcePath & sFileSpec)
48:    Do While Len(sFile) > 0
49:        FileCopy XLFolder & sSourcePath & sFile, sTargetPath    '& sFile
50:        sFile = Dir()
51:    Loop
52: End Sub

    Public Function GetXMLFromFile(sFileName As String, Optional sFolder As String = vbNullString) As String
55:    '-------------------------------------------------------------------------
56:    ' Procedure : GetXMLFromFile
57:    ' Purpose   : Gets the XML code from the foldername\filename
58:    '-------------------------------------------------------------------------
59:       If sFolder = vbNullString Then sFolder = XLFolder
60:    Dim oXMLDoc As MSXML2.DOMDocument
61:    If Len(XLFolder) = 0 Then
62:        GetXMLFromFile = vbNullString
63:    Else
64:        Set oXMLDoc = New MSXML2.DOMDocument
65:        oXMLDoc.Load sFolder & sFileName
66:        GetXMLFromFile = oXMLDoc.XML
67:        Set oXMLDoc = Nothing
68:    End If
69: End Function

    Public Function DelPartXMLFromFile(sFileName As String, TupeXML As enBaseName) As Boolean
72:    '-------------------------------------------------------------------------
73:    ' Procedure : DelPartXMLFromFile
74:    ' Purpose   : Del Part XML in XML code from the foldername\filename
75:    '-------------------------------------------------------------------------
76:    DelPartXMLFromFile = False
77:    If Len(XLFolder) <> 0 Then
78:        Dim oXMLDoc As MSXML2.DOMDocument
79:        Dim i  As Integer
80:        Dim XMLChild As String, XMLParent As String, sURI As String
81:        Set oXMLDoc = New MSXML2.DOMDocument
82:        oXMLDoc.Load XLFolder & sFileName
83:        XMLParent = oXMLDoc.XML
84:        For i = oXMLDoc.ChildNodes(1).ChildNodes.Length - 1 To 0 Step -1
85:            If BaseName(TupeXML) = oXMLDoc.ChildNodes(1).ChildNodes(i).BaseName Then
86:                XMLChild = oXMLDoc.ChildNodes(1).ChildNodes(i).XML
87:                sURI = "xmlns=" & Chr(34) & oXMLDoc.ChildNodes(1).ChildNodes(i).NamespaceURI & Chr(34) & " "
88:                XMLChild = Replace(XMLChild, sURI, "")
89:                XMLParent = Replace(XMLParent, XMLChild, vbNullString)
90:                oXMLDoc.loadXML XMLParent
91:                oXMLDoc.Save XLFolder & sFileName
92:                DelPartXMLFromFile = True
93:                Exit For
94:            End If
95:        Next i
96:        Set oXMLDoc = Nothing
97:    End If
98: End Function

     Public Sub WriteXML2File(sXml As String, sFileName As String, sXMLFolder As XMLFolder)
101:    '-------------------------------------------------------------------------
102:    ' Procedure : WriteXML2File
103:    ' Purpose   : Writes sXML to sFileName
104:    '             Modified by Ken Puls 2009-05-12
105:    '             Adjusted to add ability to write to customUI container
106:    '-------------------------------------------------------------------------
107:    Dim oXMLDoc As MSXML2.DOMDocument
108:    Set oXMLDoc = New MSXML2.DOMDocument
109:
110:    'If attempting to write a customUI component, test to see if one exists
111:
112:    'Should probably test the .rels file to see if the CustomUI relationship exists...
113:    If sXMLFolder = XMLFolder_customUI Then
114:        If Not FolderExists(XMLFolder(XMLFolder_customUI)) Then
115:
116:            MkDir XMLFolder(XMLFolder_customUI)
117:            'Write the XML to the file
118:            oXMLDoc.loadXML sXml
119:            oXMLDoc.Save XMLFolder(sXMLFolder) & sFileName
120:            'CustomUI has not been created yet.  Rels file needs to be adjusted
121:            AddCustomUIToRels
122:        End If
123:    End If
124:
125:    'Write the XML to the file
126:    oXMLDoc.loadXML sXml
127:    oXMLDoc.Save XMLFolder(sXMLFolder) & sFileName
128: End Sub

     Public Sub AddCustomUIToRels()
131:    'Macro Purpose: Add the customUI relationship to the rels file
132:
133:    Dim oXMLDoc As MSXML2.DOMDocument
134:    '    Dim oXMLElement As MSXML2.IXMLDOMElement
135:    Dim oXMLElement As MSXML2.IXMLDOMNode
136:    Dim oXMLAttrib As MSXML2.IXMLDOMAttribute
137:    Dim oNamedNodeMap As MSXML2.IXMLDOMNamedNodeMap
138:    Dim oXMLRelsList As MSXML2.IXMLDOMNodeList
139:    'Create a new XML document
140:    Set oXMLDoc = New MSXML2.DOMDocument
141:    'Attach to the root element of the .rels file
142:    oXMLDoc.Load XMLFolder(XMLFolder_rels) & ".rels"
143:
144:    'Create a new relationship element in the .rels file
145:    Set oXMLElement = oXMLDoc.createNode(1, "Relationship", _
                        "http://schemas.openxmlformats.org/package/2006/relationships")
147:    Set oNamedNodeMap = oXMLElement.Attributes
148:
149:    'Create ID attribute for the element
150:    Set oXMLAttrib = oXMLDoc.createAttribute("Id")
151:    oXMLAttrib.NodeValue = "cuID"
152:    oNamedNodeMap.setNamedItem oXMLAttrib
153:
154:    'Create Type attribute for the element
155:    Set oXMLAttrib = oXMLDoc.createAttribute("Type")
156:    oXMLAttrib.NodeValue = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
157:    oNamedNodeMap.setNamedItem oXMLAttrib
158:
159:    'Create Target element for the attribute
160:    Set oXMLAttrib = oXMLDoc.createAttribute("Target")
161:    oXMLAttrib.NodeValue = "customUI/customUI.xml"
162:    oNamedNodeMap.setNamedItem oXMLAttrib
163:
164:    'Now insert the new node at the proper location
165:    Set oXMLRelsList = oXMLDoc.SelectNodes("/Relationships")
166:    oXMLRelsList.Item(0).appendChild oXMLElement
167:    'Save the .rels file
168:    oXMLDoc.Save XMLFolder(XMLFolder_rels) & ".rels"
169:
170:    Set oXMLAttrib = Nothing
171:    Set oXMLElement = Nothing
172:    Set oXMLDoc = Nothing
173: End Sub

     Private Function GetSheetIdFromSheetName(sSheetName) As String
176:    '-------------------------------------------------------------------------
177:    ' Procedure : GetSheetIdFromSheetName
178:    ' Purpose   : Finds out what the SheetId of sSheetname is
179:    '             by reading Workbook.xml
180:    '-------------------------------------------------------------------------
181:    Dim oXMLDoc As MSXML2.DOMDocument
182:    Dim oXMLNode As MSXML2.IXMLDOMNode
183:    Dim oXMLNodeList As MSXML2.IXMLDOMNodeList
184:    If mvXLFolder <> "" And Sheet2Change <> "" Then
185:        Set oXMLDoc = New MSXML2.DOMDocument
186:        oXMLDoc.Load XLFolder & "workbook.xml"
187:        Set oXMLNodeList = oXMLDoc.SelectNodes("/workbook/sheets/sheet")
188:        For Each oXMLNode In oXMLNodeList
189:            If oXMLNode.Attributes.getNamedItem("name").NodeValue = sSheetName Then
190:                GetSheetIdFromSheetName = oXMLNode.Attributes.getNamedItem("r:id").NodeValue
191:                Exit Function
192:            End If
193:        Next
194:    End If
195: End Function

     Public Function GetSheetFileNameFromId(sSheetId As String, Optional sType As String = "Id") As String
198:    '-------------------------------------------------------------------------
199:    ' Procedure : GetSheetFileNameFromId
200:    ' Purpose   : Fetches the name of the xml file belonging to the sheet with id SheetId.
201:    '-------------------------------------------------------------------------
202:    Dim oXMLDoc As MSXML2.DOMDocument
203:    Dim oXMLNode As MSXML2.IXMLDOMNode
204:    Dim oXMLNodeList As MSXML2.IXMLDOMNodeList
205:    If mvXLFolder <> "" And Sheet2Change <> "" Then
206:        Set oXMLDoc = New MSXML2.DOMDocument
207:        oXMLDoc.Load XLFolder & mvTipeFileRels
208:
209:        Set oXMLNodeList = oXMLDoc.SelectNodes("/Relationships/Relationship")
210:        For Each oXMLNode In oXMLNodeList
211:            If sType = "Id" Then
212:                If oXMLNode.Attributes.getNamedItem("Id").NodeValue = sSheetId Then
213:                    GetSheetFileNameFromId = oXMLNode.Attributes.getNamedItem("Target").NodeValue
214:                    Exit Function
215:                End If
216:            Else
217:                If oXMLNode.Attributes.getNamedItem("Type").NodeValue Like sSheetId Then
218:                    GetSheetFileNameFromId = oXMLNode.Attributes.getNamedItem("Target").NodeValue
219:                    Exit Function
220:                End If
221:            End If
222:        Next
223:    End If
224: End Function

     Public Function GetSheetNameFromId(sId As String) As String
227:    '-------------------------------------------------------------------------
228:    ' Procedure : GetSheetNameFromId
229:    ' Purpose   : Returns the sheetname belonging to a sheetId
230:    '-------------------------------------------------------------------------
231:    Dim oXMLDoc As MSXML2.DOMDocument
232:    Dim oXMLNode As MSXML2.IXMLDOMNode
233:    Dim oXMLNodeList As MSXML2.IXMLDOMNodeList
234:    If mvXLFolder <> "" Then
235:        Set oXMLDoc = New MSXML2.DOMDocument
236:        oXMLDoc.Load XLFolder & "workbook.xml"
237:        Set oXMLNodeList = oXMLDoc.SelectNodes("/workbook/sheets/sheet")
238:        For Each oXMLNode In oXMLNodeList
239:            If oXMLNode.Attributes.getNamedItem("r:id").NodeValue = "rId" & Val(sId) + 1 Then
240:                GetSheetNameFromId = oXMLNode.Attributes.getNamedItem("name").NodeValue
241:                'Got it, get out
242:                Exit Function
243:            End If
244:        Next
245:    End If
246: End Function

     Public Sub ZipAllFilesInFolder()
249:    '-------------------------------------------------------------------------
250:    ' Procedure : ZipAllFilesInFolder
251:    ' Purpose   : Zips all files in a folder (including subfolders) whilst retaining the folder structure
252:    '-------------------------------------------------------------------------
253:    Dim oShellApp As Object
254:    Dim sDate  As String
255:    Dim sDefPath As String
256:    Dim vFileNameZip As Variant
257:    Dim FSO    As Object
258:    Dim lFileCt As Long
259:    Set FSO = CreateObject("scripting.filesystemobject")
260:
261:    'To ensure a unique filename,
262:    'append date and time to the name of the current file
263:
264:    sDate = Format(Now, " dd-mmm-yy h-mm-ss")
265:    vFileNameZip = SourceFile & sDate & ".zip"
266:
267:    'Create empty Zip File
268:    NewZip vFileNameZip
269:
270:    Set oShellApp = CreateObject("Shell.Application")
271:
272:    'Count how many items are in the "old" folder
273:    lFileCt = oShellApp.Namespace(FolderName & sUNZIPPED & FileNameXML & Application.PathSeparator).Items.Count
274:
275:    'Copy the files to the compressed folder
276:    oShellApp.Namespace(vFileNameZip).CopyHere oShellApp.Namespace(FolderName & sUNZIPPED & FileNameXML & Application.PathSeparator).Items
277:
278:    'Keep script waiting until we have same # of files in the new folder
279:    On Error Resume Next
280:    Do Until oShellApp.Namespace(vFileNameZip).Items.Count = lFileCt
281:        Application.Wait (Now + TimeValue("0:00:01"))
282:    Loop
283:    DoEvents
284:
285:    'Remove original file
286:    Kill SourceFile
287:
288:    'Rename new zipped file to same name as original file (with .zip appended)
289:    Name vFileNameZip As SourceFile
290:    On Error Resume Next
291:
292:    'Now remove old folder, just in case something went haywire
293:    FSO.deletefolder FolderName & sUNZIPPED & FileNameXML, True
294:    On Error GoTo 0
295:
296:    Set oShellApp = Nothing
297: End Sub

     Public Sub UnzipFile()
300:    '-------------------------------------------------------------------------
301:    ' Procedure : UnzipFile
302:    ' Purpose   : Unzips all files in a zip file to a designated folder
303:    '-------------------------------------------------------------------------
304:    'Courtesy www.rondebruin.nl
305:    Dim FSO    As Object
306:    Dim oShellApp As Object
307:    Set FSO = CreateObject("scripting.filesystemobject")
308:
309:    'Derive the folder to unzip to from the location of the sourcefile
310:    XMLFolderRoot = FolderName
311:
312:    'A dedicated unzip folder will be created in the same folder as the sourcefile,
313:    'called ..\Unzipped Filename\
314:    If Right(XMLFolderRoot, 1) <> Application.PathSeparator Then
315:        XMLFolderRoot = XMLFolderRoot & Application.PathSeparator & sUNZIPPED & FileNameXML & " \ """
316:    Else
317:        XMLFolderRoot = XMLFolderRoot & sUNZIPPED & FileNameXML & Application.PathSeparator
318:    End If
319:    On Error Resume Next
320:    'Remove all previous existing folders
321:    FSO.deletefolder XMLFolderRoot & "*", True
322:    Kill XMLFolderRoot & "*.*"
323:    On Error GoTo 0
324:
325:    'Create normal folder
326:    If FolderExists(XMLFolderRoot) = False Then
327:        MkDir XMLFolderRoot
328:    End If
329:
330:    XLFolder = XMLFolderRoot & mvTipeFile & Application.PathSeparator
331:
332:    Set oShellApp = CreateObject("Shell.Application")
333:    'Copy the files in the newly created folder
334:    oShellApp.Namespace(XMLFolderRoot).CopyHere oShellApp.Namespace(SourceFile).Items
335:
336:    On Error Resume Next
337:    'Clean up temp folder
338:    FSO.deletefolder Environ("Temp") & "\Temporary Directory*", True
339:
340:    Set oShellApp = Nothing
341:    Set FSO = Nothing
342:    Exit Sub
343: End Sub

     Sub NewZip(sPath)
346:    If Len(Dir(sPath)) > 0 Then Kill sPath
347:    Open sPath For Output As #1
348:    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
349:    Close #1
350: End Sub

     Public Property Get CreateBackupXML() As Boolean
353:    CreateBackupXML = mbCreateBackup
354: End Property

     Public Property Let CreateBackupXML(ByVal bCreateBackup As Boolean)
357:    mbCreateBackup = bCreateBackup
358: End Property

     Private Sub Class_Initialize()
361:    'Set defaults
362:    CreateBackupXML = True
363: End Sub

     Public Property Get SourceFile() As Variant
366:    SourceFile = mvSourceFile
367: End Property

     Public Property Let SourceFile(ByVal vSourceFile As Variant)    '''
370:    mvSourceFile = vSourceFile
371:    Call WathIsFileType(vSourceFile)
372:    If CreateBackupXML Then
373:        If Len(Dir(vSourceFile & "(backup)")) > 0 Then
374:            Kill vSourceFile & "(backup)"
375:        End If
376:        FileCopy vSourceFile, vSourceFile & "(backup)"
377:    End If
378:    If Not vSourceFile Like "*.zip" Then
379:        Name vSourceFile As vSourceFile & ".zip"
380:        mvSourceFile = mvSourceFile & ".zip"
381:        AddedZip = True
382:    End If
383: End Property

     Public Property Get FolderName() As Variant
386:    FolderName = Mid(SourceFile, 1, InStrRev(SourceFile, Application.PathSeparator))
387: End Property

     Public Property Get FileNameXML() As Variant
390:    If SourceFile <> "" Then
391:        FileNameXML = Mid(SourceFile, InStrRev(SourceFile, Application.PathSeparator) + 1, Len(SourceFile))
392:    End If
393: End Property

     Public Property Get XLFolder() As Variant
396:    XLFolder = mvXLFolder
397: End Property

     Public Property Let XLFolder(ByVal vXLFolder As Variant)
400:    mvXLFolder = vXLFolder
401: End Property

     Public Property Get XMLFolder(sXMLFolder As XMLFolder) As String
404:    Select Case sXMLFolder
        Case Is = XMLFolder_root
406:            XMLFolder = mvXMLFolderRoot
407:        Case Is = XMLFolder_customUI
408:            XMLFolder = mvXMLFolderRoot & "customUI" & Application.PathSeparator
409:        Case Is = XMLFolder_docProps
410:            XMLFolder = mvXMLFolderRoot & "docProps" & Application.PathSeparator
411:        Case Is = XMLFolder_rels
412:            XMLFolder = mvXMLFolderRoot & "_rels" & Application.PathSeparator
413:        Case Is = XMLFolder_xl
414:            XMLFolder = mvXMLFolderRoot & mvTipeFile & Application.PathSeparator
415:    End Select
416: End Property

     Public Property Get BaseName(sBaseName As enBaseName) As String
419:    Select Case sBaseName
        Case Is = dimension
421:            BaseName = "dimension"
422:        Case Is = SheetViews
423:            BaseName = "SheetViews"
424:        Case Is = sheetFormatPr
425:            BaseName = "sheetFormatPr"
426:        Case Is = sheetData
427:            BaseName = "sheetData"
428:        Case Is = sheetProtection
429:            BaseName = "sheetProtection"
430:        Case Is = pageMargins
431:            BaseName = "pageMargins"
432:        Case Is = sDrawing
433:            BaseName = "Drawing"
434:        Case Is = workbookProtection
435:            BaseName = "workbookProtection"
436:    End Select
437: End Property

     Public Property Get Sheet2Change() As String
440:    Sheet2Change = msSheet2Change
441: End Property

     Public Property Let Sheet2Change(ByVal sSheet2Change As String)
444:    msSheet2Change = sSheet2Change
445:    SheetId = GetSheetIdFromSheetName(sSheet2Change)
446:    If SheetId <> "" Then
447:        SheetFileName = GetSheetFileNameFromId(SheetId)
448:    End If
449: End Property

     Public Property Get SheetId() As String
452:    SheetId = msSheetId
453: End Property

     Public Property Let SheetId(ByVal sSheetId As String)
456:    msSheetId = sSheetId
457: End Property

     Public Property Get SheetFileName() As String
460:    SheetFileName = msSheetFileName
461: End Property

     Public Property Let SheetFileName(ByVal sSheetFileName As String)
464:    msSheetFileName = sSheetFileName
465: End Property

     Private Property Get AddedZip() As Boolean
468:    AddedZip = mbAddedZip
469: End Property

     Private Property Let AddedZip(ByVal bAddedZip As Boolean)
472:    mbAddedZip = bAddedZip
473: End Property

     Private Sub Class_Terminate()
476:    Dim FSO    As Object
477:    If AddedZip Then
478:        'Remove .zip from file again
479:        Name SourceFile As Left(SourceFile, Len(SourceFile) - 4)
480:    End If
481:    'Remove zip folder
482:    On Error Resume Next
483:    Set FSO = CreateObject("scripting.filesystemobject")
484:    'Omit trailing backslash
485:    FSO.deletefolder Left(UnzipFolder, Len(UnzipFolder) - 1), True
486:    Set FSO = Nothing
487: End Sub

     Private Property Get UnzipFolder() As Variant
490:    UnzipFolder = mvUnzipFolder
491: End Property

     Private Property Let UnzipFolder(ByVal vUnzipFolder As Variant)
494:    mvUnzipFolder = vUnzipFolder
495: End Property

     Private Property Get XMLFolderRoot() As Variant
498:    XMLFolderRoot = mvXMLFolderRoot
499: End Property

     Private Property Let XMLFolderRoot(ByVal vXMLFolderRoot As Variant)
502:    mvXMLFolderRoot = vXMLFolderRoot
503: End Property

     Private Sub WathIsFileType(ByVal sPath As String)
506:    Select Case sGetExtensionName(sPath)
        Case "docm", "dotm", "dotx", "docx", "doc", "dot":
508:            mvTipeFile = "word"
509:            mvTipeFileRels = "_rels\document.xml.rels"
510:        Case "xlsm", "xlsb", "xlam", "xla", "xlsx", "xls":
511:            mvTipeFile = "xl"
512:            mvTipeFileRels = "_rels\workbook.xml.rels"
513:        Case "pptx", "pptm", "ppt", "potx", "potm", "pot":
514:            mvTipeFile = "ppt"
515:            mvTipeFileRels = "_rels\presentation.xml.rels"
516:    End Select
517: End Sub
     Public Property Get TypeFile() As String
519:    TypeFile = mvTipeFile
520: End Property
Public Property Get TipeFileRels() As String
522:    TipeFileRels = mvTipeFileRels
End Property
