VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InfoFile2 
   Caption         =   "File Properties:"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9135
   OleObjectBlob   =   "InfoFile2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InfoFile2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : InfoFile2 - изменение свойств Last Author и Last Save Time
'* Created    : 20-07-2020 15:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

    Private Sub cmbMain_Change()
11:    Me.txtLastAuthor = X_InfoFile.GetOneProp(Workbooks(cmbMain.Value), "Last author")
12:    Me.txtLastAuthorOld = Me.txtLastAuthor
13:    Me.txtLastSaveTime = X_InfoFile.GetOneProp(Workbooks(cmbMain.Value), "Last save time")
14:    Me.txtLastSaveTimeOld = Me.txtLastSaveTime
15: End Sub

    Private Sub lbOK_Click()
18:    If Me.txtLastSaveTimeOld <> Me.txtLastSaveTime Or Me.txtLastAuthorOld <> Me.txtLastAuthor Then
19:        If IsDate(Me.txtLastSaveTime) Then
20:            Dim WB  As Workbook
21:            Dim sPath As String
22:            Set WB = Workbooks(cmbMain.Value)
23:            sPath = WB.FullName
24:            WB.Close savechanges:=True
25:            Call WriteXML(sPath, Me.txtLastAuthor.Text, CDate(Me.txtLastSaveTime.Text))
26:            Workbooks.Open Filename:=sPath
27:            Call MsgBox("Changes made to the file!", vbInformation, "Changes:")
28:            Unload Me
29:        Else
30:            Call MsgBox("The [ Last save time ] field does not contain a date!", vbCritical, "Error:")
31:        End If
32:    End If
33: End Sub

    Private Sub WriteXML(ByVal sFileName As String, ByVal LastAuthor As String, ByVal lastTime As Date, Optional bBackUp As Boolean = False)
36:    Dim cEditOpenXML As clsEditOpenXML
37:    Dim sXml        As String
38:    Dim oXMLDoc     As MSXML2.DOMDocument
39:
40:    Set oXMLDoc = New MSXML2.DOMDocument
41:
42:    Set cEditOpenXML = New clsEditOpenXML
43:    With cEditOpenXML
44:        .CreateBackupXML = bBackUp
45:        .SourceFile = sFileName
46:        .UnzipFile
47:        sXml = .GetXMLFromFile("core.xml", .XMLFolder(XMLFolder_docProps))
48:
49:        oXMLDoc.loadXML sXml
50:        With oXMLDoc.ChildNodes(1)
51:            .SelectSingleNode("cp:lastModifiedBy").nodeTypedValue = LastAuthor
52:            .SelectSingleNode("dcterms:modified").nodeTypedValue = VBA.Format$(lastTime, "yyyy\-mm\-ddThh\:mm\:ssZ")
53:        End With
54:        Call .WriteXML2File(oXMLDoc.XML, "core.xml", XMLFolder_docProps)
55:        .ZipAllFilesInFolder
56:    End With
57:
58:    Set cEditOpenXML = Nothing
59:    Set oXMLDoc = Nothing
60:
61: End Sub

    Private Sub UserForm_Activate()
64:    Dim vbProj      As VBIDE.VBProject
65:    If Workbooks.Count = 0 Then
66:        Unload Me
67:        Call MsgBox("No open ones" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
68:        Exit Sub
69:    End If
70:    With Me.cmbMain
71:        .Clear
72:        On Error Resume Next
73:        For Each vbProj In Application.VBE.VBProjects
74:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
75:        Next
76:        On Error GoTo 0
77:        .Value = ActiveWorkbook.Name
78:    End With
79: End Sub

    Private Sub cmbCancel_Click()
82:    Unload Me
83: End Sub
    Private Sub lbCancel_Click()
85:    Call cmbCancel_Click
86: End Sub
Private Sub UserForm_Initialize()
88:    Me.StartUpPosition = 0
89:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
90:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
