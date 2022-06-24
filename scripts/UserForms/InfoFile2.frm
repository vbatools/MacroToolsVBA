VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InfoFile2 
   Caption         =   "Свойства файла:"
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
13:    Me.txtLastAuthor = X_InfoFile.GetOneProp(Workbooks(cmbMain.Value), "Last author")
14:    Me.txtLastAuthorOld = Me.txtLastAuthor
15:    Me.txtLastSaveTime = X_InfoFile.GetOneProp(Workbooks(cmbMain.Value), "Last save time")
16:    Me.txtLastSaveTimeOld = Me.txtLastSaveTime
17: End Sub

    Private Sub lbOK_Click()
20:    If Me.txtLastSaveTimeOld <> Me.txtLastSaveTime Or Me.txtLastAuthorOld <> Me.txtLastAuthor Then
21:        If IsDate(Me.txtLastSaveTime) Then
22:            Dim wb  As Workbook
23:            Dim sPath As String
24:            Set wb = Workbooks(cmbMain.Value)
25:            sPath = wb.FullName
26:            wb.Close savechanges:=True
27:            Call WriteXML(sPath, Me.txtLastAuthor.Text, CDate(Me.txtLastSaveTime.Text))
28:            Workbooks.Open Filename:=sPath
29:            Call MsgBox("Изменения внесены в файл!", vbInformation, "Изменения:")
30:            Unload Me
31:        Else
32:            Call MsgBox("В поле [ Last save time ] указана не дата!", vbCritical, "Ошибка:")
33:        End If
34:    End If
35: End Sub

    Private Sub WriteXML(ByVal sfileName As String, ByVal LastAuthor As String, ByVal lastTime As Date, Optional bBackUp As Boolean = False)
38:    Dim cEditOpenXML As clsEditOpenXML
39:    Dim sXML        As String
40:    Dim oXMLDoc     As MSXML2.DOMDocument
41:
42:    Set oXMLDoc = New MSXML2.DOMDocument
43:
44:    Set cEditOpenXML = New clsEditOpenXML
45:    With cEditOpenXML
46:        .CreateBackupXML = bBackUp
47:        .SourceFile = sfileName
48:        .UnzipFile
49:        sXML = .GetXMLFromFile("core.xml", .XMLFolder(XMLFolder_docProps))
50:
51:        oXMLDoc.loadXML sXML
52:        With oXMLDoc.ChildNodes(1)
53:            .SelectSingleNode("cp:lastModifiedBy").nodeTypedValue = LastAuthor
54:            .SelectSingleNode("dcterms:modified").nodeTypedValue = VBA.Format$(lastTime, "yyyy\-mm\-ddThh\:mm\:ssZ")
55:        End With
56:        Call .WriteXML2File(oXMLDoc.XML, "core.xml", XMLFolder_docProps)
57:        .ZipAllFilesInFolder
58:    End With
59:
60:    Set cEditOpenXML = Nothing
61:    Set oXMLDoc = Nothing
62:
63: End Sub

    Private Sub UserForm_Activate()
66:    Dim vbProj      As VBIDE.VBProject
67:    If Workbooks.Count = 0 Then
68:        Unload Me
69:        Call MsgBox("Нет открытых " & Chr(34) & "Файлов Excel" & Chr(34) & "!", vbOKOnly + vbExclamation, "Ошибка:")
70:        Exit Sub
71:    End If
72:    With Me.cmbMain
73:        .Clear
74:        On Error Resume Next
75:        For Each vbProj In Application.VBE.VBProjects
76:            .AddItem C_PublicFunctions.sGetFileName(vbProj.Filename)
77:        Next
78:        On Error GoTo 0
79:        .Value = ActiveWorkbook.Name
80:    End With
81: End Sub

    Private Sub cmbCancel_Click()
84:    Unload Me
85: End Sub
    Private Sub lbCancel_Click()
87:    Call cmbCancel_Click
88: End Sub
Private Sub UserForm_Initialize()
90:    Me.StartUpPosition = 0
91:    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
92:    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
