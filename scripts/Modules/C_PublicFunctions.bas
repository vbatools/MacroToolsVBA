'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : C_PublicFunctions - ãëîáàëüíûå ôóíêöèè íàäñòðîéêè
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Modified   : Date and Time       Author              Description
'* Updated    : 07-09-2023 11:17    CalDymos

Option Explicit
Option Private Module

#If Win64 Then
Private Declare PtrSafe Function GetKeyboardState Lib "USER32" (pbKeyState As Byte) As Long
#Else
Private Declare Function GetKeyboardState Lib "USER32" (pbKeyState As Byte) As Long
#End If

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SetTextIntoClipboard - ïîìåñòèòü òåêñò â áóôåð îáìåíà
'* Created    : 08-10-2020 13:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):         Description
'*
'* ByVal txt As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub SetTextIntoClipboard(ByVal Txt As String)
       Dim MyDataObj   As New DataObject
1      MyDataObj.SetText Txt
2      MyDataObj.PutInClipboard
3     End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : SelectedLineColumnProcedure - ïîëó÷èòü íîìåðà ñòðîê è ñòîëáöîâ âûäåëåííûõ ñòðîê â ìîäóëå VBA
'* Created    : 08-10-2020 13:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function SelectedLineColumnProcedure() As String
       Dim lStartLine   As Long
       Dim lStartColumn As Long
       Dim lEndLine     As Long
       Dim lEndColumn   As Long

4      On Error GoTo ErrorHandler

5      With Application.VBE.ActiveCodePane
6          .GetSelection lStartLine, lStartColumn, lEndLine, lEndColumn
7          SelectedLineColumnProcedure = lStartLine & "|" & lStartColumn & "|" & lEndLine & "|" & lEndColumn
8      End With
9      Exit Function
ErrorHandler:
10     Select Case Err
        Case 91:
11             Debug.Print "Error!, the module for inserting code is not activated!" & vbNewLine & Err.Number & vbNewLine & Err.Description
12         Case Else:
13             Debug.Print "An error occurred in SelectedLineColumnProcedure" & vbNewLine & Err.Number & vbNewLine & Err.Description
14             Call WriteErrorLog("SelectedLineColumnProcedure")
15     End Select
16    End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DirLoadFiles - Äèàëîãîâîå îêíî âûáîðà äèðåêòîðèè
'* Created    : 08-10-2020 13:49
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sPath As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function DirLoadFiles(ByVal sPath As String) As String
17     With Application.FileDialog(msoFileDialogFolderPicker)     ' âûâîä äèàëîãîâîãî îêíà
18         .ButtonName = "Choose": .Title = "VBATools": .InitialFileName = sPath
19          If .Show <> -1 Then Exit Function    ' åñëè ïîëüçîâàòåëü îòêàçàëñÿ îò âûáîðà ïàïêè
20         DirLoadFiles = .SelectedItems(1) & Application.PathSeparator
21     End With
22    End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : Num_Not_Stable îïðåäåëåíèå ñîñòîÿíèÿ NumLock
'* Created    : 08-10-2020 13:50
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function Num_Not_Stable() As Boolean
       ' Îïðåäåëÿåò, èçìåí÷èâîå ëè ñîñòîÿíèå ó NumLock èëè íåò
       ' Âîçâðàùàåò false - ñòàáèëüíûé, true - èçìåí÷åâûé
       Dim keystat(0 To 255) As Byte
       Dim state       As String

23     GetKeyboardState keystat(0)
24     state = keystat(vbKeyNumlock)

25     If (state = 0) Then
26         Num_Not_Stable = False
27     Else
28         Num_Not_Stable = True
29     End If
30    End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : WriteErrorLog - ïðîöåäóðà âåäåíèÿ Log ôàéëà
'* Created    : 08-10-2020 13:51
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal sNameFunc As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub WriteErrorLog(ByVal sNameFunc As String)
       Dim LR          As LogRecorder
31     Set LR = New LogRecorder
32     LR.WriteErrorLog (sNameFunc)
33    End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : URLLinks - URL im Browser öffnen
'* Created    : 08-10-2020 13:51
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal url_str As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Sub URLLinks(ByVal url_str As String)
34     On Error GoTo ErrorHandler

       Dim appEX       As Object
35     Set appEX = CreateObject("Wscript.Shell")
36     appEX.Run url_str
37     Set appEX = Nothing
38     Exit Sub
ErrorHandler:
39     Select Case Err
        Case Else:
40             Call MsgBox("An error occurred in URLLinks" & vbNewLine & Err.Number & vbNewLine & Err.Description, vbOKOnly + vbCritical, "Error in URLLinks")
41             Call WriteErrorLog("URLLinks")
42     End Select
43     Set appEX = Nothing
44     Err.Clear
45    End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : FileSize - îïðåäåëèòü ðàçìåð ôàéëà
'* Created    : 08-10-2020 13:51
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):     Description
'*
'* sPath As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function FileSize(sPath As String) As Long
       'sPathFile - ñòðîêà, ïîëíûé ïóòü ê ôàéëó.
       'âîçâðàùàåò ðàçìåð ôàéëà â áàéòàõ.
       Dim sz          As Long
       Dim FSO As Object, objFile As Object
46     Set FSO = CreateObject("Scripting.FileSystemObject")
47      Set objFile = FSO.GetFile(sPath)
48      FileSize = objFile.Size
49      Set FSO = Nothing: Set objFile = Nothing
50    End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : sGetExtensionName âîçâðàùàåò ðàñøèðåíèå ïîñëåäíåãî êîìïîíåíòà â çàäàííîì ïóòè
'* Created    : 08-10-2020 13:52
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal sPathFile As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function sGetExtensionName(ByVal sPathFile As String) As String
        Dim FSO         As Object
51      Set FSO = CreateObject("Scripting.FileSystemObject")
52      sGetExtensionName = FSO.GetExtensionName(sPathFile)
53      Set FSO = Nothing
54    End Function
     Public Function sGetFileName(ByVal sPathFile As String) As String
        'sPathFile - ñòðîêà, ïóòü.
        'âîçâðàùàåò èìÿ (ñ ðàñøèðåíèåì) ïîñëåäíåãî êîìïîíåíòà â çàäàííîì ïóòè.
        Dim FSO         As Object
55      Set FSO = CreateObject("Scripting.FileSystemObject")
56      sGetFileName = FSO.GetFileName(sPathFile)
57      Set FSO = Nothing
58    End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : sGetBaseName -  âîçâðàùàåò èìÿ (áåç ðàñøèðåíèÿ) ïîñëåäíåãî êîìïîíåíòà â çàäàííîì ïóòè.
'* Created    : 04-03-2020 13:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                 Description
'*
'* ByVal sPathFile As String : ñòðîêà, ïóòü
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function sGetBaseName(ByVal sPathFile As String) As String
        Dim objFso      As Object
59      Set objFso = CreateObject("Scripting.FileSystemObject")
60      sGetBaseName = objFso.GetBaseName(sPathFile)
61      Set objFso = Nothing
62    End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : SelectedFile - äèàëîãîâîå îêíî âûáîðà ôàéëîâ èç äèðåêòîðèè
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                                 Description
'*
'* ByVal sPath As String                    :
'* Optional bMultiSelect As Boolean = True  :
'* Optional ExcelExtens As String = "*.xl*" :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function SelectedFile(ByVal sPath As String, Optional bMultiSelect As Boolean = True, Optional ExcelExtens As String = "*.xl*") As Variant
        Dim oFd         As FileDialog
        Dim s()         As String
        Dim lf          As Long
63      Set oFd = Application.FileDialog(msoFileDialogFilePicker)

64      With oFd     'èñïîëüçóåì êîðîòêîå îáðàùåíèå ê îáúåêòó
65          .AllowMultiSelect = bMultiSelect
66          .Title = "VBATools: Select an Excel file"     'çàãîëîâîê îêíà äèàëîãà
67          .Filters.Clear     'î÷èùàåì óñòàíîâëåííûå ðàíåå òèïû ôàéëîâ
68          .Filters.Add "Microsoft Excel Files", ExcelExtens, 1     'óñòàíàâëèâàåì âîçìîæíîñòü âûáîðà òîëüêî ôàéëîâ Excel
69          .InitialFileName = sPath     'íàçíà÷àåì ïàïêó îòîáðàæåíèÿ è èìÿ ôàéëà ïî óìîë÷àíèþ
70          .InitialView = msoFileDialogViewDetails     'âèä äèàëîãîâîãî îêíà(äîñòóïíî 9 âàðèàíòîâ)
71          If .Show = 0 Then
72              SelectedFile = Empty
73               Exit Function    'ïîêàçûâàåò äèàëîã
74          End If
75          ReDim Preserve s(1 To .SelectedItems.Count)
76          For lf = 1 To .SelectedItems.Count
77              s(lf) = CStr(.SelectedItems.Item(lf))     'ñ÷èòûâàåì ïîëíûé ïóòü ê ôàéëó
78          Next
79      End With
80      SelectedFile = s
81      Set oFd = Nothing
82    End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : FileHave - ïðîâåðêà ñóùåñòâîâàíèÿ ôàéëà
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                                     Description
'*
'* sPath As String                                :
'* Optional Atributes As FileAttribute = vbNormal :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function FileHave(sPath As String, Optional Atributes As FileAttribute = vbNormal) As Boolean
83      FileHave = (Dir(sPath, Atributes) <> "")
84      If sPath = vbNullString Then FileHave = False
85    End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WorkbookIsOpen - Âîçâðàùàåò ÈÑÒÈÍÀ åñëè îòêðûòà êíèãà ïîä íàçâàíèåì wname
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef wname As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function WorkbookIsOpen(ByRef wname As String) As Boolean
        Dim wb          As Workbook
86      On Error Resume Next
87      Set wb = Workbooks(wname)
88      If Err.Number = 0 Then WorkbookIsOpen = True
89    End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : IsFileOpen - åñëè ôàéë îòêðûò òî çàêðûâàåò åãî
'* Created    : 08-10-2020 13:54
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):         Description
'*
'* sFileName As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function IsFileOpen(sfileName As String) As Boolean
        Dim filenum As Integer, errnum As Integer

90      On Error Resume Next
91      filenum = FreeFile()
        ' Attempt to open the file and lock it.
92      Open sfileName For Input Lock Read As #filenum
93      Close filenum
94      errnum = Err
95      On Error GoTo 0

96      Select Case errnum
        Case 0
97              IsFileOpen = False
                ' Error number for "Permission Denied."
                ' File is already opened by another user.
98          Case 70
99              IsFileOpen = True
                ' Another error occurred.
100         Case Else
101             Error errnum
102     End Select
103   End Function
     Public Function sFolderHave(ByVal sPathFile As String) As Boolean
        'sPathFile - ñòðîêà, ïóòü.
        'Gibt True zurück, wenn der angegebene Ordner vorhanden ist, und andernfalls False.
        Dim FSO         As Object
104     Set FSO = CreateObject("Scripting.FileSystemObject")
105     sFolderHave = FSO.FolderExists(sPathFile)
106     Set FSO = Nothing
107   End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CopyFile - êîïèðîâàíèå ôàéëà
'* Created    : 04-03-2020 13:37
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                     Description
'*
'* ByVal sFileName As String    : îò êóäà
'* ByVal sNewFileName As String : êóäà
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Sub CopyFileFSO(ByVal sfileName As String, ByVal sNewFileName As String)
        Dim objFso As Object, objFile As Object

108     Set objFso = CreateObject("Scripting.FileSystemObject")
109     Set objFile = objFso.GetFile(sfileName)
110     objFile.Copy sNewFileName
111     Set objFso = Nothing
112   End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : TXTReadALLFile ñòðîêîâàÿ ïåðåìåíàÿ, âîçðàùàþùàåò òåêñò ôàéëà
'* Created    : 06-03-2020 10:07
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                         Description
'*
'* ByVal FileName As String           : ñòðîêîâàÿ ïåðåìåíàÿ, ïîëíûé ïóòü ôàéëà
'* Optional AddFile As Boolean = True : ëîãè÷åñêà ïåðåìåíàÿ, ïî óìîë÷àíèþ True, åñëè íåò ôàéëà òî ñîçäàñò åãî
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function TXTReadALLFile(ByVal sfileName As String, Optional AddFile As Boolean = True) As String

        Dim FSO         As Object
        Dim ts          As Object
113     On Error Resume Next: Err.Clear
114     Set FSO = CreateObject("scripting.filesystemobject")
115     Set ts = FSO.OpenTextFile(sfileName, 1, AddFile): TXTReadALLFile = ts.ReadAll: ts.Close
116     Set ts = Nothing: Set FSO = Nothing
117   End Function
     Public Function TXTAddIntoTXTFile(ByVal sfileName As String, ByVal Txt As String, Optional AddFile As Boolean = True) As Boolean
        'TXTAddIntoTXTFile - ëîãè÷åñêà ïåðåìåíàÿ, True - äîáàâëåíèå óäàëîñü, False - íåò
        'FileName - ñòðîêîâàÿ ïåðåìåíàÿ, ïîëíûé ïóòü ôàéëà
        'txt - òåêñò äîáàâëÿåìûé â ôàèë
        'AddFile - ëîãè÷åñêà ïåðåìåíàÿ, ïî óìîë÷àíèþ True, åñëè íåò ôàéëà òî ñîçäàñò åãî

        Dim FSO         As Object
        Dim ts          As Object
118     On Error Resume Next: Err.Clear
119     Set FSO = CreateObject("scripting.filesystemobject")
120     Set ts = FSO.OpenTextFile(sfileName, 8, AddFile): ts.Write Txt: ts.Close
121     TXTAddIntoTXTFile = Err = 0
122     Set ts = Nothing: Set FSO = Nothing
123   End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ReplceCode - function of searching in the code for names and replacing them with new ones
'* Created    : 26-03-2020 13:11
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sInCode As String : êîä ìîäóëÿ
'* sOldName As String      : ñòàðîå èìÿ
'* sNewName As String      : íîâîå èìÿ
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     Public Function ReplceCode(ByVal sInCode As String, sOldName As String, sNewName As String) As String
        Dim sCode       As String
124     sCode = sInCode
125     sCode = VBA.Replace(sCode, " " & sOldName & ".", " " & sNewName & ".", 1, -1, vbTextCompare)
126     sCode = VBA.Replace(sCode, " " & sOldName & ",", " " & sNewName & ",", 1, -1, vbTextCompare)
127     sCode = VBA.Replace(sCode, " " & sOldName & ")", " " & sNewName & ")", 1, -1, vbTextCompare)
128     sCode = VBA.Replace(sCode, "(" & sOldName & ".", "(" & sNewName & ".", 1, -1, vbTextCompare)
129     sCode = VBA.Replace(sCode, "(" & sOldName & ",", "(" & sNewName & ",", 1, -1, vbTextCompare)
130     sCode = VBA.Replace(sCode, "=" & sOldName & ".", "=" & sNewName & ".", 1, -1, vbTextCompare)
131     sCode = VBA.Replace(sCode, "=" & sOldName & vbNewLine, "=" & sNewName & vbNewLine, , , vbTextCompare)
132     sCode = VBA.Replace(sCode, "(" & sOldName & " ", "(" & sNewName & " ", 1, -1, vbTextCompare)
133     sCode = VBA.Replace(sCode, "(" & sOldName & ")", "(" & sNewName & ")", 1, -1, vbTextCompare)
134     sCode = VBA.Replace(sCode, "." & sOldName & ".", "." & sNewName & ".", 1, -1, vbTextCompare)
135     sCode = VBA.Replace(sCode, "." & sOldName & vbNewLine, "." & sNewName & vbNewLine, , , vbTextCompare)
136     sCode = VBA.Replace(sCode, " " & sOldName & "_", " " & sNewName & "_", 1, -1, vbTextCompare)
137     sCode = VBA.Replace(sCode, vbNewLine & sOldName & "_", vbNewLine & sNewName & "_", 1, -1, vbTextCompare)
138     sCode = VBA.Replace(sCode, """ & sOldName & """, """ & sNewName & """, 1, -1, vbTextCompare)
139     sCode = VBA.Replace(sCode, " " & sOldName & " ", " " & sNewName & " ", 1, -1, vbTextCompare)
140     sCode = VBA.Replace(sCode, " " & sOldName & vbNewLine, " " & sNewName & vbNewLine, 1, -1, vbTextCompare)
141     ReplceCode = sCode
142   End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : TrimSpace - óäàëåíèå âñåõ íå îäèíî÷íûõ ïðîáåëîâ
'* Created    : 08-10-2020 13:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sTxt As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function TrimSpace(ByVal sTxt As String) As String
        Dim sTemp As String
        Const LENGHT_CELLS As Long = 32767
143     sTemp = sTxt
144     If VBA.Len(sTemp) <= LENGHT_CELLS Then
145         sTemp = Application.WorksheetFunction.Trim(sTemp)
146     Else
            Dim i As Long
147         For i = 1 To VBA.Len(sTxt) Step LENGHT_CELLS
148             sTemp = sTemp & Application.WorksheetFunction.Trim(VBA.Mid$(sTxt, i, LENGHT_CELLS))
149         Next i
150     End If
151     TrimSpace = sTemp
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WorksheetExist
'* Created    : 07-09-2023 07:01
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):         Description
'*
'* strName As String :
'* wb As Workbook    :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function WorksheetExist(strName As String, wb As Workbook) As Boolean
152   On Error Resume Next
153   WorksheetExist = wb.Worksheets(strName).Index > 0
154   On Error GoTo 0
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : DelWorksheet
'* Created    : 07-09-2023 07:01
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):         Description
'*
'* strName As String :
'* wb As Workbook    :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function DelWorksheet(strName As String, wb As Workbook) As Boolean
155       Application.DisplayAlerts = False
156       If WorksheetExist(strName, wb) Then
157           wb.Worksheets(strName).Delete
158           DelWorksheet = Not WorksheetExist(strName, wb)
159       End If
160       Application.DisplayAlerts = True
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : TrimL - Removes all specified characters to the left of the string
'* Created    : 07-09-2023 11:13
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):                         Description
'*
'* ByVal str As String                 :
'* Optional ByVal Char As String = " " :
'* Optional ByVal lCount As Long = 0   :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function TrimL(ByVal str As String, Optional ByVal Char As String = " ", Optional ByVal lCount As Long = 0) As String
       
161     If Char = " " Then
162       TrimL = LTrim$(str)
163     Else
          Dim lLen As Long
       
164       lLen = Len(Char)
165       If lCount > 0 Then
            Dim i As Long
166         While Len(str) > 0 And Left$(str, lLen) = Char And i < lCount
167           str = Mid$(str, lLen + 1)
168           i = i + 1
169         Wend
170       Else
171         While Len(str) > 0 And Left$(str, lLen) = Char
172           str = Mid$(str, lLen + 1)
173         Wend
174       End If
175     End If
       
176     TrimL = str
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : TrimR - Removes all specified characters to the right of the string
'* Created    : 07-09-2023 11:14
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):                         Description
'*
'* ByVal str As String                 :
'* Optional ByVal Char As String = " " :
'* Optional ByVal lCount As Long = 0   :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function TrimR(ByVal str As String, Optional ByVal Char As String = " ", Optional ByVal lCount As Long = 0) As String
       
177     If Char = " " Then
178       TrimR = RTrim$(str)
179     Else
          Dim lLen As Long
       
180       lLen = Len(Char)
181       If lCount > 0 Then
            Dim i As Long
182         While Len(str) > 0 And Right$(str, lLen) = Char And i < lCount
183           str = Left$(str, Len(str) - lLen)
184           i = i + 1
185         Wend
186       Else
187         While Len(str) > 0 And Right$(str, lLen) = Char
188           str = Left$(str, Len(str) - lLen)
189         Wend
190       End If
191     End If
       
192     TrimR = str
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : TrimA - Removes all specified characters to the left and right of the string
'* Created    : 07-09-2023 11:14
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* Argument(s):                         Description
'*
'* ByVal str As String                 :
'* Optional ByVal Char As String = " " :
'* Optional ByVal lCount As Long = 0   :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function TrimA(ByVal str As String, Optional ByVal Char As String = " ", Optional ByVal lCount As Long = 0) As String
       
193     TrimA = TrimR(TrimL(str, Char, lCount), Char, lCount)
End Function
