Attribute VB_Name = "Y_Options"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : Y_Options - модуль создание Options
'* Created    : 17-09-2020 14:35
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module

Public Sub subOptions()
    Dim sOptions    As String
    Dim moCM        As CodeModule
    Dim vbComp      As VBIDE.VBComponent
    Dim objForm     As AddOptions
    Dim sActiveVBProject As String

    On Error Resume Next
    sActiveVBProject = Application.VBE.ActiveVBProject.Filename
    On Error GoTo 0

    On Error GoTo ErrorHandler
    Set objForm = New AddOptions

    With objForm

        If sActiveVBProject <> vbNullString Then .lbNameProject.Caption = sGetFileName(sActiveVBProject)
        .Show
        If .chOptionExplicit.Value Then
            sOptions = "Option Explicit" & vbNewLine
        End If
        If .chOptionPrivate.Value Then
            sOptions = sOptions & "Option Private Module" & vbNewLine
        End If
        If .chOptionCompare.Value Then
            sOptions = sOptions & "Option Compare Text" & vbNewLine
        End If
        If .chOptionBase.Value Then
            sOptions = sOptions & "Option Base 1" & vbNewLine
        End If
        If sOptions = vbNullString Then Exit Sub
        sOptions = VBA.Left$(sOptions, VBA.Len(sOptions) - 2)
        If sOptions = vbNullString Then Exit Sub

        If .obtnModule Then
            Set moCM = Application.VBE.ActiveCodePane.CodeModule
            Call addString(moCM, sOptions)
        Else
            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
                Set moCM = vbComp.CodeModule
                Call addString(moCM, sOptions)
            Next vbComp
        End If
    End With
    Set objForm = Nothing
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Debug.Print "Ошибка! в addOptions" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
            Call WriteErrorLog("addOptions")
    End Select
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : insertOptionsExplicitAndPrivateModule - быстрое создание толко опций Explicit и Private Module
'* Created    : 23-06-2022 11:20
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub insertOptionsExplicitAndPrivateModule()
    Dim moCM        As CodeModule
    On Error Resume Next
    Set moCM = Application.VBE.ActiveCodePane.CodeModule
    On Error GoTo 0
    If Not moCM Is Nothing Then
        Call addString(moCM, "Option Explicit" & vbNewLine & "Option Private Module")
    End If
End Sub

Private Sub addString(ByRef moCM As CodeModule, ByVal sOptions As String)
    Dim i           As Long
    Dim sLines      As String
    With moCM
        i = .CountOfDeclarationLines
        If i > 0 Then
            sLines = .Lines(1, i)
            Call .DeleteLines(1, i)
        End If
        sLines = VBA.Replace(sLines, "Option Explicit", vbNullString)
        sLines = VBA.Replace(sLines, "Option Private Module", vbNullString)
        sLines = VBA.Replace(sLines, "Option Base 1", vbNullString)
        sLines = VBA.Replace(sLines, "Option Base 0", vbNullString)
        sLines = VBA.Replace(sLines, "Option Compare Text", vbNullString)
        sLines = VBA.Replace(sLines, "Option Compare Binary", vbNullString)

        If .Parent.Type <> vbext_ct_StdModule Then
            sOptions = VBA.Replace(sOptions, "Option Private Module" & vbNewLine, vbNullString)
            sOptions = VBA.Replace(sOptions, "Option Private Module", vbNullString)
        End If

        sLines = VBA.Replace(sLines, vbNewLine & vbNewLine, "||")
        sLines = VBA.Replace(sLines, vbNewLine, "||")
        If sLines = vbNullString Then
            sLines = sOptions
        Else
            sLines = sOptions & vbNewLine & VBA.Replace(sLines, "||", vbNewLine)
        End If
        Call .InsertLines(1, sLines)
    End With
End Sub
