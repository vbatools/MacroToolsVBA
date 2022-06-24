Attribute VB_Name = "ZA_ParserString"
Option Explicit
Option Private Module

Const SH_STRING     As String = "STRING_"
Const SH_NAME_SET   As String = SH_STRING & "SET"
Const SH_NAME_FORM  As String = SH_STRING & "FORM_CONTROLS"
Const SH_NAME_UI    As String = SH_STRING & "UI"
Const SH_NAME_UI14  As String = SH_STRING & "UI14"
Const SH_NAME_CODE  As String = SH_STRING & "CODE"

Public Sub ParserStringWB()
    Dim Form        As AddStatistic
    Dim sNameWB     As String
    Dim objWB       As Workbook

    'On Error GoTo ErrStartParser
    Set Form = New AddStatistic
    With Form
        .Caption = "Сбор строковых данных:"
        .lbOK.Caption = "СОБРАТЬ"
        .chQuestion.visible = False
        .chQuestion.Value = False
        .Show
        sNameWB = .cmbMain.Value
    End With
    If sNameWB = vbNullString Then Exit Sub
    Set objWB = Workbooks(sNameWB)
    If Not objWB.FullName Like "*" & Application.PathSeparator & "*" Then
        Call MsgBox("Выбранный файл [" & sNameWB & "] не сохранен, для продолжения сохраните файл!", vbCritical, "Ошибка:")
        Exit Sub
    ElseIf objWB.VBProject.Protection = vbext_pp_locked Then
        Call MsgBox("Проект, файла [" & sNameWB & "] защищен, снимите пароль!", vbCritical, "Проект:")
        Exit Sub
    End If

    Call ParserStr(objWB, Workbooks.Add)
    Set Form = Nothing
    Exit Sub
ErrStartParser:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call MsgBox("Ошибка в ZA_ParserString.ParserStringFromWB" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
    Call WriteErrorLog("ParserStringFromWB")
End Sub

Private Sub ParserStr(ByRef WBString As Workbook, ByRef WBNew As Workbook)
    'On Error GoTo ErrStartParser
    Dim sNameFile   As String
    sNameFile = WBString.Name

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Call N_ObfParserVBA.AddShhetInWBook(SH_NAME_SET, WBNew)
    With WBNew.Worksheets(SH_NAME_SET)
        .Cells(1, 1).Value = "Full Name WB"
        .Cells(2, 1).Value = WBString.FullName
    End With

    Call ParserStrForms(WBString, WBNew)
    Call ParserStringsInCodeAdd(WBString, WBNew)
    Call ParserStrUI(WBString, WBNew, False)

    WBNew.Activate
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call MsgBox("Строковые данные книги [" & sNameFile & "] собраны!", vbInformation, "Сбор данных:")


    Exit Sub
ErrStartParser:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call MsgBox("Ошибка в ParserStringFromWB" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
End Sub

'* * * * * ParserStrForm START * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserStrForm - сбор строк UserForm
'* Created    : 30-03-2021 11:27
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ParserStrForms(ByRef wb As Workbook, ByRef WBNew As Workbook)
    Dim objVB       As VBIDE.VBProject
    Dim objVBComp   As VBIDE.VBComponent
    Dim objCont     As MSForms.control
    Dim strCapiton  As String
    Dim strValue    As String
    Dim i           As Long
    Dim arrStr()    As String

    Debug.Print "Начало - сбора строк UserForms контролов"

    Call N_ObfParserVBA.AddShhetInWBook(SH_NAME_FORM, WBNew)

    With WBNew.Worksheets(SH_NAME_FORM)
        .Cells(1, 1).Value = "НАЗВАНИЕ МОДУЛЯ"
        .Cells(1, 2).Value = "ТИП ФОРМА/КОНТРОЛ"
        .Cells(1, 3).Value = "ИМЯ КОНТРОЛА"
        .Cells(1, 4).Value = "ЗНАЧЕНИЕ"
        .Cells(1, 5).Value = "ПОДПИСЬ"
        .Cells(1, 6).Value = "CONTROLTIPTEXT"
        .Cells(1, 7).Value = "ЗНАЧЕНИЕ"
        .Cells(1, 8).Value = "ПОДПИСЬ"
        .Cells(1, 9).Value = "CONTROLTIPTEXT"
        .Columns("A:I").EntireColumn.AutoFit
        .Cells.NumberFormat = "@"
    End With

    For Each objVBComp In wb.VBProject.VBComponents
        If objVBComp.Type = vbext_ct_MSForm Then
            i = i + 1
            ReDim Preserve arrStr(1 To 6, 1 To i)

            arrStr(1, i) = objVBComp.Name
            arrStr(2, i) = "FORMA"
            arrStr(3, i) = arrStr(1, i)
            arrStr(4, i) = vbNullString
            arrStr(5, i) = GetPropertisForm(objVBComp)
            arrStr(6, i) = vbNullString

            For Each objCont In objVBComp.Designer.Controls
                With objCont
                    If PropertyIsCapiton(objCont, True) Then
                        If .Caption <> vbNullString Then
                            strCapiton = .Caption
                        End If
                    ElseIf PropertyIsCapiton(objCont, False) Then
                        If .Value <> vbNullString Then
                            strValue = .Value
                        End If
                    End If
                    If strValue & strCapiton <> vbNullString Then
                        i = i + 1
                        ReDim Preserve arrStr(1 To 6, 1 To i)
                        arrStr(1, i) = objVBComp.Name
                        arrStr(2, i) = "CONTROL"
                        arrStr(3, i) = objCont.Name
                        arrStr(4, i) = strValue
                        arrStr(5, i) = strCapiton
                        arrStr(6, i) = objCont.ControlTipText
                    End If
                    strValue = vbNullString: strCapiton = vbNullString
                End With
            Next objCont
        End If
    Next objVBComp

    If (Not Not arrStr) <> 0 Then
        With WBNew.Worksheets(SH_NAME_FORM)
            .Cells(2, 1).Resize(UBound(arrStr, 2), UBound(arrStr, 1)).Value2 = WorksheetFunction.Transpose(arrStr)
            .Columns("A:C").EntireColumn.AutoFit
        End With
        Debug.Print "Завершен - сбор строк UserForms контролов"
    Else
        Debug.Print "Завершен - сбор строк UserForms, UserForms нет в файле"
    End If
End Sub
Private Function GetPropertisForm(ByRef objVBComp As VBIDE.VBComponent) As String
    objVBComp.Activate
    GetPropertisForm = objVBComp.Properties("Caption")
End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : PropertyIsCapiton - проверка существования свойтва Caption у контрола
'* Created    : 30-03-2021 11:28
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                         Description
'*
'* ByRef objCont As MSForms.Control :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function PropertyIsCapiton(ByRef objCont As MSForms.control, Optional bCapiton As Boolean = True) As Boolean
    On Error GoTo errEnd
    Dim s           As String
    PropertyIsCapiton = True
    If bCapiton Then
        s = objCont.Caption
    Else
        s = objCont.Text
    End If
    Exit Function
errEnd:
    On Error GoTo 0
    PropertyIsCapiton = False
End Function
'* * * * * ParserStrForm END * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserStrUI
'* Created    : 30-03-2021 15:39
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook    :
'* ByRef WBNew As Workbook :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ParserStrUI(ByRef wb As Workbook, ByRef WBNew As Workbook, Optional bRenameUI As Boolean = False)

    If VBA.UCase$(wb.Name) Like "*.XLS" Then
        Debug.Print "Сбор строк UI не возможен в файлах с расширением [*.xls]"
        Debug.Print "Пресохраните файл в новый формат"
        Exit Sub
    End If

    Dim cEditOpenXML As clsEditOpenXML
    Dim sFullNameFile As String
    Dim sFullNameXML As String
    sFullNameFile = wb.FullName
    wb.Close savechanges:=True
    Set cEditOpenXML = New clsEditOpenXML
    With cEditOpenXML
        .CreateBackupXML = False
        .SourceFile = sFullNameFile
        .UnzipFile
        sFullNameXML = .XMLFolder(XMLFolder_customUI)

        If FileHave(sFullNameXML & "customUI.xml") Then
            If Not bRenameUI Then
                Debug.Print "Начало - сбора строк UI рибон панели UI"
                Call ParserStrUIMain(WBNew, SH_STRING & "UI", sFullNameXML & "customUI.xml")
                Debug.Print "Завершен - сбор строк рибон панели UI"
            Else
                Debug.Print "Начало - переименования строк рибон панели UI"
                Call ReNameStrUI(WBNew, SH_STRING & "UI", sFullNameXML & "customUI.xml")
                Debug.Print "Завершено - переименование строк рибон панели UI"
            End If
        Else
            Debug.Print "Рибон панели customUI - нет"
        End If
        If FileHave(sFullNameXML & "customUI14.xml") Then
            If Not bRenameUI Then
                Debug.Print "Начало - сбора строк UI рибон панели UI14"
                Call ParserStrUIMain(WBNew, SH_STRING & "UI14", sFullNameXML & "customUI14.xml")
                Debug.Print "Завершен - сбор строк рибон панели UI14"
            Else
                Debug.Print "Начало - переименования строк рибон панели UI14"
                Call ReNameStrUI(WBNew, SH_STRING & "UI14", sFullNameXML & "customUI14.xml")
                Debug.Print "Завершено - переименование строк рибон панели UI14"
            End If
        Else
            Debug.Print "Рибон панели customUI14 - нет"
        End If
        .ZipAllFilesInFolder
    End With
    Set cEditOpenXML = Nothing
    Workbooks.Open sFullNameFile
End Sub

'* * * * * ParserStrUI Statr * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ParserStrUI - парсер строк рибон панели
'* Created    : 30-03-2021 11:35
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByVal sPathUI As String : путь к файлу xml рибон панели
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ParserStrUIMain(ByRef WBNew As Workbook, ByVal SHName As String, ByVal sPathUI As String)
    Dim oXMLDoc     As MSXML2.DOMDocument
    Dim oXMLRelsList As MSXML2.IXMLDOMNodeList
    Dim arrStr()    As String

    Call AddShhetInWBook(SHName, WBNew)

    With WBNew.Worksheets(SHName)
        .Cells(1, 1).Value = "TYPE"
        .Cells(1, 2).Value = "ID"
        .Cells(1, 3).Value = "LABEL"
        .Cells(1, 4).Value = "SUPERTIP"
        .Cells(1, 5).Value = "SCREENTIP"
        .Cells(1, 6).Value = "TITLE"
        .Cells(1, 7).Value = "NEW " & .Cells(1, 3).Value
        .Cells(1, 8).Value = "NEW " & .Cells(1, 4).Value
        .Cells(1, 9).Value = "NEW " & .Cells(1, 5).Value
        .Cells(1, 10).Value = "NEW " & .Cells(1, 6).Value
        .Cells(1, 11).Value = "ERRORS"
        .Cells.NumberFormat = "@"
    End With

    Set oXMLDoc = New MSXML2.DOMDocument

    With oXMLDoc
        .Load sPathUI
        Set oXMLRelsList = .SelectNodes("customUI/ribbon/tabs")
        Call LookXML(arrStr, oXMLRelsList.Item(0))
    End With

    With WBNew.Worksheets(SHName)
        .Cells(2, 1).Resize(UBound(arrStr, 2), UBound(arrStr, 1)).Value2 = WorksheetFunction.Transpose(arrStr)
        .Columns("A:C").EntireColumn.AutoFit
        .Columns("F:H").EntireColumn.AutoFit
    End With
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : LookXML - чтение xml поиск значений атрибутов "id", "label", "supertip", "screentip"
'* Created    : 30-03-2021 11:36
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub LookXML(ByRef arrStr() As String, ByRef oXMLElem As MSXML2.IXMLDOMElement)
    Dim i           As Long
    With oXMLElem
        If .ChildNodes.Length = 0 Then
            Exit Sub
        Else
            For i = 0 To .ChildNodes.Length - 1
                If Not .ChildNodes(i).Attributes Is Nothing Then
                    Call ReadAtributeValue(arrStr, .ChildNodes(i), Array("id", "label", "supertip", "screentip", "title"))
                End If
                If .ChildNodes(i).NodeType = NODE_ELEMENT Then
                    Call LookXML(arrStr, .ChildNodes(i))
                End If
            Next i
        End If
    End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ReadAtributeValue - считывание значения атрибута
'* Created    : 30-03-2021 11:37
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ReadAtributeValue(ByRef arrStr() As String, ByRef oXMLElem As MSXML2.IXMLDOMElement, ByVal arrNameAtributes As Variant)
    Dim i           As Long
    Dim iCount      As Long
    With oXMLElem.Attributes

        If (Not Not arrStr) <> 0 Then
            iCount = UBound(arrStr, 2) + 1
        Else
            iCount = 1
        End If

        ReDim Preserve arrStr(1 To 6, 1 To iCount)
        arrStr(1, iCount) = GetFullNodeName(oXMLElem, oXMLElem.BaseName)
        For i = 0 To UBound(arrNameAtributes)
            If Not .getNamedItem(arrNameAtributes(i)) Is Nothing Then
                arrStr(i + 2, iCount) = .getNamedItem(arrNameAtributes(i)).Text
            End If
        Next i
    End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetFullNodeName - получение полного дерева до узла xml
'* Created    : 30-03-2021 11:37
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                             Description
'*
'* ByRef oXMLElem As MSXML2.IXMLDOMElement :
'* stxt As String                          :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function GetFullNodeName(ByRef oXMLElem As MSXML2.IXMLDOMElement, sTxt As String) As String
    With oXMLElem

        If Not oXMLElem.ParentNode.NodeType = NODE_DOCUMENT Then
            sTxt = oXMLElem.ParentNode.BaseName & "/" & sTxt
            sTxt = GetFullNodeName(oXMLElem.ParentNode, sTxt)
        Else
            GetFullNodeName = sTxt
            Exit Function
        End If
    End With
    GetFullNodeName = sTxt
End Function

Public Sub ReNameStr()
    If MsgBox("Продолжить выполнение [Переименование строковых значений] ?" & vbNewLine & "Данную операцию нельзя отменить!", vbYesNo + vbQuestion, "Переименование строк:") = vbNo Then
        Exit Sub
    End If
    Dim WBNew       As Workbook
    Set WBNew = ActiveWorkbook
    With ActiveSheet
        If .Name = SH_NAME_SET Then
            Dim sPath As String
            sPath = .Cells(2, 1).Value
            If FileHave(sPath) Then
                Dim WBString As Workbook
                Dim sWBName As String
                sWBName = C_PublicFunctions.sGetFileName(sPath)

                If C_PublicFunctions.WorkbookIsOpen(sWBName) Then
                    Set WBString = Workbooks(sWBName)
                Else
                    Set WBString = Workbooks.Open(sPath)
                End If

                If WBString.VBProject.Protection = vbext_pp_locked Then
                    Call MsgBox("Проект защищен, снимите пароль!", vbCritical, "Проект:")
                Else
                    If HaveSheetInFile(WBNew, SH_NAME_FORM) Then
                        Call ReNameFormControls(WBString, WBNew)
                    End If
                    If HaveSheetInFile(WBNew, SH_NAME_CODE) Then
                        Call ReNameParserStringsInCodeAdd(WBString, WBNew)
                    End If
                    Call ParserStrUI(WBString, WBNew, True)
                End If
            Else
                Call MsgBox("Файл не найден на листе: [" & SH_NAME_SET & "]", vbCritical, "Ошибка:")
            End If
        Else
            Call MsgBox("Создайте или перейдите на лист: [" & SH_NAME_SET & "]", vbCritical, "Поиск настроек:")
        End If
    End With
End Sub
Private Function HaveSheetInFile(ByRef wb As Workbook, ByVal SHName As String) As Boolean
    Dim SH          As Worksheet
    On Error Resume Next
    Set SH = wb.Worksheets(SHName)
    If Err.Number = 0 Then
        HaveSheetInFile = True
    Else
        HaveSheetInFile = False
        Debug.Print "Не найден лист - [" & SHName & "]"
    End If
    Err.Clear
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ReNameFormControls - изменение свойств контроллов Value, Caption, ControlTipText
'* Created    : 30-03-2021 16:07
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WB As Workbook    :
'* ByRef WBNew As Workbook :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ReNameFormControls(ByRef wb As Workbook, ByRef WBNew As Workbook)

    Dim arrData     As Variant
    Dim lLastRow    As Long

    With WBNew.Worksheets(SH_NAME_FORM)
        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        If lLastRow < 2 Then Exit Sub
        arrData = .Range(.Cells(2, 1), .Cells(lLastRow, 10)).Value2
    End With

    Dim objVB       As VBIDE.VBProject
    Dim objVBCom    As VBIDE.VBComponent
    Dim objControl  As MSForms.control
    Dim i           As Long

    Debug.Print "Начало - переименование UserForms контролов"
    Set objVB = wb.VBProject

    For i = 1 To UBound(arrData)
        If CheckVBComponent(objVB, arrData(i, 1)) Then
            Set objVBCom = objVB.VBComponents(arrData(i, 1))
            If arrData(i, 2) = "FORMA" And arrData(i, 8) <> vbNullString Then
                Call SetPropertisForm(objVBCom, arrData(i, 8))
            Else
                If CheckControlOnForm(objVBCom, arrData(i, 3)) Then
                    Set objControl = objVBCom.Designer.Controls(arrData(i, 3))
                    With objControl
                        If arrData(i, 7) <> vbNullString Then
                            If PropertyIsCapiton(objControl, False) Then .Value = arrData(i, 7)
                        End If
                        If arrData(i, 8) <> vbNullString Then
                            If PropertyIsCapiton(objControl, True) Then .Caption = arrData(i, 8)
                        End If
                        .ControlTipText = arrData(i, 9)
                    End With
                Else
                    arrData(i, 10) = "Не найден контрол"
                End If
            End If
        Else
            arrData(i, 10) = "Не найден модуль"
        End If
    Next i

    With WBNew.Worksheets(SH_NAME_FORM)
        .Cells(1, 10).Value = "ОШИБКИ"
        .Cells(2, 1).Resize(UBound(arrData, 1), UBound(arrData, 2)).Value2 = arrData
    End With
    Debug.Print "Завершено - переименование UserForms контролов"

End Sub
Private Sub SetPropertisForm(ByRef objVBComp As VBIDE.VBComponent, ByVal sVal As String)
    objVBComp.Activate
    objVBComp.Properties("Caption") = sVal
End Sub

Private Function CheckVBComponent(ByRef objVB As VBIDE.VBProject, ByVal sNameComponent As String) As Boolean
    Dim objVBCom    As VBIDE.VBComponent
    On Error GoTo endFun
    CheckVBComponent = True
    Set objVBCom = objVB.VBComponents(sNameComponent)
    Exit Function
endFun:
    Err.Clear
    CheckVBComponent = False
End Function

Private Function CheckControlOnForm(ByRef objVBCom As VBIDE.VBComponent, ByVal sNameControl As String) As Boolean
    Dim objControl  As MSForms.control
    On Error GoTo endFun
    CheckControlOnForm = True
    Set objControl = objVBCom.Designer.Controls(sNameControl)
    Exit Function
endFun:
    Err.Clear
    CheckControlOnForm = False
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ReNameStrUI - переименование элементов риббон панелей "label", "supertip", "screentip"
'* Created    : 30-03-2021 15:42
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):             Description
'*
'* ByRef WBNew As Workbook :
'* ByVal sPathUI As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ReNameStrUI(ByRef WBNew As Workbook, ByVal SHName As String, ByVal sPathUI As String)

    Dim arrData     As Variant
    Dim lLastRow    As Long

    With WBNew.Worksheets(SHName)
        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        If lLastRow < 2 Then Exit Sub
        arrData = .Range(.Cells(2, 1), .Cells(lLastRow, 11)).Value2
    End With

    Dim oXMLDoc     As MSXML2.DOMDocument
    Dim oXMLRelsList As MSXML2.IXMLDOMNodeList
    Dim i           As Long

    Set oXMLDoc = New MSXML2.DOMDocument

    oXMLDoc.Load sPathUI
    For i = 1 To UBound(arrData)
        If arrData(i, 2) <> vbNullString Then
            Set oXMLRelsList = oXMLDoc.SelectNodes(arrData(i, 1) & "[@id='" & arrData(i, 2) & "']")
            With oXMLRelsList.Item(0)
                Call ChengeAtribute(.Attributes, "label", arrData(i, 7))
                Call ChengeAtribute(.Attributes, "supertip", arrData(i, 8))
                Call ChengeAtribute(.Attributes, "screentip", arrData(i, 9))
                Call ChengeAtribute(.Attributes, "title", arrData(i, 10))
            End With
        End If
    Next i
    Call oXMLDoc.Save(sPathUI)
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ChengeAtribute - изменение значений антрибутов xml
'* Created    : 30-03-2021 15:44
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                                     Description
'*
'* ByRef oNodeMap As MSXML2.IXMLDOMNamedNodeMap :
'* ByVal sNameAtr As String                     :
'* ByVal sVal As String                         :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ChengeAtribute(ByRef oNodeMap As MSXML2.IXMLDOMNamedNodeMap, ByVal sNameAtr As String, ByVal sVal As String)
    If sVal <> vbNullString Then
        With oNodeMap
            If Not .getNamedItem(sNameAtr) Is Nothing Then
                .getNamedItem(sNameAtr).Text = sVal
            End If
        End With
    End If
End Sub

Private Sub ParserStringsInCodeAdd(ByRef wb As Workbook, ByRef WBNew As Workbook)
    Dim oVBP        As VBIDE.VBProject
    Dim oVBCom      As VBIDE.VBComponent
    Dim iLineCode   As Long
    Dim sCode       As String
    Dim arrString   As Variant
    Dim i           As Long
    Dim k           As Long
    Dim j           As Integer
    Dim sStrCode    As String
    Dim arrParser() As String
    Dim arrPartStr  As Variant

    Debug.Print "Начало - сбора строк в модулях"

    Set oVBP = wb.VBProject
    For Each oVBCom In oVBP.VBComponents
        With oVBCom.CodeModule
            iLineCode = .CountOfLines
            If iLineCode > 0 Then
                sCode = .Lines(1, iLineCode)
                If sCode <> vbNullString And sCode Like "*" & VBA.Chr$(34) & "*" Then
                    sCode = VBA.Replace(sCode, " _" & vbNewLine, vbNullString)
                    arrString = VBA.Split(sCode, vbNewLine)
                    For i = 0 To UBound(arrString)
                        sStrCode = arrString(i)
                        sStrCode = TrimSpace(sStrCode)
                        If sStrCode <> vbNullString And VBA.Left$(sStrCode, 1) <> "'" And sStrCode Like "*" & VBA.Chr$(34) & "*" Then
                            sStrCode = DeleteCommentString(sStrCode)
                            sStrCode = VBA.Replace(sStrCode, " " & VBA.Chr$(34) & VBA.Chr$(34) & " ", vbNullString)
                            sStrCode = VBA.Replace(sStrCode, " " & VBA.Chr$(34) & VBA.Chr$(34), vbNullString)
                            arrPartStr = VBA.Split(sStrCode, VBA.Chr$(34))
                            For j = 1 To UBound(arrPartStr) Step 2
                                If arrPartStr(j) <> vbNullString Then
                                    k = k + 1
                                    ReDim Preserve arrParser(1 To 2, 1 To k)
                                    arrParser(1, k) = oVBCom.Name
                                    arrParser(2, k) = arrPartStr(j)
                                End If
                            Next j
                        End If
                    Next i
                End If
            End If
        End With
    Next oVBCom

    If (Not Not arrParser) <> 0 Then
        Call AddShhetInWBook(SH_NAME_CODE, WBNew)
        With WBNew.Worksheets(SH_NAME_CODE)
            .Cells(1, 1).Value = "NAME MODULE"
            .Cells(1, 2).Value = "STRING"
            .Cells(1, 3).Value = "NEW STRING"
            .Cells(1, 4).Value = "ERRORS"
            .Cells.NumberFormat = "@"
            .Cells(2, 1).Resize(UBound(arrParser, 2), UBound(arrParser, 1)).Value2 = WorksheetFunction.Transpose(arrParser)
            .Columns("A:D").EntireColumn.AutoFit
            Debug.Print "Завершен - сбор строк в модулях"
        End With
    Else
        Debug.Print "Завершен - сбор строк в модулях, строк нет"
    End If
End Sub

Private Sub ReNameParserStringsInCodeAdd(ByRef wb As Workbook, ByRef WBNew As Workbook)
    Dim oVBP        As VBIDE.VBProject
    Dim iLineCode   As Long
    Dim sCode       As String
    Dim sCodeNew    As String
    Dim iCount      As Long
    Dim arrData     As Variant
    Dim lLastRow    As Long
    Dim i           As Long
    Dim k           As Long
    Dim oVBCMod     As VBIDE.CodeModule


    With WBNew.Worksheets(SH_NAME_CODE)
        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        If lLastRow < 2 Then Exit Sub
        arrData = .Range(.Cells(2, 1), .Cells(lLastRow, 4)).Value2
    End With

    Set oVBP = wb.VBProject
    iCount = UBound(arrData)
    For i = 1 To iCount
        k = 1
        If i = iCount Then k = 0
        If arrData(i, 1) <> vbNullString Then
            If i = 1 Then
                Set oVBCMod = oVBP.VBComponents(arrData(i, 1)).CodeModule
                sCode = GetCodeFromModule(oVBCMod)
                sCodeNew = sCode
                If arrData(i, 3) <> vbNullString Then
                    sCodeNew = VBA.Replace(sCodeNew, VBA.Chr$(34) & arrData(i, 2) & VBA.Chr$(34), VBA.Chr$(34) & arrData(i, 3) & VBA.Chr$(34))
                End If
            End If
            'если в таблице всего одна запись
            If iCount = 1 Then
                Call SetCodeInModule(oVBCMod, sCode, sCodeNew)
            Else
                If arrData(i, 3) <> vbNullString Then
                    sCodeNew = VBA.Replace(sCodeNew, VBA.Chr$(34) & arrData(i, 2) & VBA.Chr$(34), VBA.Chr$(34) & arrData(i, 3) & VBA.Chr$(34))
                End If
                If arrData(i, 1) <> arrData(i + k, 1) Or i = iCount Then
                    Call SetCodeInModule(oVBCMod, sCode, sCodeNew)
                    Set oVBCMod = oVBP.VBComponents(arrData(i + k, 1)).CodeModule
                    sCode = GetCodeFromModule(oVBCMod)
                    sCodeNew = sCode
                End If
            End If
        End If
    Next i
End Sub

Private Function GetCodeFromModule(ByRef oVBCMod As VBIDE.CodeModule) As String
    Dim iLineCode   As Long
    Dim sCode       As String
    With oVBCMod
        iLineCode = .CountOfLines
        If iLineCode > 0 Then
            sCode = .Lines(1, iLineCode)
            If sCode <> vbNullString And sCode Like "*" & VBA.Chr$(34) & "*" Then
                GetCodeFromModule = sCode
            End If
        End If
    End With
End Function

Private Sub SetCodeInModule(ByRef oVBCMod As VBIDE.CodeModule, ByVal sCode As String, ByVal sCodeNew As String)
    If sCode <> sCodeNew Then
        Dim iLineCode As Long
        With oVBCMod
            iLineCode = .CountOfLines
            If iLineCode > 0 Then
                Call .DeleteLines(1, iLineCode)
                Call .InsertLines(1, sCodeNew)
            End If
        End With
    End If
End Sub
