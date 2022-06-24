Attribute VB_Name = "B_CreateMenus"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : B_CreateMenus - создание меню в VBE
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Private Module
Option Explicit
Public ToolContextEventHandlers As New Collection

#If Win64 Then
Private Declare PtrSafe Function GetKeyboardLayoutName Lib "USER32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare PtrSafe Function LoadKeyboardLayout Lib "USER32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
#Else
Private Declare Function GetKeyboardLayoutName Lib "USER32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare Function LoadKeyboardLayout Lib "USER32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
#End If

Private Const LANG_RUSSIAN = 419
Private Const LANG_ENGLISH = 409

Private Sub Auto_Open()
    If VBAIsTrusted And ThisWorkbook.Name = C_Const.NAME_ADDIN & ".xlam" Then    '
        Call AddContextMenus
    End If
End Sub
Public Sub AddContextMenus()

    Call AddNewCommandBarMenu(C_Const.MENUMOVECONTRL)
    Call AddButtom(C_Const.MTAG5, 984, "Справка по инструменту", "HelpMoveControl", C_Const.MENUMOVECONTRL, False, True)
    Call AddButtom(C_Const.MTAG4, 38, "", "MoveControl", C_Const.MENUMOVECONTRL)
    Call AddButtom(C_Const.MTAG3, 40, "", "MoveControl", C_Const.MENUMOVECONTRL, False, True)
    Call AddButtom(C_Const.MTAG2, 39, "", "MoveControl", C_Const.MENUMOVECONTRL)
    Call AddButtom(C_Const.MTAG1, 41, "", "MoveControl", C_Const.MENUMOVECONTRL)
    Call AddComboBoxMove(C_Const.MENUMOVECONTRL)

    Call AddNewCommandBarMenu(C_Const.TOOLSMENU)
    Call AddButtom(C_Const.TAG15, 984, "Справка по надстройке", "HelpMainAddin", C_Const.TOOLSMENU, False, True)
    Call AddButtom(C_Const.TAG12, 0, "FormatBuilder", "subFormatBuilder", C_Const.TOOLSMENU, True, True)
    Call AddButtom(C_Const.TAG11, 0, "MsgBoxBuilder", "subMsgBoxGenerator", C_Const.TOOLSMENU, True, True)
    Call AddButtom(C_Const.TAG13, 0, "ProcedureBuilder", "subProcedureBuilder", C_Const.TOOLSMENU, True, True)
    Call AddButtom(C_Const.TAG23, 107, "Option's Explicit and Private Module", "insertOptionsExplicitAndPrivateModule", C_Const.TOOLSMENU, False, False)
    Call AddButtom(C_Const.TAG28, 0, "Option's", "subOptionsMenu", C_Const.TOOLSMENU, True, True)

    Call AddButtom(C_Const.TAG24, 2045, "Copy", "SetInCipBoard", C_Const.TOOLSMENU, True, False)
    Call AddButtom(C_Const.TAG25, 22, "Paste", "GetFromCipBoard", C_Const.TOOLSMENU, True, True)

    Call AddButtom(C_Const.TAG20, 1714, "Поиск не используемых переменых ", "SerchVariableUnUsedInSelectedWorkBook", C_Const.TOOLSMENU, False, False)
    Call AddButtom(C_Const.TAG19, 3838, "Закрыть все окна VBE ", "CloseAllWindowsVBE", C_Const.TOOLSMENU, False, False)
    Call AddButtom(C_Const.TAG14, 22, "Вставить класс LogRecorder ", "AddLogRecorderClass", C_Const.TOOLSMENU, False, True)

    Call AddButtom(C_Const.TAG19, 8, "Список TODO ", "ShowTODOList", C_Const.TOOLSMENU, False, False)
    Call AddButtom(C_Const.TAG18, 1972, "Создать TODO ", "sysAddTODOTop", C_Const.TOOLSMENU, False, False)
    Call AddButtom(C_Const.TAG17, 456, "Создать строку комментария обновления ", "sysAddModifiedTop", C_Const.TOOLSMENU, False, False)
    Call AddButtom(C_Const.TAG16, 1546, "Создать комментарий ", "sysAddHeaderTop", C_Const.TOOLSMENU, False, True)

    Call AddButtom(C_Const.TAG10, 3917, "Удалить форматирование Кода", "CutTab", C_Const.TOOLSMENU)
    Call AddButtom(C_Const.TAG9, 3919, "Форматировать Код", "ReBild", C_Const.TOOLSMENU, False, True)
    Call AddButtom(C_Const.TAG8, 12, "Удалить нумерацию строк", "RemoveLineNumbers_", C_Const.TOOLSMENU)
    Call AddButtom(C_Const.TAG7, 11, "Создать нумерацию строк", "AddLineNumbers_", C_Const.TOOLSMENU)
    Call AddComboBox(C_Const.TOOLSMENU)
    Call AddButtom(C_Const.TAG27, 210, "Сортировка процедур по алфавиту", "AlphabetizeProcedure", C_Const.TOOLSMENU, False, True)
    Call AddButtom(C_Const.TAG6, 47, "Очистить окно [Immediate]", "ClearImmediateWindow", C_Const.TOOLSMENU, False, True)
    Call AddButtom(C_Const.TAG5, 2059, "Создать легенду", "AddLegend", C_Const.TOOLSMENU)
    Call AddButtom(C_Const.TAG4, 21, "Удалить модуль", "DeleteSnippetEnumModule", C_Const.TOOLSMENU)
    Call AddButtom(C_Const.TAG3, 1753, "Вставить модуль", "AddSnippetEnumModule", C_Const.TOOLSMENU)
    Call AddButtom(C_Const.TAG2, 22, "Вставить код", "InsertCode", C_Const.TOOLSMENU, False, False)

    Call AddButtom(C_Const.TAG26, 9634, "Поменять местами относ [=]", "SwapEgual", C_Const.POPMENU, True, False)
    Call AddButtom(C_Const.TAG21, 0, "UPPER Case", "toUpperCase", C_Const.POPMENU, True, False)
    Call AddButtom(C_Const.TAG22, 0, "lower Case", "toLowerCase", C_Const.POPMENU, True, False)
    Call AddButtom(C_Const.TAG1, 22, "Вставить код", "InsertCode", C_Const.POPMENU, True, False)

    Call AddButtom(C_Const.RTAG1, 162, "ReName Control", "RenameControl", C_Const.RENAMEMENU, True)
    Call AddButtom(C_Const.RTAG2, 22, "Paste Style", "PasteStyleControl", C_Const.RENAMEMENU, True)
    Call AddButtom(C_Const.RTAG3, 1076, "Copy Style", "CopyStyleControl", C_Const.RENAMEMENU, True)
    Call AddButtom(C_Const.RTAG4, 704, "Paste Icon", "AddIcon", C_Const.RENAMEMENU, True, True)
    Call AddButtom(C_Const.RTAG5, 0, "UPPER Case", "UperTextInControl", C_Const.RENAMEMENU, True, False)
    Call AddButtom(C_Const.RTAG6, 0, "lower Case", "LowerTextInControl", C_Const.RENAMEMENU, True, False)
    
    Call AddButtom(C_Const.CTAG1, 2045, "Copy Module", "CopyModyleVBE", C_Const.COPYMODULE, True, False)
    
    Call AddButtom(C_Const.RTAG2, 22, "Paste Style", "PasteStyleControl", C_Const.mMSFORMS, True)
    Call AddButtom(C_Const.RTAG3, 1076, "Copy Style", "CopyStyleControl", C_Const.mMSFORMS, True)
    Call AddButtom(C_Const.RTAG5, 0, "UPPER Case", "UperTextInForm", C_Const.mMSFORMS, True, False)
    Call AddButtom(C_Const.RTAG6, 0, "lower Case", "LowerTextInForm", C_Const.mMSFORMS, True, False)
End Sub
Private Sub AddNewCommandBarMenu(ByVal sNameCommandBar As String)
    Dim myCommandBar As CommandBar
    On Error GoTo AddNewCommandBar
    Set myCommandBar = Application.VBE.CommandBars(sNameCommandBar)
    If myCommandBar Is Nothing Then
AddNewCommandBar:
        Set myCommandBar = Application.VBE.CommandBars.Add(Name:=sNameCommandBar, Position:=msoBarTop)
        myCommandBar.visible = True
        myCommandBar.RowIndex = 3
    End If
End Sub
Private Sub AddButtom( _
        ByVal sTag As String, _
        ByVal Face As Long, _
        ByVal Capitan As String, _
        ByVal sOnAction As String, _
        ByVal sMenu As String, _
        Optional ByRef VisibleCapiton As Boolean = False, _
        Optional ByVal Begin_Group As Boolean = False, _
        Optional ByVal ShortcutText As String = vbNullString, _
        Optional ByVal Before As Byte = 1)
    Dim btn         As CommandBarButton
    Dim evtContextMenu As VBECommandHandler
    Set btn = Application.VBE.CommandBars(sMenu).Controls.Add(Type:=msoControlButton, Before:=Before)
    With btn
        .FaceId = Face
        If VisibleCapiton Then .Caption = Capitan
        .TooltipText = Capitan
        .Tag = sTag
        .OnAction = "'" & ThisWorkbook.Name & "'!" & sOnAction
        .Style = msoButtonIconAndCaption
        .BeginGroup = Begin_Group
        .ShortcutText = ShortcutText
    End With
    Set evtContextMenu = New VBECommandHandler
    Set evtContextMenu.EvtHandler = btn
    ToolContextEventHandlers.Add evtContextMenu
End Sub
Private Sub AddComboBox(ByVal sMenu As String)
    Dim combox      As CommandBarComboBox
    Set combox = Application.VBE.CommandBars(sMenu).Controls.Add(Type:=msoControlComboBox, Before:=1)
    With combox
        .Tag = C_Const.TAGCOM
        .AddItem C_Const.SELECTEDMODULE
        .AddItem C_Const.ALLVBAPROJECT
        .Text = C_Const.SELECTEDMODULE
    End With
End Sub
Private Sub AddComboBoxMove(ByVal sMenu As String)
    Dim combox      As CommandBarComboBox
    Set combox = Application.VBE.CommandBars(sMenu).Controls.Add(Type:=msoControlComboBox, Before:=1)
    With combox
        .Tag = C_Const.MTAGCOM
        .AddItem C_Const.MOVECONT
        .AddItem C_Const.MOVECONTTOPLEFT
        .AddItem C_Const.MOVECONTBOTTOMRIGHT
        .Text = C_Const.MOVECONT
    End With
End Sub
Private Sub Auto_Close()
    If VBAIsTrusted Then
        Call DeleteContextMenus
    End If
End Sub
Public Sub DeleteContextMenus()
    Dim myCommandBar As CommandBar
    On Error GoTo ErrorHandler

    Call DeleteButton(C_Const.TAG1, C_Const.POPMENU)
    Call DeleteButton(C_Const.TAG26, C_Const.POPMENU)
    Call DeleteButton(C_Const.TAG21, C_Const.POPMENU)
    Call DeleteButton(C_Const.TAG22, C_Const.POPMENU)
    
    Call DeleteButton(C_Const.CTAG1, C_Const.COPYMODULE)
    
    Call DeleteButton(C_Const.RTAG1, C_Const.RENAMEMENU)
    Call DeleteButton(C_Const.RTAG2, C_Const.RENAMEMENU)
    Call DeleteButton(C_Const.RTAG3, C_Const.RENAMEMENU)
    Call DeleteButton(C_Const.RTAG4, C_Const.RENAMEMENU)
    Call DeleteButton(C_Const.RTAG5, C_Const.RENAMEMENU)
    Call DeleteButton(C_Const.RTAG6, C_Const.RENAMEMENU)
    
    Call DeleteButton(C_Const.RTAG2, C_Const.mMSFORMS)
    Call DeleteButton(C_Const.RTAG3, C_Const.mMSFORMS)
    Call DeleteButton(C_Const.RTAG5, C_Const.mMSFORMS)
    Call DeleteButton(C_Const.RTAG6, C_Const.mMSFORMS)

    Set myCommandBar = Application.VBE.CommandBars(C_Const.TOOLSMENU)
    If Not myCommandBar Is Nothing Then
        myCommandBar.Delete
    End If

    Set myCommandBar = Application.VBE.CommandBars(C_Const.MENUMOVECONTRL)
    If Not myCommandBar Is Nothing Then
        myCommandBar.Delete
    End If

    'очистка колекции
    Do Until ToolContextEventHandlers.Count = 0
        ToolContextEventHandlers.Remove 1
    Loop

    Exit Sub
ErrorHandler:

    Select Case Err
        Case 5:
            Err.Clear
        Case Else:
            Debug.Print "Ошибка! в DeleteContextMenus" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
            Call WriteErrorLog("DeleteContextMenus")
    End Select
    Err.Clear
End Sub
Private Sub DeleteButton(ByRef sTag As String, ByVal sMenu As String)
    Dim Cbar        As CommandBar
    Dim Ctrl        As CommandBarControl
    On Error GoTo ErrorHandler
    Set Cbar = Application.VBE.CommandBars(sMenu)
    For Each Ctrl In Cbar.Controls
        If Ctrl.Tag = sTag Then
            Ctrl.Delete
            'Exit Sub
        End If
    Next Ctrl
    Exit Sub
ErrorHandler:
    Debug.Print "Ошибка! в DeleteButton" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
    Call WriteErrorLog("DeleteButton")
    Err.Clear
    Resume Next
End Sub
Public Function VBAIsTrusted() As Boolean
    On Error GoTo ErrorHandler
    Dim sTxt As String
    sTxt = Application.VBE.Version
    VBAIsTrusted = True
    Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 1004:
            'If ThisWorkbook.Name = C_Const.NAME_ADDIN & ".xlam" Then
            Call MsgBox("Предупреждение! " & C_Const.NAME_ADDIN & vbLf & vbNewLine & _
                    "Отключено: [Доверять доступ к объектной модели VBE]" & vbLf & _
                    "Для включения перейдите: Файл->Параметры->Центр управления безопасностью->Параметры макросов" & _
                    vbLf & vbNewLine & "И перезапустите Excel", vbCritical, "Предупреждение:")
        Case Else:
            Debug.Print "Ошибка! в VBAIsTrusted" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
            Call WriteErrorLog("VBAIsTrusted")
    End Select
    Err.Clear
    VBAIsTrusted = False
End Function
Public Function WhatIsTextInComboBoxHave() As String
    Dim myCommandBar As CommandBar
    Dim cntrl       As CommandBarControl

    Set myCommandBar = Application.VBE.CommandBars(C_Const.TOOLSMENU)
    For Each cntrl In myCommandBar.Controls
        If cntrl.Tag = C_Const.TAGCOM Then
            WhatIsTextInComboBoxHave = cntrl.Text
            Exit Function
        End If
    Next cntrl
End Function
Public Sub ClearImmediateWindow()
    Dim KeybLayoutName As String * 8
    KeybLayoutName = String(8, "0")
    GetKeyboardLayoutName KeybLayoutName
    KeybLayoutName = Val(KeybLayoutName)

    Select Case Val(KeybLayoutName)
        Case LANG_ENGLISH
            Call ClearImmediateWindowFunction
            Call ClearImmediateWindowFunction
        Case LANG_RUSSIAN
            ' Переключение на английскую раскладку
            Call LoadKeyboardLayout("00000409", &H1)
            Call ClearImmediateWindowFunction
            Call LoadKeyboardLayout("00000419", &H1)
        Case Else
            Call MsgBox("Переключите раскладку клавиатуры на Английскую!", vbInformation, "Переключение раскладки клавиатуры")
    End Select
End Sub
Private Sub ClearImmediateWindowFunction()
    Call SendKeys("^g")
    Call SendKeys("^a")
    Call SendKeys("{DEL}")
End Sub
Public Sub RefreshMenu()
    Call B_CreateMenus.DeleteContextMenus
    Call B_CreateMenus.AddContextMenus
    Call MsgBox("Перезагрузка надстройки " & C_Const.NAME_ADDIN & " прошла!", vbInformation, "Перезагрузка надстройки " & C_Const.NAME_ADDIN & ":")
End Sub
Private Sub subMsgBoxGenerator()
    MsgBoxGenerator.Show
End Sub
Private Sub subFormatBuilder()
    BilderFormat.Show
End Sub
Private Sub subProcedureBuilder()
    BilderProcedure.Show
End Sub
Private Sub subOptionsMenu()
    Call Y_Options.subOptions
End Sub
