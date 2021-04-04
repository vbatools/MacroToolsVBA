Attribute VB_Name = "V_BlackAndWiteTheme"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : V_BlackAndWiteTheme - смена темы редактора VBE светлая и темная
'* Created    : 19-02-2020 12:57
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

Private Const REG As String = "HKEY_CURRENT_USER\Software\Microsoft\VBA\"
Private Const REG_BACK_COLOR As String = "\Common\CodeBackColors"
Private Const REG_FORE_COLOR As String = "\Common\CodeForeColors"
Private Const BACK_COLOR_BLACK_THEME As String = "4 0 4 7 6 4 4 4 11 4 0 0 0 0 0 0"
Private Const FORE_COLOR_BLACK_THEME As String = "1 0 5 14 1 9 11 15 4 1 0 0 0 0 0 0"
Private Const BACK_COLOR_WHITE_THEME As String = "0 0 0 7 6 0 0 0 0 0 0 0 0 0 0 0"
Private Const FORE_COLOR_WHITE_THEME As String = "0 0 5 0 1 10 14 0 0 0 0 0 0 0 0 0"

    Public Sub ChangeColorWhiteTheme()
20:    Call ChangeColorTheme(BACK_COLOR_WHITE_THEME, FORE_COLOR_WHITE_THEME, "Включена Светлая тема, перезагрузите MS Excel")
21: End Sub

    Public Sub ChangeColorDarkTheme()
24:    Call ChangeColorTheme(BACK_COLOR_BLACK_THEME, FORE_COLOR_BLACK_THEME, "Включена Тёмная тема, перезагрузите MS Excel")
25: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ChangeColorTheme - Основная процедура смены темы
'* Created    : 19-02-2020 19:12
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                     Description
'*
'* ByVal sBackColorTheme As String : фон темы
'* ByVal sForeColorTheme As String : стиль темы
'* sMsg As String                  : строка сообщения
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub ChangeColorTheme(ByVal sBackColorTheme As String, ByVal sForeColorTheme As String, sMsg As String)
41:    Dim BackColor As String
42:    Dim ForeColor As String
43:
44:    On Error GoTo ErrorHandler
45:
46:    BackColor = REG & GetVersionVBE & REG_BACK_COLOR
47:    ForeColor = REG & GetVersionVBE & REG_FORE_COLOR
48:
49:    With CreateObject("WScript.Shell")
50:        .RegWrite BackColor, sBackColorTheme, "REG_SZ"
51:        .RegWrite ForeColor, sForeColorTheme, "REG_SZ"
52:    End With
53:    Call MsgBox(sMsg, vbInformation, "Смена темы:")
54:
55:    Exit Sub
ErrorHandler:
57:    Call MsgBox("Ошибка! в V_BlackAndWiteTheme.ChangeColorTheme" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl, vbCritical, "Ошибка:")
58:    Call WriteErrorLog("V_BlackAndWiteTheme.ChangeColorTheme")
59: End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetVersionVBE - возращает номер версии VBA используемой в системе
'* Created    : 19-02-2020 19:12
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function GetVersionVBE() As String
69:    Dim sVersion As String
70:
71:    sVersion = VBA.Replace(Application.VBE.Version, 0, vbNullString)
72:    If VBA.Right$(sVersion, 1) = "." Then
73:        sVersion = sVersion & "0"
74:    End If
75:    GetVersionVBE = sVersion
End Function
