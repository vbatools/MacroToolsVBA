Attribute VB_Name = "N_Obfuscation"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : N_Obfuscation - удаление форматировани€ кода
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

'”дал€ем Option Explicit изо всех модулей
Public Sub Remove_OptionExplicit(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim i           As Long
    Dim strLine     As String

    For i = CurCodeModule.CountOfLines To 1 Step -1
        strLine = Trim(CurCodeModule.Lines(i, 1))
        If InStr(1, strLine, "Option Explicit") <> 0 Then
            CurCodeModule.DeleteLines i    'удал€ем всю строку
        End If
    Next i
End Sub
'”дал€ем пустые строки:
Public Sub Remove_EmptyLines(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim i&, strLine$

    For i = CurCodeModule.CountOfLines To 1 Step -1
        strLine = Trim(CurCodeModule.Lines(i, 1))
        If strLine = vbCrLf Or strLine = Chr(10) Or strLine = "" Then CurCodeModule.DeleteLines i    'удал€ем пустую строку
    Next i
End Sub
'”дал€ем комментарии:
Public Sub Remove_Comments(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim i           As Long
    Dim strLine     As String
    Dim pos         As Long
    Dim iCount      As Long

    Rem (!) ¬спомогательные переменные
    Dim bMultiLine  As Boolean
    Dim s           As String

    With CurCodeModule
        For i = .CountOfLines To 1 Step -1
            Rem (!) обрезаем только справа, чтобы не удал€ть отступ
            strLine = RTrim(.Lines(i, 1))
            pos = 1
try_again:
            pos = InStr(pos, strLine, Chr(39))    'позици€ апострофа
            If pos > 0 Then    'есть апостроф

                Rem (!) ≈сли в строке выше есть перенос, то переходим с обработке этой строки
                If i > 1 Then
                    s = RTrim(.Lines(i - 1, 1))    'строка выше
                    If Right(s, 2) = " _" Then GoTo next_i
                End If

                Rem (!) ≈сли справа строки есть перенос, то запоминаем, что это многострочный комментарий
                If Right(RTrim(strLine), 2) = " _" Then
                    bMultiLine = True
                Else
                    bMultiLine = False
                End If

                'ѕровер€ем не в строке ли апостроф:
                'считаем сколько кавычек слева от апострофа
                iCount = CountChrInString(Left(strLine, pos - 1), """")
                'этот апостроф в строке, значит он не метка комментари€
                If iCount Mod 2 = 1 Then pos = pos + 1: GoTo try_again
                strLine = RTrim(Left(strLine, pos - 1))
                .ReplaceLine i, strLine
                '(!) запоминаем строку
                s = strLine
                Rem (!) ≈сли многострочный коментарий
                If bMultiLine Then
                    Do
                        .DeleteLines i
                        strLine = Trim(.Lines(i, 1))
                    Loop While Right(strLine, 2) = " _"
                    'последнюю строку замен€ем на ту, что запомнили
                    .ReplaceLine i, s
                End If
                If Trim(s) = "" Then .DeleteLines i
            End If
next_i:
        Next i
    End With
End Sub
'—колько раз встречаетс€ символ char в строке str:
Private Function CountChrInString(sSTR As String, char As String) As Long
    Dim iResult     As Long
    Dim sParts()    As String

    sParts = Split(sSTR, char)
    iResult = UBound(sParts, 1): If (iResult = -1) Then iResult = 0
    CountChrInString = iResult
End Function
'”дал€ем крайние табул€ции и пробелы (все строки прижимаютс€ к левому краю):
Public Sub TrimLinesTabAndSpase(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim i           As Long
    Dim strLine     As String
    Dim strLine2    As String

    For i = CurCodeModule.CountOfLines To 1 Step -1
        strLine = CurCodeModule.Lines(i, 1)
        strLine2 = Trim(strLine)
        If strLine <> strLine2 Then
            On Error Resume Next
            CurCodeModule.ReplaceLine i, strLine2
            On Error GoTo 0
        End If
    Next i
End Sub
'удаление строку переноса кода
Public Sub RemoveBreaksLineInCode(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim strVar      As String
    With CurCodeModule
        If .CountOfLines = 0 Then Exit Sub
        strVar = .Lines(1, .CountOfLines)
        strVar = Replace(strVar, " _" & vbNewLine, " ")
        .DeleteLines StartLine:=1, Count:=.CountOfLines
        .InsertLines Line:=1, String:=strVar
    End With
End Sub

