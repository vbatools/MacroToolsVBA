Attribute VB_Name = "N_Obfuscation"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : N_Obfuscation - удаление форматирования кода
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

'Удаляем Option Explicit изо всех модулей
Public Sub Remove_OptionExplicit(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim i           As Long
    Dim strLine     As String

    For i = CurCodeModule.CountOfLines To 1 Step -1
        strLine = Trim(CurCodeModule.Lines(i, 1))
        If InStr(1, strLine, "Option Explicit") <> 0 Then
            CurCodeModule.DeleteLines i    'удаляем всю строку
        End If
    Next i
End Sub
'Удаляем пустые строки:
Public Sub Remove_EmptyLines(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim i&, strLine$

    For i = CurCodeModule.CountOfLines To 1 Step -1
        strLine = Trim(CurCodeModule.Lines(i, 1))
        If strLine = vbCrLf Or strLine = Chr(10) Or strLine = "" Then CurCodeModule.DeleteLines i    'удаляем пустую строку
    Next i
End Sub
'Удаляем комментарии:
Public Sub Remove_Comments(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim i           As Long
    Dim strLine     As String
    Dim pos         As Long
    Dim iCount      As Long
    With CurCodeModule
        For i = .CountOfLines To 1 Step -1
            strLine = Trim(.Lines(i, 1))
            pos = 1
try_again:
            pos = InStr(pos, strLine, Chr(39))    'позиция апострофа
            If pos > 0 Then    'есть апостроф
                If pos > 1 Then    'перед апострофом есть текст - удаляем апостроф и текст правее его
                    'проверяем не в строке ли апостроф:
                    iCount = CountChrInString(Left(strLine, pos - 1), """")    'считаем сколько кавычек слева от апострофа
                    If iCount Mod 2 = 1 Then pos = pos + 1: GoTo try_again    'этот апостроф в строке, значит он не метка комментария
                    strLine = RTrim(Left(strLine, pos - 1))
                    .ReplaceLine i, strLine
                Else    'апостроф с начала строки - удаляем всю строку
                    'однострочный коментарий
                    .DeleteLines i
                    strLine = Trim(.Lines(i, 1))
                    'многострочный коментарий
                    If Right(strLine, 2) = " _" Then
                        Do
                            .DeleteLines i
                            strLine = Trim(.Lines(i, 1))
                        Loop While Right(strLine, 2) = " _"
                        'удаление последней строки
                        .DeleteLines i
                    End If
                End If
            End If
        Next i
    End With
End Sub
'Сколько раз встречается символ char в строке str:
Private Function CountChrInString(sSTR As String, char As String) As Long
    Dim iResult     As Long
    Dim sParts()    As String

    sParts = Split(sSTR, char)
    iResult = UBound(sParts, 1): If (iResult = -1) Then iResult = 0
    CountChrInString = iResult
End Function
'Удаляем крайние табуляции и пробелы (все строки прижимаются к левому краю):
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
        strVar = Replace(strVar, " _" & vbNewLine, "")
        .DeleteLines StartLine:=1, Count:=.CountOfLines
        .InsertLines Line:=1, String:=strVar
    End With
End Sub

