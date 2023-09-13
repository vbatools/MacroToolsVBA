VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAddProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : CAddProc
'* Created    : 12-09-2023 06:09
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private m_strModuleName As String
Private m_strModuleCipherName As String
Private m_strName As String
Private m_strCipherName As String
Private m_strCodeLines() As Variant
Private m_eBehavProcExists As enumBehavProcExist

Public Property Get BehavProcExists() As enumBehavProcExist

    BehavProcExists = m_eBehavProcExists

End Property

Public Property Let BehavProcExists(ByVal eNewValue As enumBehavProcExist)

    m_eBehavProcExists = eNewValue

End Property

Public Property Get CodeLines() As Variant()

    CodeLines = m_strCodeLines

End Property

Public Property Let CodeLines(strNewValue() As Variant)

    m_strCodeLines = strNewValue

End Property

Public Property Get ModuleName() As String

    ModuleName = m_strModuleName

End Property

Public Property Let ModuleName(ByVal strNewValue As String)

    m_strModuleName = strNewValue

End Property

Public Function GetCodeLine(index As Long) As String

    GetCodeLine = m_strCodeLines(index)

End Function

Public Sub SetCodeLine(strNewValue As String, index As Long)

    m_strCodeLines(index) = strNewValue

End Sub

Public Sub AddCodeLine(strNewValue As String)

    If IsArrayEmpty(m_strCodeLines()) Then
        ReDim m_strCodeLines(0)
    Else
        ReDim Preserve m_strCodeLines(UBound(m_strCodeLines) + 1)
    End If
    
    m_strCodeLines(UBound(m_strCodeLines)) = strNewValue

End Sub

Public Sub SetCodeLinesSize(UpBound As Long)

    If IsArrayEmpty(m_strCodeLines()) Then
        ReDim m_strCodeLines(UpBound)
    Else
        ReDim Preserve m_strCodeLines(UpBound)
    End If

End Sub

Public Property Get ModuleCipherName() As String

    ModuleCipherName = m_strModuleCipherName

End Property

Public Property Let ModuleCipherName(ByVal strNewValue As String)

    m_strModuleCipherName = strNewValue

End Property

Public Property Get CipherName() As String

    CipherName = m_strCipherName

End Property

Public Property Let CipherName(ByVal strNewValue As String)

    m_strCipherName = strNewValue

End Property

Public Property Get Name() As String

    Name = m_strName

End Property

Public Property Let Name(ByVal strNewValue As String)

    m_strName = strNewValue

End Property
