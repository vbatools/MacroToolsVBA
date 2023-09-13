VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAddGlobalVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : CAddGlobalVar
'* Created    : 12-09-2023 06:09
'* Author     : CalDymos
'* Copyright  : Byte Ranger Software
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private m_strModuleName As String
Private m_strModuleCipherName As String
Private m_strName As String
Private m_strCipherName As String
Private m_strValue As String
Private m_eDateType As enumDataType
Private m_eVisibility As enumVisibility
Private m_blnIsConstant As Boolean

Public Property Get ModuleName() As String

    ModuleName = m_strModuleName

End Property

Public Property Let ModuleName(ByVal strNewValue As String)

    m_strModuleName = strNewValue

End Property

Public Property Get Name() As String

    Name = m_strName

End Property

Public Property Let Name(ByVal strNewValue As String)

    m_strName = strNewValue

End Property

Public Property Get Value() As String

    Value = m_strValue

End Property

Public Property Let Value(ByVal strNewValue As String)

    m_strValue = strNewValue

End Property

Public Property Get DateType() As enumDataType

    DateType = m_eDateType

End Property

Public Property Let DateType(ByVal eNewValue As enumDataType)

    m_eDateType = eNewValue

End Property

Public Property Get Visibility() As enumVisibility

    Visibility = m_eVisibility

End Property

Public Property Let Visibility(ByVal eNewValue As enumVisibility)

    m_eVisibility = eNewValue

End Property

Public Property Get IsConstant() As Boolean

    IsConstant = m_blnIsConstant

End Property

Public Property Let IsConstant(ByVal blnNewValue As Boolean)

    m_blnIsConstant = blnNewValue

End Property

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
