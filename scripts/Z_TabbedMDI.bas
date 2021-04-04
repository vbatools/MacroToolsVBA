Attribute VB_Name = "Z_TabbedMDI"
Option Explicit

Private Const S_OK = 0
Private hResult     As Long

#If Win64 Then
Private Const DLLNAME = "vbemdi64.dll"
Private Declare PtrSafe Function Connect Lib "vbemdi64.dll" (ByVal r As Object) As Long
Private Declare PtrSafe Function Disconnect Lib "vbemdi64.dll" () As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private m_hDll      As LongPtr
#Else
Private Const DLLNAME = "vbemdi.dll"
Private Declare Function Connect Lib "vbemdi.dll" (ByVal r As Object) As Long
Private Declare Function Disconnect Lib "vbemdi.dll" () As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private m_hDll      As Long
#End If

Public Sub AddTabbed()
    Dim dllPath     As String
    dllPath = Environ("APPDATA") & "\Microsoft\AddIns\" & DLLNAME
    If Dir(dllPath) <> "" Then
        'Debug.Assert (m_hDll = 0)
        If (m_hDll = 0) Then m_hDll = LoadLibrary(dllPath)
        'Debug.Assert (m_hDll <> 0)
        If m_hDll <> 0 Then
            hResult = Connect(Application)
            'Debug.Assert (hResult = S_OK)
            'Debug.Print DLLNAME & "::Connect()"
        End If
    Else
        Debug.Print DLLNAME & " file not found", vbCritical
    End If
End Sub

Public Sub CloseTabed()
    If (m_hDll <> 0) Then
        hResult = Disconnect()
        'Debug.Assert (hResult = S_OK)
        FreeLibrary m_hDll
        m_hDll = 0
    End If
End Sub

