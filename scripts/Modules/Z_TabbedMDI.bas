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
23:    Dim dllPath     As String
24:    dllPath = Environ("APPDATA") & "\Microsoft\AddIns\" & DLLNAME
25:    If Dir(dllPath) <> "" Then
26:        'Debug.Assert (m_hDll = 0)
27:        If (m_hDll = 0) Then m_hDll = LoadLibrary(dllPath)
28:        'Debug.Assert (m_hDll <> 0)
29:        If m_hDll <> 0 Then
30:            hResult = Connect(Application)
31:            'Debug.Assert (hResult = S_OK)
32:            'Debug.Print DLLNAME & "::Connect()"
33:        End If
34:    Else
35:        Debug.Print DLLNAME & " file not found", vbCritical
36:    End If
37: End Sub

    Public Sub CloseTabed()
40:    If (m_hDll <> 0) Then
41:        hResult = Disconnect()
42:        'Debug.Assert (hResult = S_OK)
43:        FreeLibrary m_hDll
44:        m_hDll = 0
45:    End If
46: End Sub

