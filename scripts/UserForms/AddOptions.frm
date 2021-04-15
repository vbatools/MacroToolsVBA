VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddOptions 
   Caption         =   "Option:"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   OleObjectBlob   =   "AddOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : addOptions - создание OPTIONs в модулях проекта
'* Created    : 17-09-2020 14:06
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

    Private Sub chAll_Change()
11:    Dim bFlag       As Boolean
12:    bFlag = chAll.Value
13:    chOptionExplicit.Value = bFlag
14:    chOptionPrivate.Value = bFlag
15:    chOptionCompare.Value = bFlag
16:    chOptionBase.Value = bFlag
17: End Sub

    Private Sub lbOK_Click()
20:    Unload Me
21: End Sub

    Private Sub lbBase_Click()
24:    Dim sTxt        As String
25:    sTxt = "Used at the module level to declare the lower bound of arrays, by default." & vbNewLine & vbNewLine
26:    sTxt = sTxt & "Syntax" & vbNewLine & "Option Base { 0 | 1 }" & vbNewLine & vbNewLine
27:    sTxt = sTxt & "Since Option Base is 0 by default, the Option Base statement is never used. The operator must be in the module before all procedures." & vbNewLine
28:    sTxt = sTxt & "The Option Base statement can only be specified once in a module and must precede array declarations that include dimensions." & vbNewLine & vbNewLine
29:    sTxt = sTxt & "Note" & vbNewLine & vbNewLine
30:    sTxt = sTxt & "The To clause in the Dim, Private, Public, ReDim, and Static statements provides a more flexible way to manage the range of array indexes." & vbNewLine
31:    sTxt = sTxt & "However, if the lower bound of the indexes is not explicitly set in the To clause, you can use the Option Base statement," & vbNewLine
32:    sTxt = sTxt & "to set the default lower bound of indexes to 1. Lower bound of array index values," & vbNewLine
33:    sTxt = sTxt & "created using the Array function, always equal to zero, regardless of the Option Base statement."
34:    sTxt = sTxt & vbNewLine & vbNewLine & "The Option Base statement acts on the lower bound of array indexes only of the module in which the instruction itself is located."
35:    Debug.Print sTxt
36: End Sub
    Private Sub lbCompare_Click()
38:    Dim sTxt        As String
39:    sTxt = "Used at the module level to declare the default comparison method to be used when comparing string data." & vbNewLine & vbNewLine
40:    sTxt = sTxt & "Syntax" & vbNewLine & "Option Compare { Binary | Text | Database }" & vbNewLine & vbNewLine
41:    sTxt = sTxt & "Note" & vbNewLine & vbNewLine
42:    sTxt = sTxt & "The Option Compare statement, when used, must be in the module before any procedure." & vbNewLine
43:    sTxt = sTxt & "The Option Compare statement specifies how to compare strings (Binary, Text, or Database) for the module." & vbNewLine
44:    sTxt = sTxt & "If the module does not contain an Option Compare statement, the Binary comparison method is used by default." & vbNewLine
45:    sTxt = sTxt & "The Option Compare Binary statement sets the string comparison based on the sort order determined by the internal binary representation of the characters." & vbNewLine
46:    sTxt = sTxt & "In Microsoft Windows, the sort order is determined by the character code page." & vbNewLine
47:    sTxt = sTxt & "The following example shows a typical binary sort order result:" & vbNewLine & vbNewLine
48:    sTxt = sTxt & "A < B < E < Z < a < b < e < z < Б < Л < Ш < б < л < ш" & vbNewLine & vbNewLine
49:    sTxt = sTxt & "The Option Compare Text statement sets a case-insensitive string comparison based on the system's national setting." & vbNewLine
50:    sTxt = sTxt & "The same character as above, when sorting with the Option Compare Text statement, corresponds to the following order:" & vbNewLine & vbNewLine
51:    sTxt = sTxt & "(A=a) < (B=b) < (E=e) < (Z=z) < (Б=б) < (Л=л) < (Ш=ш)" & vbNewLine & vbNewLine
52:    sTxt = sTxt & "The Option Compare Database statement can only be used in Microsoft Access. It sets the string comparison based on the sort order.," & vbNewLine
53:    sTxt = sTxt & "defined by the national setting of the database in which the rows are compared."
54:    Debug.Print sTxt
55: End Sub
    Private Sub lbExplicit_Click()
57:    Dim sTxt        As String
58:    sTxt = "Used at the module level to force explicit declaration of all variables in this module." & vbNewLine & vbNewLine
59:    sTxt = sTxt & "Syntax" & vbNewLine & "Option Explicit" & vbNewLine & vbNewLine
60:    sTxt = sTxt & "Note" & vbNewLine & vbNewLine
61:    sTxt = sTxt & "The Option Explicit statement, when used, must be in the module before any procedure." & vbNewLine
62:    sTxt = sTxt & "When using the Option Explicit statement, you must explicitly describe all variables using the Dim, Private, Public, ReDim, or Static statements." & vbNewLine
63:    sTxt = sTxt & "When trying to use an undescribed variable name, a compile-time error occurs." & vbNewLine
64:    sTxt = sTxt & "When the Option Explicit statement is not used, all undescribed variables are of the Variant type, unless the default data type is set using the deftype statement." & vbNewLine
65:    sTxt = sTxt & "Use the Option Explicit statement to avoid incorrect input of the name of an existing variable or the risk of conflicts in the program when the scope of the variable definition is not entirely clear."
66:    Debug.Print sTxt
67: End Sub
    Private Sub lbPrivate_Click()
69:    Dim sTxt        As String
70:    sTxt = "Used at the module level to prohibit references to the module content from outside the project." & vbNewLine & vbNewLine
71:    sTxt = sTxt & "Syntax" & vbNewLine & "Option Private Module" & vbNewLine & vbNewLine
72:    sTxt = sTxt & "Note" & vbNewLine & vbNewLine
73:    sTxt = sTxt & "When a module contains an Option Private Module statement, common elements, such as variables, objects, and user-defined types, are described at the module level," & vbNewLine
74:    sTxt = sTxt & "they remain available within the project that contains this module, but are not available for other applications or projects." & vbNewLine
75:    sTxt = sTxt & "Microsoft Excel supports uploading multiple projects. In this case, the Option Private Module statement allows you to limit the mutual visibility of projects."
76:    Debug.Print sTxt
77: End Sub

    Private Sub cmbCancel_Click()
80:    Unload Me
81: End Sub
    Private Sub lbCancel_Click()
83:    Call cmbCancel_Click
84: End Sub

Private Sub UserForm_Activate()
87:    On Error GoTo ErrorHandler
88:
89:    lbExplicit.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
90:    lbPrivate.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
91:    lbCompare.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
92:    lbBase.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
93:
94:    lbModule.Caption = Application.VBE.ActiveCodePane.CodeModule.Parent.Name
95:
96:    Exit Sub
ErrorHandler:
98:    Select Case Err.Number
        Case 91:
100:            Unload Me
101:            Debug.Print "There is no active module, go to the code module!"
102:            Exit Sub
103:        Case 76:
104:            Exit Sub
105:        Case Else:
106:            Debug.Print "Error in add Options" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line " & Erl
107:            Call WriteErrorLog("addOptions")
108:    End Select
109:    Err.Clear
End Sub
