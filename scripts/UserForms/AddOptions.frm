VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddOptions 
   Caption         =   "OPTION:"
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
24:    Dim stxt        As String
25:    stxt = "Используется на уровне модуля для объявления нижней границы массивов, по умолчанию." & vbNewLine & vbNewLine
26:    stxt = stxt & "Синтаксис" & vbNewLine & "Option Base { 0 | 1 }" & vbNewLine & vbNewLine
27:    stxt = stxt & "Поскольку Option Base по умолчанию равна 0, оператор Option Base никогда не используется. Оператор должен находиться в модуле до всех процедур." & vbNewLine
28:    stxt = stxt & "Оператор Option Base может указываться в модуле только один раз и должен предшествовать объявлениям массивов, включающим размерности." & vbNewLine & vbNewLine
29:    stxt = stxt & "Примечание" & vbNewLine & vbNewLine
30:    stxt = stxt & "Предложение To в инструкциях Dim, Private, Public, ReDim и Static предоставляет более гибкий способ управления диапазоном индексов массива." & vbNewLine
31:    stxt = stxt & "Однако если нижняя граница индексов не задается явно в предложении To, можно воспользоваться инструкцией Option Base," & vbNewLine
32:    stxt = stxt & "чтобы установить используемую по умолчанию нижнюю границу индексов, равную 1. Нижняя граница значений индексов массивов," & vbNewLine
33:    stxt = stxt & "создаваемых с помощью функции Array, всегда равняется нулю; вне зависимости от инструкции Option Base."
34:    stxt = stxt & vbNewLine & vbNewLine & "Инструкция Option Base действует на нижнюю границу индексов массивов только того модуля, в котором расположена сама эта инструкция."
35:    Debug.Print stxt
36: End Sub
    Private Sub lbCompare_Click()
38:    Dim stxt        As String
39:    stxt = "Используется на уровне модуля для объявления метода сравнения по умолчанию, который будет использоваться при сравнении строковых данных." & vbNewLine & vbNewLine
40:    stxt = stxt & "Синтаксис" & vbNewLine & "Option Compare { Binary | Text | Database }" & vbNewLine & vbNewLine
41:    stxt = stxt & "Примечание" & vbNewLine & vbNewLine
42:    stxt = stxt & "Инструкция Option Compare при ее использовании должна находиться в модуле перед любой процедурой." & vbNewLine
43:    stxt = stxt & "Инструкция Option Compare указывает способ сравнения строк (Binary, Text или Database) для модуля." & vbNewLine
44:    stxt = stxt & "Если модуль не содержит инструкцию Option Compare, по умолчанию используется способ сравнения Binary." & vbNewLine
45:    stxt = stxt & "Инструкция Option Compare Binary задает сравнение строк на основе порядка сортировки, определяемого внутренним двоичным представлением символов." & vbNewLine
46:    stxt = stxt & "В Microsoft Windows порядок сортировки определяется кодовой страницей символов." & vbNewLine
47:    stxt = stxt & "В следующем примере представлен типичный результат двоичного порядка сортировки:" & vbNewLine & vbNewLine
48:    stxt = stxt & "A < B < E < Z < a < b < e < z < Б < Л < Ш < б < л < ш" & vbNewLine & vbNewLine
49:    stxt = stxt & "Инструкция Option Compare Text задает сравнение строк без учета регистра символов на основе системной национальной настройки." & vbNewLine
50:    stxt = stxt & "Тем же символам, что и выше, при сортировке с инструкцией Option Compare Text соответствует следующий порядок: " & vbNewLine & vbNewLine
51:    stxt = stxt & "(A=a) < (B=b) < (E=e) < (Z=z) < (Б=б) < (Л=л) < (Ш=ш)" & vbNewLine & vbNewLine
52:    stxt = stxt & "Инструкция Option Compare Database может использоваться только в Microsoft Access. При этом задает сравнение строк на основе порядка сортировки," & vbNewLine
53:    stxt = stxt & "определяемого национальной настройкой базы данных, в которой производится сравнение строк. "
54:    Debug.Print stxt
55: End Sub
    Private Sub lbExplicit_Click()
57:    Dim stxt        As String
58:    stxt = "Используется на уровне модуля для принудительного явного объявления всех переменных в этом модуле." & vbNewLine & vbNewLine
59:    stxt = stxt & "Синтаксис" & vbNewLine & "Option Explicit" & vbNewLine & vbNewLine
60:    stxt = stxt & "Примечание" & vbNewLine & vbNewLine
61:    stxt = stxt & "Инструкция Option Explicit при ее использовании должна находиться в модуле до любой процедуры." & vbNewLine
62:    stxt = stxt & "При использовании инструкции Option Explicit необходимо явно описать все переменные с помощью инструкций Dim, Private, Public, ReDim или Static." & vbNewLine
63:    stxt = stxt & "При попытке использовать неописанное имя переменной возникает ошибка во время компиляции." & vbNewLine
64:    stxt = stxt & "Когда инструкция Option Explicit не используется, все неописанные переменные имеют тип Variant, если используемый по умолчанию тип данных не задается с помощью инструкции DefТип." & vbNewLine
65:    stxt = stxt & "Используйте инструкцию Option Explicit, чтобы избежать неверного ввода имени имеющейся переменной или риска конфликтов в программе, когда область определения переменной не совсем ясна."
66:    Debug.Print stxt
67: End Sub
    Private Sub lbPrivate_Click()
69:    Dim stxt        As String
70:    stxt = "Используется на уровне модуля для запрета ссылок на контент модуля извне проекта." & vbNewLine & vbNewLine
71:    stxt = stxt & "Синтаксис" & vbNewLine & "Option Private Module" & vbNewLine & vbNewLine
72:    stxt = stxt & "Примечание" & vbNewLine & vbNewLine
73:    stxt = stxt & "Когда модуль содержит инструкцию Option Private Module, общие элементы, например, переменные, объекты и определяемые пользователем типы, описанные на уровне модуля," & vbNewLine
74:    stxt = stxt & "остаются доступными внутри проекта, содержащего этот модуль, но недоступными для других приложений или проектов." & vbNewLine
75:    stxt = stxt & "Microsoft Excel поддерживает загрузку нескольких проектов. В этом случае инструкция Option Private Module позволяет ограничить взаимную видимость проектов."
76:    Debug.Print stxt
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
101:            Debug.Print "Нет активного модуля, перейдите в модуль кода!"
102:            Exit Sub
103:        Case 76:
104:            Exit Sub
105:        Case Else:
106:            Debug.Print "Ошибка! в addOptions" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
107:            Call WriteErrorLog("addOptions")
108:    End Select
109:    Err.Clear
End Sub
