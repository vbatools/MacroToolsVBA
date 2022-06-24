Attribute VB_Name = "C_Const"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : C_Const - глобальные константы надстройки
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Private Module
Option Explicit

Public Const NAME_VERSION As String = "Обновление 2020.11"

Public FlagVisible As Boolean

Public Const URL_MAIN As String = "https://vbatools.ru/"
Public Const URL_ADDIN As String = URL_MAIN & "macro-tools-vba-addin-excel/"

Public Const URL_BILD As String = URL_ADDIN & "code-builders/"
Public Const URL_BILD_FOFMAT As String = URL_BILD & "format-builder/"
Public Const URL_BILD_PROC As String = URL_BILD & "procedure-builder/"
Public Const URL_BILD_MSG As String = URL_BILD & "msgbox-builder/"

Public Const URL_STYLE As String = URL_ADDIN & "code-tools/"
Public Const URL_STYLE_STYLE As String = URL_STYLE & "oformlenie-stilya-koda/"
Public Const URL_STYLE_SNNIP As String = URL_STYLE & "code-library-snippets/"

Public Const URL_FILE As String = URL_ADDIN & "operaczii-s-fajlom/"
Public Const URL_FILE_PROTECT As String = URL_FILE & "udalit-paroli-vba-i-excel-remove-passvord-vba-excel/"
Public Const URL_FILE_OBFS As String = URL_FILE & "obfuskacziya-koda/"

Public Const URL_MOVE_CNTR As String = URL_ADDIN & "controls-rename-move/"

Public Const URL_CONTACT As String = URL_ADDIN & "contacts/"

Public Const URL_UPDATE As String = URL_MAIN & "vbatools/updateversion.txt"
Public Const URL_DOWNLOAD As String = URL_MAIN & "skachat/"

Public Const URL_VK As String = "https://vk.com/vbatools"
Public Const URL_FB As String = "https://www.facebook.com/groups/VBAToolsExcel/"

Public Const sMSGVBA1 As String = "Отключено: [Доверять доступ к объектной модели VBE]" & vbLf & "Для включения перейдите: Файл->Параметры->Центр управления безопасностью->Параметры макросов" & vbLf & "И перезапустите Excel"
Public Const sMSGVBA2 As String = "Нет доступа к объектной модели VBE"

Public Const SH_SNIPPETS As String = "SHSNIPPETS"
Public Const SH_STATISTICA As String = "Статистика"

Public Const TB_SNIPPETS As String = "SNIPPETS"
'Public Const TB_SNIPPETS_PRE As String = "SNIPPETS_PRE"
Public Const TB_DESCRIPTION As String = "DESCRIPTION"
Public Const TB_SERVICEWORDS As String = "SERVICEWORDS"
Public Const TB_OPTIONSIDEDENT As String = "OptionsIdedent"
Public Const TB_LOG_CODE As String = "LOG_CODE"
Public Const TB_UPDATE As String = "Update"
Public Const TB_COMMENT As String = "Comments"
Public Const TB_HOT_KEYS As String = "TB_HOT_KEYS"

Public Const MOD_ENUM_NAME As String = "SNIPPET_ENUM"

Public Const CLS_LOG_NAME As String = "LogRecorder"

Public Const NAME_ADDIN As String = "MACROTools"

Public Const TAG1 As String = NAME_ADDIN & "_VBE_TAG1"
Public Const TAG2 As String = NAME_ADDIN & "_VBE_TAG2"
Public Const TAG3 As String = NAME_ADDIN & "_VBE_TAG3"
Public Const TAG4 As String = NAME_ADDIN & "_VBE_TAG4"
Public Const TAG5 As String = NAME_ADDIN & "_VBE_TAG5"
Public Const TAG6 As String = NAME_ADDIN & "_VBE_TAG6"
Public Const TAG7 As String = NAME_ADDIN & "_VBE_TAG7"
Public Const TAG8 As String = NAME_ADDIN & "_VBE_TAG8"
Public Const TAG9 As String = NAME_ADDIN & "_VBE_TAG9"
Public Const TAG10 As String = NAME_ADDIN & "_VBE_TAG10"
Public Const TAG11 As String = NAME_ADDIN & "_VBE_TAG11"
Public Const TAG12 As String = NAME_ADDIN & "_VBE_TAG12"
Public Const TAG13 As String = NAME_ADDIN & "_VBE_TAG13"
Public Const TAG14 As String = NAME_ADDIN & "_VBE_TAG14"
Public Const TAG15 As String = NAME_ADDIN & "_VBE_TAG15"
Public Const TAG16 As String = NAME_ADDIN & "_VBE_TAG16"
Public Const TAG17 As String = NAME_ADDIN & "_VBE_TAG17"
Public Const TAG18 As String = NAME_ADDIN & "_VBE_TAG18"
Public Const TAG19 As String = NAME_ADDIN & "_VBE_TAG19"
Public Const TAG20 As String = NAME_ADDIN & "_VBE_TAG20"
Public Const TAG21 As String = NAME_ADDIN & "_VBE_TAG21"
Public Const TAG22 As String = NAME_ADDIN & "_VBE_TAG22"
Public Const TAG23 As String = NAME_ADDIN & "_VBE_TAG23"
Public Const TAG24 As String = NAME_ADDIN & "_VBE_TAG24"
Public Const TAG25 As String = NAME_ADDIN & "_VBE_TAG25"
Public Const TAG26 As String = NAME_ADDIN & "_VBE_TAG26"
Public Const TAG27 As String = NAME_ADDIN & "_VBE_TAG27"
Public Const TAG28 As String = NAME_ADDIN & "_VBE_TAG28"

Public Const TAGCOM As String = NAME_ADDIN & "_VBE_TAGCOM"

Public Const MTAG1 As String = NAME_ADDIN & "_VBE_MOVE_TAG1"
Public Const MTAG2 As String = NAME_ADDIN & "_VBE_MOVE_TAG2"
Public Const MTAG3 As String = NAME_ADDIN & "_VBE_MOVE_TAG3"
Public Const MTAG4 As String = NAME_ADDIN & "_VBE_MOVE_TAG4"
Public Const MTAG5 As String = NAME_ADDIN & "_VBE_MOVE_TAG5"
Public Const MTAGCOM As String = NAME_ADDIN & "_VBE_MOVE_TAGCOM"

Public Const RTAG1 As String = NAME_ADDIN & "_VBE_RENAME_TAG1"
Public Const RTAG2 As String = NAME_ADDIN & "_VBE_RENAME_TAG2"
Public Const RTAG3 As String = NAME_ADDIN & "_VBE_RENAME_TAG3"
Public Const RTAG4 As String = NAME_ADDIN & "_VBE_RENAME_TAG4"
Public Const RTAG5 As String = NAME_ADDIN & "_VBE_RENAME_TAG5"
Public Const RTAG6 As String = NAME_ADDIN & "_VBE_RENAME_TAG6"

Public Const POPMENU As String = "Code Window"
Public Const TOOLSMENU As String = NAME_ADDIN & " ToolBar"
Public Const MENUMOVECONTRL As String = NAME_ADDIN & " ControlBar"
Public Const RENAMEMENU As String = "MSForms Control"

Public Const CTAG1 As String = NAME_ADDIN & "_VBE_COPY_TAG1"
Public Const COPYMODULE As String = "Project Window"

Public Const SELECTEDMODULE As String = "Selected Module"
Public Const ALLVBAPROJECT As String = "All VBAProject"
Public Const mMSFORMS As String = "MSForms"

Public Const MOVECONT As String = "Control"
Public Const MOVECONTTOPLEFT As String = "Top Left"
Public Const MOVECONTBOTTOMRIGHT As String = "Bottom Right"

Public Const NAME_SH          As String = "DATA_OBF_VBATools"
Public Const NAME_SH_STR      As String = "STRING_OBF_VBATools"

Public Enum enumAnchorStyles
    enumAnchorStyleNone = 0
    enumAnchorStyleTop = 1
    enumAnchorStyleBottom = 2
    enumAnchorStyleLeft = 4
    enumAnchorStyleRight = 8
End Enum
