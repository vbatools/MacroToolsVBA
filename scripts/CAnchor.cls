VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAnchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : CAnchor
'* Created    : 15-09-2019 15:53
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Public cnt     As MSForms.control
Public AnchorStyle As enumAnchorStyles
Public MinimumLeft As Single
Public MinimumTop As Single
Public MinimumWidth As Single
Public MinimumHeight As Single
Public OrigLeft As Single
Public OrigTop As Single
Public OrigWidth As Single
Public OrigHeight As Single
