VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSLPEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Enums

Public Enum slpBorderStyles
    None = 0
    FixedSingle = 1
End Enum

Public Enum slpBackStyle
    Opaque = 1
    Transparent = 0
End Enum

Public Enum slpScrollText_DropShadowDirection
    slpDropShadowDefault = 0
    slpDropShadowNorthWest = 1
    slpDropShadowNorthEast = 2
    slpDropShadowSouthWest = 3
    slpDropShadowSouthEast = 4
End Enum

Public Enum slpScrollText_ScrollDirection
    slpVertical = 1
    slpHorizontal = 2
End Enum

Public Enum slpScrollText_TextJustify
    slpJustifyDefault = 0
    slpJustifyCenter = 1
    slpJustifyRight = 2
    slpJustifyLeft = 3
End Enum
    
