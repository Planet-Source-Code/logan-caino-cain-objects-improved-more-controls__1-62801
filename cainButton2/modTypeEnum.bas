Attribute VB_Name = "modTypeEnum"
Option Explicit

Public Enum DRAWSTATE_OPTION
  
  DST_COMPLEX = &H0
  DST_TEXT = &H1
  DST_PREFIXTEXT = &H2
  DST_ICON = &H3
  DST_BITMAP = &H4
  DSS_NORMAL = &H0
  DSS_UNION = &H10
  DSS_DISABLED = &H20
  DSS_MONO = &H80
  DSS_RIGHT = &H8000
  
End Enum

Public Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Public Type RECT

  Left        As Long
  Top         As Long
  Right       As Long
  Bottom      As Long

End Type

Public Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Public Type RGB
    R As Integer
    G As Integer
    B As Integer
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum ButtonState
    
    bNormal = 0
    bHovered = 1
    bPressed = 2
    bDisabled = 3
    bUnselected = 4

End Enum

Public Enum SlideState

    Slide_Normal = 0
    Slide_TabSelected = 1
    Slide_Hover = 2
    Slide_Clicked = 3
    Slide_Disabled = 4
    Slide_Unselected = 5

End Enum

Public Enum TabSelected

    tbNormal = 0
    tbTabed = 1

End Enum

Public Type ColorSet

    csFrontColor As Long
    csBackColor As Long
    csColor1(1 To 10) As Long

End Type

Public Enum StartWindowState
    START_HIDDEN = 0
    START_NORMAL = 4
    START_MINIMIZED = 2
    START_MAXIMIZED = 3
End Enum

Public Enum eControlIDs
    
    id_Button = 0
    id_ColorPicker = 1
    id_DropDown = 2
    id_Harmonizer = 4
    id_Label = 5
    id_Listbox = 6
    id_Monthview = 7
    id_Progressbar = 8
    id_PUMenu = 9
    id_Scrollbar = 10
    id_Slider = 11
    id_Tab = 12
    id_Textbox = 13
    id_Toolbar = 14
    id_XButton = 15
    
End Enum
