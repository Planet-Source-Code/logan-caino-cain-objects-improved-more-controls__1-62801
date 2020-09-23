VERSION 5.00
Begin VB.UserControl cainXButton 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "cainXButton.ctx":0000
   Begin VB.Timer MousePos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "cainXButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Standard-Eigenschaftswerte:
'Const m_def_ImageIndex = 0
'Const m_def_TabSelectColor = &H96E7&
Const m_def_BackColor = &HFFFFFF
Const m_def_ForeColor = &HC00000
'Eigenschaftsvariablen:
'Dim m_ImageList As Object
'Dim m_ImageIndex As Integer
'Dim m_TabSelectColor As OLE_COLOR
Dim m_Caption As String
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
'Ereignisdeklarationen:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Dim Mouse As POINTAPI
Dim MouseOverMe As POINTAPI
Dim MouseClick As Boolean

Dim MyState As ButtonState
Dim MyTabState As TabSelected

Public Function ID() As Integer
    ID = eControlIDs.id_XButton
End Function

Private Sub DrawButton()

    UserControl.Cls

    If UserControl.Enabled = True Then
        If MyState = bDisabled Then MyState = bNormal
    Else
        MyState = bDisabled
    End If
           
    If MyState = bNormal And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        DrawFace GetColorSetNormal(m_BackColor, m_ForeColor)
        
    ElseIf MyState = bHovered And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        DrawFace GetColorSetHovered(m_BackColor, m_ForeColor)
        
    ElseIf MyState = bPressed And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        DrawFace GetColorSetClicked(m_BackColor, m_ForeColor)
        
    ElseIf MyState = bDisabled And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        DrawFace GetColorSetDisabled(m_BackColor, m_ForeColor)
        
    ElseIf MyState = bUnselected And MyTabState = tbNormal Then
        DrawFace GetColorSetNormal(m_BackColor, m_ForeColor)
        
    ElseIf MyState = bUnselected And MyTabState = tbTabed Then
        DrawFace GetColorSetNormal(m_BackColor, m_ForeColor) 'GetColorSetTabbed(m_BackColor, m_TabSelectColor, m_ForeColor)
        
    End If
    
End Sub

Private Sub UserControl_EnterFocus()
    
    MyTabState = tbNormal 'tbTabed
    DrawButton
    
End Sub

Private Sub UserControl_ExitFocus()
    
    MyTabState = tbNormal
    DrawButton

End Sub

Private Sub UserControl_Initialize()
    MyState = bUnselected
    MyTabState = tbNormal
    'DrawButton
End Sub

Private Sub DrawCaption(fntFont As Font, strText As String, FntColor As Long, XYOffset As Integer)

    Set UserControl.Font = fntFont
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) + XYOffset
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight(strText) / 2))) + XYOffset - 1
        
    UserControl.ForeColor = FntColor
    UserControl.Print strText

End Sub

Private Sub DrawFace(tColorSet As ColorSet)
   
    UserControl.BackColor = m_BackColor
    
    'background
    GradientCy UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tColorSet.csColor1(1), tColorSet.csColor1(2), tColorSet.csColor1(3), pbHorizontal

    'Borders
    GradientLine UserControl.hdc, 1, 0, UserControl.ScaleWidth - 2, pbHorizontal, tColorSet.csColor1(4), tColorSet.csColor1(5)
    GradientLine UserControl.hdc, 0, 1, UserControl.ScaleHeight - 2, pbVertical, tColorSet.csColor1(4), tColorSet.csColor1(6)
    UserControl.Line (1, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), tColorSet.csColor1(6)
    GradientLine UserControl.hdc, UserControl.ScaleWidth - 1, 1, UserControl.ScaleHeight - 2, pbVertical, tColorSet.csColor1(5), tColorSet.csColor1(6)
       
    If UserControl.Enabled = False Then
        MouseClick = False
        DrawCaption UserControl.Font, m_Caption, tColorSet.csColor1(6), 0
    Else
        DrawCaption UserControl.Font, m_Caption, tColorSet.csColor1(7), 0
    End If
    
End Sub

Private Function GetParentBackcolor() As Long
    GetParentBackcolor = GetLngColor(UserControl.Parent.BackColor)
End Function

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    DrawButton
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    DrawButton
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    DrawButton
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Schriftart"
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    DrawButton
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If KeyCode = 32 Or KeyCode = 13 Then
        
        MyState = bPressed
        DrawButton
        RaiseEvent Click
    
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    If KeyCode = 32 Or KeyCode = 13 Then
        MyState = bUnselected
        DrawButton
    End If

End Sub

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()

    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Caption = Ambient.DisplayName
    Set UserControl.Font = Ambient.Font
'    m_TabSelectColor = m_def_TabSelectColor
'    m_ImageIndex = m_def_ImageIndex
    
    DrawButton
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
'    m_TabSelectColor = PropBag.ReadProperty("TabSelectColor", m_def_TabSelectColor)
'    Set m_ImageList = PropBag.ReadProperty("ImageList", Nothing)
'    m_ImageIndex = PropBag.ReadProperty("ImageIndex", m_def_ImageIndex)
    
    DrawButton
    
End Sub

Private Sub UserControl_Resize()
    DrawButton
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
'    Call PropBag.WriteProperty("TabSelectColor", m_TabSelectColor, m_def_TabSelectColor)
'    Call PropBag.WriteProperty("ImageList", m_ImageList, Nothing)
'    Call PropBag.WriteProperty("ImageIndex", m_ImageIndex, m_def_ImageIndex)
    
    DrawButton
    
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    DrawButton
End Property

Private Sub MousePos_Timer()

    GetCursorPos Mouse

    If ((MouseOverMe.X > Mouse.X - 2) And (MouseOverMe.X < Mouse.X + 2)) And ((MouseOverMe.Y > Mouse.Y - 2) And (MouseOverMe.Y < Mouse.Y + 2)) Then
        If MouseClick = True Then
            MyState = bPressed
            DrawButton
        Else
            MyState = bHovered
            DrawButton
        End If
    Else
        MyState = bUnselected
        DrawButton
        MousePos.Enabled = False
    End If

    DoEvents
    
    'RaiseEvent MouseOver

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    MouseClick = True
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GetCursorPos MouseOverMe
        If MousePos.Enabled = False Then MousePos.Enabled = True
    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 1 Then Exit Sub
    If MouseClick = True Then RaiseEvent Click
    MouseClick = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub
''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
''MemberInfo=10,0,0,0
'Public Property Get TabSelectColor() As OLE_COLOR
'    TabSelectColor = m_TabSelectColor
'End Property
'
'Public Property Let TabSelectColor(ByVal New_TabSelectColor As OLE_COLOR)
'    m_TabSelectColor = New_TabSelectColor
'    PropertyChanged "TabSelectColor"
'    DrawButton
'End Property
''
'''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'''MemberInfo=9,0,0,0
''Public Property Get ImageList() As Object
''    Set ImageList = m_ImageList
''End Property
''
''Public Property Set ImageList(ByVal New_ImageList As Object)
''    Set m_ImageList = New_ImageList
''    PropertyChanged "ImageList"
''    DrawButton
''End Property
''
'''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'''MemberInfo=7,0,0,0
''Public Property Get ImageIndex() As Integer
''    ImageIndex = m_ImageIndex
''End Property
''
''Public Property Let ImageIndex(ByVal New_ImageIndex As Integer)
''    m_ImageIndex = New_ImageIndex
''    PropertyChanged "ImageIndex"
''    DrawButton
''End Property
''
''
