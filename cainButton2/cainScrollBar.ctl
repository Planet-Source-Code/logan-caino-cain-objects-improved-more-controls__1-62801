VERSION 5.00
Begin VB.UserControl cainScrollBar 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin cainObjects.cainXButton cainButton1 
      Height          =   375
      Index           =   0
      Left            =   1200
      Top             =   2040
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin VB.Timer MousePos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   0
   End
   Begin cainObjects.cainXButton cainButton1 
      Height          =   375
      Index           =   1
      Left            =   1200
      Top             =   2520
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
End
Attribute VB_Name = "cainScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Standard-Eigenschaftswerte:
Const m_def_ForeColor = &HC00000
Const m_def_BackColor = &HFFFFFF
Const m_def_Min = 1
Const m_def_Max = 100
Const m_def_Value = 50
'Const m_def_SelectionColor = &HFF0000
Const m_def_Orientation = 2
'Eigenschaftsvariablen:
Dim m_ForeColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_Min As Long
Dim m_Max As Long
Dim m_Value As Long
'Dim m_SelectionColor As OLE_COLOR
Dim m_Orientation As MyOrientation
Dim m_TabbedColor As OLE_COLOR
'Ereignisdeklarationen:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste drückt, während ein Objekt den Fokus hat."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste losläßt, während ein Objekt den Fokus hat."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Tritt auf, wenn der Benutzer die Maus bewegt."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Tritt auf, wenn der Benutzer eine Maustaste über einem Objekt drückt und wieder losläßt."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Tritt auf, wenn der Benutzer eine ANSI-Taste drückt und losläßt."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Tritt auf, wenn der Benutzer eine Taste losläßt, während ein Objekt den Fokus hat."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Tritt auf, wenn der Benutzer eine Taste drückt, während ein Objekt den Fokus besitzt."
Event Scroll()
Event ValueChanged()

Dim MyState As SlideState
'Dim MyTabState As TabSelected
Dim MyColorSet As ColorSet

Dim Mouse As POINTAPI
Dim MouseOverMe As POINTAPI
Dim MouseClick As Boolean

Dim OldValue As Long
Dim MyLeft As Long
Dim lCurrentDelta As Long

Public Function ID() As Integer
    ID = eControlIDs.id_Scrollbar
End Function

Private Function GetLeftFromValue(lValue As Long) As Long

    Dim i As Long
    Dim i2 As Long
    
    i = lValue - m_Min 'real Value
    
    If i = 0 Then
        i2 = 0 'if the value = 0 then we know it is 0 %
    Else
        i2 = (i / (m_Max - m_Min)) * 100 'Percent of Value from max
    End If
    If m_Orientation = pbHorizontal Then
        i = (UserControl.ScaleWidth - GetSlideHeight - (GetSlideOffset * 2)) 'Slider's sliding space in pixel
    ElseIf m_Orientation = pbVertical Then
        i = (UserControl.ScaleHeight - GetSlideHeight - (GetSlideOffset * 2)) 'Slider's sliding space in pixel
    End If
    MyLeft = Fix((i2 * i) / 100)
        
    GetLeftFromValue = MyLeft

End Function

Private Function GetValueFromLeft(SliderPos As Single) As Long

    Dim i2 As Long
    Dim i As Long
    
    If SliderPos = 0 Then
        i2 = 0 'if the value = 0 then we know it is 0 %
    Else
        If m_Orientation = pbHorizontal Then
            i2 = (SliderPos / (UserControl.ScaleWidth - GetSlideHeight - (GetSlideOffset * 2))) * 100 'Percent of Value from max
        ElseIf m_Orientation = pbVertical Then
            i2 = (SliderPos / (UserControl.ScaleHeight - GetSlideHeight - (GetSlideOffset * 2))) * 100 'Percent of Value from max
        End If
    End If
    
    i = Fix((i2 * (m_Max - m_Min)) / 100)
    i = i + m_Min
    If i < m_Min Then i = m_Min
    If i > m_Max Then i = m_Max
    
    GetValueFromLeft = i

End Function

Private Function GetSlideHeight() As Long
    If m_Orientation = pbHorizontal Then
        GetSlideHeight = (UserControl.ScaleWidth / 2) - GetSlideOffset
    ElseIf m_Orientation = pbVertical Then
        GetSlideHeight = (UserControl.ScaleHeight / 2) - GetSlideOffset
    End If
End Function

Private Function GetSlideOffset() As Long
    GetSlideOffset = cainButton1(0).Height + 5
End Function

Private Sub RefreshColor()
    
    Dim i As Integer
            
    If UserControl.Enabled = True Then
        If MyState = Slide_Disabled Then MyState = Slide_Normal
    Else
         MyState = Slide_Disabled
    End If
           
    If MyState = Slide_Normal Or MyState = Slide_Unselected Or MyState = Slide_TabSelected Then 'And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        MyColorSet = GetColorSetNormal(m_BackColor, m_ForeColor)
        
    ElseIf MyState = Slide_Hover Then 'And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        MyColorSet = GetColorSetHovered(m_BackColor, m_ForeColor)
        
    ElseIf MyState = Slide_Clicked Then 'And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        MyColorSet = GetColorSetHovered(m_BackColor, m_ForeColor)
        
    ElseIf MyState = Slide_Disabled Then 'And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        MyColorSet = GetColorSetDisabled(m_BackColor, m_ForeColor)
        
'    ElseIf MyState = Slide_Unselected And MyTabState = tbNormal Then
'        MyColorSet = GetColorSetNormal(m_BackColor, m_ForeColor)
'
'    ElseIf MyState = Slide_Unselected And MyTabState = tbTabed Then
'        MyColorSet = GetColorSetTabbed(m_BackColor, m_ForeColor, m_TabbedColor)
        
    End If
    
    For i = 0 To 1
        cainButton1(i).BackColor = m_BackColor
        cainButton1(i).ForeColor = m_ForeColor
    Next i

End Sub

Private Sub DrawScrollbar()
    
    UserControl.Cls
    UserControl.BackColor = MyColorSet.csColor1(6)
    
    If m_Orientation = pbHorizontal Then
        DrawGradient UserControl.hdc, 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, GetRGBColors(MyColorSet.csColor1(3)), GetRGBColors(MyColorSet.csColor1(2)), 1
    
        UserControl.Line (1, 1)-(1, UserControl.ScaleHeight - 2), MyColorSet.csColor1(1)
        UserControl.Line (1, 1)-(cainButton1(0).Width + 2, 1), MyColorSet.csColor1(1)
        UserControl.Line (cainButton1(0).Width + 2, 1)-(cainButton1(0).Width + 2, UserControl.ScaleHeight - 2), MyColorSet.csColor1(1)
        UserControl.Line (1, UserControl.ScaleHeight - 2)-(cainButton1(0).Width + 3, UserControl.ScaleHeight - 2), MyColorSet.csColor1(1)

        UserControl.Line (UserControl.ScaleWidth - 2, 1)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2), MyColorSet.csColor1(1)
        UserControl.Line (UserControl.ScaleWidth - 2, 1)-(UserControl.ScaleWidth - (cainButton1(0).Height + 3), 1), MyColorSet.csColor1(1)
        UserControl.Line (UserControl.ScaleWidth - (cainButton1(0).Height + 3), 1)-(UserControl.ScaleWidth - (cainButton1(0).Height + 3), UserControl.ScaleHeight - 1), MyColorSet.csColor1(1)
        UserControl.Line (UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2)-(UserControl.ScaleWidth - (cainButton1(0).Height + 3), UserControl.ScaleHeight - 2), MyColorSet.csColor1(1)

        UserControl.Line (cainButton1(0).Width + 3, 1)-(cainButton1(0).Width + 3, UserControl.ScaleHeight - 1), MyColorSet.csColor1(6)
        UserControl.Line (UserControl.ScaleWidth - (cainButton1(0).Height + 4), 1)-(UserControl.ScaleWidth - (cainButton1(0).Height + 4), UserControl.ScaleWidth - 1), MyColorSet.csColor1(6)
    
    ElseIf m_Orientation = pbVertical Then
        DrawGradient UserControl.hdc, 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2, GetRGBColors(MyColorSet.csColor1(3)), GetRGBColors(MyColorSet.csColor1(2)), 0
    
        UserControl.Line (1, 1)-(UserControl.ScaleWidth - 1, 1), MyColorSet.csColor1(1)
        UserControl.Line (1, 1)-(1, cainButton1(0).Height + 2), MyColorSet.csColor1(1)
        UserControl.Line (1, cainButton1(0).Height + 2)-(UserControl.ScaleWidth - 1, cainButton1(0).Height + 2), MyColorSet.csColor1(1)
        UserControl.Line (UserControl.ScaleWidth - 2, 1)-(UserControl.ScaleWidth - 2, cainButton1(0).Height + 2), MyColorSet.csColor1(1)
    
        UserControl.Line (1, UserControl.ScaleHeight - 2)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 2), MyColorSet.csColor1(1)
        UserControl.Line (1, UserControl.ScaleHeight - 2)-(1, UserControl.ScaleHeight - (cainButton1(0).Height + 3)), MyColorSet.csColor1(1)
        UserControl.Line (1, UserControl.ScaleHeight - (cainButton1(0).Height + 3))-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - (cainButton1(0).Height + 3)), MyColorSet.csColor1(1)
        UserControl.Line (UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - (cainButton1(0).Height + 3)), MyColorSet.csColor1(1)
    
        UserControl.Line (1, cainButton1(0).Height + 3)-(UserControl.ScaleWidth - 1, cainButton1(0).Height + 3), MyColorSet.csColor1(6)
        UserControl.Line (1, UserControl.ScaleHeight - (cainButton1(0).Height + 4))-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - (cainButton1(0).Height + 4)), MyColorSet.csColor1(6)

    End If

    If UserControl.Enabled = True Then DrawSlide
    
'    If m_Orientation = pbHorizontal Then
'
'        UserControl.Line (MyLeft, 0)-(MyLeft, UserControl.ScaleHeight), vbBlack
'        UserControl.Line (MyLeft + lCurrentDelta, 0)-(MyLeft + lCurrentDelta, UserControl.ScaleHeight), vbRed
'        UserControl.Line (MyLeft + GetSlideHeight, 0)-(MyLeft + GetSlideHeight, UserControl.ScaleHeight), vbBlue
'
'        UserControl.Line (MyLeft + GetSlideOffset, 0)-(MyLeft + GetSlideOffset, UserControl.ScaleHeight), vbGreen
'        UserControl.Line (MyLeft + GetSlideHeight + GetSlideOffset, 0)-(MyLeft + GetSlideHeight + GetSlideOffset, UserControl.ScaleHeight), vbGreen
'
'    ElseIf m_Orientation = pbVertical Then
'
'        UserControl.Line (0, MyLeft)-(UserControl.ScaleWidth, MyLeft), vbBlack
'        UserControl.Line (0, MyLeft + lCurrentDelta)-(UserControl.ScaleWidth, MyLeft + lCurrentDelta), vbRed
'        UserControl.Line (0, MyLeft + GetSlideHeight)-(UserControl.ScaleWidth, MyLeft + GetSlideHeight), vbBlue
'
'        UserControl.Line (0, MyLeft + GetSlideOffset)-(UserControl.ScaleWidth, MyLeft + GetSlideOffset), vbGreen
'        UserControl.Line (0, MyLeft + GetSlideHeight + GetSlideOffset)-(UserControl.ScaleWidth, MyLeft + GetSlideHeight + GetSlideOffset), vbGreen
'
'    End If

End Sub

Private Sub DrawSlide()

    Dim X As Long
    Dim Slide_Height As Long
    
    X = GetLeftFromValue(m_Value) + GetSlideOffset 'temp
    
    If OldValue <> m_Value Then
        RaiseEvent ValueChanged
        OldValue = m_Value
    End If
    
    Slide_Height = GetSlideHeight
    
    If m_Orientation = pbHorizontal Then
        
        GradientCy UserControl.hdc, X, 2, Slide_Height, UserControl.ScaleHeight - 2, MyColorSet.csColor1(1), MyColorSet.csColor1(2), MyColorSet.csColor1(3), pbHorizontal
'
'        'Borders
        GradientLine UserControl.hdc, X, 2, Slide_Height - 1, pbHorizontal, MyColorSet.csColor1(4), MyColorSet.csColor1(5)
        GradientLine UserControl.hdc, X, 2, UserControl.ScaleHeight - 4, pbVertical, MyColorSet.csColor1(4), MyColorSet.csColor1(6)
        UserControl.Line (X, UserControl.ScaleHeight - 3)-(X + Slide_Height - 1, UserControl.ScaleHeight - 3), MyColorSet.csColor1(6)
        GradientLine UserControl.hdc, X + Slide_Height - 1, 2, UserControl.ScaleHeight - 4, pbVertical, MyColorSet.csColor1(5), MyColorSet.csColor1(6)

        UserControl.Line (X - 1, 1)-(X - 1, UserControl.ScaleHeight - 1), MyColorSet.csColor1(1)
        UserControl.Line (X - 1, 1)-(X + Slide_Height, 1), MyColorSet.csColor1(1)
        UserControl.Line (X + Slide_Height, 1)-(X + Slide_Height, UserControl.ScaleHeight - 1), MyColorSet.csColor1(1)
        UserControl.Line (X + Slide_Height, UserControl.ScaleHeight - 2)-(X - 1, UserControl.ScaleHeight - 2), MyColorSet.csColor1(1)
        
        If Slide_Height > 10 Then
            UserControl.Line (X + ((Slide_Height / 2) - 1), 5)-(X + ((Slide_Height / 2) - 1), UserControl.ScaleHeight - 4), MyColorSet.csColor1(6)
            UserControl.Line (X + ((Slide_Height / 2) - 3), 5)-(X + ((Slide_Height / 2) - 3), UserControl.ScaleHeight - 4), MyColorSet.csColor1(6)
            UserControl.Line (X + ((Slide_Height / 2) + 1), 5)-(X + ((Slide_Height / 2) + 1), UserControl.ScaleHeight - 4), MyColorSet.csColor1(6)
            UserControl.Line (X + ((Slide_Height / 2) + 3), 5)-(X + ((Slide_Height / 2) + 3), UserControl.ScaleHeight - 4), MyColorSet.csColor1(6)
        
            UserControl.Line (X + ((Slide_Height / 2) - 2), 4)-(X + ((Slide_Height / 2) - 2), UserControl.ScaleHeight - 5), MyColorSet.csColor1(1)
            UserControl.Line (X + ((Slide_Height / 2) - 4), 4)-(X + ((Slide_Height / 2) - 4), UserControl.ScaleHeight - 5), MyColorSet.csColor1(1)
            UserControl.Line (X + ((Slide_Height / 2)), 4)-(X + ((Slide_Height / 2)), UserControl.ScaleHeight - 5), MyColorSet.csColor1(1)
            UserControl.Line (X + ((Slide_Height / 2) + 2), 4)-(X + ((Slide_Height / 2) + 2), UserControl.ScaleHeight - 5), MyColorSet.csColor1(1)
        End If
    
    ElseIf m_Orientation = pbVertical Then
        
        GradientCy UserControl.hdc, 2, X, UserControl.ScaleWidth - 2, Slide_Height, MyColorSet.csColor1(1), MyColorSet.csColor1(2), MyColorSet.csColor1(3), pbVertical
        
        'Borders
        GradientLine UserControl.hdc, 2, X, UserControl.ScaleWidth - 4, pbHorizontal, MyColorSet.csColor1(4), MyColorSet.csColor1(5)
        GradientLine UserControl.hdc, 2, X, Slide_Height - 1, pbVertical, MyColorSet.csColor1(4), MyColorSet.csColor1(6)
        UserControl.Line (3, X + Slide_Height - 1)-(UserControl.ScaleWidth - 3, X + Slide_Height - 1), MyColorSet.csColor1(6)
        GradientLine UserControl.hdc, UserControl.ScaleWidth - 3, X + 1, Slide_Height - 2, pbVertical, MyColorSet.csColor1(5), MyColorSet.csColor1(6)
    
        UserControl.Line (1, X - 1)-(UserControl.ScaleWidth - 1, X - 1), MyColorSet.csColor1(1)
        UserControl.Line (1, X - 1)-(1, X + Slide_Height), MyColorSet.csColor1(1)
        UserControl.Line (1, X + Slide_Height)-(UserControl.ScaleWidth - 1, X + Slide_Height), MyColorSet.csColor1(1)
        UserControl.Line (UserControl.ScaleWidth - 2, X + Slide_Height)-(UserControl.ScaleWidth - 2, X - 1), MyColorSet.csColor1(1)
    
        If Slide_Height > 10 Then
            UserControl.Line (5, X + ((Slide_Height / 2) - 1))-(UserControl.ScaleWidth - 4, X + ((Slide_Height / 2) - 1)), MyColorSet.csColor1(6)
            UserControl.Line (5, X + ((Slide_Height / 2) - 3))-(UserControl.ScaleWidth - 4, X + ((Slide_Height / 2) - 3)), MyColorSet.csColor1(6)
            UserControl.Line (5, X + ((Slide_Height / 2) + 1))-(UserControl.ScaleWidth - 4, X + ((Slide_Height / 2) + 1)), MyColorSet.csColor1(6)
            UserControl.Line (5, X + ((Slide_Height / 2) + 3))-(UserControl.ScaleWidth - 4, X + ((Slide_Height / 2) + 3)), MyColorSet.csColor1(6)
        
            UserControl.Line (4, X + ((Slide_Height / 2) - 2))-(UserControl.ScaleWidth - 5, X + ((Slide_Height / 2) - 2)), MyColorSet.csColor1(1)
            UserControl.Line (4, X + ((Slide_Height / 2) - 4))-(UserControl.ScaleWidth - 5, X + ((Slide_Height / 2) - 4)), MyColorSet.csColor1(1)
            UserControl.Line (4, X + ((Slide_Height / 2)))-(UserControl.ScaleWidth - 5, X + ((Slide_Height / 2))), MyColorSet.csColor1(1)
            UserControl.Line (4, X + ((Slide_Height / 2) + 2))-(UserControl.ScaleWidth - 5, X + ((Slide_Height / 2) + 2)), MyColorSet.csColor1(1)
        End If
        
    End If

End Sub

Private Sub cainButton1_Click(Index As Integer)
    
    Select Case Index
        Case 0
            Value = m_Value - 1
        Case 1
            Value = m_Value + 1
    End Select
    
End Sub

Private Sub cainButton1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub cainButton1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub cainButton1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub MousePos_Timer()

    GetCursorPos Mouse

    If ((MouseOverMe.X > Mouse.X - 2) And (MouseOverMe.X < Mouse.X + 2)) And ((MouseOverMe.Y > Mouse.Y - 2) And (MouseOverMe.Y < Mouse.Y + 2)) Then
        If MouseClick = True Then
            MyState = Slide_Clicked
            RefreshColor
            DrawScrollbar
            RaiseEvent Click
        Else
            MyState = Slide_Hover
            RefreshColor
            DrawScrollbar
        End If
    Else
        MyState = Slide_Unselected
        RefreshColor
        DrawScrollbar
        MousePos.Enabled = False
    End If

    DoEvents
    
    'RaiseEvent MouseOver

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If Button = 1 Then
        If m_Orientation = pbHorizontal Then
            If (((X > MyLeft + GetSlideOffset) And (X < (MyLeft + GetSlideHeight + GetSlideOffset)) And ((Y > 2) And (Y < UserControl.ScaleWidth - 2)))) Then
                MouseClick = True
                lCurrentDelta = X - MyLeft - GetSlideOffset
            ElseIf X < MyLeft + GetSlideOffset Then
                Value = m_Min
            ElseIf X > MyLeft + GetSlideOffset Then
                Value = m_Max
            End If
        ElseIf m_Orientation = pbVertical Then
            'Value = GetValueFromLeft(Y - 10 - cainButton1(0).Height)
            
            If (((Y > MyLeft + GetSlideOffset) And (Y < (MyLeft + GetSlideHeight + GetSlideOffset)) And ((X > 2) And (X < UserControl.ScaleWidth - 2)))) Then
                MouseClick = True
                lCurrentDelta = Y - MyLeft - GetSlideOffset
            ElseIf (Y < MyLeft + GetSlideOffset) Then
                Value = m_Min
            ElseIf (Y > MyLeft + GetSlideOffset) Then
                Value = m_Max
            End If
        End If
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    If m_Orientation = pbHorizontal Then
        If (((X > MyLeft + GetSlideOffset) And (X < (MyLeft + GetSlideHeight + GetSlideOffset)) And ((Y > 2) And (Y < UserControl.ScaleWidth - 2)))) Or MouseClick = True Then
            GetCursorPos MouseOverMe
        
                If MousePos.Enabled = False Then MousePos.Enabled = True
            
            If Button = 1 Then
                Value = GetValueFromLeft(X - GetSlideOffset - lCurrentDelta)
                RaiseEvent Scroll
            End If
        End If
    ElseIf m_Orientation = pbVertical Then
        If (((Y > MyLeft + GetSlideOffset) And (Y < (MyLeft + GetSlideHeight + GetSlideOffset)) And ((X > 2) And (X < UserControl.ScaleWidth - 2)))) Or MouseClick = True Then
            GetCursorPos MouseOverMe
        
                If MousePos.Enabled = False Then MousePos.Enabled = True
            
            If Button = 1 Then
                'lCurrentDelta = Y - MyLeft '+ GetSlideOffset
                Value = GetValueFromLeft(Y - GetSlideOffset - lCurrentDelta)
                RaiseEvent Scroll
            End If
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 1 Then Exit Sub
    If MouseClick = True Then RaiseEvent Click
    MouseClick = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub

'
'
''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
''MappingInfo=UserControl,UserControl,-1,BackColor
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    UserControl.BackColor() = New_BackColor
'    PropertyChanged "BackColor"
'    DrawScrollbar
'End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt einen Wert zurück, der bestimmt, ob ein Objekt auf vom Benutzer erzeugte Ereignisse reagieren kann, oder legt diesen fest."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
    RefreshColor
    DrawScrollbar
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    
    RefreshColor
    DrawScrollbar
    
End Sub

Private Sub ButtonInit()
    
    Dim i As Integer

    For i = 0 To 1
        cainButton1(i).Font.Name = "Webdings"
        cainButton1(i).Font.Size = 8
    Next i
    
    If m_Orientation = pbHorizontal Then
        cainButton1(0).Caption = 3
        cainButton1(1).Caption = 4
    
    ElseIf m_Orientation = pbVertical Then
        cainButton1(0).Caption = 5
        cainButton1(1).Caption = 6
        
    End If

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=8,0,0,0
Public Property Get Min() As Long
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Verschiedenes"
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
    
    If New_Min > m_Max Then New_Min = m_Max - 1
    If New_Min = m_Max Then New_Min = m_Max - 1
    
    m_Min = New_Min
    PropertyChanged "Min"
    
    If m_Value < m_Min Then
        m_Value = New_Min
        PropertyChanged "Value"
    End If

    DrawScrollbar
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=8,0,0,0
Public Property Get Max() As Long
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Verschiedenes"
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)

    If New_Max < m_Min Then New_Max = m_Min + 1
    If New_Max = m_Min Then New_Max = m_Min + 1
   
    m_Max = New_Max
    PropertyChanged "Max"
    
    If m_Value > m_Max Then
        m_Value = New_Max
        PropertyChanged "Value"
    End If
    
    DrawScrollbar
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Verschiedenes"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    
    'Prevents Slider from executing this line again and again
    If OldValue = New_Value Then Exit Property
    
    If New_Value < m_Min Then New_Value = m_Min
    If New_Value > m_Max Then New_Value = m_Max
    
    OldValue = m_Value
    
    m_Value = New_Value
    PropertyChanged "Value"
    GetLeftFromValue m_Value
    DrawScrollbar
        
End Property
'
''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
''MemberInfo=10,0,0,0
'Public Property Get SelectionColor() As OLE_COLOR
'    SelectionColor = m_SelectionColor
'End Property
'
'Public Property Let SelectionColor(ByVal New_SelectionColor As OLE_COLOR)
'    m_SelectionColor = New_SelectionColor
'    PropertyChanged "SelectionColor"
'    DrawScrollbar
'End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,0
Public Property Get Orientation() As MyOrientation
Attribute Orientation.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As MyOrientation)
    m_Orientation = New_Orientation
    PropertyChanged "Orientation"
    ButtonInit
    UserControl_Resize
    DrawScrollbar
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
'Public Property Get TabbedColor() As OLE_COLOR
'    TabbedColor = m_TabbedColor
'End Property
'
'Public Property Let TabbedColor(ByVal New_TabbedColor As OLE_COLOR)
'    m_TabbedColor = New_TabbedColor
'    PropertyChanged "TabbedColor"
'    RefreshColor
'    DrawScrollbar
'End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()

    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
'    m_SelectionColor = m_def_SelectionColor
    m_Orientation = m_def_Orientation
    
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    
    UserControl_Resize
    
    RefreshColor
    DrawScrollbar
    
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
'    m_SelectionColor = PropBag.ReadProperty("SelectionColor", m_def_SelectionColor)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
   ' m_TabbedColor = PropBag.ReadProperty("TabbedColor", m_def_TabbedColor)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    
    UserControl_Resize
    ButtonInit
    
    RefreshColor
    DrawScrollbar
End Sub

Private Sub UserControl_Resize()

    If m_Orientation = pbHorizontal Then
    
        If UserControl.Height > (25 * Screen.TwipsPerPixelX) Then UserControl.Height = (25 * Screen.TwipsPerPixelX)
        If UserControl.Height < (15 * Screen.TwipsPerPixelX) Then UserControl.Height = (15 * Screen.TwipsPerPixelX)
    
        cainButton1(0).Height = UserControl.ScaleHeight - 4
        cainButton1(0).Width = UserControl.ScaleHeight - 4
    
        cainButton1(1).Height = UserControl.ScaleHeight - 4
        cainButton1(1).Width = UserControl.ScaleHeight - 4
    
        cainButton1(0).Top = 2
        cainButton1(1).Top = 2
        
        cainButton1(0).Left = 2
        cainButton1(1).Left = UserControl.ScaleWidth - (cainButton1(1).Height + 2)
    
    ElseIf m_Orientation = pbVertical Then
    
        If UserControl.Width > (25 * Screen.TwipsPerPixelX) Then UserControl.Width = (25 * Screen.TwipsPerPixelX)
        If UserControl.Width < (15 * Screen.TwipsPerPixelX) Then UserControl.Width = (15 * Screen.TwipsPerPixelX)
    
        cainButton1(0).Height = UserControl.ScaleWidth - 4
        cainButton1(0).Width = UserControl.ScaleWidth - 4
    
        cainButton1(1).Height = UserControl.ScaleWidth - 4
        cainButton1(1).Width = UserControl.ScaleWidth - 4
    
        cainButton1(0).Left = 2
        cainButton1(1).Left = 2
        
        cainButton1(0).Top = 2
        cainButton1(1).Top = UserControl.ScaleHeight - (cainButton1(1).Height + 2)
    
    End If

    DrawScrollbar
    
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
'    Call PropBag.WriteProperty("SelectionColor", m_SelectionColor, m_def_SelectionColor)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    'Call PropBag.WriteProperty("TabbedColor", m_TabbedColor, m_def_TabbedColor)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    
    UserControl_Resize
    ButtonInit
    RefreshColor
    DrawScrollbar
    
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,&HFFFFFF
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    RefreshColor
    DrawScrollbar
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=14,0,0,&HC00000
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    RefreshColor
    DrawScrollbar
End Property

