VERSION 5.00
Begin VB.UserControl cainLabel 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer MousePos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "cainLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Standard-Eigenschaftswerte:
Const m_def_AutoSize = 0
Const m_def_TextWrap = False
Const m_def_BackColor = &HFFFFFF
Const m_def_ForeColor = &HC00000
Const m_def_ImageIndex = 0
Const m_def_Hyperlink = ""
Const m_def_SelectionColor = &H80FF&
'Eigenschaftsvariablen:
Dim m_AutoSize As Boolean
Dim m_TextWrap As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
'Dim m_Font As Font
Dim m_ImageList As Object
Dim m_ImageIndex As Integer
Dim m_Caption As String
Dim m_Hyperlink As String
Dim m_SelectionColor As OLE_COLOR
'Ereignisdeklarationen:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste drückt, während ein Objekt den Fokus hat."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Tritt auf, wenn der Benutzer die Maus bewegt."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste losläßt, während ein Objekt den Fokus hat."
Event Click()
Event HyperlinkClick(sHyperlink As String, Cancel As Integer)

Dim tColorSet As ColorSet

Dim Mouse As POINTAPI
Dim MouseOverMe As POINTAPI
Dim MouseClick As Boolean

Dim MyState As ButtonState

Public Function ID() As Integer
    ID = eControlIDs.id_Label
End Function

Private Sub RefreshColor()

    tColorSet = GetColorSetNormal(m_BackColor, m_ForeColor, m_SelectionColor)
    UserControl.ForeColor = tColorSet.csColor1(7)
    UserControl.BackColor = BlendColor(m_BackColor, m_ForeColor, 240)
    
    DrawLabel
    
End Sub

Private Sub DrawLabel()
    
    UserControl.Cls
    
    If UserControl.Enabled = True Then
        If (Len(m_Hyperlink) <> 0) And MyState = bHovered Then
            UserControl.ForeColor = tColorSet.csColor1(10)
        Else
            UserControl.ForeColor = tColorSet.csColor1(7)
        End If
    Else
        UserControl.ForeColor = tColorSet.csColor1(6)
    End If
    
    If (m_ImageList Is Nothing) Or (Not (m_ImageList Is Nothing) And m_ImageIndex = 0) Then
        
        If m_TextWrap = True Then
            WrapText m_Caption, UserControl.ScaleHeight - 8, UserControl.ScaleWidth - 2, 0, 2
        Else
            UserControl.CurrentX = 0
            UserControl.CurrentY = 2 '(UserControl.ScaleHeight / 2) - (UserControl.TextHeight(m_Caption) / 2)
            UserControl.Print ShortWord(Trim(Replace(Replace(m_Caption, vbCr, "ë"), vbLf, "")), UserControl.ScaleWidth - 2)
        End If
        
    Else
    
        UserControl.CurrentX = m_ImageList.ImageWidth + 4
        
        
        If m_TextWrap = True Then
            WrapText m_Caption, UserControl.ScaleHeight - 8, UserControl.ScaleWidth - (m_ImageList.ImageWidth + 4), m_ImageList.ImageWidth + 4, 2
        Else
        
            If m_ImageList.ImageHeight > UserControl.TextHeight("I") Then
                UserControl.CurrentY = 2 + ((m_ImageList.ImageHeight / 2) - (UserControl.TextHeight(m_Caption) / 2)) '(UserControl.ScaleHeight / 2) - (UserControl.TextHeight(m_Caption) / 2)
            Else
                UserControl.CurrentY = 2
            End If
            
            UserControl.Print ShortWord(Trim(Replace(Replace(m_Caption, vbCr, "ë"), vbLf, "")), UserControl.ScaleWidth - (m_ImageList.ImageWidth + 4))
            
        End If
        
        If UserControl.Enabled = False Then
            Call DrawState(UserControl.hdc, 0, 0, _
            m_ImageList.ListImages(m_ImageIndex).ExtractIcon, 0, _
             2, 2, _
            m_ImageList.ImageWidth, m_ImageList.ImageHeight, DST_ICON Or DSS_MONO)
        
        Else
            Call DrawState(UserControl.hdc, 0, 0, _
            m_ImageList.ListImages(m_ImageIndex).ExtractIcon, 0, _
            2, 2, _
            m_ImageList.ImageWidth, m_ImageList.ImageHeight, DST_ICON Or DSS_NORMAL)
        End If
        
    End If
    
End Sub

Private Sub WrapText(sString As String, H As Single, W As Single, X As Single, Y As Single, Optional iAlignment As Integer = 0)

    Dim tmpString As String
    Dim tmpLong As Long
    Dim xString As String
    Dim sText As String
    
    tmpString = Trim(Replace(Replace(sString, vbCr, "ë"), vbLf, ""))
    UserControl.CurrentY = Y
    
    Do
        If ((UserControl.TextHeight("I") / 2) + UserControl.CurrentY) > (Y + H) Then Exit Do
        sText = ShortWord(GetWrapLine(tmpString, W * 1, tmpLong), W * 1)
        Select Case iAlignment
        Case 0
            UserControl.CurrentX = X
        
        Case 1
            UserControl.CurrentX = X + ((W) / 2) - (UserControl.TextWidth(sText) / 2)
        
        Case 2
            UserControl.CurrentX = X + (W - UserControl.TextWidth(sText))
        
        End Select
        UserControl.Print sText '& "ë"
        tmpString = Right(tmpString, Len(tmpString) - tmpLong)
        If tmpString = "" Then Exit Do
    
    Loop

End Sub

Private Function GetWrapLine(sText As String, RefLenght As Long, lLenght As Long) As String

    Dim tmpLong As Long
    Dim tmpLong2 As Long
    Dim tmpString As String
    Dim sString As String
    
    sString = Left(sText, 1000)
    tmpString = sString
    tmpLong2 = Len(sString)
        
    Do
        
        If UserControl.TextWidth(tmpString) < RefLenght Then Exit Do
        
        tmpLong = InStrRev(sString, " ", tmpLong2)
        If tmpLong = 0 Then Exit Do
        tmpString = Left(sString, tmpLong)
        tmpLong2 = tmpLong - 1
        
    Loop
       
    tmpLong = InStr(1, tmpString, "ë")
    If tmpLong <> 0 Then
        tmpString = Left(tmpString, tmpLong)
    End If
            
    lLenght = Len(tmpString)
    GetWrapLine = Replace(Trim(tmpString), "ë", "")

End Function

Private Function ShortWord(sString As String, RefLenght As Long) As String
    
    Dim tmpString As String
    
    tmpString = sString
    
    If UserControl.TextWidth(tmpString) > RefLenght Then
    
        Do Until UserControl.TextWidth(tmpString) < RefLenght
        
            If Len(tmpString) <= 3 Then Exit Do
            tmpString = Left(tmpString, Len(tmpString) - 4) & "..."
            
        Loop
        
    End If
    
    ShortWord = tmpString

End Function

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    RefreshColor
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Vordergrundfarbe zurück, die zum Anzeigen von Text und Grafiken in einem Objekt verwendet wird, oder legt diese fest."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    
    RefreshColor
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt einen Wert zurück, der bestimmt, ob ein Objekt auf vom Benutzer erzeugte Ereignisse reagieren kann, oder legt diesen fest."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    
    DrawLabel
End Property
'
''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
''MemberInfo=6,0,0,0
'Public Property Get Font() As Font
'    Set Font = m_Font
'End Property
'
'Public Property Set Font(ByVal New_Font As Font)
'    Set m_Font = New_Font
'    PropertyChanged "Font"
'
'    UserControl_Resize
'End Property

Private Sub MousePos_Timer()

    GetCursorPos Mouse

    If ((MouseOverMe.X > Mouse.X - 2) And (MouseOverMe.X < Mouse.X + 2)) And ((MouseOverMe.Y > Mouse.Y - 2) And (MouseOverMe.Y < Mouse.Y + 2)) Then
        If MouseClick = True Then
            MyState = bPressed
            DrawLabel
        Else
            MyState = bHovered
            DrawLabel
        End If
    Else
        MyState = bUnselected
        DrawLabel
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
    
    If Len(m_Hyperlink) <> 0 Then
        GetCursorPos MouseOverMe
            If MousePos.Enabled = False Then MousePos.Enabled = True
    End If
        
    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 1 Then Exit Sub
    If MouseClick = True Then
        RaiseEvent Click
        
        If Len(m_Hyperlink) <> 0 Then
            Dim iCancel As Integer
            iCancel = 0
            RaiseEvent HyperlinkClick(m_Hyperlink, iCancel)
            If iCancel = 0 Then
                ShellDocument m_Hyperlink
            End If
        End If
        
    End If
    MouseClick = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=9,0,0,0
Public Property Get ImageList() As Object
Attribute ImageList.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    Set ImageList = m_ImageList
End Property

Public Property Set ImageList(ByVal New_ImageList As Object)
    Set m_ImageList = New_ImageList
    PropertyChanged "ImageList"
    
    UserControl_Resize
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,0
Public Property Get ImageIndex() As Integer
Attribute ImageIndex.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    ImageIndex = m_ImageIndex
End Property

Public Property Let ImageIndex(ByVal New_ImageIndex As Integer)
    m_ImageIndex = New_ImageIndex
    PropertyChanged "ImageIndex"
    
    DrawLabel
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    
    DrawLabel
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,0
Public Property Get Hyperlink() As String
Attribute Hyperlink.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    Hyperlink = m_Hyperlink
End Property

Public Property Let Hyperlink(ByVal New_Hyperlink As String)
    m_Hyperlink = New_Hyperlink
    PropertyChanged "Hyperlink"
    
    DrawLabel
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get SelectionColor() As OLE_COLOR
Attribute SelectionColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    SelectionColor = m_SelectionColor
End Property

Public Property Let SelectionColor(ByVal New_SelectionColor As OLE_COLOR)
    m_SelectionColor = New_SelectionColor
    PropertyChanged "SelectionColor"
    
    RefreshColor
End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
'    Set m_Font = Ambient.Font
    m_ImageIndex = m_def_ImageIndex
    m_Caption = Ambient.DisplayName
    m_Hyperlink = m_def_Hyperlink
    m_SelectionColor = m_def_SelectionColor
    m_TextWrap = m_def_TextWrap
    m_AutoSize = m_def_AutoSize
    Set UserControl.Font = Ambient.Font
    
    RefreshColor
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
'    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set m_ImageList = PropBag.ReadProperty("ImageList", Nothing)
    m_ImageIndex = PropBag.ReadProperty("ImageIndex", m_def_ImageIndex)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_Hyperlink = PropBag.ReadProperty("Hyperlink", m_def_Hyperlink)
    m_SelectionColor = PropBag.ReadProperty("SelectionColor", m_def_SelectionColor)
    m_TextWrap = PropBag.ReadProperty("TextWrap", m_def_TextWrap)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    
    RefreshColor
End Sub

Private Sub UserControl_Resize()

    If m_AutoSize = True Then
        If m_ImageList Is Nothing Then
            UserControl.Height = (UserControl.TextHeight("I") + 4) * Screen.TwipsPerPixelY
            UserControl.Width = (UserControl.TextWidth(m_Caption) + 4) * Screen.TwipsPerPixelX
        Else
            If m_ImageList.ImageHeight > UserControl.TextHeight("I") Then
                UserControl.Height = (m_ImageList.ImageHeight + 4) * Screen.TwipsPerPixelY
            Else
                UserControl.Height = (UserControl.TextHeight("I") + 4) * Screen.TwipsPerPixelY
            End If
    
            UserControl.Width = (m_ImageList.ImageWidth + UserControl.TextWidth(m_Caption) + 8) * Screen.TwipsPerPixelX
    
        End If
    End If
    
    DrawLabel

End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
'    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("ImageList", m_ImageList, Nothing)
    Call PropBag.WriteProperty("ImageIndex", m_ImageIndex, m_def_ImageIndex)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Hyperlink", m_Hyperlink, m_def_Hyperlink)
    Call PropBag.WriteProperty("SelectionColor", m_SelectionColor, m_def_SelectionColor)
    Call PropBag.WriteProperty("TextWrap", m_TextWrap, m_def_TextWrap)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    
    RefreshColor
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,False
Public Property Get TextWrap() As Boolean
Attribute TextWrap.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    TextWrap = m_TextWrap
End Property

Public Property Let TextWrap(ByVal New_TextWrap As Boolean)
    m_TextWrap = New_TextWrap
    PropertyChanged "TextWrap"
    m_AutoSize = False
    PropertyChanged "AutoSize"
    DrawLabel
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,0
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    
    m_TextWrap = False
    PropertyChanged "TextWrap"
    UserControl_Resize
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Gibt ein Font-Objekt zurück."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Schriftart"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    
    UserControl_Resize
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Gibt den Typ des Mauszeigers zurück, der angezeigt wird, wenn dieser sich über einem Teil eines Objekts befindet, oder legt diesen fest."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Legt ein benutzerdefiniertes Maussymbol fest."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

