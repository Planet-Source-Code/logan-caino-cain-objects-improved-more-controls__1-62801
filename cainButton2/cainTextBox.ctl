VERSION 5.00
Begin VB.UserControl cainTextBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin cainObjects.cainPUMenu cainPUMenu1 
      Height          =   705
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1244
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Text            =   "Dummy Text"
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "cainTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum DataFormats

    df_AllChars = 0
    df_AlphaOnly = 1
    df_NumOnly = 2
    df_NumAndChars = 3
    df_NumAndAlpha = 4
    df_NumAndAlphaChars = 5
    df_UCase = 6
    df_UCaseAlphaOnly = 7
    df_UCaseNumAndAlpha = 8
    df_UCaseNumAndAlphaChars = 9
    df_LCase = 10
    df_LCaseAlphaOnly = 11
    df_LCaseNumAndAlpha = 12
    df_LCaseNumAndAlphaChars = 13
    df_AlphaAndChars = 14
    df_UCaseAlphaAndChars = 15
    df_LCaseAlphaAndChars = 16

End Enum

'Standard-Eigenschaftswerte:
'Const m_def_Title = ""
Const m_def_textCut = "Cut"
Const m_def_textCopy = "Copy"
Const m_def_textPaste = "Paste"
Const m_def_textDelete = "Delete"
Const m_def_textSelectAll = "Select all"
Const m_def_SelectionColor = &H80FF&
Const m_def_BackColor = &HFFFFFF
Const m_def_ForeColor = &HC00000
Const m_def_AutoSelect = 0
Const m_def_DataFormat = 0
'Eigenschaftsvariablen:
'Dim m_Title As String
Dim m_textCut As String
Dim m_textCopy As String
Dim m_textPaste As String
Dim m_textDelete As String
Dim m_textSelectAll As String
Dim m_SelectionColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Font As Font
Dim m_AutoSelect As Boolean
Dim m_DataFormat As DataFormats
'Ereignisdeklarationen:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyDown
Attribute KeyDown.VB_Description = "Tritt auf, wenn der Benutzer eine Taste drückt, während ein Objekt den Fokus besitzt."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Text1,Text1,-1,KeyPress
Attribute KeyPress.VB_Description = "Tritt auf, wenn der Benutzer eine ANSI-Taste drückt und losläßt."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyUp
Attribute KeyUp.VB_Description = "Tritt auf, wenn der Benutzer eine Taste losläßt, während ein Objekt den Fokus hat."

Dim tColorSet As ColorSet

Dim oldSelStart As Long
Dim oldSelLenght As Long
Dim bMenuOpened As Boolean

Public Function ID() As Integer
    ID = eControlIDs.id_Textbox
End Function

Private Sub DrawFace()
        
    tColorSet = GetColorSetNormal(m_BackColor, m_ForeColor)
    
    UserControl.BackColor = tColorSet.csColor1(1)
    Text1.BackColor = tColorSet.csColor1(1)
    If UserControl.Enabled = True Then
        Text1.ForeColor = tColorSet.csColor1(7)
    Else
        Text1.ForeColor = tColorSet.csColor1(6)
    End If
    
    UserControl.Cls
    
    'background
    'GradientCy UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tColorSet.csColor1(1), tColorSet.csColor1(2), tColorSet.csColor1(3), pbHorizontal

    'Borders
    GradientLine UserControl.hdc, 1, 0, UserControl.ScaleWidth - 2, pbHorizontal, tColorSet.csColor1(4), tColorSet.csColor1(5)
    GradientLine UserControl.hdc, 0, 1, UserControl.ScaleHeight - 2, pbVertical, tColorSet.csColor1(4), tColorSet.csColor1(6)
    UserControl.Line (1, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), tColorSet.csColor1(6)
    GradientLine UserControl.hdc, UserControl.ScaleWidth - 1, 1, UserControl.ScaleHeight - 2, pbVertical, tColorSet.csColor1(5), tColorSet.csColor1(6)

End Sub

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
    DrawFace
    SetMenuColor
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
    DrawFace
    SetMenuColor
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
    
    DrawFace
    
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Gibt ein Font-Objekt zurück."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Schriftart"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
    
    SetFont
    SetMenuColor
    
End Property

Private Sub SetFont()
    
    UserControl.Font = m_Font
    Text1.FontBold = m_Font.Bold
    Text1.FontItalic = m_Font.Italic
    Text1.FontName = m_Font.Name
    Text1.FontSize = m_Font.Size
    Text1.FontStrikethru = m_Font.Strikethrough
    Text1.FontUnderline = m_Font.Underline
    
    UserControl_Resize

End Sub

Private Sub cainPUMenu1_Closed()
    On Error Resume Next
    
    Text1.Enabled = True
    Text1.SetFocus
    Text1.SelStart = oldSelStart
    Text1.SelLength = oldSelLenght
End Sub

Private Sub cainPUMenu1_GotFocus()
    bMenuOpened = True
End Sub

Private Sub cainPUMenu1_ItemClick(ItemIndex As Integer, ItemKey As String)

    On Error Resume Next
    
    Select Case ItemKey
        
        Case "a"
            Clipboard.Clear
            Clipboard.SetText Mid(Text1.Text, oldSelStart + 1, oldSelLenght)
            Text1.Text = Left(Text1.Text, oldSelStart) & Right(Text1.Text, Len(Text1.Text) - oldSelStart - oldSelLenght)
            oldSelLenght = 0
            
        Case "b"
            Clipboard.Clear
            Clipboard.SetText Mid(Text1.Text, oldSelStart + 1, oldSelLenght)
            
        Case "c"
        
            Dim i As Long
            Dim cText As String
            Dim tASCII As Integer
        
            'We have to check if the clipboard data has the right format, if not... Clean it!!
            cText = Clipboard.GetText
            
            For i = 1 To Len(cText)
                tASCII = InputCheck(Asc(Mid(cText, i, 1)))
                If tASCII = 0 Then tASCII = Asc(" ")
                
                Mid(cText, i, 1) = Chr(tASCII)
            Next i
    
            Text1.Text = Left(Text1.Text, oldSelStart) & cText & Right(Text1.Text, Len(Text1.Text) - oldSelStart - oldSelLenght)
            oldSelStart = oldSelStart + Len(cText)
            oldSelLenght = 0
            
        Case "d"
            Text1.Text = Left(Text1.Text, oldSelStart) & Right(Text1.Text, Len(Text1.Text) - oldSelStart - oldSelLenght)
            oldSelLenght = 0
        
        Case "f"
            oldSelStart = 0
            oldSelLenght = Len(Text1.Text)
            
    
    End Select

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    KeyAscii = InputCheck(KeyAscii)

    RaiseEvent KeyPress(KeyAscii)
    
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Function InputCheck(strdfText As Integer) As Integer

    Dim tmpStr As String

    tmpStr = Chr(strdfText)

    
    Select Case DataFormat
    
    Case DataFormats.df_AllChars
        InputCheck = strdfText
        Exit Function
        
    Case DataFormats.df_AlphaOnly
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = strdfText
            Exit Function
        End Select
        
    Case DataFormats.df_LCase
        InputCheck = Asc(LCase(Chr(strdfText)))
        Exit Function
        
    Case DataFormats.df_LCaseAlphaOnly
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(LCase(Chr(strdfText)))
            Exit Function
        End Select
        
    Case DataFormats.df_LCaseNumAndAlpha
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), 0 To 9, "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(LCase(Chr(strdfText)))
            Exit Function
        End Select
        
    Case DataFormats.df_LCaseNumAndAlphaChars
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), 0 To 9, ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(LCase(Chr(strdfText)))
            Exit Function
        End Select
        
    Case DataFormats.df_NumAndAlpha
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), 0 To 9, "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = strdfText
            Exit Function
        End Select
        
    Case DataFormats.df_NumAndAlphaChars
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), 0 To 9, ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = strdfText
            Exit Function
        End Select
        
    Case DataFormats.df_NumAndChars
        Select Case tmpStr
        Case Chr(8), 0 To 9, ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = strdfText
            Exit Function
        End Select
        
    Case DataFormats.df_NumOnly
        Select Case tmpStr
        Case Chr(8), 0 To 9
            InputCheck = strdfText
            Exit Function
        End Select
        
    Case DataFormats.df_UCase
        InputCheck = Asc(UCase(Chr(strdfText)))
        Exit Function
    
    Case DataFormats.df_UCaseAlphaOnly
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(UCase(Chr(strdfText)))
            Exit Function
        End Select
    
    Case DataFormats.df_UCaseAlphaAndChars
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(UCase(Chr(strdfText)))
            Exit Function
        End Select
    
    Case DataFormats.df_LCaseAlphaAndChars
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(LCase(Chr(strdfText)))
            Exit Function
        End Select
    
    Case DataFormats.df_AlphaAndChars
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = strdfText
            Exit Function
        End Select
        
    Case DataFormats.df_UCaseNumAndAlpha
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), 0 To 9, "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(UCase(Chr(strdfText)))
            Exit Function
        End Select
        
        
    Case DataFormats.df_UCaseNumAndAlphaChars
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), 0 To 9, ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(UCase(Chr(strdfText)))
            Exit Function
        End Select
    
    End Select
    
    InputCheck = 0
    
End Function

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Text1,Text1,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Gibt die Ausrichtung eines Kontrollkästchens, eines Optionsfeldes oder eines Steuerelementtextes zurück oder legt sie fest."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    Alignment = Text1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    Text1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Text1,Text1,-1,MultiLine
'Public Property Get MultiLine() As Boolean
'    MultiLine = Text1.MultiLine
'End Property
'
'Public Property Let MultiLine(ByVal New_MultiLine As Boolean)
'    Text1.MultiLine = New_MultiLine
'End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Gibt den Text zurück, der im Steuerelement enthalten ist, oder legt diesen fest."
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,0
Public Property Get AutoSelect() As Boolean
Attribute AutoSelect.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    AutoSelect = m_AutoSelect
End Property

Public Property Let AutoSelect(ByVal New_AutoSelect As Boolean)
    m_AutoSelect = New_AutoSelect
    PropertyChanged "AutoSelect"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Text1,Text1,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Gibt die Anzahl der ausgewählten Zeichen zurück oder legt diese fest."
Attribute SelLength.VB_ProcData.VB_Invoke_Property = ";Verschiedenes"
    SelLength = Text1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    Text1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Text1,Text1,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Gibt den Anfangspunkt des ausgewählten Textes zurück oder legt diesen fest."
Attribute SelStart.VB_ProcData.VB_Invoke_Property = ";Verschiedenes"
    SelStart = Text1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    Text1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Text1,Text1,-1,ScrollBars
'Public Property Get ScrollBars() As Integer
'    ScrollBars = Text1.ScrollBars
'End Property
'
'Public Property Let ScrollBars(ByVal New_ScrollBars As Integer)
'    Text1.ScrollBars = New_ScrollBars
'End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Text1,Text1,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Gibt einen Wert zurück, der bestimmt, ob Zeichen, die von einem Benutzer eingegeben werden, oder Platzhalterzeichen in einem Steuerelement angezeigt werden."
Attribute PasswordChar.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    PasswordChar = Text1.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    Text1.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Text1,Text1,-1,RightToLeft
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Bestimmt die Ausrichtung des angezeigten Textes und steuert die Darstellung auf einem bidirektionalen System."
Attribute RightToLeft.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    RightToLeft = Text1.RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    Text1.RightToLeft() = New_RightToLeft
    PropertyChanged "RightToLeft"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Text1,Text1,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Bestimmt, ob ein Steuerelement bearbeitet werden kann."
Attribute Locked.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    Locked = Text1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Text1.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Text1,Text1,-1,HideSelection
'Public Property Get HideSelection() As Boolean
'    HideSelection = Text1.HideSelection
'End Property
'
'Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
'    Text1.HideSelection = New_HideSelection
'End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,0
Public Property Get DataFormat() As DataFormats
Attribute DataFormat.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    DataFormat = m_DataFormat
End Property

Public Property Let DataFormat(ByVal New_DataFormat As DataFormats)
    m_DataFormat = New_DataFormat
    PropertyChanged "DataFormat"
End Property

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        
        oldSelStart = Text1.SelStart
        oldSelLenght = Text1.SelLength
        
        Text1.Enabled = False
        cainPUMenu1.CreateMenu
    End If

End Sub

Private Sub UserControl_EnterFocus()
    On Error Resume Next
    Text1.SetFocus
End Sub

Private Sub UserControl_GotFocus()
    On Error Resume Next
    Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
    
    If AutoSelect = False Then Exit Sub
    
    If bMenuOpened = False Then
        On Error Resume Next
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
    End If
    
    bMenuOpened = False

End Sub

Private Sub UserControl_Initialize()
    
    cainPUMenu1.Top = -cainPUMenu1.Height - 10
    cainPUMenu1.Left = -cainPUMenu1.Width - 10
    
    cainPUMenu1.MenuItem.Add "a", m_textCut 'Ausschneiden
    cainPUMenu1.MenuItem.Add "b", m_textCopy 'Kopieren
    cainPUMenu1.MenuItem.Add "c", m_textPaste 'Einfügen
    cainPUMenu1.MenuItem.Add "d", m_textDelete 'Löschen
    cainPUMenu1.MenuItem.Add "e", , , , , , itPlaceholder 'Trenner
    cainPUMenu1.MenuItem.Add "f", m_textSelectAll 'Alles Markieren
    
    bMenuOpened = False

End Sub

Private Sub Create_Menu()

    cainPUMenu1.MenuItem("a").Caption = m_textCut  'Ausschneiden
    cainPUMenu1.MenuItem("b").Caption = m_textCopy 'Kopieren
    cainPUMenu1.MenuItem("c").Caption = m_textPaste 'Einfügen
    cainPUMenu1.MenuItem("d").Caption = m_textDelete 'Löschen
    cainPUMenu1.MenuItem("f").Caption = m_textSelectAll 'Alles Markieren

End Sub

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    m_AutoSelect = m_def_AutoSelect
    m_DataFormat = m_def_DataFormat
    m_SelectionColor = m_def_SelectionColor

    m_textCut = m_def_textCut
    m_textCopy = m_def_textCopy
    m_textPaste = m_def_textPaste
    m_textDelete = m_def_textDelete
    m_textSelectAll = m_def_textSelectAll
    
'    m_Title = m_def_Title
    UserControl_Resize
    
    Create_Menu
    
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Text1.Alignment = PropBag.ReadProperty("Alignment", 0)
    Text1.Text = PropBag.ReadProperty("Text", "Text1")
    m_AutoSelect = PropBag.ReadProperty("AutoSelect", m_def_AutoSelect)
    Text1.SelLength = PropBag.ReadProperty("SelLength", 0)
    Text1.SelStart = PropBag.ReadProperty("SelStart", 0)
    Text1.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    Text1.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    Text1.Locked = PropBag.ReadProperty("Locked", False)
    m_DataFormat = PropBag.ReadProperty("DataFormat", m_def_DataFormat)
    m_SelectionColor = PropBag.ReadProperty("SelectionColor", m_def_SelectionColor)
    m_textCut = PropBag.ReadProperty("textCut", m_def_textCut)
    m_textCopy = PropBag.ReadProperty("textCopy", m_def_textCopy)
    m_textPaste = PropBag.ReadProperty("textPaste", m_def_textPaste)
    m_textDelete = PropBag.ReadProperty("textDelete", m_def_textDelete)
    m_textSelectAll = PropBag.ReadProperty("textSelectAll", m_def_textSelectAll)
'    m_Title = PropBag.ReadProperty("Title", m_def_Title)
    Text1.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    
    SetFont
    Create_Menu
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    Text1.Left = 2
    Text1.Width = UserControl.ScaleWidth - 4
    Text1.Top = 2
    Text1.Height = UserControl.TextHeight("I")
    UserControl.Height = (Text1.Height + 4) * Screen.TwipsPerPixelY
    
    DrawFace

End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Alignment", Text1.Alignment, 0)
    Call PropBag.WriteProperty("Text", Text1.Text, "Text1")
    Call PropBag.WriteProperty("AutoSelect", m_AutoSelect, m_def_AutoSelect)
    Call PropBag.WriteProperty("SelLength", Text1.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", Text1.SelStart, 0)
    Call PropBag.WriteProperty("PasswordChar", Text1.PasswordChar, "")
    Call PropBag.WriteProperty("RightToLeft", Text1.RightToLeft, False)
    Call PropBag.WriteProperty("Locked", Text1.Locked, False)
    Call PropBag.WriteProperty("DataFormat", m_DataFormat, m_def_DataFormat)
    Call PropBag.WriteProperty("SelectionColor", m_SelectionColor, m_def_SelectionColor)
    Call PropBag.WriteProperty("textCut", m_textCut, m_def_textCut)
    Call PropBag.WriteProperty("textCopy", m_textCopy, m_def_textCopy)
    Call PropBag.WriteProperty("textPaste", m_textPaste, m_def_textPaste)
    Call PropBag.WriteProperty("textDelete", m_textDelete, m_def_textDelete)
    Call PropBag.WriteProperty("textSelectAll", m_textSelectAll, m_def_textSelectAll)
'    Call PropBag.WriteProperty("Title", m_Title, m_def_Title)
    Call PropBag.WriteProperty("MaxLength", Text1.MaxLength, 0)
    
    SetFont
    Create_Menu
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get SelectionColor() As OLE_COLOR
Attribute SelectionColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    SelectionColor = m_SelectionColor
End Property

Public Property Let SelectionColor(ByVal New_SelectionColor As OLE_COLOR)
    m_SelectionColor = New_SelectionColor
    PropertyChanged "SelectionColor"
    
    SetMenuColor
    
End Property

Private Sub SetMenuColor()
    
    cainPUMenu1.BackColor = m_BackColor
    cainPUMenu1.ForeColor = m_ForeColor
    cainPUMenu1.SelectionColor = m_SelectionColor
    cainPUMenu1.Font = UserControl.Font
    
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,Cut
Public Property Get textCut() As String
Attribute textCut.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    textCut = m_textCut
End Property

Public Property Let textCut(ByVal New_textCut As String)
    m_textCut = New_textCut
    PropertyChanged "textCut"
    Create_Menu
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,Copy
Public Property Get textCopy() As String
Attribute textCopy.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    textCopy = m_textCopy
End Property

Public Property Let textCopy(ByVal New_textCopy As String)
    m_textCopy = New_textCopy
    PropertyChanged "textCopy"
    Create_Menu
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,Paste
Public Property Get textPaste() As String
Attribute textPaste.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    textPaste = m_textPaste
End Property

Public Property Let textPaste(ByVal New_textPaste As String)
    m_textPaste = New_textPaste
    PropertyChanged "textPaste"
    Create_Menu
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,Delete
Public Property Get textDelete() As String
Attribute textDelete.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    textDelete = m_textDelete
End Property

Public Property Let textDelete(ByVal New_textDelete As String)
    m_textDelete = New_textDelete
    PropertyChanged "textDelete"
    Create_Menu
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,Select all
Public Property Get textSelectAll() As String
Attribute textSelectAll.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    textSelectAll = m_textSelectAll
End Property

Public Property Let textSelectAll(ByVal New_textSelectAll As String)
    m_textSelectAll = New_textSelectAll
    PropertyChanged "textSelectAll"
    Create_Menu
End Property
'
''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
''MemberInfo=13,0,0,
'Public Property Get Title() As String
'    Title = m_Title
'End Property
'
'Public Property Let Title(ByVal New_Title As String)
'    m_Title = New_Title
'    PropertyChanged "Title"
'    SetFont
'
'End Property
'
'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Text1,Text1,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Gibt die maximale Anzahl an Zeichen zurück, die in einem Steuerelement eingegeben werden können, oder legt diese fest."
    MaxLength = Text1.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    Text1.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

