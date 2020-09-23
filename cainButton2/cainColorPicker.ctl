VERSION 5.00
Begin VB.UserControl cainColorPicker 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "cainColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim selX As Single
Dim selY As Single
'Standard-Eigenschaftswerte:
Const m_def_Color = &H96E7&
'Eigenschaftsvariablen:
Dim m_Color As OLE_COLOR
'Ereignisdeklarationen:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste drückt, während ein Objekt den Fokus hat."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Tritt auf, wenn der Benutzer die Maus bewegt."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste losläßt, während ein Objekt den Fokus hat."

Dim lSelectedcolor As Long

Public Function ID() As Integer
    ID = eControlIDs.id_ColorPicker
End Function

Private Sub DrawFace()
    
    Dim RGBcols As RGB
    
    RGBcols = SetColor(lSelectedcolor)
    
    UserControl.Cls
    UserControl.BackColor = vbWhite

    GradientCy UserControl.hdc, 0, 0, 10, UserControl.ScaleHeight, vbWhite, BlendColor(vbWhite, vbBlack), vbBlack, pbHorizontal
    GradientCy UserControl.hdc, 10, 0, 10, UserControl.ScaleHeight, vbWhite, vbRed, vbBlack, pbHorizontal
    GradientCy UserControl.hdc, 20, 0, 10, UserControl.ScaleHeight, vbWhite, BlendColor(vbRed, vbYellow), vbBlack, pbHorizontal
    GradientCy UserControl.hdc, 30, 0, 10, UserControl.ScaleHeight, vbWhite, vbYellow, vbBlack, pbHorizontal
    GradientCy UserControl.hdc, 40, 0, 10, UserControl.ScaleHeight, vbWhite, vbGreen, vbBlack, pbHorizontal
    GradientCy UserControl.hdc, 50, 0, 10, UserControl.ScaleHeight, vbWhite, BlendColor(vbGreen, vbBlue), vbBlack, pbHorizontal
    GradientCy UserControl.hdc, 60, 0, 10, UserControl.ScaleHeight, vbWhite, vbBlue, vbBlack, pbHorizontal
    GradientCy UserControl.hdc, 70, 0, 10, UserControl.ScaleHeight, vbWhite, BlendColor(vbBlue, vbRed), vbBlack, pbHorizontal
    
    UserControl.Line (selX - 3, selY - 3)-(selX + 4, selY - 3), vbBlack
    UserControl.Line (selX - 3, selY - 3)-(selX - 3, selY + 3), vbBlack
    UserControl.Line (selX - 3, selY + 3)-(selX + 3, selY + 3), vbBlack
    UserControl.Line (selX + 3, selY + 3)-(selX + 3, selY - 3), vbBlack
    
    UserControl.Line (selX - 2, selY - 2)-(selX + 3, selY - 2), vbWhite
    UserControl.Line (selX - 2, selY - 2)-(selX - 2, selY + 2), vbWhite
    UserControl.Line (selX - 2, selY + 2)-(selX + 2, selY + 2), vbWhite
    UserControl.Line (selX + 2, selY + 2)-(selX + 2, selY - 2), vbWhite
    
    UserControl.Line (90, 10)-(111, 10), vbBlack
    UserControl.Line (110, 30)-(110, 10), vbBlack
    UserControl.Line (110, 30)-(90, 30), vbBlack
    UserControl.Line (90, 30)-(90, 10), vbBlack
    
    UserControl.Line (90, 40)-(111, 40), vbBlack
    UserControl.Line (110, 60)-(110, 40), vbBlack
    UserControl.Line (110, 60)-(90, 60), vbBlack
    UserControl.Line (90, 60)-(90, 40), vbBlack
    
    DrawGradient UserControl.hdc, 91, 11, 19, 19, GetRGBColors(lSelectedcolor), GetRGBColors(lSelectedcolor)
    DrawGradient UserControl.hdc, 91, 41, 19, 19, GetRGBColors(m_Color), GetRGBColors(m_Color)
    
    UserControl.ForeColor = vbBlack
    
    UserControl.CurrentX = 115
    UserControl.CurrentY = 10
    UserControl.Print "R: " & RGBcols.R
    
    UserControl.CurrentX = 115
    UserControl.CurrentY = 25
    UserControl.Print "G: " & RGBcols.G
    
    UserControl.CurrentX = 115
    UserControl.CurrentY = 40
    UserControl.Print "B: " & RGBcols.B
    
    UserControl.CurrentX = 90
    UserControl.CurrentY = 65
    UserControl.Print HexEx(lSelectedcolor)
    
End Sub

Private Sub UserControl_Initialize()
    selX = 3
    selY = 3
    lSelectedcolor = vbWhite
    DrawFace
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If Button <> 1 Or X < 3 Or Y < 3 Or Y > (UserControl.ScaleHeight - 4) Or X > (UserControl.ScaleWidth - 4) Then Exit Sub
    selX = X
    selY = Y
    
    lSelectedcolor = UserControl.Point(selX, selY)
    
    DrawFace
    
End Sub

Private Sub UserControl_Resize()
    DrawFace
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get Color() As OLE_COLOR
    Color = m_Color
End Property

Public Property Let Color(ByVal New_Color As OLE_COLOR)
    m_Color = New_Color
    PropertyChanged "Color"
    
    DrawFace
End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_Color = m_def_Color
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Color = PropBag.ReadProperty("Color", m_def_Color)
    
    DrawFace
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Color", m_Color, m_def_Color)
    
    DrawFace
End Sub

