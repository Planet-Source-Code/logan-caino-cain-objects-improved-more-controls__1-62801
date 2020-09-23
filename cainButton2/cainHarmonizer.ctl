VERSION 5.00
Begin VB.UserControl cainHarmonizer 
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   870
   InvisibleAtRuntime=   -1  'True
   Picture         =   "cainHarmonizer.ctx":0000
   ScaleHeight     =   825
   ScaleWidth      =   870
End
Attribute VB_Name = "cainHarmonizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Standard-Eigenschaftswerte:
Const m_def_ParentBackColor = 0
Const m_def_BackColor = &HFFFFFF
Const m_def_ForeColor = &HC00000
Const m_def_SelectionColor = &H80FF&
Const m_def_TabSelectColor = &H96E7&
'Eigenschaftsvariablen:
Dim m_ParentBackColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Font As Font
Dim m_SelectionColor As OLE_COLOR
Dim m_TabSelectColor As OLE_COLOR

Public Function ID() As Integer
    ID = eControlIDs.id_Harmonizer
End Function

Public Sub Harmonize()

    Dim i As Integer
    Dim iID As Integer
    
    With UserControl.Parent
        
        For i = 1 To .Count - 1
            On Error Resume Next
            iID = .Controls(i).ID
            If Err = 0 Then
                
                Select Case iID
                
                Case id_Button, id_Slider
                    .Controls(i).BackColor = m_BackColor
                    .Controls(i).ForeColor = m_ForeColor
                    .Controls(i).TabSelectColor = m_TabSelectColor
                
                Case id_ColorPicker
                
                Case id_DropDown
                    .Controls(i).BackColor = m_BackColor
                    .Controls(i).ForeColor = m_ForeColor
                    .Controls(i).TabSelectColor = m_TabSelectColor
                    .Controls(i).SelectionColor = m_SelectionColor
                
                Case id_Harmonizer
                
                Case id_Label, id_Listbox, id_Monthview, id_PUMenu, id_Textbox, id_Toolbar
                    .Controls(i).BackColor = m_BackColor
                    .Controls(i).ForeColor = m_ForeColor
                    .Controls(i).SelectionColor = m_SelectionColor
                    
                Case id_Progressbar, id_Scrollbar, id_XButton
                    .Controls(i).BackColor = m_BackColor
                    .Controls(i).ForeColor = m_ForeColor
                    
                Case id_Tab
                    
                End Select
                
                Set .Controls(i).Font = m_Font
                
            End If
        Next i
    
    End With
    
    On Error Resume Next
    m_ParentBackColor = BlendColor(m_BackColor, m_ForeColor, 240)
    UserControl.Parent.BackColor = m_ParentBackColor

End Sub

Private Sub UserControl_Resize()
    
    UserControl.Height = 32 * Screen.TwipsPerPixelY
    UserControl.Width = 32 * Screen.TwipsPerPixelX
    
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
    
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Gibt ein Font-Objekt zurück."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Darstellung"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
    
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
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get TabSelectColor() As OLE_COLOR
Attribute TabSelectColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    TabSelectColor = m_TabSelectColor
End Property

Public Property Let TabSelectColor(ByVal New_TabSelectColor As OLE_COLOR)
    m_TabSelectColor = New_TabSelectColor
    PropertyChanged "TabSelectColor"
End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    m_SelectionColor = m_def_SelectionColor
    m_TabSelectColor = m_def_TabSelectColor
    m_ParentBackColor = m_def_ParentBackColor
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_SelectionColor = PropBag.ReadProperty("SelectionColor", m_def_SelectionColor)
    m_TabSelectColor = PropBag.ReadProperty("TabSelectColor", m_def_TabSelectColor)
    m_ParentBackColor = PropBag.ReadProperty("ParentBackColor", m_def_ParentBackColor)
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("SelectionColor", m_SelectionColor, m_def_SelectionColor)
    Call PropBag.WriteProperty("TabSelectColor", m_TabSelectColor, m_def_TabSelectColor)
    Call PropBag.WriteProperty("ParentBackColor", m_ParentBackColor, m_def_ParentBackColor)
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,1,1,0
Public Property Get ParentBackColor() As OLE_COLOR
Attribute ParentBackColor.VB_MemberFlags = "400"
    ParentBackColor = m_ParentBackColor
End Property

Public Property Let ParentBackColor(ByVal New_ParentBackColor As OLE_COLOR)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_ParentBackColor = New_ParentBackColor
    PropertyChanged "ParentBackColor"
End Property

