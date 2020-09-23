VERSION 5.00
Begin VB.UserControl SuperLabel 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1560
   PropertyPages   =   "SuperLabel.ctx":0000
   ScaleHeight     =   615
   ScaleWidth      =   1560
   ToolboxBitmap   =   "SuperLabel.ctx":002D
   Begin VB.Label L1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SuperLabel1"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   900
   End
   Begin VB.Label L2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SuperLabel1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "SuperLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Dim Rastoqnie As SLFont3DConstants
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "  Occurs when the user presses and then releases a mouse button over an object.\n"
Event Change()
Attribute Change.VB_Description = " Occurs when the contents of a control have changed.\n"
Event DblClick()
Attribute DblClick.VB_Description = " Occurs when the user presses and releases a mouse button and then presses and releases it again over an object.\n"
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = " Occurs when the user presses the mouse button while an object has the focus.\n"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse.\n"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = " Occurs when the user releases the mouse button while an object has the focus.\n"

Public Enum SLBorderStyleConstants
   None = 0
   FixedSingle = 1
End Enum

Public Enum SLBackStyleConstants
   Transparent = 0
   Opaque = 1
End Enum

Public Enum SLFont3DConstants
   None3D = 0
   Raised3D = 1
   Insert3D = 2
End Enum

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = " Returns/sets the background color used to display text and graphics in an object.\n"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    L2.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = " Returns/sets the foreground color used to display text and graphics in an object.\n"
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "  Returns/sets a value that determines whether an object can respond to user-generated events.\n"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    L1.Enabled = New_Enabled
    L2.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object.\n"
    Set Font = L1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set L1.Font = New_Font
    Set L2.Font = New_Font
    PropertyChanged "Font"
    Call UserControl_Resize
End Property

Public Property Get Font3D() As SLFont3DConstants
Attribute Font3D.VB_Description = " Returns/sets whether or not an object is painted at run time with 3-D effects.\n"
    Font3D = Rastoqnie
End Property

Public Property Let Font3D(New_Font3D As SLFont3DConstants)
   Rastoqnie = New_Font3D
   Call RePosition
   PropertyChanged "Font3D"
End Property

Public Property Get BackStyle() As SLBackStyleConstants
Attribute BackStyle.VB_Description = " Indicates whether a Label or the background of a Shape is transparent or opaque.\n"
    BackStyle = L2.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As SLBackStyleConstants)
'Opa!!! It's a little problem hire
'When UserControl.BackStyle = Transparent the UserControl it's not visible any more
'Sorry!
    UserControl.BackStyle = New_BackStyle
    L2.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As SLBorderStyleConstants
Attribute BorderStyle.VB_Description = " Returns/sets the border style for an object.\n"
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As SLBorderStyleConstants)
   ' L2.BorderStyle() = New_BorderStyle
    UserControl.BorderStyle() = New_BorderStyle
    Call UserControl_Resize
    PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object.\n"
    UserControl.Refresh
    L1.Refresh
    L2.Refresh
End Sub

Private Sub L1_Change()
   RaiseEvent Change
   Call UserControl_Resize
End Sub

Private Sub L1_Click()
   RaiseEvent Click
End Sub

Private Sub L1_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub L1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub L1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub L1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub L2_Change()
   RaiseEvent Change
End Sub

Private Sub L2_Click()
   RaiseEvent Click
End Sub

Private Sub L2_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub L2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub L2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub L2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
   Call SuperLabelResize
   Rastoqnie = Raised3D
   Font3D = Raised3D
   BackStyle = Opaque
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = " Determines whether a control is automatically resized to display its entire contents.\n"
    AutoSize = L1.AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    L1.AutoSize() = New_AutoSize
    L2.AutoSize() = New_AutoSize
    Call UserControl_Resize
    Call RePosition
    PropertyChanged "AutoSize"
End Property

Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Return/sets alignment of control text."
    Alignment = L1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    L1.Alignment() = New_Alignment
    L2.Alignment() = New_Alignment
    Call RePosition
    PropertyChanged "Alignment"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon.\n"
    Caption = L1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    L1.Caption() = New_Caption
    L2.Caption() = New_Caption
    PropertyChanged "Caption"
    Call UserControl_Resize
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = " Returns/sets bold font styles.\n"
    FontBold = L1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    L1.FontBold() = New_FontBold
    L2.FontBold() = New_FontBold
    PropertyChanged "FontBold"
    Call UserControl_Resize
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = " Returns/sets italic font styles.\n"
    FontItalic = L1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    L1.FontItalic() = New_FontItalic
    L2.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
    Call UserControl_Resize
End Property

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level.\n"
    FontName = L1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    L1.FontName() = New_FontName
    L2.FontName() = New_FontName
    PropertyChanged "FontName"
    Call UserControl_Resize
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level.\n"
    FontSize = L1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    L1.FontSize() = New_FontSize
    L2.FontSize() = New_FontSize
    PropertyChanged "FontSize"
    Call UserControl_Resize
End Property

Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles.\n"
    FontStrikethru = L1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    L1.FontStrikethru() = New_FontStrikethru
    L2.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles.\n"
    FontUnderline = L1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    L1.FontUnderline() = New_FontUnderline
    L2.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = " Returns/sets a value that determines whether a control expands to fit the text in its Caption.\n"
    WordWrap = L1.WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    L1.WordWrap() = New_WordWrap
    L2.WordWrap() = New_WordWrap
    PropertyChanged "WordWrap"
End Property

Public Property Get FRWColor() As OLE_COLOR
Attribute FRWColor.VB_Description = " Returns/sets the foreground color used to display text and graphics in an object.\n"
    FRWColor = L1.ForeColor
End Property

Public Property Let FRWColor(ByVal New_FRWColor As OLE_COLOR)
    L1.ForeColor = New_FRWColor
    PropertyChanged "FRWColor"
End Property

Public Property Get BGColor() As OLE_COLOR
Attribute BGColor.VB_Description = " Returns/sets the background color used to display text and graphics in an object.\n"
    BGColor = L2.ForeColor
End Property

Public Property Let BGColor(ByVal New_BGColor As OLE_COLOR)
    L2.ForeColor() = New_BGColor
    PropertyChanged "BGColor"
End Property

Private Sub SuperLabelResize()
    L2.Move 0, 0
    L1.Move 20, 20
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set L1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set L2.Font = L1.Font
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    L1.AutoSize = PropBag.ReadProperty("AutoSize", True)
    L2.AutoSize = L1.AutoSize
    L1.Alignment = PropBag.ReadProperty("Alignment", 0)
    L2.Alignment = L1.Alignment
    L1.Caption = PropBag.ReadProperty("Caption", "Label1")
    L2.Caption = L1.Caption
    L1.FontBold = PropBag.ReadProperty("FontBold", 0)
    L2.FontBold = L1.FontBold
    L1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    L2.FontItalic = L1.FontItalic
    L1.FontName = PropBag.ReadProperty("FontName", "")
    L2.FontName = L1.FontName
    L1.FontSize = PropBag.ReadProperty("FontSize", 0)
    L2.FontSize = L1.FontSize
    L1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    L2.FontStrikethru = L1.FontStrikethru
    L1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    L2.FontUnderline = L1.FontUnderline
    L1.WordWrap = PropBag.ReadProperty("WordWrap", False)
    L2.WordWrap = L1.WordWrap
    L1.ForeColor = PropBag.ReadProperty("FRWColor", &H80000012)
    L2.ForeColor = PropBag.ReadProperty("BGColor", &HFFFFFF)
    Rastoqnie = PropBag.ReadProperty("Font3D", Rastoqnie)
End Sub

Private Sub UserControl_Resize()
    If L1.AutoSize = True Then
       If UserControl.BorderStyle = vbFixedSingle Then
          UserControl.Width = L1.Width + 70
          UserControl.Height = L1.Height + 70
       Else
          UserControl.Width = L1.Width + 20
          UserControl.Height = L1.Height + 20
       End If
    Else
       L1.Width = UserControl.Width
       L1.Height = UserControl.Height
       L2.Width = UserControl.Width
       L2.Height = UserControl.Height
    End If
    UserControl.Refresh
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", L1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("AutoSize", L1.AutoSize, True)
    Call PropBag.WriteProperty("Alignment", L1.Alignment, 0)
    Call PropBag.WriteProperty("Caption", L1.Caption, "Label1")
    Call PropBag.WriteProperty("FontBold", L1.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", L1.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", L1.FontName, "")
    Call PropBag.WriteProperty("FontSize", L1.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", L1.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", L1.FontUnderline, 0)
    Call PropBag.WriteProperty("WordWrap", L1.WordWrap, False)
    Call PropBag.WriteProperty("FRWColor", L1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BGColor", L2.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("Font3D", Rastoqnie, 1)
End Sub

Private Sub RePosition()
    'Sometimes we need to Re-Position the labels
    Select Case Rastoqnie
          Case 0: L1.Move 0, 0
                  L2.Move 0, 0
          Case 1: L1.Move 20, 20
                  L2.Move 0, 0
          Case 2: L1.Move 0, 0
                  L2.Move 20, 20
    End Select
End Sub
