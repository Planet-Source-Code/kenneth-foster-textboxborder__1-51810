VERSION 5.00
Begin VB.UserControl TxtBoxBorder 
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   ControlContainer=   -1  'True
   ScaleHeight     =   645
   ScaleWidth      =   615
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   405
      Width           =   240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   330
      Left            =   0
      Top             =   15
      Width           =   525
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "TxtBoxBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Property Get FocusColor() As OLE_COLOR
    
    FocusColor = Shape1.BorderColor
    
End Property

Public Property Let FocusColor(ByVal NewFocusColor As OLE_COLOR)

    Shape1.BorderColor = NewFocusColor
    PropertyChanged "FocusColor"
    
End Property
Public Property Get NonFocusColor() As OLE_COLOR

    NonFocusColor = Shape2.BorderColor
    
End Property
Public Property Let NonFocusColor(ByVal NewNonFocusColor As OLE_COLOR)
    Shape2.BorderColor = NewNonFocusColor
    PropertyChanged "NonFocusColor"
End Property

Public Property Get Locked() As Boolean
     Locked = Text1.Locked
End Property
Public Property Let Locked(ByVal NewLocked As Boolean)
     Text1.Locked = NewLocked
     PropertyChanged "Locked"
End Property
Public Property Get Text() As String
      Text = Text1.Text
End Property
Public Property Let Text(ByVal NewText As String)
    Text1.Text = NewText
    PropertyChanged "Text"
End Property
Private Sub Text1_GotFocus()
   Shape1.Visible = True
   Shape2.Visible = False
End Sub

Private Sub Text1_LostFocus()
   Shape1.Visible = False
   Shape2.Visible = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Shape1.BorderColor = PropBag.ReadProperty("FocusColor", vbBlue)
    Shape2.BorderColor = PropBag.ReadProperty("NonFocusColor", vbRed)
    Text1.Text = PropBag.ReadProperty("Text", "Text")
    Text1.Locked = PropBag.ReadProperty("Locked", False)
     Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Text1.FontSize = PropBag.ReadProperty("FontSize", 8.25)
    Text1.FontBold = PropBag.ReadProperty("FontBold", 0)
End Sub

Private Sub UserControl_Resize()
    
    Shape1.Left = 0
    Shape1.Top = 0
    Shape1.Width = UserControl.Width
    Shape1.Height = UserControl.Height
    Shape2.Left = 0
    Shape2.Top = 0
    Shape2.Width = UserControl.Width
    Shape2.Height = UserControl.Height
    Text1.Left = Shape1.Left + 10
    Text1.Top = Shape1.Top + 10
    Text1.Width = UserControl.Width - 25
    Text1.Height = UserControl.Height - 25
    
End Sub

Private Sub UserControl_Show()
    
    On Error GoTo ResizeErr
    
    If UserControl.ContainedControls.Count > 0 Then
    
        With UserControl.ContainedControls.Item(0)
            .Top = 10
            .Left = 10
            .Height = UserControl.Height - 25
            .Width = UserControl.Width - 25
        End With

    End If
    
ResizeErr:
    Exit Sub

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    PropBag.WriteProperty "FocusColor", Shape1.BorderColor, vbBlue
    PropBag.WriteProperty "NonFocusColor", Shape2.BorderColor, vbRed
    PropBag.WriteProperty "Locked", Text1.Locked, False
    PropBag.WriteProperty "Text", Text1.Text, "Text"
    Call PropBag.WriteProperty("FontSize", Text1.FontSize, 0)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", Text1.FontBold, 0)
    
End Sub
Public Property Get Font() As Font
    Set Font = Text1.Font
End Property
    
Public Property Set Font(ByVal New_Font As Font)
    Set Text1.Font = New_Font
    PropertyChanged "Font"
End Property
    
Public Property Get FontSize() As Single
    FontSize = Text1.FontSize
    UserControl_Resize
End Property
    
Public Property Let FontSize(ByVal New_FontSize As Single)
    Text1.FontSize = New_FontSize
    PropertyChanged "FontSize"
End Property
    
Public Property Get FontBold() As Boolean
    FontBold = Text1.FontBold
End Property
    
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Text1.FontBold = New_FontBold
    PropertyChanged "FontBold"
End Property
