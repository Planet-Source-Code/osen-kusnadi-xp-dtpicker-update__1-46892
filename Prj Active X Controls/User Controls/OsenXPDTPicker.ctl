VERSION 5.00
Begin VB.UserControl OsenXPDTPicker 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3195
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   360
   ScaleWidth      =   3195
   ToolboxBitmap   =   "OsenXPDTPicker.ctx":0000
   Begin VB.PictureBox ImgUp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      Picture         =   "OsenXPDTPicker.ctx":0312
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   30
      Width           =   255
   End
   Begin VB.TextBox TxtSpin 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Text            =   "0"
      Top             =   60
      Width           =   720
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   3
      Left            =   690
      Picture         =   "OsenXPDTPicker.ctx":06C8
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   2
      Left            =   720
      Picture         =   "OsenXPDTPicker.ctx":0A7E
      Top             =   1710
      Width           =   255
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   1
      Left            =   750
      Picture         =   "OsenXPDTPicker.ctx":0E34
      Top             =   1410
      Width           =   255
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   0
      Left            =   750
      Picture         =   "OsenXPDTPicker.ctx":11EA
      Top             =   1140
      Width           =   255
   End
   Begin VB.Shape ShapeBorder 
      BorderColor     =   &H00B99D7F&
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   1125
   End
End
Attribute VB_Name = "OsenXPDTPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINTAPI
        X As Long
        Y As Long
End Type


'Default Property Values:
Const m_def_Value = 0
Const m_def_FormatDate = "yyyy/mm/dd"
Dim m_Value As Long
Dim m_FormatDate As String
Dim m_ToolTipText As String
Dim m_Font As Font
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event Change() 'MappingInfo=TxtSpin,TxtSpin,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."

Private Sub RePos()
Dim I As Integer
    If Width < 400 Then Width = 400
    ShapeBorder.Width = Width
    ImgUp.Left = Width - 285
    TxtSpin.Width = Width - 375
    Height = 315
    TxtSpin.Top = 60
    If TxtSpin.FontSize > 8 Then
        I = TxtSpin.FontSize - 8
        I = I * 15
        TxtSpin.Top = TxtSpin.Top - I
    End If
End Sub

Sub ResetPic()
    If ImgUp.Picture <> Img(0).Picture Then
        ImgUp.Picture = Img(0).Picture
    End If
End Sub

Private Sub ImgUp_Click()
On Error GoTo ErrXYZ
    Dim MyApi As POINTAPI
    Dim MyRect As RECT
    Dim MyLeft As Long, MyTop As Long
    
    GetCursorPos MyApi
    GetWindowRect ImgUp.hWnd, MyRect
    
    MyTop = (MyRect.Top * 15) + ImgUp.Height + 45
    If Screen.Height - MyTop < 3000 Then
        MyTop = MyTop - 2835 - 30 - Height
    End If
    
    MyLeft = (MyRect.Left * 15) - ImgUp.Left
    If Screen.Width - MyLeft < 2900 Then
        MyLeft = (MyRect.Left * 15) + ImgUp.Width + 45 - 2760
    End If
    
    If Value = 0 Then Value = Val(Format(Date, "#"))
    
    ResultDate = Value
    DoEvents
    
    With MyForm
        .InitValue = Value
        DoEvents
        .Left = MyLeft
        .Top = MyTop
        .Height = 2835
        .Width = 2745
        .Show 1
    End With
    
    ResetPic
    Value = ResultDate
ErrXYZ:
End Sub

Private Sub ImgUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgUp.Picture = Img(2).Picture
    ImgUp_Click
End Sub

Private Sub ImgUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ImgUp.Picture <> Img(1).Picture Then ImgUp.Picture = Img(1).Picture
End Sub

Private Sub TxtSpin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetPic
End Sub


Private Sub UserControl_InitProperties()
    BackColor = vbWhite
    Set m_Font = Ambient.Font
    m_Value = m_def_Value
    m_FormatDate = m_def_FormatDate
    Value = Format(Date, "#")
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetPic
End Sub

Private Sub UserControl_Resize()
    RePos
    
End Sub

Private Sub TxtSpin_Change()
    RaiseEvent Change
End Sub
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    TxtSpin.Enabled = New_Enabled
    ImgUp.Enabled = New_Enabled
    If New_Enabled = False Then
        ImgUp.Picture = Img(3).Picture
        ShapeBorder.BorderColor = &HC0C0C0
    Else
        ResetPic
        ShapeBorder.BorderColor = &HB99D7F
    End If
    PropertyChanged "Enabled"
End Property
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = TxtSpin.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    TxtSpin.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = TxtSpin.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    TxtSpin.Text() = New_Text
    PropertyChanged "Text"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    TxtSpin.Locked = PropBag.ReadProperty("Locked", False)
    TxtSpin.Text = PropBag.ReadProperty("Text", "Text1")
    Set TxtSpin.Font = PropBag.ReadProperty("Font", Ambient.Font)
    

    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_FormatDate = PropBag.ReadProperty("FormatDate", m_def_FormatDate)
    
    RePos

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    With PropBag
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("Locked", TxtSpin.Locked, False)
        Call .WriteProperty("Text", TxtSpin.Text, "Text1")
        Call .WriteProperty("Font", TxtSpin.Font, Ambient.Font)
        Call .WriteProperty("Value", m_Value, m_def_Value)
        Call .WriteProperty("FormatDate", m_FormatDate, m_def_FormatDate)
    End With
    
End Sub
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = TxtSpin.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    If New_Font.Size > 10 Then New_Font.Size = 10
    Set TxtSpin.Font = New_Font
    PropertyChanged "Font"
    RePos
End Property
Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    m_Value = New_Value
    Text = Format(m_Value, FormatDate)
    PropertyChanged "Value"
End Property
Public Property Get FormatDate() As String
    FormatDate = m_FormatDate
End Property

Public Property Let FormatDate(ByVal New_FormatDate As String)
    m_FormatDate = New_FormatDate
    Value = Value
    PropertyChanged "FormatDate"
End Property
