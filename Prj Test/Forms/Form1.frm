VERSION 5.00
Object = "*\A..\..\Prj Active X Controls\XP Date Picker Ctl.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample XP DatePicker"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin OsenXPControls2.CommandButton CommandButton1 
      Height          =   405
      Left            =   2010
      TabIndex        =   16
      Top             =   2550
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Vote Me !!!"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   5790
      TabIndex        =   3
      Top             =   0
      Width           =   5790
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":058A
         Height          =   585
         Left            =   630
         TabIndex        =   5
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XP DatePicker Control"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   4
         Top             =   120
         Width           =   2130
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   4920
         Picture         =   "Form1.frx":0615
         Top             =   180
         Width           =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   30
         X2              =   5790
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   30
         X2              =   5790
         Y1              =   1020
         Y2              =   1020
      End
   End
   Begin OsenXPControls2.OsenXPDTPicker OsenXPDTPicker3 
      Height          =   315
      Left            =   4140
      TabIndex        =   2
      Top             =   1530
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Text            =   "14/07/2003"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   37816
      FormatDate      =   "dd/mm/yyyy"
   End
   Begin OsenXPControls2.OsenXPSpin OsenXPSpin1 
      Height          =   330
      Left            =   330
      TabIndex        =   1
      Top             =   4170
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Text            =   "0"
      Max             =   50
      Min             =   -50
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
   Begin OsenXPControls2.OsenXPDTPicker OsenXPDTPicker1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   1530
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      Text            =   "July 19,2003"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   37821
      FormatDate      =   "mmmm dd,yyyy"
   End
   Begin OsenXPControls2.OsenXPSpin OsenXPSpin2 
      Height          =   330
      Left            =   3150
      TabIndex        =   12
      Top             =   4140
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      Text            =   "0"
      Max             =   10
      Min             =   0
      LargeChange     =   0.25
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max: 10"
      Height          =   195
      Index           =   5
      Left            =   3180
      TabIndex        =   15
      Top             =   3660
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min: 0"
      Height          =   195
      Index           =   4
      Left            =   3180
      TabIndex        =   14
      Top             =   3870
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Large Change: 0.25"
      Height          =   195
      Index           =   3
      Left            =   3180
      TabIndex        =   13
      Top             =   3450
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max: 50"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   3690
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min: -50"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   3900
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Large Change: 1"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XP SPin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Format date: dd/mm/yyyy"
      Height          =   195
      Left            =   3570
      TabIndex        =   7
      Top             =   1260
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Format date: mmmm dd,yyyy"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   1260
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*********************************************
'/ Created Date : 2003/07/14
'/ Author       : Osen Kusnadi<okusnadi@cikarang.actaris.com>
'/**********************************************
'
' If you want to make your form layout like XP Visual Style
' please download at :  > http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=44773&lngWId=1
'                       > http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=46629&lngWId=1
'
' Dont forget to vote me !!!
'/**********************************************************************************************************
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub CommandButton1_Click()
    ShellExecute hwnd, "open", "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=46892&lngWId=1", vbNullString, vbNullString, conSwNormal
End Sub
