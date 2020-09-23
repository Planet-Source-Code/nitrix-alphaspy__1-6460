VERSION 4.00
Begin VB.Form frmColor 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3255
   ClientLeft      =   1230
   ClientTop       =   1995
   ClientWidth     =   1725
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   3660
   Left            =   1170
   LinkTopic       =   "Form1"
   Picture         =   "frmColor.frx":0000
   ScaleHeight     =   3255
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   Top             =   1650
   Width           =   1845
   Begin VB.Frame Frame4 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   0.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Left            =   60
      TabIndex        =   20
      Top             =   2550
      Width           =   1635
      Begin VB.Label Label2 
         Caption         =   "-"
         Height          =   225
         Left            =   780
         TabIndex        =   24
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         Height          =   225
         Left            =   780
         TabIndex        =   23
         Top             =   60
         Width           =   795
      End
      Begin VB.Label Label14 
         Caption         =   "vb color -"
         Height          =   225
         Left            =   90
         TabIndex        =   22
         Top             =   60
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "qb color -"
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   0.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   885
      Left            =   60
      TabIndex        =   11
      Top             =   1560
      Width           =   1605
      Begin VB.Label Label18 
         Caption         =   "-"
         Height          =   165
         Left            =   900
         TabIndex        =   19
         Top             =   450
         Width           =   675
      End
      Begin VB.Label Label19 
         Caption         =   "-"
         Height          =   195
         Left            =   900
         TabIndex        =   18
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label20 
         Caption         =   "-"
         Height          =   195
         Left            =   900
         TabIndex        =   17
         Top             =   630
         Width           =   675
      End
      Begin VB.Label Label7 
         Caption         =   "-"
         Height          =   195
         Left            =   720
         TabIndex        =   16
         Top             =   60
         Width           =   765
      End
      Begin VB.Label Label9 
         Caption         =   "r-value -"
         Height          =   225
         Left            =   210
         TabIndex        =   15
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "g-value -"
         Height          =   225
         Left            =   180
         TabIndex        =   14
         Top             =   450
         Width           =   795
      End
      Begin VB.Label Label17 
         Caption         =   "b-value -"
         Height          =   225
         Left            =   180
         TabIndex        =   13
         Top             =   630
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "rgb -  #"
         Height          =   225
         Left            =   150
         TabIndex        =   12
         Top             =   60
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   0.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   60
      TabIndex        =   6
      Top             =   960
      Width           =   1605
      Begin VB.Label Label6 
         Caption         =   "-"
         Height          =   225
         Left            =   930
         TabIndex        =   10
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         Height          =   225
         Left            =   930
         TabIndex        =   9
         Top             =   60
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "decimal -"
         Height          =   225
         Left            =   90
         TabIndex        =   8
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "hex -"
         Height          =   225
         Left            =   360
         TabIndex        =   7
         Top             =   270
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   0.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   1125
      Begin VB.Label Label13 
         Caption         =   "y pos -"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label12 
         Caption         =   "x pos -"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   60
         Width           =   525
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "-"
         Height          =   225
         Left            =   690
         TabIndex        =   3
         Top             =   60
         Width           =   345
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "-"
         Height          =   255
         Left            =   660
         TabIndex        =   2
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   1260
      ScaleHeight     =   465
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   360
      Width           =   405
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3360
      Top             =   660
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   1440
      TabIndex        =   27
      Top             =   120
      Width           =   165
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   1320
      TabIndex        =   26
      Top             =   150
      Width           =   165
   End
   Begin VB.Shape Shape1 
      Height          =   225
      Left            =   60
      Top             =   90
      Width           =   1575
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   1770
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   1710
      X2              =   1710
      Y1              =   3630
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   1860
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3630
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C0C0&
      Caption         =   " alphaSpy             -  x"
      ForeColor       =   &H00F86800&
      Height          =   225
      Left            =   60
      TabIndex        =   25
      Top             =   90
      Width           =   1575
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
form_center Me
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
form_drag Me
End Sub


Private Sub Label22_Click()
Me.WindowState = 1
End Sub

Private Sub Label23_Click()
End
End Sub


Private Sub Timer1_Timer()
'put color in picture box
Picture1.BackColor = spy_color
'decimal color
Label1.Caption = spy_color
'qbasic color
Label2.Caption = spy_colorqb
'vb color
Label3.Caption = spy_colorvb
'cursor positions
Label5.Caption = spy_cursory
Label4.Caption = spy_cursorx
'convert decimal to hex
Label6.Caption = Hex(CDbl(spy_color))
'convert hex to rgb color
Label7.Caption = hex_2rgb(Label6.Caption)
'add the necessary zeros to the rgb number
Label7.Caption = hex_fill(Label7.Caption)
rgb_get (spy_color)
End Sub


