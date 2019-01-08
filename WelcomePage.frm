VERSION 5.00
Begin VB.Form frmSplashScreen 
   Caption         =   "Welcome"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Height          =   14500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15570
      Begin VB.Timer Timer2 
         Left            =   840
         Top             =   9360
      End
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   120
         Top             =   7800
      End
      Begin VB.Timer Timer3 
         Interval        =   3000
         Left            =   840
         Top             =   7920
      End
      Begin VB.Image Image2 
         Height          =   3255
         Left            =   9720
         Picture         =   "WelcomePage.frx":0000
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   5055
      End
      Begin VB.Image Image1 
         Height          =   3255
         Left            =   1560
         Picture         =   "WelcomePage.frx":30BB
         Stretch         =   -1  'True
         Top             =   4920
         Width           =   8295
      End
      Begin VB.Label dtt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   6120
         Width           =   4575
      End
      Begin VB.Label lblHotelName 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "WELCOME TO ROCKET SEARCH INT'L LTD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   555
         Left            =   3000
         TabIndex        =   3
         Top             =   2040
         Width           =   9930
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   11640
         TabIndex        =   2
         Top             =   8520
         Width           =   1275
      End
      Begin VB.Label LblProjectname 
         BackColor       =   &H00800000&
         Caption         =   "    SYSTEM  MANAGEMENT SOFTWARE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   2160
         TabIndex        =   1
         Top             =   3000
         Width           =   12015
      End
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
'Unload Me
End Sub


Private Sub Timer1_Timer()
Timer1.Enabled = False
    frmLogin.Show
    Unload Me
End Sub
