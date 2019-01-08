VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login Page"
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
      Caption         =   "Staff Login"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.CommandButton cmdLogin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4920
         Width           =   2100
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4920
         Width           =   2100
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         IMEMode         =   3  'DISABLE
         Left            =   4560
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   3000
         Width           =   3075
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4560
         TabIndex        =   1
         Top             =   1800
         Width           =   3075
      End
      Begin VB.Label lblLogin 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "KINDLY ENTER USERNAME AND PASSWORD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   6615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   4
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   1215
         Left            =   1560
         Shape           =   4  'Rounded Rectangle
         Top             =   4560
         Width           =   7095
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   2655
         Left            =   1560
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean

Private Sub cmdExit_Click()
Dim res As Integer
res = MsgBox("Do you want to cancel", vbInformation + vbYesNo, "Exit")
If res = vbYes Then
End
Else
End If
End Sub

Private Sub cmdLogin_Click()
'check for correct password

    If txtUserName.Text = "Admin" And txtPassword.Text = "Service2018" Then
       LoginSucceeded = True
       Me.Hide
       frmMain.Show
    Else
        MsgBox "Username or Password did not match. Try Again!", , "Login"
        txtPassword.SetFocus
        
    End If
End Sub

