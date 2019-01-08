VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Menu"
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
      Caption         =   "MAKE YOUR SELECTION"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.OptionButton optBankDeposit 
         BackColor       =   &H00800000&
         Caption         =   "Bank Deposit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   500
         Left            =   2040
         TabIndex        =   7
         Top             =   3360
         Width           =   3315
      End
      Begin VB.OptionButton optExit 
         BackColor       =   &H00800000&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   4560
         Width           =   2055
      End
      Begin VB.OptionButton optOthers 
         BackColor       =   &H00800000&
         Caption         =   "Other Sales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   500
         Left            =   2040
         TabIndex        =   4
         Top             =   3960
         Width           =   3315
      End
      Begin VB.OptionButton optStartimes 
         BackColor       =   &H00800000&
         Caption         =   "Startimes Bouquets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   500
         Left            =   2040
         TabIndex        =   3
         Top             =   2880
         Width           =   3195
      End
      Begin VB.OptionButton optGoTV 
         BackColor       =   &H00800000&
         Caption         =   "GoTV Package"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   500
         Left            =   2040
         TabIndex        =   2
         Top             =   2280
         Width           =   2715
      End
      Begin VB.OptionButton optDSTV 
         BackColor       =   &H00800000&
         Caption         =   "DSTV Package"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   500
         Left            =   2040
         TabIndex        =   1
         Top             =   1560
         Width           =   3435
      End
      Begin VB.Image Image2 
         Height          =   2895
         Left            =   9120
         Picture         =   "Main_menu.frx":0000
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   " SELECT  TYPE OF SERVICE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   3840
         TabIndex        =   5
         Top             =   960
         Width           =   5655
      End
      Begin VB.Image Image1 
         Height          =   2895
         Left            =   720
         Picture         =   "Main_menu.frx":30BB
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   8415
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub optBankDeposit_Click()
Me.Hide
frmBankDeposit.Show
End Sub

Private Sub optDSTV_Click()
Me.Hide
frmDSTV.Show
End Sub

Private Sub optExit_Click()
Dim res As Integer
res = MsgBox("Do you want to cancel", vbInformation + vbYesNo, "Exit")
If res = vbYes Then
End
Else
End If
End Sub

Private Sub optGoTV_Click()
Me.Hide
frmGoTV.Show
End Sub

Private Sub optOthers_Click()
Me.Hide
frmOthers.Show
End Sub

Private Sub optStartimes_Click()
Me.Hide
frmStartimes.Show
End Sub
