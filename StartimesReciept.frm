VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOtherReciept 
   Caption         =   "Other Reciept"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reciept"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   5520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblHotelAddress 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"StartimesReciept.frx":0000
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   3855
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "2 ROCKET SEARCH INT'L LTD.  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Amount 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   225
         Left            =   1440
         TabIndex        =   14
         Top             =   3960
         Width           =   2040
      End
      Begin VB.Label ServiceName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   225
         Left            =   1080
         TabIndex        =   13
         Top             =   3600
         Width           =   3225
         WordWrap        =   -1  'True
      End
      Begin VB.Label CustomerName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   240
         Left            =   1080
         TabIndex        =   12
         Top             =   2520
         Width           =   3585
         WordWrap        =   -1  'True
      End
      Begin VB.Label DatePaid 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   225
         Left            =   1320
         TabIndex        =   11
         Top             =   2880
         Width           =   2280
      End
      Begin VB.Label CustomerPhoneno 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   225
         Left            =   1200
         TabIndex        =   10
         Top             =   3240
         Width           =   2040
      End
      Begin VB.Label lblAmount 
         BackColor       =   &H8000000E&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label lblDescription 
         BackColor       =   &H8000000E&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label lblSerialNumber 
         BackColor       =   &H8000000E&
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblTime 
         BackColor       =   &H8000000E&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblOrderFrom 
         BackColor       =   &H8000000E&
         Caption         =   "Phone Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Thank you for patronizing us. Visit us again"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   4920
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000014&
         Caption         =   "Name of staff"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label StaffName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   4320
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmOtherReciept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdPrint_Click()
CommonDialog1.PrinterDefault = True
CommonDialog1.ShowPrinter
frmOtherReciept.PrintForm
Printer.EndDoc
End Sub


