VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStartimesReciept 
   Caption         =   "STARTIMES RECIEPT"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6540
   LinkTopic       =   "Form2"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Startimes Reciept"
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
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
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
         Left            =   4800
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5400
         Top             =   4800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label TransactionID 
         BackColor       =   &H8000000E&
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
         Left            =   1440
         TabIndex        =   20
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Transaction id"
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
         TabIndex        =   19
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label lblHotelAddress 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"OtherReciept.frx":0000
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label SubscriptionAmount 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   225
         Left            =   1440
         TabIndex        =   16
         Top             =   4200
         Width           =   1800
      End
      Begin VB.Label SubscriptionType 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   225
         Left            =   1080
         TabIndex        =   15
         Top             =   3840
         Width           =   3000
         WordWrap        =   -1  'True
      End
      Begin VB.Label CustomerName 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   240
         Left            =   1200
         TabIndex        =   14
         Top             =   2400
         Width           =   2985
         WordWrap        =   -1  'True
      End
      Begin VB.Label DatePaid 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   225
         Left            =   1680
         TabIndex        =   13
         Top             =   2760
         Width           =   2280
      End
      Begin VB.Label CustomerPhoneno 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   225
         Left            =   1320
         TabIndex        =   12
         Top             =   3120
         Width           =   2160
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
         TabIndex        =   11
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label lblDescription 
         BackColor       =   &H8000000E&
         Caption         =   "Startimes"
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
         TabIndex        =   10
         Top             =   3840
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
         TabIndex        =   9
         Top             =   2400
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
         TabIndex        =   8
         Top             =   2760
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
         TabIndex        =   7
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Smartcard Number"
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
         TabIndex        =   6
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label SmartCardNumber 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   225
         Left            =   1800
         TabIndex        =   5
         Top             =   3480
         Width           =   2535
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
         Top             =   5280
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
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label StaffName 
         BackColor       =   &H8000000E&
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
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   4560
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmStartimesReciept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
On Error GoTo cancel_error
    'Show the printer dialog box
    CommonDialog1.PrinterDefault = True
    CommonDialog1.ShowPrinter
    frmStartimesReciept.PrintForm
    Printer.EndDoc

Exit Sub

cancel_error:
     If Err.Number <> vbCancel Then
          MsgBox Err.Description
     End If

End Sub


