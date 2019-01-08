VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBankDeposit 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtAccountNo 
      DataField       =   "Account_no"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      MaxLength       =   11
      TabIndex        =   3
      Top             =   1920
      Width           =   4000
   End
   Begin VB.Frame DSTV 
      BackColor       =   &H00C00000&
      Caption         =   "BANK DEPOSIT"
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
      Begin VB.TextBox txtStaffName 
         DataField       =   "Staff_name"
         DataSource      =   "Adodc1"
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
         Left            =   4440
         TabIndex        =   29
         Top             =   6480
         Width           =   4215
      End
      Begin VB.CommandButton cmdReferenceID 
         Caption         =   "Ref"
         Height          =   495
         Left            =   7920
         TabIndex        =   27
         Top             =   4440
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DatePicker1 
         Height          =   375
         Left            =   7920
         TabIndex        =   26
         Top             =   5880
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         Format          =   92536835
         CurrentDate     =   43343
      End
      Begin VB.TextBox txtDate 
         DataField       =   "Transaction_date"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   5880
         Width           =   3495
      End
      Begin VB.ComboBox cboBankName 
         DataField       =   "Bank_name"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "BankDeposit.frx":0000
         Left            =   4440
         List            =   "BankDeposit.frx":0046
         TabIndex        =   7
         Text            =   "Select Bank Name"
         Top             =   5160
         Width           =   4095
      End
      Begin VB.TextBox txtAccountName 
         DataField       =   "Account_name"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   4
         Top             =   2760
         Width           =   4000
      End
      Begin VB.TextBox txtCustomerName 
         DataField       =   "Customer_name"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   1
         Top             =   360
         Width           =   4000
      End
      Begin VB.TextBox txtCustomerPhoneNo 
         DataField       =   "Phone_number"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         MaxLength       =   11
         TabIndex        =   2
         Top             =   1080
         Width           =   4000
      End
      Begin VB.TextBox txtAmount 
         DataField       =   "Amount"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   5
         Top             =   3600
         Width           =   4000
      End
      Begin VB.TextBox txtTransactionID 
         DataField       =   "Transaction_ID"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   6
         Top             =   4440
         Width           =   3525
      End
      Begin VB.CommandButton cmdMain 
         Caption         =   "Main menu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   6120
         TabIndex        =   15
         Top             =   7800
         Width           =   1455
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
         Height          =   500
         Left            =   7800
         TabIndex        =   14
         Top             =   7800
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
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
         Height          =   500
         Left            =   9120
         TabIndex        =   13
         Top             =   7800
         Width           =   975
      End
      Begin VB.CommandButton frmAdd 
         Caption         =   "Add New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   720
         TabIndex        =   12
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2280
         TabIndex        =   11
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CommandButton cmdBankdepositHis 
         Caption         =   "View Record"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3840
         TabIndex        =   10
         Top             =   7800
         Width           =   1815
      End
      Begin VB.CommandButton cmdReciept 
         Caption         =   "Reciept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   7200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   3960
         Top             =   7200
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Rocket_Search_Application\RocketDB.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Rocket_Search_Application\RocketDB.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "BankDeposit"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C00000&
         Caption         =   "Staff Name"
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
         Height          =   495
         Left            =   840
         TabIndex        =   28
         Top             =   6480
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   2775
         Left            =   9120
         Picture         =   "BankDeposit.frx":0240
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   5775
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C00000&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   840
         TabIndex        =   25
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         Caption         =   "Account Number"
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
         Height          =   495
         Left            =   840
         TabIndex        =   23
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Account Name"
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
         Height          =   495
         Left            =   840
         TabIndex        =   21
         Top             =   2760
         Width           =   2505
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Customer Name"
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
         Height          =   495
         Left            =   840
         TabIndex        =   20
         Top             =   360
         Width           =   2505
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Customer Phone Number"
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
         Height          =   495
         Left            =   840
         TabIndex        =   19
         Top             =   1080
         Width           =   2985
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Amount"
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
         Height          =   495
         Left            =   840
         TabIndex        =   18
         Top             =   3720
         Width           =   2505
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "Transaction ID"
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
         Height          =   495
         Left            =   840
         TabIndex        =   17
         Top             =   4440
         Width           =   2505
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Bank Name"
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
         Height          =   495
         Left            =   840
         TabIndex        =   16
         Top             =   5160
         Width           =   2505
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "Bank Name"
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
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   2505
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C00000&
      Caption         =   "Subscription Type"
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
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   2505
   End
End
Attribute VB_Name = "frmBankDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub cmdBankdepositHis_Click()
Me.Hide
frmBankDepositHistory.Show
End Sub

Private Sub cmdExit_Click()
Dim res As Integer
res = MsgBox("Do you want to cancel", vbInformation + vbYesNo, "Exit")
If res = vbYes Then
End
Else
End If
End Sub

Private Sub cmdMain_Click()
Me.Hide
frmMain.Show
End Sub

Private Sub cmdPrint_Click()
cmdReciept = True
End Sub

Private Sub cmdReciept_Click()
Dim CustomerName As Object
Dim CustomerPhoneno As Object
Dim AccountNo As Object
Dim AccountName As Object
Dim TransactionID As Object
Dim Amount As Object
Dim BankName1 As Object
Dim TransactionDate As Object
Dim StaffName As Object
cmdReciept.Enabled = False
Set StaffName = txtCustomerName
Set CustomerPhoneno = txtCustomerPhoneNo
Set AccountNo = txtAccountNo
Set AccountName = txtAccountName
Set TransactionID = txtTransactionID
Set Amount = txtAmount
Set BankName1 = cboBankName
Set TransactionDate = txtDate
Set StaffName = txtStaffName
With frmBankDepositReciept
    .Show
    .CustomerName = txtCustomerName
    .CustomerPhoneno = txtCustomerPhoneNo
    .AccountNo = txtAccountNo
    .AccountName = txtAccountName
    .TransactionID = txtTransactionID
    .Amount = txtAmount
    .BankName1 = cboBankName
    .TransactionDate = txtDate
    .StaffName = txtStaffName
End With
End Sub

Private Sub cmdReferenceID_Click()
Dim intResult As Integer
Randomize
intResult = Int((10000 * Rnd) + 1)
txtTransactionID.Text = ("REFO" & intResult)
End Sub

Private Sub cmdSubmit_Click()
Dim res As Integer
On Error GoTo trap:
Adodc1.Recordset.Update
res = MsgBox("Record Saved", vbInformation, "Record was successfully Saved")
Adodc1.Refresh
Exit Sub
trap: res = MsgBox("Empty field cann't Saved", vbInformation, "Saved")
Adodc1.Refresh
End Sub


Private Sub DatePicker1_Change()
txtDate.Text = DatePicker1.Value
txtDate.Text = Format(txtDate.Text, "dd-MM-yyyy")
End Sub

Private Sub Form_Load()
' open the connection
   con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\RocketDB.mdb;Persist Security Info=False"
 
  'create a recordset
   rs.Open "Select * from BankDeposit", con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub frmAdd_Click()
Dim res As Integer
On Error GoTo trap:
Adodc1.Recordset.AddNew
Exit Sub
trap:
End Sub


