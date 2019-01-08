VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStartimes 
   Caption         =   "Startimes Bouquets"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   9720
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "Startimes.frx":0000
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Frame frmStartimes 
      BackColor       =   &H00C00000&
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
      Height          =   9000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.TextBox txtTransactionID 
         DataField       =   "Transaction_id"
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
         Left            =   5160
         TabIndex        =   25
         Top             =   5400
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker DatePicker1 
         Height          =   495
         Left            =   8520
         TabIndex        =   22
         Top             =   6240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         _Version        =   393216
         Format          =   92864513
         CurrentDate     =   43337
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
         Left            =   480
         TabIndex        =   21
         Top             =   6960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   3840
         Top             =   6960
         Width           =   3975
         _ExtentX        =   7011
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
         RecordSource    =   "Startimes"
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
      Begin VB.CommandButton Command6 
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
         Left            =   3600
         TabIndex        =   20
         Top             =   7680
         Width           =   1575
      End
      Begin VB.CommandButton cmdSubmitt 
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
         Left            =   2040
         TabIndex        =   19
         Top             =   7680
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
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
         Left            =   600
         TabIndex        =   18
         Top             =   7680
         Width           =   1215
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
         Height          =   495
         Left            =   8520
         TabIndex        =   10
         Top             =   7680
         Width           =   975
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
         Left            =   7080
         TabIndex        =   9
         Top             =   7680
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
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
         Height          =   495
         Left            =   5400
         TabIndex        =   8
         Top             =   7680
         Width           =   1455
      End
      Begin VB.TextBox txtDate 
         DataField       =   "Payment_date"
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
         Left            =   5160
         TabIndex        =   7
         Top             =   6240
         Width           =   3405
      End
      Begin VB.TextBox txtStaffName 
         DataField       =   "Staff_name"
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
         Left            =   5160
         TabIndex        =   6
         Top             =   4680
         Width           =   4000
      End
      Begin VB.TextBox txtSmartCardNo 
         DataField       =   "Smart_cardNo"
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
         Left            =   5160
         MaxLength       =   11
         TabIndex        =   5
         Top             =   3840
         Width           =   4000
      End
      Begin VB.TextBox txtSubscriptionAmount 
         DataField       =   "Subscription_amount"
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
         Left            =   5160
         TabIndex        =   4
         Top             =   3000
         Width           =   4000
      End
      Begin VB.TextBox txtCustomerPhoneNo 
         DataField       =   "Customer_phoneNo"
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
         Left            =   5160
         MaxLength       =   11
         TabIndex        =   3
         Top             =   1440
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
         Left            =   5160
         TabIndex        =   2
         Top             =   720
         Width           =   4000
      End
      Begin VB.ComboBox cboSubscriptionType 
         DataField       =   "Subscription_type"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "Startimes.frx":00DE
         Left            =   5160
         List            =   "Startimes.frx":00F1
         TabIndex        =   1
         Text            =   "   SELECT STARTIMES  BOUQUET TYPE"
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C00000&
         Caption         =   "Transaction id"
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
         Left            =   1560
         TabIndex        =   24
         Top             =   5400
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   9840
         Picture         =   "Startimes.frx":0146
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   4575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Date of Payment"
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
         Left            =   1560
         TabIndex        =   17
         Top             =   6240
         Width           =   2505
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "Name of Staff"
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
         Left            =   1560
         TabIndex        =   16
         Top             =   4680
         Width           =   2505
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Smart Card Number"
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
         Left            =   1560
         TabIndex        =   15
         Top             =   3840
         Width           =   2505
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Subscription Amount"
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
         Left            =   1560
         TabIndex        =   14
         Top             =   3000
         Width           =   2505
      End
      Begin VB.Label Label3 
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
         Left            =   1560
         TabIndex        =   13
         Top             =   2160
         Width           =   2505
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Customer Phone Number"
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
         Left            =   1560
         TabIndex        =   12
         Top             =   1440
         Width           =   2505
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Customer Name"
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
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   2505
      End
   End
End
Attribute VB_Name = "frmStartimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdAdd_Click()
Dim res As Integer
On Error GoTo trap:
Adodc1.Recordset.AddNew
Exit Sub
trap:
End Sub

Private Sub cmdExit_Click()
Dim res As Integer
res = MsgBox("Do you want to cancel", vbInformation + vbYesNo, "Exit")
If res = vbYes Then
End
Else
End If
End Sub

Private Sub cmdPrint_Click()
cmdReciept = True
End Sub

Private Sub cmdReciept_Click()
Dim CustomerName As Object
Dim CustomerPhoneno As Object
Dim SubscriptionType As Object
Dim SubscriptionAmount As Object
Dim DatePaid As Object
Dim SmartCardNumber As Object
Dim StaffName As Object
Dim TransactionID As Object

cmdReciept.Enabled = False
Set StaffName = txtCustomerName
Set CustomerPhoneno = txtCustomerPhoneNo
Set SubscriptionType = cboSubscriptionType
Set SubscriptionAmount = txtSubscriptionAmount
Set DatePaid = txtDate
Set SmartCardNumber = txtSmartCardNo
Set StaffName = txtStaffName
Set TransactionID = txtTransactionID

With frmStartimesReciept
    .Show
    .CustomerName = txtCustomerName
    .CustomerPhoneno = txtCustomerPhoneNo
    .SubscriptionType = cboSubscriptionType
    .SubscriptionAmount = txtSubscriptionAmount
    .DatePaid = txtDate
    .SmartCardNumber = txtSmartCardNo
    .StaffName = txtStaffName
    .TransactionID = txtTransactionID
End With
End Sub

Private Sub cmdSubmitt_Click()
Dim res As Integer
On Error GoTo trap:
Adodc1.Recordset.Update
res = MsgBox("Record Saved", vbInformation, "Record was successfully Saved")
Adodc1.Refresh
Exit Sub
trap: res = MsgBox("Empty field cann't Saved", vbInformation, "Saved")
Adodc1.Refresh
End Sub

Private Sub Command1_Click()
Me.Hide
frmMain.Show
End Sub

Private Sub Command6_Click()
Me.Hide
frmStartimesPaymentHistory.Show
End Sub

Private Sub DatePicker1_Change()
txtDate.Text = DatePicker1.Value
txtDate.Text = Format(txtDate.Text, "dd-MM-yyyy")
End Sub

Private Sub Form_Load()
' open the connection
   con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\RocketDB.mdb;Persist Security Info=False"
 
  'create a recordset
   rs.Open "Select * from Startimes", con, adOpenDynamic, adLockPessimistic
End Sub
