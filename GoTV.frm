VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGoTV 
   Caption         =   "GoTV Package"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frmGoTV 
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
      Height          =   8415
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
         Top             =   5520
         Width           =   3975
      End
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
         Height          =   2415
         Left            =   10560
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "GoTV.frx":0000
         Top             =   1560
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
         Format          =   92340225
         CurrentDate     =   43337
      End
      Begin VB.CommandButton cmdReciept 
         Caption         =   "Reciept"
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
         Left            =   240
         TabIndex        =   21
         Top             =   6960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   2160
         Top             =   6960
         Width           =   5175
         _ExtentX        =   9128
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
         RecordSource    =   "GOTV"
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
         Left            =   3480
         TabIndex        =   20
         Top             =   7680
         Width           =   1815
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
         Left            =   2160
         TabIndex        =   19
         Top             =   7680
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
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
         TabIndex        =   18
         Top             =   7680
         Width           =   1215
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
         ItemData        =   "GoTV.frx":00AB
         Left            =   5160
         List            =   "GoTV.frx":00BB
         TabIndex        =   10
         Text            =   "   SELECT GOTV SUBSCRIPTION  PACKAGE"
         Top             =   2280
         Width           =   3855
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
         TabIndex        =   9
         Top             =   720
         Width           =   4000
      End
      Begin VB.TextBox txtCustomerPhoneNo 
         DataField       =   "Customer_Phoneno"
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
         TabIndex        =   8
         Top             =   1440
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
         TabIndex        =   7
         Top             =   3120
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
         TabIndex        =   6
         Top             =   3960
         Width           =   4000
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
         TabIndex        =   5
         Top             =   4800
         Width           =   4000
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
         TabIndex        =   4
         Top             =   6240
         Width           =   3405
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
         Left            =   5520
         TabIndex        =   3
         Top             =   7680
         Width           =   1455
      End
      Begin VB.CommandButton cmd 
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
         Left            =   7320
         TabIndex        =   2
         Top             =   7680
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
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
         Left            =   8640
         TabIndex        =   1
         Top             =   7680
         Width           =   975
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
         Left            =   1440
         TabIndex        =   24
         Top             =   5520
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   2655
         Left            =   9960
         Picture         =   "GoTV.frx":00EE
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   4695
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
         TabIndex        =   17
         Top             =   720
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
         TabIndex        =   16
         Top             =   1440
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
         TabIndex        =   15
         Top             =   2280
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
         Top             =   3120
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
         TabIndex        =   13
         Top             =   3960
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
         TabIndex        =   12
         Top             =   4800
         Width           =   2505
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
         TabIndex        =   11
         Top             =   6240
         Width           =   2505
      End
   End
End
Attribute VB_Name = "frmGoTV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub cmd_Click()
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

With frmGOTVReciept
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

Private Sub Command1_Click()
Me.Hide
frmMain.Show
End Sub

Private Sub Command3_Click()
Dim res As Integer
res = MsgBox("Do you want to cancel", vbInformation + vbYesNo, "Exit")
If res = vbYes Then
End
Else
End If
End Sub

Private Sub Command4_Click()
Dim res As Integer
On Error GoTo trap:
Adodc1.Recordset.AddNew
Exit Sub
trap:
End Sub

Private Sub Command6_Click()
Me.Hide
frmGOTVPaymentHistory.Show
End Sub


Private Sub DatePicker1_Change()
txtDate.Text = DatePicker1.Value
txtDate.Text = Format(txtDate.Text, "dd-MM-yyyy")
End Sub


Private Sub Form_Load()
' open the connection
   con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\RocketDB.mdb;Persist Security Info=False"
 
  'create a recordset
   rs.Open "Select * from GOTV", con, adOpenDynamic, adLockPessimistic
End Sub
