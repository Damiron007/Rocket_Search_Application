VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOthers 
   Caption         =   "Other Sales"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frmDSTV 
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
         Left            =   360
         TabIndex        =   21
         Top             =   6480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DatePicker1 
         Height          =   495
         Left            =   8640
         TabIndex        =   20
         Top             =   5520
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         _Version        =   393216
         Format          =   92864513
         CurrentDate     =   43337
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "Add"
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
         Left            =   960
         TabIndex        =   19
         Top             =   7440
         Width           =   975
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   2760
         Top             =   6480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1085
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
         RecordSource    =   "Others"
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
      Begin VB.TextBox txtServiceName 
         DataField       =   "Service_name"
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
         TabIndex        =   18
         Top             =   2640
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
         TabIndex        =   11
         Top             =   720
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
         TabIndex        =   10
         Top             =   1680
         Width           =   4000
      End
      Begin VB.TextBox txtAmount 
         DataField       =   "Amountpaid"
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
         Top             =   3600
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
         TabIndex        =   8
         Top             =   4560
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
         TabIndex        =   7
         Top             =   5520
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
         Left            =   5760
         TabIndex        =   6
         Top             =   7440
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
         Left            =   7560
         TabIndex        =   5
         Top             =   7440
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
         Left            =   8880
         TabIndex        =   4
         Top             =   7440
         Width           =   975
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
         Left            =   5160
         TabIndex        =   3
         Top             =   6600
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
         Left            =   2160
         TabIndex        =   2
         Top             =   7440
         Width           =   1215
      End
      Begin VB.CommandButton cmdRecord 
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
         Left            =   3720
         TabIndex        =   1
         Top             =   7440
         Width           =   1815
      End
      Begin VB.Image Image3 
         Height          =   2415
         Left            =   10320
         Picture         =   "OTHERS.frx":0000
         Stretch         =   -1  'True
         Top             =   5760
         Width           =   3855
      End
      Begin VB.Image Image2 
         Height          =   2415
         Left            =   10320
         Picture         =   "OTHERS.frx":188C
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   4095
      End
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   10320
         Picture         =   "OTHERS.frx":543D
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3855
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
         Top             =   1680
         Width           =   2505
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Product"
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
         Top             =   2640
         Width           =   2505
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Amount"
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
         Top             =   3600
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
         TabIndex        =   13
         Top             =   4560
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
         TabIndex        =   12
         Top             =   5520
         Width           =   2505
      End
   End
End
Attribute VB_Name = "frmOthers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command3_Click()
Dim res As Integer
res = MsgBox("Do you want to cancel", vbInformation + vbYesNo, "Exit")
If res = vbYes Then
End
Else
End If
End Sub

Private Sub cmdAddNew_Click()
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
Dim ServiceName As Object
Dim Amount As Object
Dim DatePaid As Object
Dim StaffName As Object
cmdReciept.Enabled = False
Set StaffName = txtCustomerName
Set CustomerPhoneno = txtCustomerPhoneNo
Set ServiceName = txtServiceName
Set Amount = txtAmount
Set DatePaid = txtDate
Set StaffName = txtStaffName

With frmOtherReciept
    .Show
    .CustomerName = txtCustomerName
    .CustomerPhoneno = txtCustomerPhoneNo
    .ServiceName = txtServiceName
    .Amount = txtAmount
    .DatePaid = txtDate
    .StaffName = txtStaffName
End With
End Sub

Private Sub cmdRecord_Click()
Me.Hide
frmOtherServices.Show
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
   rs.Open "Select * from Others", con, adOpenDynamic, adLockPessimistic
End Sub
