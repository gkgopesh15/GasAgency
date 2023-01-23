VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEmployeePayment 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   $"frmEmployeePayment.frx":0000
   ClientHeight    =   9240
   ClientLeft      =   5136
   ClientTop       =   3120
   ClientWidth     =   16824
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   16824
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text9 
      DataField       =   "EmployeeID"
      DataSource      =   "Adodc2"
      Height          =   288
      Left            =   7080
      TabIndex        =   24
      Text            =   "Text9"
      Top             =   8880
      Visible         =   0   'False
      Width           =   972
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   312
      Left            =   9480
      Top             =   8880
      Visible         =   0   'False
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   550
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=DsnAgency"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "DsnAgency"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from EmployeePayment"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text8 
      DataField       =   "No"
      DataSource      =   "Adodc1"
      Height          =   612
      Left            =   15600
      TabIndex        =   23
      Text            =   "Text8"
      Top             =   8160
      Visible         =   0   'False
      Width           =   732
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   12720
      Top             =   8880
      Visible         =   0   'False
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   550
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=DsnAgency"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "DsnAgency"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from EmployeeEntry"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFF00&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8160
      Width           =   2532
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFF00&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8160
      Width           =   2652
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFF00&
      Caption         =   "Delete "
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8160
      Width           =   2652
   End
   Begin VB.CommandButton cmdMakePayment 
      BackColor       =   &H00FFFF00&
      Caption         =   "Make Payment"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8160
      Width           =   2772
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   14880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Width           =   732
   End
   Begin VB.TextBox txtSearchID 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   12480
      TabIndex        =   17
      Top             =   4800
      Width           =   2172
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmEmployeePayment.frx":00B1
      Height          =   2532
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   16572
      _ExtentX        =   29231
      _ExtentY        =   4466
      _Version        =   393216
      AllowUpdate     =   -1  'True
      DefColWidth     =   208
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   30
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "EmployeeID"
         Caption         =   "EmployeeID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "EmployeeName"
         Caption         =   "EmployeeName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Branch"
         Caption         =   "Branch"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Salary"
         Caption         =   "Salary"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "PaymentDetails"
         Caption         =   "PaymentDetails"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PaymentDate"
         Caption         =   "PaymentDate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2508.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2508.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2508.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2508.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2508.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2508.095
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPaymentDate 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   14280
      TabIndex        =   13
      Top             =   3120
      Width           =   2172
   End
   Begin VB.TextBox txtPaymentDetails 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   11040
      TabIndex        =   12
      Top             =   3120
      Width           =   2652
   End
   Begin VB.TextBox txtSalary 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   9000
      TabIndex        =   11
      Top             =   3120
      Width           =   1572
   End
   Begin VB.TextBox txtBranch 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   6360
      TabIndex        =   10
      Top             =   3120
      Width           =   2292
   End
   Begin VB.TextBox txtEmployeeName 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   3240
      TabIndex        =   9
      Top             =   3120
      Width           =   2652
   End
   Begin VB.TextBox txtEmployeeID 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   360
      TabIndex        =   8
      Top             =   3120
      Width           =   2172
   End
   Begin VB.Image Image2 
      Height          =   1812
      Left            =   14640
      Picture         =   "frmEmployeePayment.frx":00C6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1692
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   10200
      TabIndex        =   16
      Top             =   4800
      Width           =   2040
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee List :"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   480
      TabIndex        =   15
      Top             =   4800
      Width           =   2436
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   14280
      TabIndex        =   7
      Top             =   2400
      Width           =   2184
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Details"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   11040
      TabIndex        =   6
      Top             =   2400
      Width           =   2568
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   9120
      TabIndex        =   5
      Top             =   2400
      Width           =   1008
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Branch"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   6360
      TabIndex        =   4
      Top             =   2400
      Width           =   1176
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   3240
      TabIndex        =   3
      Top             =   2400
      Width           =   2496
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   2040
   End
   Begin VB.Image Image1 
      Height          =   1812
      Left            =   600
      Picture         =   "frmEmployeePayment.frx":26A8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1692
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee List  and Payment details"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   744
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   10944
   End
   Begin VB.Label Label1 
      BackColor       =   &H002448A8&
      ForeColor       =   &H00420CB4&
      Height          =   1812
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16812
   End
End
Attribute VB_Name = "frmEmployeePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClear_Click()
ResetData
End Sub

Private Sub cmdDelete_Click()
ans2 = MsgBox(" Do you want to confirm to Delate.", vbQuestion + vbYesNo, " Message")
If (ans1 - vbYes) Then
Adodc2.Recordset.Delete
End If
End Sub

Private Sub cmdExit_Click()
Exit Sub
End Sub

Private Sub cmdMakePayment_Click()
If txtEmployeeID.Text = "" Or txtPaymentDetails.Text = "" Or txtPaymentDate.Text = "" Then
    MsgBox " Please enter payment Details and payment Date information .", vbInformation, "message"
Else
ans1 = MsgBox(" Do you want to confirm  the Payment", vbQuestion + vbYesNo, " Message")
If (ans1 - vbYes) Then
Adodc2.Recordset.AddNew
With Adodc2.Recordset
        .Fields(0).Value = txtEmployeeID.Text
        .Fields(1).Value = txtEmployeeName.Text
        .Fields(3).Value = txtSalary.Text
        .Fields(2).Value = txtBranch.Text
        .Fields(4).Value = txtPaymentDetails.Text
        .Fields(5).Value = txtPaymentDate.Text
    End With
    Adodc2.Recordset.Update
    Adodc2.Refresh
    MsgBox " Payment Make Succefully!!", vbInformation + vbOKOnly, "Saved Successfully"
    End If
     txtEmployeeName.Text = ""
    ResetData
    txtEmployeeID.SetFocus
End If
End Sub
Public Sub ResetData()
txtEmployeeID.Text = ""
txtEmployeeName.Text = ""
txtSalary.Text = ""
txtBranch.Text = ""
txtPaymentDetails.Text = ""
txtPaymentDetails.Text = ""

End Sub

Private Sub txtEmployeeID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Adodc1.RecordSource = " select *from EmployeeEntry where EmployeeID= '" & txtEmployeeID.Text & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
        
          MsgBox " Sorry Employee Id is wrong , Please check your consumer Id", vbInformation, "Instruction"
    Else
          loadData
End If
End If
End Sub
Public Sub loadData()
    txtEmployeeName.Text = Adodc1.Recordset.Fields(2).Value
    txtSalary.Text = Adodc1.Recordset.Fields(5).Value
    txtBranch.Text = Adodc1.Recordset.Fields(7).Value
End Sub

Private Sub txtPaymentDetails_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtPaymentDetails.Text = "" Then
    MsgBox "Please enter payment details . ", vbInformation, "message"
    txtPaymentDetails.SetFocus
    Else
    txtPaymentDate.SetFocus
    End If
End If
End Sub
