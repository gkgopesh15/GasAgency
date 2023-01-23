VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBooking 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   $"frmBooking.frx":0000
   ClientHeight    =   10728
   ClientLeft      =   4920
   ClientTop       =   2388
   ClientWidth     =   17016
   ForeColor       =   &H002448A8&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10728
   ScaleWidth      =   17016
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   312
      Left            =   3600
      Top             =   1560
      Visible         =   0   'False
      Width           =   1692
      _ExtentX        =   2985
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
      RecordSource    =   "select * from Stock"
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
   Begin VB.TextBox Text3 
      DataField       =   "Booked"
      DataSource      =   "Adodc2"
      Height          =   288
      Left            =   4320
      TabIndex        =   38
      Text            =   "Text3"
      Top             =   2280
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.TextBox Text2 
      DataField       =   "No"
      DataSource      =   "AdoBook"
      Height          =   288
      Left            =   14760
      TabIndex        =   37
      Text            =   "Text2"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1092
   End
   Begin MSAdodcLib.Adodc AdoBook 
      Height          =   312
      Left            =   11880
      Top             =   1920
      Visible         =   0   'False
      Width           =   1932
      _ExtentX        =   3408
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
      RecordSource    =   "Select *from BookingTable"
      Caption         =   "adobook"
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
   Begin VB.TextBox Text1 
      DataField       =   "ConsumerNo"
      DataSource      =   "Adodc1"
      Height          =   372
      Left            =   3480
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   372
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   720
      Top             =   1800
      Visible         =   0   'False
      Width           =   2172
      _ExtentX        =   3831
      _ExtentY        =   656
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
      RecordSource    =   "select *from NewConnection"
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
   Begin VB.Frame Frame1 
      Height          =   7812
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   16452
      Begin VB.TextBox txtConnectionType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   3600
         TabIndex        =   36
         Top             =   2640
         Width           =   2052
      End
      Begin VB.TextBox txtStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   14160
         TabIndex        =   34
         Top             =   2640
         Width           =   2052
      End
      Begin MSComCtl2.DTPicker DTPBookingDate 
         Height          =   492
         Left            =   9600
         TabIndex        =   33
         Top             =   2640
         Width           =   1932
         _ExtentX        =   3408
         _ExtentY        =   868
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110231553
         CurrentDate     =   43282
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFF00&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   6360
         Width           =   3132
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   5160
         Width           =   3132
      End
      Begin VB.CommandButton cmdBooking 
         BackColor       =   &H00FFFF00&
         Caption         =   "Booking"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3960
         Width           =   3132
      End
      Begin VB.TextBox txtPhoneNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   3960
         TabIndex        =   28
         Top             =   6480
         Width           =   5892
      End
      Begin VB.TextBox txtPinCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   3960
         TabIndex        =   27
         Top             =   5640
         Width           =   5892
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   3960
         TabIndex        =   26
         Top             =   4800
         Width           =   5892
      End
      Begin VB.TextBox txtConsumerName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   3960
         TabIndex        =   25
         Top             =   3960
         Width           =   5892
      End
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   12240
         TabIndex        =   20
         Top             =   2640
         Width           =   1692
      End
      Begin VB.ComboBox cmbNoofCylinder 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   444
         Left            =   6720
         TabIndex        =   19
         Top             =   2640
         Width           =   1812
      End
      Begin VB.ComboBox cmbConsumerID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   444
         Left            =   600
         TabIndex        =   18
         Top             =   2640
         Width           =   1812
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         Height          =   372
         Left            =   0
         TabIndex        =   32
         Top             =   7440
         Width           =   16452
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   480
         TabIndex        =   24
         Top             =   6480
         Width           =   2112
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   480
         TabIndex        =   23
         Top             =   5640
         Width           =   1308
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   480
         TabIndex        =   22
         Top             =   4800
         Width           =   1164
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consumer Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   480
         TabIndex        =   21
         Top             =   3960
         Width           =   2352
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   14640
         TabIndex        =   17
         Top             =   1800
         Width           =   876
      End
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   12480
         TabIndex        =   16
         Top             =   1800
         Width           =   1092
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Booking Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   9600
         TabIndex        =   15
         Top             =   1800
         Width           =   1896
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Cylinder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   6600
         TabIndex        =   14
         Top             =   1800
         Width           =   2100
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cylinder Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   3600
         TabIndex        =   13
         Top             =   1800
         Width           =   1968
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consumer ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   600
         TabIndex        =   12
         Top             =   1800
         Width           =   1824
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         Height          =   372
         Left            =   0
         TabIndex        =   11
         Top             =   3360
         Width           =   16452
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   14400
         TabIndex        =   10
         Top             =   720
         Width           =   96
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   12480
         TabIndex        =   9
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblToday 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   14040
         TabIndex        =   8
         Top             =   240
         Width           =   96
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   348
         Left            =   12480
         TabIndex        =   7
         Top             =   240
         Width           =   1056
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Height          =   372
         Left            =   0
         TabIndex        =   6
         Top             =   1200
         Width           =   16452
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cylinder  Booking  Screen"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00420CB4&
      Height          =   564
      Left            =   5880
      TabIndex        =   4
      Top             =   1680
      Width           =   5208
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Pyar  bnayen  Pyar  prosen)"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   444
      Left            =   8760
      TabIndex        =   3
      Top             =   840
      Width           =   4068
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Bharat Gas)"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002448A8&
      Height          =   744
      Left            =   9840
      TabIndex        =   2
      Top             =   120
      Width           =   4212
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bhagalpur Gas Agency "
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002448A8&
      Height          =   744
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   6828
   End
   Begin VB.Image Image1 
      Height          =   1212
      Left            =   840
      Picture         =   "frmBooking.frx":00B2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Height          =   1452
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17052
   End
End
Attribute VB_Name = "frmBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbConsumerID_Click()
Adodc1.RecordSource = " select *from NewConnection where ConsumerID= '" & cmbConsumerID.Text & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
        
          MsgBox " Sorry Consumer Id is wrong , Please check your consumer Id", vbInformation, "Instruction"
    Else
          loadData
End If
AdoBook.RecordSource = " select *from BookingTable where ConsumerID= '" & cmbConsumerID.Text & "'"
AdoBook.Refresh
If AdoBook.Recordset.EOF Then
    txtRemark.Text = ""
    txtStatus.Text = ""
Else

On Error Resume Next
txtRemark.Text = AdoBook.Recordset.Fields(5).Value
txtStatus.Text = AdoBook.Recordset.Fields(6).Value
DTPBookingDate.Value = AdoBook.Recordset.Fields(4).Value
cmbNoofCylinder.Text = AdoBook.Recordset.Fields(3).Value
End If
End Sub
Public Sub loadData()
    txtConsumerName.Text = Adodc1.Recordset.Fields(2).Value
    txtPhoneNumber.Text = Adodc1.Recordset.Fields(4).Value
    txtAddress.Text = Adodc1.Recordset.Fields(6).Value
    txtPinCode.Text = Adodc1.Recordset.Fields(7).Value
    txtConnectionType.Text = Adodc1.Recordset.Fields(9).Value
    If Adodc1.Recordset.Fields(10).Value = "Single Connection" Then
        cmbNoofCylinder.Text = 1
    Else
        cmbNoofCylinder.Text = 2
    End If
    
End Sub

Private Sub cmdBooking_Click()
If cmbConsumerID.Text = "" Or cmbNoofCylinder.Text = "" Or txtRemark.Text = "" Or txtStatus.Text = "" Then
    MsgBox " Please Enter Consumer id or other fields first ", vbInformation, "Instruction"
Else

Adodc1.RecordSource = " select *from BookingTable where ConsumerID= '" & cmbConsumerID.Text & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
ans = MsgBox(" Do you want to confirm the booking", vbQuestion + vbYesNo, "Message")
If (ans = vbYes) Then
         AdoBook.Recordset.AddNew
With AdoBook.Recordset
        .Fields(1).Value = cmbConsumerID.Text
        .Fields(2).Value = txtConsumerName.Text
        .Fields(3).Value = cmbNoofCylinder.Text
        .Fields(4).Value = DTPBookingDate.Value
        .Fields(5).Value = txtRemark.Text
        .Fields(6).Value = txtStatus.Text
    End With
    AdoBook.Recordset.Update
    AdoBook.Refresh
    MsgBox " Congratulation your Booking is Succefull!!", vbInformation + vbOKOnly, "Saved Successfully"
    Adodc2.Recordset.MoveFirst
    book = Adodc2.Recordset.Fields(0).Value
    book = book + cmbNoofCylinder.Text
    Adodc2.Recordset.Update
    Adodc2.Recordset.Fields(0).Value = book
    Adodc2.Recordset.Update
    cmbConsumerID.Text = ""
    ResetData
    cmbConsumerID.SetFocus
End If
Else
          MsgBox " Booking already exixt .", vbInformation, "Information"
End If
End If
End Sub
Public Sub ResetData()
txtConsumerName.Text = ""
cmbNoofCylinder.Text = ""
DTPBookingDate.Value = Date
txtRemark.Text = ""
txtStatus.Text = ""
txtAddress.Text = ""
txtPinCode.Text = ""
txtPhoneNumber.Text = ""
txtConnectionType.Text = ""
End Sub

Private Sub Form_Load()
lblToday.Caption = Date
lblTime.Caption = Time
cmbNoofCylinder.AddItem "1"
cmbNoofCylinder.AddItem "2"
Adodc1.RecordSource = " Select ConsumerID from NewConnection"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
While Adodc1.Recordset.EOF = False
cmbConsumerID.AddItem Adodc1.Recordset.Fields(0).Value
Adodc1.Recordset.MoveNext
Wend
End If
book = 0
AdoBook.Recordset.MoveLast
book = AdoBook.Recordset.Fields(0).Value
book = book + 1

End Sub

Private Sub cmdCancel_Click()
ans1 = MsgBox(" Do you want to cancel the booking ", vbQuestion + vbYesNo, "Message")
If (ans1 = vbYes) Then
AdoBook.Recordset.Delete
MsgBox " Your booking is cancel succefully .", vbInformation, "Information"
 Adodc2.Recordset.MoveFirst
    book = Adodc2.Recordset.Fields(0).Value
    book = book - cmbNoofCylinder.Text
    Adodc2.Recordset.Update
    Adodc2.Recordset.Fields(0).Value = book
    Adodc2.Recordset.Update
cmbConsumerID.Text = ""
ResetData
End If
End Sub
