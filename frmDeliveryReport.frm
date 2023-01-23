VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDeliveryReport 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   $"frmDeliveryReport.frx":0000
   ClientHeight    =   9912
   ClientLeft      =   3960
   ClientTop       =   3024
   ClientWidth     =   17940
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9912
   ScaleWidth      =   17940
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      DataField       =   "ConsumerID"
      DataSource      =   "Adodc1"
      Height          =   372
      Left            =   480
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   972
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   120
      Top             =   2280
      Visible         =   0   'False
      Width           =   2652
      _ExtentX        =   4678
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
      RecordSource    =   "select *from DeliveryTable"
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
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00FFFF00&
      Caption         =   "View All Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9240
      Width           =   4212
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFF00&
      Caption         =   "Back"
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
      Left            =   15960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9240
      Width           =   1452
   End
   Begin VB.TextBox txtID 
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
      Left            =   2760
      TabIndex        =   6
      Top             =   3720
      Width           =   3012
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmDeliveryReport.frx":00BD
      Height          =   4692
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   17532
      _ExtentX        =   30925
      _ExtentY        =   8276
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16711680
      DefColWidth     =   142
      HeadLines       =   1
      RowHeight       =   26
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "No"
         Caption         =   "No"
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
         DataField       =   "ConsumerID"
         Caption         =   "ConsumerID"
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
         DataField       =   "ConsumerName"
         Caption         =   "ConsumerName"
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
         DataField       =   "CylinderNo"
         Caption         =   "CylinderNo"
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
         DataField       =   "BookingDate"
         Caption         =   "BookingDate"
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
         DataField       =   "DeliveryDate"
         Caption         =   "DeliveryDate"
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
      BeginProperty Column06 
         DataField       =   "PinCode"
         Caption         =   "PinCode"
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
      BeginProperty Column07 
         DataField       =   "PhoneNumber"
         Caption         =   "PhoneNumber"
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
      BeginProperty Column08 
         DataField       =   "Address"
         Caption         =   "Address"
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
      BeginProperty Column09 
         DataField       =   "Status"
         Caption         =   "Status"
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
            ColumnWidth     =   1716.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1716.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1716.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1716.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1716.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1716.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1716.095
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1716.095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1716.095
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1716.095
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consumer ID "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   444
      Left            =   360
      TabIndex        =   5
      Top             =   3720
      Width           =   1944
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cylinder Delivery Report"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   744
      Left            =   5520
      TabIndex        =   3
      Top             =   2400
      Width           =   7596
   End
   Begin VB.Image Image1 
      Height          =   2052
      Left            =   15600
      Picture         =   "frmDeliveryReport.frx":00D2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2412
   End
   Begin VB.Image Image2 
      Height          =   2052
      Left            =   0
      Picture         =   "frmDeliveryReport.frx":260C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2292
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bharat Gas Service"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002448A8&
      Height          =   744
      Left            =   6240
      TabIndex        =   2
      Top             =   120
      Width           =   5820
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Bhagalpur Gas Agency )"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   28.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002448A8&
      Height          =   576
      Left            =   6360
      TabIndex        =   1
      Top             =   960
      Width           =   6000
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Height          =   2052
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18012
   End
End
Attribute VB_Name = "frmDeliveryReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAll_Click()
Adodc1.RecordSource = "Select * from DeliveryTable"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub txtID_Change()
If KeyAscii = vbKeyReturn Then
    If txtID.Text = "" Then
    MsgBox " Please ender ID ", vbInformation
    txtID.SetFocus
Else
Adodc1.RecordSource = "Select * from DeliveryTable where ConsumerID='" + txtID.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End If
End If
End Sub
