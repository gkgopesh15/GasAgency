VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDelivery 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12516
   ClientLeft      =   2232
   ClientTop       =   -312
   ClientWidth     =   18168
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12516
   ScaleWidth      =   18168
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   312
      Left            =   12960
      Top             =   2040
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
      RecordSource    =   "select *from DeliveryTable"
      Caption         =   "AdoDelivery"
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
   Begin VB.TextBox Text2 
      DataField       =   "ConsumerNo"
      DataSource      =   "Adodc2"
      Height          =   612
      Left            =   4920
      TabIndex        =   41
      Text            =   "Text2"
      Top             =   1800
      Visible         =   0   'False
      Width           =   492
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   312
      Left            =   2640
      Top             =   1800
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
      RecordSource    =   "select *from NewConnection"
      Caption         =   "AdoConnection"
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
      DataField       =   "ConsumerID"
      DataSource      =   "Adodc1"
      Height          =   288
      Left            =   720
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   852
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   360
      Top             =   1800
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
      RecordSource    =   "select *from BookingTable"
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
   Begin VB.TextBox txtCylinderNo 
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
      Left            =   6720
      TabIndex        =   39
      Top             =   4800
      Width           =   2052
   End
   Begin VB.Frame Frame1 
      Height          =   10092
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   17892
      Begin VB.TextBox Text4 
         DataField       =   "Delivered"
         DataSource      =   "Adodc4"
         Height          =   288
         Left            =   5040
         TabIndex        =   44
         Text            =   "Text4"
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   312
         Left            =   2280
         Top             =   360
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
         RecordSource    =   "select * from Stock"
         Caption         =   "Adodc4"
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
         Left            =   15840
         TabIndex        =   43
         Top             =   2400
         Width           =   1812
      End
      Begin VB.TextBox Text3 
         Height          =   288
         Left            =   16680
         TabIndex        =   42
         Text            =   "Text3"
         Top             =   0
         Visible         =   0   'False
         Width           =   1092
      End
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
         Left            =   3360
         TabIndex        =   38
         Top             =   2400
         Width           =   2052
      End
      Begin VB.CommandButton cmdBill 
         BackColor       =   &H00FFFF00&
         Caption         =   "Generate Bill"
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
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   5400
         Width           =   2892
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFF00&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   22.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   5400
         Width           =   2532
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmDelivery.frx":0000
         Height          =   2292
         Left            =   120
         TabIndex        =   34
         Top             =   7080
         Width           =   17532
         _ExtentX        =   30925
         _ExtentY        =   4043
         _Version        =   393216
         AllowUpdate     =   -1  'True
         DefColWidth     =   208
         ForeColor       =   16711680
         HeadLines       =   1
         RowHeight       =   24
         RowDividerStyle =   5
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
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
            DataField       =   "Remark"
            Caption         =   "Remark"
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
            BeginProperty Column06 
               ColumnWidth     =   2508.095
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdBookingList 
         BackColor       =   &H00FFFF00&
         Caption         =   "Booking List"
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
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   5400
         Width           =   2892
      End
      Begin MSComCtl2.DTPicker DTPBooking 
         Height          =   492
         Left            =   9840
         TabIndex        =   30
         Top             =   2400
         Width           =   2052
         _ExtentX        =   3620
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
         CalendarForeColor=   16711680
         CalendarTitleForeColor=   -2147483635
         Format          =   209911809
         CurrentDate     =   43277
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
         Left            =   480
         TabIndex        =   11
         Top             =   2400
         Width           =   1932
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
         TabIndex        =   10
         Top             =   3720
         Width           =   3852
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
         TabIndex        =   9
         Top             =   4560
         Width           =   3852
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
         Left            =   12360
         TabIndex        =   8
         Top             =   3720
         Width           =   3852
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
         Left            =   12360
         TabIndex        =   7
         Top             =   4560
         Width           =   3852
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
         Height          =   492
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5400
         Width           =   2892
      End
      Begin VB.CommandButton cmdConfirmDelivery 
         BackColor       =   &H00FFFF00&
         Caption         =   "Confirm Delivery"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
         Width           =   2892
      End
      Begin MSComCtl2.DTPicker DTPDelivery 
         Height          =   492
         Left            =   13080
         TabIndex        =   31
         Top             =   2400
         Width           =   2052
         _ExtentX        =   3620
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
         CalendarForeColor=   16711680
         Format          =   209911809
         CurrentDate     =   43277
      End
      Begin VB.Label Label22 
         BackColor       =   &H8000000D&
         Height          =   372
         Left            =   0
         TabIndex        =   35
         Top             =   9720
         Width           =   17892
      End
      Begin VB.Label lblAllList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "All Booked Cunsumer list"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   348
         Left            =   1320
         TabIndex        =   33
         Top             =   6600
         Width           =   3648
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
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
         Left            =   13080
         TabIndex        =   29
         Top             =   1800
         Width           =   1884
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Height          =   372
         Left            =   0
         TabIndex        =   27
         Top             =   1200
         Width           =   17892
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
         TabIndex        =   26
         Top             =   240
         Width           =   1056
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
         TabIndex        =   25
         Top             =   240
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
         TabIndex        =   24
         Top             =   720
         Width           =   900
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
         TabIndex        =   23
         Top             =   720
         Width           =   96
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         Height          =   372
         Left            =   0
         TabIndex        =   22
         Top             =   3120
         Width           =   17892
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
         Left            =   480
         TabIndex        =   21
         Top             =   1800
         Width           =   1824
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connection Type"
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
         Left            =   3360
         TabIndex        =   20
         Top             =   1800
         Width           =   2388
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
         TabIndex        =   19
         Top             =   1800
         Width           =   2100
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
         Left            =   9840
         TabIndex        =   18
         Top             =   1800
         Width           =   1896
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
         Left            =   15840
         TabIndex        =   17
         Top             =   1800
         Width           =   876
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
         TabIndex        =   16
         Top             =   3720
         Width           =   2352
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
         TabIndex        =   15
         Top             =   4560
         Width           =   1164
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
         Left            =   9600
         TabIndex        =   14
         Top             =   3720
         Width           =   1308
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
         Left            =   9600
         TabIndex        =   13
         Top             =   4560
         Width           =   2112
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         Height          =   372
         Left            =   0
         TabIndex        =   12
         Top             =   6120
         Width           =   17892
      End
   End
   Begin VB.Image Image2 
      Height          =   1212
      Left            =   16320
      Picture         =   "frmDelivery.frx":0015
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1452
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cylinder  Delivery  Entry  Screen"
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
      Left            =   5520
      TabIndex        =   28
      Top             =   1680
      Width           =   6480
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
      Left            =   12000
      TabIndex        =   3
      Top             =   960
      Width           =   4068
   End
   Begin VB.Image Image1 
      Height          =   1212
      Left            =   360
      Picture         =   "frmDelivery.frx":254F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1452
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
      Left            =   11760
      TabIndex        =   2
      Top             =   240
      Width           =   4332
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bhagalpur Gas Agency "
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002448A8&
      Height          =   996
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   9360
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Height          =   1572
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18132
   End
End
Attribute VB_Name = "frmDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbConsumerID_Click()

        Adodc1.RecordSource = " select *from BookingTable where ConsumerID= '" & cmbConsumerID.Text & "'"
        Adodc1.Refresh
    If Adodc1.Recordset.EOF Then
        
          MsgBox " Sorry Consumer Id is wrong , Please check your consumer Id", vbInformation, "Instruction"
    Else
          loadData
    End If
    Adodc2.RecordSource = "select *from NewConnection where ConsumerID= '" & cmbConsumerID.Text & "'"
    Adodc2.Refresh
    If Adodc2.Recordset.EOF Then
        MsgBox " Sorry consumer id doesn't exist ", vbInformation, "Oops"
    Else
        txtPinCode.Text = Adodc2.Recordset.Fields(7).Value
        txtAddress.Text = Adodc2.Recordset.Fields(6).Value
        txtPhoneNumber.Text = Adodc2.Recordset.Fields(4).Value
        txtConnectionType.Text = Adodc2.Recordset.Fields(9).Value
    End If
     Adodc3.RecordSource = "select *from DeliveryTable where ConsumerID= '" & cmbConsumerID.Text & "'"
    Adodc3.Refresh
     If Adodc3.Recordset.EOF Then
         txtStatus.Text = "Pending"
         DTPDelivery.Value = Date
     Else
         txtStatus.Text = "Delivered"
         DTPDelivery.Value = Adodc3.Recordset.Fields(5).Value
    End If
End Sub


Public Sub loadData()
    txtConsumerName.Text = Adodc1.Recordset.Fields(2).Value
    txtCylinderNo.Text = Adodc1.Recordset.Fields(3).Value
    DTPBooking.Value = Adodc1.Recordset.Fields(4).Value
End Sub

Private Sub cmdBill_Click()
If txtStatus.Text = "Pending" Then
    MsgBox " Delivery is pending please confirm delivery first "
Else
frmBill.Show
Unload Me
End If
End Sub

Private Sub cmdCancel_Click()
ans1 = MsgBox(" Do you want to cancel the delivery.", vbQuestion + vbYesNo, "Message")
If (ans1 = vbYes) Then
Adodc3.Recordset.Delete
MsgBox " Your delivery cancel Succefully ", vbInformation, "Succefull Message"
Adodc4.Recordset.MoveFirst
    deli = Adodc4.Recordset.Fields(1).Value
    deli = deli - txtCylinderNo.Text
    Adodc4.Recordset.Update
    Adodc4.Recordset.Fields(1).Value = deli
    Adodc4.Recordset.Update
ResetData
End If
End Sub

Private Sub cmdConfirmDelivery_Click()
If txtStatus.Text = "Pending" Then
MsgBox " Delivery is pending Please Confirm the delivery status "
Else
If cmbConsumerID.Text = "" Then
    MsgBox " Please Enter Consumer id or other fields first ", vbInformation, "Instruction"
Else
Adodc3.RecordSource = " select *from DeliveryTable where ConsumerID= '" & cmbConsumerID.Text & "'"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
ans = MsgBox(" Do you want to confirm the delivery ", vbQuestion + vbYesNo, "Message")
If (ans = vbYes) Then
         Adodc3.Recordset.AddNew
With Adodc3.Recordset
        .Fields(1).Value = cmbConsumerID.Text
        .Fields(2).Value = txtConsumerName.Text
        .Fields(3).Value = txtConnectionType.Text
        .Fields(4).Value = DTPBooking.Value
        .Fields(5).Value = DTPDelivery.Value
        .Fields(6).Value = txtPinCode.Text
        .Fields(7).Value = txtAddress.Text
        .Fields(8).Value = txtPhoneNumber.Text
    End With
    Adodc3.Recordset.Update
    Adodc3.Refresh
    MsgBox " Congratulation your order is delivered Succefull!!", vbInformation + vbOKOnly, "Saved Successfully"
     cmbConsumerID.Text = ""
    Adodc4.Recordset.MoveFirst
    deli = Adodc4.Recordset.Fields(1).Value
    deli = deli + txtCylinderNo.Text
    Adodc4.Recordset.Update
    Adodc4.Recordset.Fields(1).Value = deli
    Adodc4.Recordset.Update
    ResetData
    cmbConsumerID.SetFocus
End If
Else
          MsgBox " Delivery already confirm .", vbCritical, "Information"
End If
End If
End If
End Sub
Public Sub ResetData()
    cmbConsumerID.Text = ""
    txtConsumerName.Text = ""
    txtConnectionType.Text = ""
    DTPBooking.Value = Date
    DTPDelivery.Value = Date
    txtPinCode.Text = ""
    txtAddress.Text = ""
    txtPhoneNumber.Text = ""
    txtStatus.Text = ""
    txtCylinderNo.Text = ""
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Adodc1.RecordSource = " Select ConsumerID from BookingTable"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
While Adodc1.Recordset.EOF = False
cmbConsumerID.AddItem Adodc1.Recordset.Fields(0).Value
Adodc1.Recordset.MoveNext
Wend
End If
End Sub

Private Sub lblAllList_Click()
Adodc1.RecordSource = "Select * from BookingTable"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub
