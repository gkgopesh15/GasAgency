VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBookingReport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   $"frmBookingReport.frx":0000
   ClientHeight    =   10512
   ClientLeft      =   4284
   ClientTop       =   3024
   ClientWidth     =   17820
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10512
   ScaleWidth      =   17820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdViewAll 
      BackColor       =   &H0000FF00&
      Caption         =   "View all Records"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   672
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9000
      Width           =   4452
   End
   Begin VB.TextBox Text3 
      DataField       =   "No"
      DataSource      =   "Adodc1"
      Height          =   288
      Left            =   480
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1212
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   2172
      _ExtentX        =   3831
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
      Height          =   612
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9000
      Width           =   1812
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
      Left            =   9000
      TabIndex        =   7
      Top             =   3120
      Width           =   2892
   End
   Begin VB.TextBox txtConsumerID 
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
      Left            =   2760
      TabIndex        =   6
      Top             =   3120
      Width           =   2172
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmBookingReport.frx":009B
      Height          =   4812
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   17412
      _ExtentX        =   30713
      _ExtentY        =   8488
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   12632256
      DefColWidth     =   208
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   26
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
   Begin VB.Image Image2 
      Height          =   1452
      Left            =   14640
      Picture         =   "frmBookingReport.frx":00B0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2412
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Reports "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   864
      Left            =   7080
      TabIndex        =   10
      Top             =   1680
      Width           =   4368
   End
   Begin VB.Label Label4 
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
      Left            =   6120
      TabIndex        =   5
      Top             =   3120
      Width           =   2352
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
      TabIndex        =   4
      Top             =   3120
      Width           =   1824
   End
   Begin VB.Image Image1 
      Height          =   1452
      Left            =   720
      Picture         =   "frmBookingReport.frx":25EA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2412
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bharat Gas  Service"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   744
      Left            =   6000
      TabIndex        =   2
      Top             =   0
      Width           =   5784
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
      ForeColor       =   &H0000FF00&
      Height          =   744
      Left            =   5880
      TabIndex        =   1
      Top             =   840
      Width           =   6828
   End
   Begin VB.Label Label1 
      BackColor       =   &H00420CB4&
      Height          =   1692
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17892
   End
End
Attribute VB_Name = "frmBookingReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdSearchbtn_Click()
Adodc1.RecordSource = "select *from BookingTable where BookingDate= " & DTPDate1.Value & ""
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox " Record Not found", vbCritical, "Warning"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End Sub

Private Sub cmdViewAll_Click()
Adodc1.RecordSource = "Select * from BookingTable"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DTPDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub txtConsumerID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtConsumerID.Text = "" Then
    MsgBox " Please ender ID ", vbInformation
    txtConsumerID.SetFocus
Else
Adodc1.RecordSource = "Select * from BookingTable where ConsumerID='" + txtConsumerID.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End If
End If
    
End Sub

Private Sub txtConsumerName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtConsumerName.Text = "" Then
    MsgBox " Please enter name ", vbInformation
    txtConsumerName.SetFocus
Else
Adodc1.RecordSource = "Select * from BookingTable where ConsumerName='" + txtConsumerName.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End If
End If
End Sub
