VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmStock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   $"frmStock.frx":0000
   ClientHeight    =   8940
   ClientLeft      =   5880
   ClientTop       =   3540
   ClientWidth     =   12972
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Palette         =   "frmStock.frx":0088
   Picture         =   "frmStock.frx":5318C
   ScaleHeight     =   8940
   ScaleWidth      =   12972
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   9360
      Top             =   3960
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
      RecordSource    =   "select * from Stock"
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
   Begin VB.TextBox Text1 
      Height          =   732
      Left            =   11640
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Add Stock"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3492
      Left            =   1800
      TabIndex        =   9
      Top             =   4800
      Width           =   9012
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   19.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2520
         Width           =   2052
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   19.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2520
         Width           =   2052
      End
      Begin VB.TextBox txtQuantity 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
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
         Left            =   5040
         TabIndex        =   11
         Top             =   720
         Width           =   2172
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   444
         Left            =   1200
         TabIndex        =   10
         Top             =   720
         Width           =   1272
      End
   End
   Begin VB.TextBox txtAvailable 
      DataField       =   "Available"
      DataSource      =   "Adodc1"
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
      Left            =   8880
      TabIndex        =   8
      Top             =   2640
      Width           =   2172
   End
   Begin VB.TextBox txtDeliver 
      DataField       =   "Delivered"
      DataSource      =   "Adodc1"
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
      Left            =   5760
      TabIndex        =   6
      Top             =   2640
      Width           =   2172
   End
   Begin VB.TextBox txtBooked 
      DataField       =   "Booked"
      DataSource      =   "Adodc1"
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
      Left            =   2640
      TabIndex        =   5
      Top             =   2640
      Width           =   2172
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   480
      Left            =   8880
      TabIndex        =   7
      Top             =   1920
      Width           =   2076
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivered"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   480
      Left            =   5760
      TabIndex        =   4
      Top             =   1920
      Width           =   2004
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booked "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   480
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   1716
   End
   Begin VB.Image Image2 
      Height          =   1212
      Left            =   480
      Picture         =   "frmStock.frx":B2C22
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1452
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "( Bharat  Gas  Service ) "
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002448A8&
      Height          =   456
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   4464
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bhagalpur Gas Agency "
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002448A8&
      Height          =   540
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   5172
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Height          =   1452
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12972
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label6_Click()

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim a As Integer, b As Integer
a = txtAvailable.Text
b = a + txtQuantity
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Update
Adodc1.Recordset.Fields(2).Value = b
Adodc1.Recordset.Update
End Sub

Private Sub Form_Load()
Adodc1.Recordset.MoveFirst
    avai = Adodc1.Recordset.Fields(2).Value
    avai = avai - txtDeliver.Text
    txtAvailable.Text = avai
End Sub

