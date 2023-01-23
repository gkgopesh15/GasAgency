VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewConnection 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"frmNewConnection.frx":0000
   ClientHeight    =   12372
   ClientLeft      =   5772
   ClientTop       =   276
   ClientWidth     =   13872
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12372
   ScaleWidth      =   13872
   Begin VB.TextBox txtState 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3120
      TabIndex        =   44
      Top             =   7080
      Width           =   3975
   End
   Begin VB.TextBox txtDistrict 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3120
      TabIndex        =   43
      Top             =   6240
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12840
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtConsumerNo 
      DataField       =   "ConsumerNo"
      DataSource      =   "Adodc1"
      Height          =   732
      Left            =   12720
      TabIndex        =   40
      Text            =   "Text8"
      Top             =   3120
      Visible         =   0   'False
      Width           =   612
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   10200
      Top             =   2040
      Visible         =   0   'False
      Width           =   2532
      _ExtentX        =   4466
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
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFF00&
      Caption         =   "<>  Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   11160
      Width           =   3015
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H00FFFF00&
      Caption         =   "@  SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   11160
      Width           =   3015
   End
   Begin VB.ComboBox cmbCylinderNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   444
      Left            =   3120
      TabIndex        =   34
      Text            =   "Combo1"
      Top             =   10440
      Width           =   3975
   End
   Begin VB.CommandButton cmdUpload 
      BackColor       =   &H00FFFF00&
      Caption         =   ">>  Upload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CheckBox chkComercial 
      BackColor       =   &H8000000E&
      Caption         =   "Comercial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   5280
      TabIndex        =   29
      Top             =   9600
      Width           =   1815
   End
   Begin VB.CheckBox chkDomestic 
      BackColor       =   &H8000000E&
      Caption         =   "Domestic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3240
      TabIndex        =   28
      Top             =   9600
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   3240
      TabIndex        =   25
      Top             =   3000
      Width           =   3735
      Begin VB.OptionButton optFemale 
         BackColor       =   &H8000000E&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   1800
         TabIndex        =   27
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton optMale 
         BackColor       =   &H8000000E&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   0
         TabIndex        =   26
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.TextBox txtDepositeAmount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3120
      TabIndex        =   24
      Top             =   11280
      Width           =   3975
   End
   Begin VB.TextBox txtDealerName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3120
      TabIndex        =   23
      Top             =   8760
      Width           =   3975
   End
   Begin VB.TextBox txtPinCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3120
      TabIndex        =   22
      Top             =   7920
      Width           =   3975
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3120
      TabIndex        =   21
      Top             =   5400
      Width           =   3975
   End
   Begin VB.TextBox txtEmailID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3120
      TabIndex        =   20
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox txtMobileNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3120
      TabIndex        =   19
      Top             =   3960
      Width           =   3975
   End
   Begin VB.TextBox txtConsumerName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3120
      TabIndex        =   18
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   4215
      Left            =   7320
      TabIndex        =   9
      Top             =   6720
      Width           =   6375
      Begin VB.CheckBox chkElectricBill 
         BackColor       =   &H8000000E&
         Caption         =   "Electric Bill"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   3480
         TabIndex        =   39
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CheckBox chkPassport 
         BackColor       =   &H8000000E&
         Caption         =   "Passport"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   3480
         TabIndex        =   16
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox chkDrivingLicence 
         BackColor       =   &H8000000E&
         Caption         =   "Driving Licence"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   3480
         TabIndex        =   15
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox chkRationCard 
         BackColor       =   &H8000000E&
         Caption         =   "Ration Card"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   480
         TabIndex        =   14
         Top             =   3120
         Width           =   2412
      End
      Begin VB.CheckBox chkVoterID 
         BackColor       =   &H8000000E&
         Caption         =   "Voter Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CheckBox chkAadhar 
         BackColor       =   &H8000000E&
         Caption         =   "Aadhar(UI)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Left            =   4320
         TabIndex        =   17
         Top             =   360
         Width           =   75
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proof of Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Left            =   2400
         TabIndex        =   12
         Top             =   1080
         Width           =   2085
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Left            =   3000
         TabIndex        =   11
         Top             =   360
         Width           =   570
      End
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   348
      Left            =   480
      TabIndex        =   42
      Top             =   6240
      Width           =   876
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   348
      Left            =   480
      TabIndex        =   41
      Top             =   7080
      Width           =   660
   End
   Begin VB.Image Image2 
      Height          =   1812
      Left            =   0
      Picture         =   "frmNewConnection.frx":0089
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3852
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New  Connection  Registration"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   732
      Left            =   4560
      TabIndex        =   38
      Top             =   480
      Width           =   7068
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C000&
      Height          =   1815
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   13935
   End
   Begin VB.Label lblConsumerID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BH00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   348
      Left            =   11400
      TabIndex        =   32
      Top             =   6120
      Width           =   708
   End
   Begin VB.Label lebel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consumer ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   9120
      TabIndex        =   31
      Top             =   6120
      Width           =   1650
   End
   Begin VB.Image imgPhoto 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Cylinder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   480
      TabIndex        =   30
      Top             =   10440
      Width           =   1860
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deposite Amount "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   480
      TabIndex        =   8
      Top             =   11280
      Width           =   2250
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   480
      TabIndex        =   7
      Top             =   9600
      Width           =   2172
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   480
      TabIndex        =   6
      Top             =   8760
      Width           =   1668
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pin Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   480
      TabIndex        =   5
      Top             =   7920
      Width           =   1176
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   480
      TabIndex        =   4
      Top             =   5400
      Width           =   1056
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   480
      TabIndex        =   3
      Top             =   4680
      Width           =   1032
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   480
      TabIndex        =   2
      Top             =   3840
      Width           =   1380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   480
      TabIndex        =   1
      Top             =   3120
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cunsumer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   2160
   End
End
Attribute VB_Name = "frmNewConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AddProof As String
Dim C_no As Integer
Dim Address As String


Private Sub chkComercial_Click()
If chkComercial.Value = 1 Then
    chkDomestic.Value = 0
Else
    chkDomestic.Value = 1
End If
End Sub

Private Sub chkComercial_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmbCylinderNo.SetFocus
 End If
End Sub

Private Sub chkDomestic_Click()
If chkDomestic.Value = 1 Then
    chkComercial.Value = 0
Else
    chkComercial.Value = 1
End If
End Sub

Private Sub chkDomestic_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
    cmbCylinderNo.SetFocus
 End If
End Sub

Private Sub cmbCylinderNo_Click()
If cmbCylinderNo.Text = "Single Connection" Then
 txtDepositeAmount.Text = Val(2800)
 Else
 txtDepositeAmount.Text = Val(4400)
 End If
End Sub

Private Sub cmbCylinderNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If cmbCylinderNo.Text = "" Then
        MsgBox " Please select No. of Cylinder connection.", Message
        cmbCylinderNo.SetFocus
        Else
        txtDepositeAmount.SetFocus
        End If
    End If
End Sub

Private Sub cmdSubmit_Click()
AddProof = ""
        If chkAadhar.Value = 1 Then
        AddProof = AddProof & "  Aadhar  "
        End If
        If chkDrivingLicence.Value = 1 Then
        AddProof = AddProof & "  Driving Licence "
        End If
        If chkVoterID.Value = 1 Then
        AddProof = AddProof & "  Voter Id  "
        End If
        If chkPassport.Value = 1 Then
        AddProof = AddProof & "  Passport  "
        End If
        If chkRationCard.Value = 1 Then
        AddProof = AddProof & "  Ration Card  "
        End If
        If chkElectricBill.Value = 1 Then
        AddProof = AddProof & "  Electric Bill  "
        End If
If txtConsumerName.Text = "" Or txtMobileNo.Text = "" Or txtEmailID.Text = "" Or txtAddress.Text = "" _
Or txtPinCode.Text = "" Or txtDealerName.Text = "" Or cmbCylinderNo.Text = "" Or AddProof = "" Then
    MsgBox " Opps something is missing please check all fields !!!", vbExclamation, "instruction"
Else
Adodc1.Recordset.AddNew
With Adodc1.Recordset
        .Fields(0).Value = C_no
        .Fields(1).Value = lblConsumerID.Caption
        .Fields(2).Value = txtConsumerName.Text
        If optMale.Value = True Then
            .Fields(3).Value = "Male"
        Else
            .Fields(3).Value = "Female"
        End If
        .Fields(4).Value = txtMobileNo.Text
        .Fields(5).Value = txtEmailID.Text
        Address = txtAddress.Text & ", " & txtDistrict.Text & ", " & txtState.Text
        .Fields(6).Value = Address
        .Fields(7).Value = txtPinCode.Text
        .Fields(8).Value = txtDealerName.Text
        If chkDomestic.Value = 1 Then
            .Fields(9).Value = "Domestic"
        Else
            .Fields(9).Value = "Commercial"
        End If
        .Fields(10).Value = cmbCylinderNo.Text
        .Fields(11).Value = txtDepositeAmount.Text
        AddProof = ""
        If chkAadhar.Value = 1 Then
        AddProof = AddProof & "  Aadhar  "
        End If
        If chkDrivingLicence.Value = 1 Then
        AddProof = AddProof & "  Driving Licence "
        End If
        If chkVoterID.Value = 1 Then
        AddProof = AddProof & "  Voter Id  "
        End If
        If chkPassport.Value = 1 Then
        AddProof = AddProof & "  Passport  "
        End If
        If chkRationCard.Value = 1 Then
        AddProof = AddProof & "  Ration Card  "
        End If
        If chkElectricBill.Value = 1 Then
        AddProof = AddProof & "  Electric Bill  "
        End If
        .Fields(12).Value = AddProof
        .Fields(13).Value = lblDate.Caption
    End With
    SavePicture
    Adodc1.Recordset.Update
    Adodc1.Refresh
    MsgBox " New Consumer Register Succefully!!", vbInformation + vbOKOnly, "Saved Successfully"
     txtConsumerName.Text = ""
    ResetData
    txtConsumerName.SetFocus
End If
End Sub

Private Sub cmdUpload_Click()
CommonDialog1.Filter = "jpeg|*.jpg"
CommonDialog1.ShowOpen
imgPhoto.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub cmdUpload_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        chkAadhar.SetFocus
    End If
End Sub

Private Sub Form_Load()
chkDomestic.Value = 1
lblDate.Caption = Date
cmbCylinderNo.AddItem "Single Connection"
cmbCylinderNo.AddItem "Double Connection"
optMale.Value = True
Adodc1.Recordset.MoveLast
C_no = Adodc1.Recordset.Fields(0).Value
C_no = C_no + 1
lblConsumerID.Caption = lblConsumerID.Caption & C_no
End Sub

Private Sub optFemale_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        txtMobileNo.SetFocus
    End If
End Sub
Private Sub optMale_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtMobileNo.SetFocus
    End If
End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtAddress.Text = "" Then
        MsgBox " Please enter Address.", Message
        txtAddress.SetFocus
        Else
        txtDistrict.SetFocus
        End If
    End If
End Sub

Private Sub txtConsumerName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtConsumerName.Text = "" Then
        MsgBox " Please enter Consumer name.", Message
        txtConsumerName.SetFocus
        Else
        optMale.SetFocus
        End If
    End If
End Sub
Private Sub txtDealerName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtDealerName.Text = "" Then
        MsgBox " Please enter Dealer name.", Message
        txtDealerName.SetFocus
        Else
        chkDomestic.SetFocus
        End If
    End If
End Sub

Private Sub txtDepositeAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtDepositeAmount.Text = "" Then
        MsgBox " Please enter deposite amount.", Message
        txtDepositeAmount.SetFocus
        Else
        cmdUpload.SetFocus
        End If
    End If
End Sub

Private Sub txtDistirict_Change()

End Sub

Private Sub txtDistrict_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtDistrict.Text = "" Then
        MsgBox " Please enter District.", Message
        txtDistrict.SetFocus
        Else
        txtState.SetFocus
        End If
    End If
End Sub

Private Sub txtEmailID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtEmailID.Text = "" Then
        MsgBox " Please enter email id.", Message
        txtEmailID.SetFocus
        Else
        txtAddress.SetFocus
        End If
    End If
End Sub

Private Sub txtMobileNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtMobileNo.Text = "" Then
        MsgBox " Please enter Mobile No.", vbInformation, "Message"
        txtMobileNo.SetFocus
        Else
        txtEmailID.SetFocus
        End If
    End If
End Sub
Private Sub txtPinCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtPinCode.Text = "" Then
        MsgBox " Please enter Pin code.", Message
        txtPinCode.SetFocus
        Else
        txtDealerName.SetFocus
        End If
    End If
End Sub
Public Sub SavePicture()
    On Error Resume Next
    Dim st As ADODB.Stream
    Set st = New ADODB.Stream
    st.Type = adTypeBinary
    st.Open
    st.LoadFromFile (CommonDialog1.FileName)
    Adodc1.Recordset.Fields(14).Value = st.Read
    st.Close
End Sub
Public Sub ResetData()
    optMale.Value = True
    txtMobileNo.Text = ""
    txtEmailID.Text = ""
    txtAddress.Text = ""
    txtPinCode.Text = ""
    txtDealerName.Text = ""
    chkDomestic.Value = 0
    chkComercial.Value = 0
    cmbCylinderNo.Text = ""
    txtDepositeAmount.Text = ""
    chkAadhar.Value = 0
    chkDrivingLicence.Value = 0
    chkVoterID.Value = 0
    chkRationCard.Value = 0
    chkPassport.Value = 0
    chkElectricBill.Value = 0
    txtDistrict.Text = ""
    txtState.Text = ""
    C_no = C_no + 1
    lblConsumerID.Caption = lblConsumerID.Caption & C_no
    imgPhoto.Picture = LoadPicture("")
End Sub

Private Sub txtState_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
        If txtState.Text = "" Then
        MsgBox " Please enter State.", Message
        txtState.SetFocus
        Else
        txtPinCode.SetFocus
        End If
    End If
End Sub
