VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBill 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "fgh"
   ClientHeight    =   11496
   ClientLeft      =   9240
   ClientTop       =   1680
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11496
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFF00&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9000
      Width           =   1692
   End
   Begin VB.TextBox Text2 
      DataField       =   "Comercial"
      DataSource      =   "Adodc1"
      Height          =   372
      Left            =   1320
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   0
      Visible         =   0   'False
      Width           =   1332
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   492
      Left            =   480
      Top             =   6840
      Visible         =   0   'False
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   868
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
      RecordSource    =   "select *from SetPriceTable"
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
   Begin VB.Image Image3 
      Height          =   1212
      Left            =   6720
      Picture         =   "frmBill.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gopesh Kumar"
      BeginProperty Font 
         Name            =   "Kunstler Script"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   432
      Left            =   360
      TabIndex        =   36
      Top             =   8520
      Width           =   2220
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Signature"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   444
      Left            =   360
      TabIndex        =   35
      Top             =   7920
      Width           =   2352
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " GST No.      :     Bharat8051gstin"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   288
      Left            =   240
      TabIndex        =   34
      Top             =   1560
      Width           =   3216
   End
   Begin VB.Label lblGST 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   6360
      TabIndex        =   33
      Top             =   6720
      Width           =   96
   End
   Begin VB.Label lblPrice 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   6360
      TabIndex        =   32
      Top             =   6120
      Width           =   96
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GST  8%  :-"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   444
      Left            =   3720
      TabIndex        =   31
      Top             =   6720
      Width           =   1632
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price        :- "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   444
      Left            =   3720
      TabIndex        =   30
      Top             =   6120
      Width           =   1764
   End
   Begin VB.Label lblConnectionType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   396
      Left            =   3600
      TabIndex        =   28
      Top             =   4560
      Width           =   84
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection Type"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   396
      Left            =   240
      TabIndex        =   27
      Top             =   4560
      Width           =   2076
   End
   Begin VB.Label lblContactNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   396
      Left            =   3600
      TabIndex        =   25
      Top             =   5520
      Width           =   84
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   396
      Left            =   240
      TabIndex        =   24
      Top             =   5520
      Width           =   2232
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   324
      Left            =   5520
      TabIndex        =   23
      Top             =   2040
      Width           =   72
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   324
      Left            =   5640
      TabIndex        =   22
      Top             =   1560
      Width           =   72
   End
   Begin VB.Shape Shape1 
      Height          =   12
      Left            =   0
      Top             =   2520
      Width           =   7932
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8051687313"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   324
      Left            =   1920
      TabIndex        =   21
      Top             =   2040
      Width           =   1320
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No. :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   324
      Left            =   240
      TabIndex        =   20
      Top             =   2040
      Width           =   1476
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Signature "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   444
      Left            =   4680
      TabIndex        =   19
      Top             =   7920
      Width           =   2832
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   6360
      TabIndex        =   18
      Top             =   7320
      Width           =   96
   End
   Begin VB.Label lblNoofCylinder 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   396
      Left            =   3600
      TabIndex        =   17
      Top             =   5040
      Width           =   84
   End
   Begin VB.Label lblBookingDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   396
      Left            =   3720
      TabIndex        =   16
      Top             =   4080
      Width           =   84
   End
   Begin VB.Label lblDeliveryDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   396
      Left            =   3720
      TabIndex        =   15
      Top             =   3600
      Width           =   84
   End
   Begin VB.Label lblConsumerName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   396
      Left            =   3720
      TabIndex        =   14
      Top             =   3120
      Width           =   84
   End
   Begin VB.Label lblConsumerID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   396
      Left            =   3720
      TabIndex        =   13
      Top             =   2640
      Width           =   84
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   324
      Left            =   4320
      TabIndex        =   12
      Top             =   2040
      Width           =   528
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   324
      Left            =   4320
      TabIndex        =   11
      Top             =   1560
      Width           =   696
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Thank    You  "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   588
      Left            =   2160
      TabIndex        =   10
      Top             =   8880
      Width           =   2580
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total       :- "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   444
      Left            =   3720
      TabIndex        =   9
      Top             =   7320
      Width           =   1728
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. of  Cylinder :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   396
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   2208
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BookingDate   :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   396
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   2016
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   396
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   1956
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consumer Name :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   396
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   2184
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consumer ID :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   396
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1836
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Cook Love  Serve  Love)"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   3804
   End
   Begin VB.Image Image2 
      Height          =   1932
      Left            =   0
      Picture         =   "frmBill.frx":253A
      Stretch         =   -1  'True
      Top             =   9480
      Width           =   7932
   End
   Begin VB.Image Image1 
      Height          =   1212
      Left            =   120
      Picture         =   "frmBill.frx":53AA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Bharat Gas)"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   25.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002448A8&
      Height          =   540
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   2880
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bhagalpur Gas Agency "
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
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   4296
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Height          =   1452
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7932
   End
End
Attribute VB_Name = "frmBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Amount As Integer

Private Sub cmdPrint_Click()
cmdPrint.Visible = False
PrintForm
End Sub

Private Sub Form_Load()
lblDate.Caption = Date
lblTime.Caption = Time
Bill
If lblConnectionType.Caption = "Comercial" Then
    Amount = Adodc1.Recordset.Fields(0).Value
Else
    Amount = Adodc1.Recordset.Fields(1).Value
End If
lblPrice.Caption = Amount * Val(lblNoofCylinder.Caption)
lblGST.Caption = (Val(lblPrice.Caption) * 8) \ 100
lblTotal.Caption = Val(lblPrice.Caption) + Val(lblGST.Caption)
End Sub

