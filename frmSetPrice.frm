VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSetPrice 
   BackColor       =   &H00420CB4&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   $"frmSetPrice.frx":0000
   ClientHeight    =   9420
   ClientLeft      =   4920
   ClientTop       =   2808
   ClientWidth     =   16800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   16800
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      DataField       =   "Comercial"
      DataSource      =   "Adodc1"
      Height          =   732
      Left            =   13920
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   8640
      Visible         =   0   'False
      Width           =   732
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   432
      Left            =   14640
      Top             =   8760
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   762
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
      Height          =   972
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7320
      Width           =   3132
   End
   Begin VB.CommandButton cmdSet 
      BackColor       =   &H00FFFF00&
      Caption         =   "Set Price"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7320
      Width           =   3132
   End
   Begin VB.TextBox txtDomestic 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   732
      Left            =   8400
      TabIndex        =   5
      Top             =   5880
      Width           =   5052
   End
   Begin VB.TextBox txtComercial 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   780
      Left            =   8400
      TabIndex        =   4
      Top             =   4920
      Width           =   5052
   End
   Begin VB.Image Image3 
      Height          =   2292
      Left            =   2640
      Picture         =   "frmSetPrice.frx":00B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11412
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Set  Price of Cylinder  "
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   28.2
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   576
      Left            =   5400
      TabIndex        =   8
      Top             =   3600
      Width           =   5628
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domestic "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   28.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   672
      Left            =   3120
      TabIndex        =   3
      Top             =   5760
      Width           =   2076
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comercial "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   28.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   672
      Left            =   3120
      TabIndex        =   2
      Top             =   4800
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2292
      Left            =   14040
      Picture         =   "frmSetPrice.frx":21874
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2772
   End
   Begin VB.Image Image2 
      Height          =   2292
      Left            =   0
      Picture         =   "frmSetPrice.frx":23DAE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2652
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bhagalpur Gas Agency "
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   744
      Left            =   3120
      TabIndex        =   1
      Top             =   2400
      Width           =   7080
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Height          =   2292
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16812
   End
End
Attribute VB_Name = "frmSetPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSet_Click()
If txtComercial.Text = "" Or txtDomestic.Text = "" Then
MsgBox " Please enter New Price", vbCritical, "Oops"
Else
On Error Resume Next
Adodc1.Recordset.Update
Adodc1.Recordset.Fields(0).Value = txtComercial.Text
Adodc1.Recordset.Fields(1).Value = txtDomestic.Text
Adodc1.Recordset.Update
MsgBox " Price is Update succefully", vbInformation, "Congratulation"
End If
End Sub

Private Sub txtComercial_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtComercial.Text = "" Then
        MsgBox " Please Enter price", vbCritical, "Instruction"
        txtComercial.SetFocus
    Else
        txtDomestic.SetFocus
    End If
End If
End Sub

Private Sub txtDomestic_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtDomestic.Text = "" Then
        MsgBox " Please Enter price", vbCritical, "Instruction"
        txtDomestic.SetFocus
    Else
        cmdSet.SetFocus
    End If
End If
End Sub
