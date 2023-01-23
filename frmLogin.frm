VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                                                                        Login"
   ClientHeight    =   8472
   ClientLeft      =   6300
   ClientTop       =   2172
   ClientWidth     =   11172
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8472
   ScaleWidth      =   11172
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4440
      Top             =   7920
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3360
      Top             =   7800
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   840
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3620
      _ExtentY        =   593
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
      Connect         =   "DSN=DSNgas"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "DSNgas"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from LoginTable"
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
   Begin VB.CommandButton cmdLogin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      Picture         =   "frmLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   2052
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Picture         =   "frmLogin.frx":1898
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   2052
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   516
      IMEMode         =   3  'DISABLE
      Left            =   6360
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   4800
      Width           =   3612
   End
   Begin VB.ComboBox cmbUser 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      ItemData        =   "frmLogin.frx":30C9
      Left            =   6360
      List            =   "frmLogin.frx":30CB
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Image Image4 
      Height          =   2295
      Left            =   1800
      Picture         =   "frmLogin.frx":30CD
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   7335
   End
   Begin VB.Label lblRegister 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   7920
      TabIndex        =   11
      Top             =   7920
      Width           =   1035
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "New User                      Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   7920
      Width           =   5175
   End
   Begin VB.Label lblForgotPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   7440
      TabIndex        =   9
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Label lblHidePassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hide  Password"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   7680
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label lblShowPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Password"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   7680
      TabIndex        =   6
      Top             =   5640
      Width           =   2250
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   4320
      TabIndex        =   4
      Top             =   4920
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   4320
      TabIndex        =   2
      Top             =   3840
      Width           =   1635
   End
   Begin VB.Image Image3 
      Height          =   3375
      Left            =   720
      Picture         =   "frmLogin.frx":5200
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BHAGALPUR GAS AGENCY"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   6975
   End
   Begin VB.Image Image2 
      Height          =   2055
      Left            =   9240
      Picture         =   "frmLogin.frx":9677
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   0
      Picture         =   "frmLogin.frx":BBB1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbUser_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtPassword.SetFocus
End If

End Sub

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdLogin_Click()
Adodc1.RecordSource = " select *From LoginTable where UserName = '" + cmbUser.Text + _
"'And Password = '" + txtPassword.Text + "'"
    Adodc1.Refresh
    If Adodc1.Recordset.EOF Then
    MsgBox "Login failed, Try Again..!!!", vbCritical, "Please enter correct UserName and Password"
    txtPassword.Text = ""
    txtPassword.SetFocus
    frmLogin.Show
    Else
        frmSplash.Show
        Unload Me
    End If
    Adodc1.Refresh
End Sub

Private Sub cmdLogin_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdLogin_Click
End If

End Sub

Private Sub Form_Load()
Adodc1.RecordSource = " Select UserName from LoginTable"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
While Adodc1.Recordset.EOF = False
cmbUser.AddItem Adodc1.Recordset.Fields(0).Value
Adodc1.Recordset.MoveNext
Wend
End If
txtPassword.Text = ""
End Sub

Private Sub lblForgotPass_Click()
frmForgot.Show
lblForgot
frmLogin.Hide
End Sub

Private Sub lblHidePassword_Click()
txtPassword.PasswordChar = "*"
lblHidePassword.Visible = False
lblShowPassword.Visible = True
End Sub

Private Sub lblRegister_Click()
frmNewUser.Show
Unload Me
End Sub

Private Sub lblShowPassword_Click()
If txtPassword.Text = "" Then
MsgBox "Please Enter Password First"
lblShowPassword.Visible = True
Else
 txtPassword.PasswordChar = ""
 lblShowPassword.Visible = False
 lblHidePassword.Visible = True
End If
End Sub

Private Sub Timer1_Timer()
lblRegister.ForeColor = vbRed
End Sub

Private Sub Timer2_Timer()
lblRegister.ForeColor = vbBlue
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdLogin.SetFocus
End If

End Sub
