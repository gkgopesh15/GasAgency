VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmReset 
   Caption         =   "                                                                                     Reset Password"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3960
      Top             =   7560
      Width           =   2535
      _ExtentX        =   4466
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
      Connect         =   "DSN=DSNgas"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "DSNgas"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from LoginTable"
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
      Height          =   975
      Left            =   5520
      Picture         =   "frmReset.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   2655
   End
   Begin VB.CommandButton cmdConfirm 
      Height          =   975
      Left            =   1200
      Picture         =   "frmReset.frx":1629
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   2655
   End
   Begin VB.TextBox txtReNewPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5040
      TabIndex        =   7
      Top             =   4680
      Width           =   3615
   End
   Begin VB.TextBox txtNewPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label lblMessage3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Don't match please enter Same as new password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   5340
   End
   Begin VB.Label lblMessage2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Please enter new pasword"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   6000
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Label lblMessage1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Please enter new pasword"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   6000
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Label lblUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gopesh Kumar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   5040
      TabIndex        =   5
      Top             =   2400
      Width           =   3600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Enter New Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   4305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   3675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   120
      Picture         =   "frmReset.frx":2C76
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reset Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   28.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   690
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   4920
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirm_Click()
Adodc1.RecordSource = "select * from LoginTable where UserName='" _
& lblUserName.Caption & "'"
Adodc1.Refresh
If txtNewPass.Text = txtReNewPass.Text Then
Adodc1.Recordset.Fields("Password") = txtNewPass.Text
Adodc1.Recordset.Update
MsgBox "Your Password is Succefully changed"
Else
lblMessage3.Visible = True
txtReNewPass.Text = ""
txtReNewPass.SetFocus
End If
End Sub

Private Sub cmdConfirm_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdConfirm_Click
End If

End Sub

Private Sub Form_Load()
lblReset
End Sub

Private Sub txtNewPass_Click()
lblMessage1.Visible = False
End Sub

Private Sub txtNewPass_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtNewPass.Text = "" Then
    lblMessage1.Visible = True
    txtNewPass.SetFocus
    Else
    txtReNewPass.SetFocus
    End If
End If
End Sub

Private Sub txtReNewPass_Click()
lblMessage2.Visible = False
lblMessage3.Visible = False
End Sub

Private Sub txtReNewPass_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 If txtReNewPass.Text = "" Then
    lblMessage2.Visible = True
    txtReNewPass.SetFocus
 Else
    cmdConfirm.SetFocus
 End If
End If
End Sub
