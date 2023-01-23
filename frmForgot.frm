VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmForgot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                                                                          Password Recovery"
   ClientHeight    =   8760
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11760
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   7200
      Top             =   7920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "select * from QuesTable"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5640
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2138
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
      Height          =   855
      Left            =   8760
      Picture         =   "frmForgot.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton cmdOk 
      Height          =   855
      Left            =   4440
      Picture         =   "frmForgot.frx":1629
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7560
      Width           =   2415
   End
   Begin VB.TextBox txtAns2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5760
      TabIndex        =   11
      Top             =   6240
      Width           =   5655
   End
   Begin VB.TextBox txtAns1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5760
      TabIndex        =   10
      Top             =   3840
      Width           =   5655
   End
   Begin VB.ComboBox cmbQ2 
      BackColor       =   &H00C0FFC0&
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
      Height          =   555
      Left            =   5520
      TabIndex        =   9
      Top             =   5160
      Width           =   5895
   End
   Begin VB.ComboBox cmbQ1 
      BackColor       =   &H00C0FFC0&
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
      Height          =   555
      Left            =   5520
      TabIndex        =   5
      Top             =   2760
      Width           =   5895
   End
   Begin VB.Label lblwarning2 
      BackStyle       =   0  'Transparent
      Caption         =   "Oops wrong answer please enter write answer "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6600
      TabIndex        =   13
      Top             =   6960
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblwarning1 
      BackStyle       =   0  'Transparent
      Caption         =   "Oops wrong answer please enter write answer "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Ans1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   21.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   4440
      TabIndex        =   8
      Top             =   3840
      Width           =   1260
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Q2."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   21.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   4440
      TabIndex        =   7
      Top             =   5160
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Ans2."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   21.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   4440
      TabIndex        =   6
      Top             =   6240
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Q1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   21.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   4440
      TabIndex        =   4
      Top             =   2760
      Width           =   765
   End
   Begin VB.Label lbluserName 
      BackColor       =   &H00C0FFFF&
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
      Height          =   615
      Left            =   6720
      TabIndex        =   3
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Label Label3 
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
      Left            =   4440
      TabIndex        =   2
      Top             =   1800
      Width           =   1980
   End
   Begin VB.Image Image3 
      Height          =   7335
      Left            =   0
      Picture         =   "frmForgot.frx":2894
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   120
      Picture         =   "frmForgot.frx":498A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password Recovery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   28.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   690
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   5940
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
   End
   Begin VB.Image Image1 
      Height          =   7335
      Left            =   4080
      Picture         =   "frmForgot.frx":B999
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   7635
   End
End
Attribute VB_Name = "frmForgot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbQ1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 txtAns1.SetFocus
End If
End Sub

Private Sub cmbQ2_Click()
Adodc2.RecordSource = " select *From LoginTable where Ques1= '" + cmbQ1.Text + _
"'And Ans1 = '" + txtAns1.Text + "'"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount = 1 Then
    txtAns2.SetFocus
    Else
    txtAns1.Text = ""
    txtAns1.SetFocus
    lblwarning1.Visible = True
    cmbQ2.Text = ""
    End If
    Adodc1.Refresh
    If txtAns1.Text = "" Then
    lblwarning1.Visible = True
    txtAns1.SetFocus
    End If
End Sub

Private Sub cmbQ2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 txtAns2.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
frmLogin.Show
frmForgot.Hide
End Sub

Private Sub cmdOk_Click()
 Adodc1.RecordSource = " select *From LoginTable where Ques2= '" + cmbQ2.Text + _
 "' And Ans2 = '" + txtAns2.Text + "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount = 1 Then
    frmReset.Show
    lblReset
    frmForgot.Hide
    Else
    txtAns2.Text = ""
    lblwarning2.Visible = True
    End If
    Adodc1.Refresh
    If cmbQ1.Text = "" Or txtAns1.Text = "" Or cmbQ2.Text = "" Or txtAns2.Text = "" Then
    MsgBox " Please Chose Questions and give the Answer ", vbCritical
    txtAns1.Text = ""
    txtAns2.Text = ""
    End If
End Sub

Private Sub cmdOk_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdOk_Click
End If

End Sub

Private Sub Form_Load()
Adodc2.RecordSource = " Select Ques1 from QuesTable"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
While Adodc2.Recordset.EOF = False
cmbQ1.AddItem Adodc2.Recordset.Fields(0).Value
Adodc2.Recordset.MoveNext
Wend
End If
Adodc2.RecordSource = " Select Ques2 from QuesTable"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
While Adodc2.Recordset.EOF = False
cmbQ2.AddItem Adodc2.Recordset.Fields(0).Value
Adodc2.Recordset.MoveNext
Wend
End If

End Sub

Private Sub txtAns1_Click()
lblwarning1.Visible = False
txtAns2.Text = ""
cmbQ2.Text = ""
End Sub

Private Sub txtAns1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 cmbQ2.SetFocus
End If
End Sub

Private Sub txtAns2_Click()
lblwarning2.Visible = False
lblwarning1.Visible = False
 If txtAns1.Text = "" Then
    lblwarning1.Visible = True
    txtAns1.SetFocus
 End If
End Sub

Private Sub txtAns2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
 cmdOk.SetFocus
End If
End Sub
