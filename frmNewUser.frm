VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNewUser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   $"frmNewUser.frx":0000
   ClientHeight    =   10512
   ClientLeft      =   3960
   ClientTop       =   3444
   ClientWidth     =   18096
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewUser.frx":00C2
   ScaleHeight     =   10512
   ScaleWidth      =   18096
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      DataField       =   "DOB"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   6600
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   612
      Left            =   3720
      Top             =   9120
      Visible         =   0   'False
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   1080
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
      Height          =   492
      Left            =   480
      Top             =   9000
      Visible         =   0   'False
      Width           =   2772
      _ExtentX        =   4890
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
      Height          =   612
      Left            =   15960
      Picture         =   "frmNewUser.frx":6F5D8
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9000
      Width           =   1692
   End
   Begin VB.OptionButton optAgree 
      BackColor       =   &H000000FF&
      Caption         =   "    I   agree   all    the    tearm     and      conditions "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   480
      TabIndex        =   25
      Top             =   8520
      Width           =   6852
   End
   Begin VB.CommandButton cmdSubmit 
      Height          =   612
      Left            =   13320
      Picture         =   "frmNewUser.frx":70A5E
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9000
      Width           =   1572
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4440
      TabIndex        =   23
      Top             =   3840
      Width           =   3852
      Begin VB.OptionButton optFemale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   1920
         TabIndex        =   28
         Top             =   0
         Width           =   1575
      End
      Begin VB.OptionButton optMale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSComCtl2.DTPicker DTPDob 
      Height          =   492
      Left            =   4440
      TabIndex        =   22
      Top             =   2640
      Width           =   3852
      _ExtentX        =   6795
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
      Format          =   146407425
      CurrentDate     =   43269
   End
   Begin VB.ComboBox cmbQ2 
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
      Height          =   528
      Left            =   12480
      TabIndex        =   21
      Top             =   5040
      Width           =   4935
   End
   Begin VB.ComboBox cmbQ1 
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
      Height          =   528
      Left            =   12480
      TabIndex        =   20
      Top             =   2640
      Width           =   4932
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
      Height          =   552
      Left            =   12480
      TabIndex        =   19
      Top             =   6240
      Width           =   4932
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
      Height          =   552
      Left            =   12480
      TabIndex        =   18
      Top             =   3840
      Width           =   4932
   End
   Begin VB.TextBox txtEmail 
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
      Height          =   552
      Left            =   12480
      TabIndex        =   17
      Top             =   1440
      Width           =   4932
   End
   Begin VB.TextBox txtNumber 
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
      Height          =   495
      Left            =   4440
      TabIndex        =   16
      Top             =   7680
      Width           =   3855
   End
   Begin VB.TextBox txtConfirmPassword 
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
      Height          =   495
      Left            =   4440
      TabIndex        =   15
      Top             =   6360
      Width           =   3855
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      Top             =   5040
      Width           =   3855
   End
   Begin VB.TextBox txtName 
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
      Height          =   495
      Left            =   4440
      TabIndex        =   13
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   12120
      Width           =   11295
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
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
      Height          =   432
      Left            =   9720
      TabIndex        =   12
      Top             =   6240
      Width           =   1308
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Security Ques2"
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
      Height          =   432
      Left            =   9720
      TabIndex        =   11
      Top             =   5040
      Width           =   2676
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
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
      Height          =   432
      Left            =   9840
      TabIndex        =   10
      Top             =   3840
      Width           =   1308
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Security Ques1"
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
      Height          =   432
      Left            =   9840
      TabIndex        =   9
      Top             =   2760
      Width           =   2676
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email "
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
      Height          =   432
      Left            =   9840
      TabIndex        =   8
      Top             =   1440
      Width           =   1116
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Number"
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
      Height          =   432
      Left            =   600
      TabIndex        =   7
      Top             =   7680
      Width           =   2712
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
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
      Height          =   432
      Left            =   600
      TabIndex        =   6
      Top             =   6360
      Width           =   3216
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Height          =   432
      Left            =   600
      TabIndex        =   5
      Top             =   5040
      Width           =   1740
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Height          =   432
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOB "
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
      Height          =   432
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   948
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
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New User Regisration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   552
      Left            =   7680
      TabIndex        =   1
      Top             =   240
      Width           =   4968
   End
   Begin VB.Image Image1 
      Height          =   972
      Left            =   4920
      Picture         =   "frmNewUser.frx":71EF2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2412
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1212
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18132
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSubmit_Click()
If txtName.Text = "" Or txtPassword.Text = "" Or txtConfirmPassword.Text = "" _
Or txtNumber.Text = "" Or txtEmail.Text = "" Or cmbQ1.Text = "" Or _
txtAns1.Text = "" Or cmbQ2.Text = "" Or txtAns2.Text = "" Then
    MsgBox " Opps something went wrong please check all data fields ", vbInformation, "Instruction"
    ResetData
    txtName.SetFocus
Else
    
   If optAgree.Value = False Then
        MsgBox " Do you agree with all terms and Conditions If yes Then chose first", vbInformation, "Instruction"
        optAgree.SetFocus
        optAgree.Value = False
    Else
    Adodc2.Recordset.AddNew
    
    With Adodc2.Recordset
    .Fields(0).Value = txtName.Text
    .Fields(1).Value = DTPDob.Value
    If optMale.Value = True Then
    .Fields(5).Value = "Male"
    Else
    .Fields(5).Value = "Female"
    End If
    .Fields(2).Value = txtPassword.Text
    .Fields(3).Value = txtNumber.Text
    .Fields(4).Value = txtEmail.Text
    .Fields(6).Value = cmbQ1.Text
    .Fields(7).Value = txtAns1.Text
    .Fields(8).Value = cmbQ2.Text
    .Fields(9).Value = txtAns2.Text
    End With
    Adodc2.Recordset.Update
    Adodc2.Refresh
    MsgBox "User Registered Succefully", vbInformation, "Message"
    
   End If
End If
    ResetData
    txtName.SetFocus
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtName.Text = "" Then
        MsgBox " Enter Name First", vbInformation, "Instruction"
        txtName.SetFocus
    Else
        DTPDob.SetFocus
    End If
End If
End Sub

Private Sub cmbQ1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtAns1.SetFocus
End If
End Sub

Private Sub cmbQ2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtAns2.SetFocus
End If
End Sub

Private Sub DTPDob_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    optMale.SetFocus
End If
End Sub

Private Sub Form_Load()
optMale.Value = True
Adodc1.RecordSource = " Select Ques1 from QuesTable"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
While Adodc1.Recordset.EOF = False
cmbQ1.AddItem Adodc1.Recordset.Fields(0).Value
Adodc1.Recordset.MoveNext
Wend
End If
Adodc1.RecordSource = " Select Ques2 from QuesTable"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
While Adodc1.Recordset.EOF = False
cmbQ2.AddItem Adodc1.Recordset.Fields(0).Value
Adodc1.Recordset.MoveNext
Wend
End If

End Sub

Private Sub optAgree_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdSubmit.SetFocus
End If
End Sub

Private Sub optFemale_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtPassword.SetFocus
End If
End Sub

Private Sub optMale_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtPassword.SetFocus
End If
End Sub

Private Sub txtAns1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtAns1.Text = "" Then
        MsgBox " Enter answer First", vbInformation, "Instruction"
        txtAns1.SetFocus
    Else
        cmbQ2.SetFocus
    End If
End If
End Sub

Private Sub txtAns2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtAns2.Text = "" Then
        MsgBox " Enter answer First", vbInformation, "Instruction"
        txtAns2.SetFocus
    Else
        optAgree.SetFocus
    End If
End If
End Sub

Private Sub txtConfirmPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtConfirmPassword.Text = "" Then
        MsgBox " Enter password First", vbInformation, "Instruction"
        txtConfirmPassword.SetFocus
    Else
        txtNumber.SetFocus
    End If
End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtEmail.Text = "" Then
        MsgBox " Enter Email First", vbInformation, "Instruction"
        txtEmail.SetFocus
    Else
        cmbQ1.SetFocus
    End If
End If
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtNumber.Text = "" Then
        MsgBox " Enter Number First", vbInformation, "Instruction"
        txtNumber.SetFocus
    Else
        txtEmail.SetFocus
    End If
End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If txtPassword.Text = "" Then
        MsgBox " Enter password First", vbInformation, "Instruction"
        txtPassword.SetFocus
    Else
        txtConfirmPassword.SetFocus
    End If
End If
End Sub
Public Sub ResetData()
txtName.Text = ""
DTPDob.Value = Date
optMale.Value = True
txtPassword.Text = ""
txtConfirmPassword.Text = ""
txtNumber.Text = ""
txtEmail.Text = ""
txtAns1.Text = ""
txtAns2.Text = ""

End Sub
