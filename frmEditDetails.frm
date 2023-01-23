VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditConsumerDetails 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   $"frmEditDetails.frx":0000
   ClientHeight    =   10476
   ClientLeft      =   5664
   ClientTop       =   1860
   ClientWidth     =   16764
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10476
   ScaleWidth      =   16764
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   1092
      Left            =   240
      TabIndex        =   23
      Top             =   2400
      Width           =   5292
      Begin VB.TextBox txtConsumerID 
         BackColor       =   &H00C0E0FF&
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
         Left            =   2760
         TabIndex        =   25
         Top             =   360
         Width           =   2172
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consumer Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   348
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   1788
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14880
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text7 
      DataField       =   "ConsumerNo"
      DataSource      =   "Adodc1"
      Height          =   492
      Left            =   15360
      TabIndex        =   22
      Text            =   "Text7"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1092
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   1320
      Top             =   9600
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
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdEditPhoto 
      BackColor       =   &H00FFFF00&
      Caption         =   "Edit Photo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6960
      Width           =   2652
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "<<  Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9120
      Width           =   1932
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H000000FF&
      Caption         =   "[ ] Delete "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7800
      Width           =   1932
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H000000FF&
      Caption         =   "< >Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7800
      Width           =   2052
   End
   Begin VB.CheckBox chkComercial 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Comercial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   17
      Top             =   9000
      Width           =   2055
   End
   Begin VB.CheckBox chkDomestic 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Domestic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   9000
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6120
      TabIndex        =   13
      Top             =   4800
      Width           =   4095
      Begin VB.OptionButton optFemale 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   15
         Top             =   120
         Width           =   1575
      End
      Begin VB.OptionButton optMale 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.TextBox txtPinCode 
      BackColor       =   &H00C0E0FF&
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
      Height          =   510
      Left            =   6360
      TabIndex        =   12
      Top             =   8040
      Width           =   4335
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00C0E0FF&
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
      Height          =   510
      Left            =   6360
      TabIndex        =   11
      Top             =   7200
      Width           =   4335
   End
   Begin VB.TextBox txtEmailID 
      BackColor       =   &H00C0E0FF&
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
      Height          =   510
      Left            =   6360
      TabIndex        =   10
      Top             =   6360
      Width           =   4335
   End
   Begin VB.TextBox txtMobileNo 
      BackColor       =   &H00C0E0FF&
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
      Height          =   510
      Left            =   6360
      TabIndex        =   9
      Top             =   5520
      Width           =   4335
   End
   Begin VB.TextBox txtConsumerName 
      BackColor       =   &H00C0E0FF&
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
      Height          =   510
      Left            =   6360
      TabIndex        =   8
      Top             =   4080
      Width           =   4335
   End
   Begin VB.Image imgPhoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   2772
      Left            =   12360
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   2652
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   432
      Left            =   1920
      TabIndex        =   7
      Top             =   8880
      Width           =   2712
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   432
      Left            =   1920
      TabIndex        =   6
      Top             =   4920
      Width           =   1188
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pin Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   432
      Left            =   1920
      TabIndex        =   5
      Top             =   7920
      Width           =   1476
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   432
      Left            =   1920
      TabIndex        =   4
      Top             =   7200
      Width           =   1332
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   432
      Left            =   1920
      TabIndex        =   3
      Top             =   6480
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   432
      Left            =   1920
      TabIndex        =   2
      Top             =   5760
      Width           =   1728
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consumer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   432
      Left            =   1920
      TabIndex        =   1
      Top             =   4080
      Width           =   2652
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Consumer Details"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1164
      Left            =   6000
      TabIndex        =   0
      Top             =   2280
      Width           =   8148
   End
   Begin VB.Image Image1 
      Height          =   2292
      Left            =   0
      Picture         =   "frmEditDetails.frx":00B2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16812
   End
End
Attribute VB_Name = "frmEditConsumerDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub loadData()
    txtConsumerName.Text = Adodc1.Recordset.Fields(2).Value
    If Adodc1.Recordset.Fields(3).Value = "Male" Then
        optMale.Value = True
    Else
        optFemale.Value = True
    End If
    txtMobileNo.Text = Adodc1.Recordset.Fields(4).Value
    txtEmailID.Text = Adodc1.Recordset.Fields(5).Value
    txtAddress.Text = Adodc1.Recordset.Fields(6).Value
    txtPinCode.Text = Adodc1.Recordset.Fields(7).Value
    If Adodc1.Recordset.Fields(9).Value = "Domestic" Then
    chkDomestic.Value = 1
    Else
    chkComercial.Value = 1
    End If
    LoadImageFile
End Sub
Public Sub LoadImageFile()
    Dim st As ADODB.Stream
    Set st = New ADODB.Stream
    st.Type = adTypeBinary
    st.Open
    If IsNull(Adodc1.Recordset.Fields(14).Value) = False Then
        st.Write (Adodc1.Recordset.Fields(14).Value)
        st.SaveToFile App.Path & "\aa.jpg", adSaveCreateOverWrite
        imgPhoto.Picture = LoadPicture(App.Path & "\aa.jpg")
    End If
    st.Close
End Sub

Private Sub chkComercial_Click()
If chkComercial.Value = 1 Then
    chkDomestic.Value = 0
End If
End Sub

Private Sub chkDomestic_Click()
If chkDomestic.Value = 1 Then
    chkComercial.Value = 0
End If
End Sub

Private Sub cmdDelete_Click()
ans2 = MsgBox(" Do you want to delete the record", vbQuestion + vbYesNo, " Message")
If (ans2 - vbYes) Then
Adodc1.Recordset.Delete
MsgBox " Records deleted succesfully ...!", vbInformation, "Instruction"
ResetData
End If
End Sub

Private Sub cmdEditPhoto_Click()
CommonDialog1.Filter = "jpeg|*.jpg"
CommonDialog1.ShowOpen
imgPhoto.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdUpdate_Click()
ans1 = MsgBox(" Do you want to confirm the change the record", vbQuestion + vbYesNo, " Message")
If (ans1 = vbYes) Then
Adodc1.Recordset.Update
Adodc1.Recordset.Fields(1).Value = txtConsumerID.Text
Adodc1.Recordset.Fields(2).Value = txtConsumerName.Text
If optMale.Value = True Then
    Adodc1.Recordset.Fields(3).Value = "Male"
Else
    Adodc1.Recordset.Fields(3).Value = "Female"
End If
Adodc1.Recordset.Fields(4).Value = txtMobileNo.Text
Adodc1.Recordset.Fields(5).Value = txtEmailID.Text
Adodc1.Recordset.Fields(6).Value = txtAddress.Text
Adodc1.Recordset.Fields(7).Value = txtPinCode.Text
If chkDomestic.Value = 1 Then
    Adodc1.Recordset.Fields(9).Value = "Domestic"
Else
    Adodc1.Recordset.Fields(9).Value = "Comercial"
End If
SavePicture
Adodc1.Recordset.Update
MsgBox "Records updated succefully. ", vbExclamation, "Instruction"
End If
txtConsumerID.Text = ""
 ResetData
    txtConsumerID.SetFocus
End Sub
Public Sub SavePicture()
    Dim st As ADODB.Stream
    Set st = New ADODB.Stream
    st.Type = adTypeBinary
    st.Open
    st.LoadFromFile (CommonDialog1.FileName)
    Adodc1.Recordset.Fields(14).Value = st.Read
    st.Close
End Sub

Private Sub txtConsumerID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Adodc1.RecordSource = " select *from NewConnection where ConsumerID= '" & txtConsumerID.Text & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
        
          MsgBox " Sorry Consumer Id is wrong , Please check your consumer Id", vbInformation, "Instruction"
    Else
          loadData
End If
End If
End Sub

Public Sub ResetData()
    txtConsumerName.Text = ""
    optMale.Value = False
    optFemale.Value = False
    txtMobileNo.Text = ""
    txtEmailID.Text = ""
    txtAddress.Text = ""
    txtPinCode.Text = ""
    chkDomestic.Value = 0
    chkComercial.Value = 0
    imgPhoto.Picture = LoadPicture("")
End Sub
