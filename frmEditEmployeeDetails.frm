VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditEmployeeDetails 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   $"frmEditEmployeeDetails.frx":0000
   ClientHeight    =   9960
   ClientLeft      =   2160
   ClientTop       =   804
   ClientWidth     =   16548
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   16548
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14280
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      DataField       =   "No"
      DataSource      =   "Adodc1"
      Height          =   612
      Left            =   14760
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   492
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   372
      Left            =   13080
      Top             =   2160
      Visible         =   0   'False
      Width           =   3132
      _ExtentX        =   5525
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
      RecordSource    =   "select * from EmployeeEntry"
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
   Begin VB.CommandButton cmdChangePhoto 
      BackColor       =   &H00FFFF80&
      Caption         =   "Change Photo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6600
      Width           =   2772
   End
   Begin VB.TextBox txtEmployeeName 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   29
      Top             =   4080
      Width           =   3975
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "<<Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8760
      Width           =   1692
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFC0&
      Caption         =   "[ ]Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8760
      Width           =   1692
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "< >Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8760
      Width           =   1692
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   24
      Top             =   8880
      Width           =   3975
   End
   Begin VB.TextBox txtState 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8760
      TabIndex        =   22
      Top             =   6360
      Width           =   3975
   End
   Begin VB.TextBox txtDistrict 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8760
      TabIndex        =   20
      Top             =   5280
      Width           =   3975
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8760
      TabIndex        =   18
      Top             =   4080
      Width           =   3975
   End
   Begin VB.TextBox txtContact 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   17
      Top             =   7680
      Width           =   3975
   End
   Begin VB.TextBox txtSalary 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8760
      TabIndex        =   16
      Top             =   7680
      Width           =   3975
   End
   Begin VB.TextBox txtQualification 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   6360
      Width           =   3975
   End
   Begin VB.TextBox txtBranch 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8760
      TabIndex        =   14
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox txtEmployeeID 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   2760
      Width           =   3975
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   4800
      TabIndex        =   11
      Top             =   5280
      Width           =   1575
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   2640
      TabIndex        =   10
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   348
      Index           =   8
      Left            =   360
      TabIndex        =   28
      Top             =   4080
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   372
      Index           =   11
      Left            =   360
      TabIndex        =   23
      Top             =   8880
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   348
      Index           =   10
      Left            =   7320
      TabIndex        =   21
      Top             =   6360
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   348
      Index           =   9
      Left            =   7320
      TabIndex        =   19
      Top             =   5280
      Width           =   876
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   372
      Index           =   7
      Left            =   7320
      TabIndex        =   12
      Top             =   4080
      Width           =   1092
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   372
      Index           =   6
      Left            =   360
      TabIndex        =   9
      Top             =   5280
      Width           =   972
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact no."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   372
      Index           =   5
      Left            =   360
      TabIndex        =   8
      Top             =   7680
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sallary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   372
      Index           =   4
      Left            =   7320
      TabIndex        =   7
      Top             =   7680
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qualification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   372
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   6360
      Width           =   1608
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Branch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   372
      Index           =   2
      Left            =   7320
      TabIndex        =   5
      Top             =   2880
      Width           =   936
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   348
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   1656
   End
   Begin VB.Image imgPhoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   3372
      Left            =   13560
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2772
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Employee Information"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   852
      Left            =   4320
      TabIndex        =   3
      Top             =   1680
      Width           =   7416
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adampur chowk Bhagalpur-812001"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   8280
      TabIndex        =   2
      Top             =   1200
      Width           =   3516
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bhagalpur Gas Agency"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   756
      Index           =   0
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   6948
   End
   Begin VB.Image Image1 
      Height          =   1692
      Left            =   0
      Picture         =   "frmEditEmployeeDetails.frx":00B9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2172
   End
   Begin VB.Image Image2 
      Height          =   1692
      Left            =   14400
      Picture         =   "frmEditEmployeeDetails.frx":25F3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2172
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Height          =   1692
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   17172
   End
End
Attribute VB_Name = "frmEditEmployeeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdChangePhoto_Click()
CommonDialog1.Filter = "jpeg|*.jpg"
CommonDialog1.ShowOpen
imgPhoto.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub cmdUpdate_Click()
ans1 = MsgBox(" Do you want to confirm the change the record", vbQuestion + vbYesNo, " Message")
If (ans1 = vbYes) Then
Adodc1.Recordset.Update
Adodc1.Recordset.Fields(1).Value = txtEmployeeID.Text
Adodc1.Recordset.Fields(2).Value = txtEmployeeName.Text
If optMale.Value = True Then
    Adodc1.Recordset.Fields(3).Value = "Male"
Else
    Adodc1.Recordset.Fields(3).Value = "Female"
End If
Adodc1.Recordset.Fields(4).Value = txtQualification.Text
Adodc1.Recordset.Fields(5).Value = txtSalary.Text
Adodc1.Recordset.Fields(6).Value = txtContact.Text
Adodc1.Recordset.Fields(7).Value = txtBranch.Text
Adodc1.Recordset.Fields(8).Value = txtAddress.Text
Adodc1.Recordset.Fields(9).Value = txtDistrict.Text
Adodc1.Recordset.Fields(10).Value = txtState.Text
Adodc1.Recordset.Fields(11).Value = txtEmail.Text
SavePicture
Adodc1.Recordset.Update
MsgBox "Records updated succefully. ", vbExclamation, "Instruction"
txtEmployeeID.Text = ""
End If
 ResetData
    txtEmployeeID.SetFocus
End Sub

Private Sub txtEmployeeID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Adodc1.RecordSource = " select *from EmployeeEntry where EmployeeID= '" & txtEmployeeID.Text & "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
        
          MsgBox " Sorry Employee Id is wrong , Please check your consumer Id", vbInformation, "Instruction"
    Else
          loadData
End If
End If
End Sub

Public Sub loadData()
    txtEmployeeName.Text = Adodc1.Recordset.Fields(2).Valuen
    If Adodc1.Recordset.Fields(3).Value = "Male" Then
        optMale.Value = True
    Else
        optFemale.Value = True
    End If
    txtQualification.Text = Adodc1.Recordset.Fields(4).Value
    txtSalary.Text = Adodc1.Recordset.Fields(5).Value
    txtContact.Text = Adodc1.Recordset.Fields(6).Value
    txtBranch.Text = Adodc1.Recordset.Fields(7).Value
    txtEmail.Text = Adodc1.Recordset.Fields(14).Value
    txtAddress.Text = Adodc1.Recordset.Fields(8).Value
    txtDistrict.Text = Adodc1.Recordset.Fields(9).Value
    txtState.Text = Adodc1.Recordset.Fields(10).Value
    LoadImageFile
End Sub
Public Sub LoadImageFile()
    Dim st As ADODB.Stream
    Set st = New ADODB.Stream
    st.Type = adTypeBinary
    st.Open
    If IsNull(Adodc1.Recordset.Fields(15).Value) = False Then
        st.Write (Adodc1.Recordset.Fields(15).Value)
        st.SaveToFile App.Path & "\aa.jpg", adSaveCreateOverWrite
        imgPhoto.Picture = LoadPicture(App.Path & "\aa.jpg")
    End If
    st.Close
End Sub

Private Sub cmdDelete_Click()
ans2 = MsgBox(" Do you want to Delete the record", vbQuestion + vbYesNo, " Message")
If (ans2 - vbYes) Then
Adodc1.Recordset.Delete
MsgBox " Records deleted succesfully ...!", vbInformation, "Instruction"
ResetData
End If
End Sub
Public Sub ResetData()
    On Error Resume Next
    txtContact.Text = ""
    txtEmail.Text = ""
    txtAddress.Text = ""
    txtPinCode.Text = ""
    txtDistrict.Text = ""
    txtState.Text = ""
    txtQualification.Text = ""
    txtSalary.Text = ""
    txtBranch.Text = ""
    txtEmail.Text = ""
    imgPhoto.Picture = LoadPicture("")
End Sub


Public Sub SavePicture()
    Dim st As ADODB.Stream
    Set st = New ADODB.Stream
    st.Type = adTypeBinary
    st.Open
    st.LoadFromFile (CommonDialog1.FileName)
    Adodc1.Recordset.Fields(15).Value = st.Read
    st.Close
End Sub

