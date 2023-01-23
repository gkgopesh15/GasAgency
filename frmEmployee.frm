VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEmployee 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   $"frmEmployee.frx":0000
   ClientHeight    =   9372
   ClientLeft      =   5664
   ClientTop       =   3120
   ClientWidth     =   13860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9372
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      DataField       =   "No"
      DataSource      =   "Adodc1"
      Height          =   288
      Left            =   10200
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1452
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   492
      Left            =   840
      Top             =   2400
      Visible         =   0   'False
      Width           =   1812
      _ExtentX        =   3196
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
      RecordSource    =   "select *from employeeEntry"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Height          =   6015
      Left            =   8400
      TabIndex        =   17
      Top             =   3120
      Width           =   5295
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   492
         Left            =   2280
         TabIndex        =   34
         Top             =   3600
         Width           =   2892
         _ExtentX        =   5101
         _ExtentY        =   868
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Modern No. 20"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111673345
         CurrentDate     =   43281
      End
      Begin VB.CommandButton cmdAddEmployee 
         BackColor       =   &H00FFFF00&
         Caption         =   "(*)  Add Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   5160
         Width           =   4935
      End
      Begin VB.TextBox txtDistrict 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   22
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtExpDetails 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   21
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtPinCode 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   20
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txtState 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   19
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   18
         Top             =   4440
         Width           =   2895
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "District"
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
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   696
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin Code"
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
         Left            =   240
         TabIndex        =   27
         Top             =   2880
         Width           =   972
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of joining"
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
         Left            =   240
         TabIndex        =   26
         Top             =   3720
         Width           =   1536
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp. Details"
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
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   1260
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State "
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
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   612
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
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
         Left            =   240
         TabIndex        =   23
         Top             =   4440
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   6015
      Left            =   2640
      TabIndex        =   5
      Top             =   3120
      Width           =   5655
      Begin VB.OptionButton optFemaile 
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3720
         TabIndex        =   33
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton optMale 
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2280
         TabIndex        =   32
         Top             =   1320
         Width           =   1452
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   29
         Top             =   5040
         Width           =   3255
      End
      Begin VB.TextBox txtContactNo 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   16
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox txtSalary 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   15
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox txtQualification 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   14
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox txtBranch 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   13
         Top             =   4320
         Width           =   3255
      End
      Begin VB.TextBox txtEmployeeName 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2280
         TabIndex        =   12
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   240
         TabIndex        =   28
         Top             =   5160
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   816
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No."
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
         Left            =   240
         TabIndex        =   10
         Top             =   3600
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
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
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qualification"
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
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
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
         Left            =   240
         TabIndex        =   7
         Top             =   4440
         Width           =   768
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1740
      End
   End
   Begin VB.CommandButton cmdUpload 
      BackColor       =   &H00FFFF00&
      Caption         =   ">> Upload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label lblEmployeeID 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "EMP00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   840
      TabIndex        =   38
      Top             =   8040
      Width           =   1488
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFF00&
      Height          =   492
      Left            =   240
      TabIndex        =   37
      Top             =   7920
      Width           =   2292
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee  Entry"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   750
      Left            =   4320
      TabIndex        =   31
      Top             =   2040
      Width           =   5145
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      Caption         =   "   Employee ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   240
      TabIndex        =   3
      Top             =   7560
      Width           =   2292
   End
   Begin VB.Image imgPhoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   3012
      Left            =   240
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   2292
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "( Adampur  Chock  Bhagalpur - 802001 )"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   5880
      TabIndex        =   2
      Top             =   1320
      Width           =   4035
   End
   Begin VB.Image Image2 
      Height          =   1935
      Left            =   11520
      Picture         =   "frmEmployee.frx":0099
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   0
      Picture         =   "frmEmployee.frx":267B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bhagalpur Gas Agency"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   990
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   6645
   End
   Begin VB.Label Lebel 
      BackColor       =   &H00FFFFC0&
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13815
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C_no As Integer

Private Sub cmdAddEmployee_Click()
If txtEmployeeName.Text = "" Or txtQualification.Text = "" Or txtSalary.Text = "" Or txtContactNo.Text = "" _
Or txtBranch.Text = "" Or txtAddress.Text = "" Or txtDistrict.Text = "" Or txtState.Text = "" Or txtExpDetails.Text = "" _
Or txtPinCode.Text = "" Or txtEmail.Text = "" Then
    MsgBox " Opps something is missing please check all fields !!!", vbExclamation, "instruction"
Else
ans1 = MsgBox(" Do you want to add the employee. ", vbQuestion + vbYesNo, " Message")
If (ans1 - vbYes) Then
Adodc1.Recordset.AddNew
With Adodc1.Recordset
        .Fields(0).Value = C_no
        .Fields(1).Value = lblEmployeeID.Caption
        .Fields(2).Value = txtEmployeeName.Text
        If optMale.Value = True Then
            .Fields(3).Value = "Male"
        Else
            .Fields(3).Value = "Female"
        End If
        .Fields(4).Value = txtQualification.Text
        .Fields(5).Value = txtSalary.Text
        .Fields(6).Value = txtContactNo.Text
        .Fields(7).Value = txtBranch.Text
        .Fields(8).Value = txtAddress.Text
        .Fields(9).Value = txtDistrict.Text
        .Fields(10).Value = txtState.Text
        .Fields(11).Value = txtExpDetails.Text
        .Fields(12).Value = txtPinCode.Text
        .Fields(13).Value = dtpDate.Value
        .Fields(14).Value = txtEmail.Text
    End With
    SavePicture
    Adodc1.Recordset.Update
    Adodc1.Refresh
    MsgBox " New Employee Register Succefully!!", vbInformation + vbOKOnly, "Saved Successfully"
    End If
     txtEmployeeName.Text = ""
    ResetData
    txtEmployeeName.SetFocus
End If
End Sub
Public Sub SavePicture()
    On Error Resume Next
    Dim st As ADODB.Stream
    Set st = New ADODB.Stream
    st.Type = adTypeBinary
    st.Open
    st.LoadFromFile (CommonDialog1.FileName)
    Adodc1.Recordset.Fields(15).Value = st.Read
    st.Close
End Sub
Public Sub ResetData()
    optMale.Value = True
    txtContactNo.Text = ""
    txtEmail.Text = ""
    txtAddress.Text = ""
    txtPinCode.Text = ""
    txtDistrict.Text = ""
    txtState.Text = ""
    txtQualification.Text = ""
    txtSalary.Text = ""
    txtBranch.Text = ""
    txtExpDetails.Text = ""
    dtpDate.Value = Date
    txtEmail.Text = ""
    C_no = C_no + 1
    lblEmployeeID.Caption = lblEmployeeID.Caption & C_no
    imgPhoto.Picture = LoadPicture("")
End Sub

Private Sub cmdAddEmployee_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
cmdAddEmployee_Click
End If
End Sub

Private Sub cmdUpload_Click()
CommonDialog1.Filter = "jpeg|*.jpg"
CommonDialog1.ShowOpen
imgPhoto.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub cmdUpload_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
cmdAddEmployee.SetFocus
End If
End Sub

Private Sub dtpDate_Click()
If KeyAscii = vbKeyReturn Then
        txtEmail.SetFocus
End If
End Sub

Private Sub Form_Load()
optMale.Value = True
Adodc1.Recordset.MoveLast
C_no = Adodc1.Recordset.Fields(0).Value
C_no = C_no + 1
lblEmployeeID.Caption = lblEmployeeID.Caption & C_no
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtAddress.Text = "" Then
        MsgBox " Please enter address.", Message
        txtAddress.SetFocus
        Else
        txtDistrict.SetFocus
        End If
    End If
End Sub

Private Sub txtBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtBranch.Text = "" Then
        MsgBox " Please enter branch.", Message
        txtBranch.SetFocus
        Else
        txtAddress.SetFocus
        End If
    End If
End Sub

Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtContactNo.Text = "" Then
        MsgBox " Please enter Contact no.", Message
        txtContactNo.SetFocus
        Else
        txtBranch.SetFocus
        End If
    End If
End Sub

Private Sub txtDistrict_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtDistrict.Text = "" Then
        MsgBox " Please enter district.", Message
        txtDistrict.SetFocus
        Else
        txtState.SetFocus
        End If
    End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtEmail.Text = "" Then
        MsgBox " Please enter salary.", Message
        txtEmail.SetFocus
        Else
        cmdUpload.SetFocus
        End If
    End If
End Sub

Private Sub txtEmployeeName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtEmployeeName.Text = "" Then
        MsgBox " Please Enter employee name.", Message
        txtEmployeeName.SetFocus
        Else
        optMale.SetFocus
        End If
    End If
End Sub

Private Sub txtExpDetails_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtExpDetails.Text = "" Then
        MsgBox " Please enter Experence Details .", Message
        txtExpDetails.SetFocus
        Else
        txtPinCode.SetFocus
        End If
    End If
End Sub

Private Sub txtPinCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtPinCode.Text = "" Then
        MsgBox " Please enter pin code.", Message
        txtPinCode.SetFocus
        Else
        dtpDate.SetFocus
        End If
    End If
End Sub

Private Sub txtQualification_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtQualification.Text = "" Then
        MsgBox " Please enter qualification.", Message
        txtQualification.SetFocus
        Else
        txtSalary.SetFocus
        End If
    End If
End Sub

Private Sub txtSalary_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtSalary.Text = "" Then
        MsgBox " Please enter salary.", Message
        txtSalary.SetFocus
        Else
        txtContactNo.SetFocus
        End If
    End If
End Sub

Private Sub txtState_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        If txtState.Text = "" Then
        MsgBox " Please enter state .", Message
        txtState.SetFocus
        Else
        txtExpDetails.SetFocus
        End If
    End If
End Sub
