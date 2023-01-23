VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   $"frmSplash.frx":0000
   ClientHeight    =   9144
   ClientLeft      =   4920
   ClientTop       =   2064
   ClientWidth     =   13368
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9144
   ScaleWidth      =   13368
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   12840
      Top             =   4200
   End
   Begin VB.Label lblPercent 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   552
      Left            =   5400
      TabIndex        =   6
      Top             =   7800
      Width           =   156
   End
   Begin VB.Label lblLoading 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   552
      Left            =   3120
      TabIndex        =   5
      Top             =   7800
      Width           =   2280
   End
   Begin VB.Shape Shape1 
      Height          =   612
      Left            =   960
      Top             =   8520
      Width           =   10000
   End
   Begin VB.Label lblWidth 
      BackColor       =   &H0000FF00&
      Height          =   612
      Left            =   960
      TabIndex        =   4
      Top             =   8520
      Width           =   200
   End
   Begin VB.Image Image3 
      Height          =   4452
      Left            =   8760
      Picture         =   "frmSplash.frx":0087
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   4572
   End
   Begin VB.Image Image4 
      Height          =   4452
      Left            =   4320
      Picture         =   "frmSplash.frx":4AC75
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   4452
   End
   Begin VB.Image Image2 
      Height          =   4452
      Left            =   0
      Picture         =   "frmSplash.frx":4D1E4
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   4332
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Number  :  7488167982 "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   588
      Left            =   5400
      TabIndex        =   3
      Top             =   4080
      Width           =   5856
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adampur Chock Bhagalpur , 812001 "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   588
      Left            =   4200
      TabIndex        =   2
      Top             =   3480
      Width           =   6780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bhagalpur  Gas  Agency"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   996
      Left            =   1800
      TabIndex        =   1
      Top             =   2400
      Width           =   9612
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Height          =   2172
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   13332
   End
   Begin VB.Image Image1 
      Height          =   2532
      Left            =   0
      Picture         =   "frmSplash.frx":4F7FF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
lblWidth.Width = lblWidth.Width + 200
lblPercent.Caption = Val(lblWidth.Width * 200 / 20000) & "%"
If lblWidth.Width > 10000 Then
Timer1.Enabled = False
Unload Me
MDIForm1.Show
End If
End Sub

