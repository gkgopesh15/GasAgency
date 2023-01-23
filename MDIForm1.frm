VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0FFC0&
   Caption         =   $"MDIForm1.frx":0000
   ClientHeight    =   12216
   ClientLeft      =   132
   ClientTop       =   780
   ClientWidth     =   17748
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":00BF
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   612
      Left            =   0
      ScaleHeight     =   564
      ScaleWidth      =   17700
      TabIndex        =   8
      Top             =   11604
      Width           =   17748
      Begin VB.PictureBox StatusBar1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   0
         ScaleHeight     =   564
         ScaleWidth      =   17724
         TabIndex        =   9
         Top             =   0
         Width           =   17772
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFC0&
      Height          =   11604
      Left            =   0
      ScaleHeight     =   11556
      ScaleWidth      =   3780
      TabIndex        =   0
      Top             =   0
      Width           =   3828
      Begin VB.CommandButton cmdEditEmployeeDetails 
         BackColor       =   &H008080FF&
         Caption         =   "Edit Employee Details"
         DownPicture     =   "MDIForm1.frx":4ACAD
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   16.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6960
         Width           =   3492
      End
      Begin VB.CommandButton cmdEmployeeEntry 
         BackColor       =   &H00420CB4&
         Caption         =   "Employee Entry"
         DownPicture     =   "MDIForm1.frx":4C141
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   16.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6120
         Width           =   3492
      End
      Begin VB.CommandButton cmdStockReport 
         BackColor       =   &H0080FF80&
         Caption         =   "Stock Report"
         DownPicture     =   "MDIForm1.frx":4D5D5
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   16.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5280
         Width           =   3492
      End
      Begin VB.CommandButton cmdDeliveryReport 
         BackColor       =   &H000000FF&
         Caption         =   "Delivery Report"
         DownPicture     =   "MDIForm1.frx":4EA69
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   16.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4440
         Width           =   3492
      End
      Begin VB.CommandButton cmdBookingReport 
         BackColor       =   &H00FF80FF&
         Caption         =   "Booking Report"
         DownPicture     =   "MDIForm1.frx":4FEFD
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   16.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3600
         Width           =   3492
      End
      Begin VB.CommandButton cmdCylinderDelivery 
         BackColor       =   &H000080FF&
         Caption         =   "Cylinder Delivery"
         DownPicture     =   "MDIForm1.frx":51391
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   16.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2760
         Width           =   3492
      End
      Begin VB.CommandButton cmdCylinderBooking 
         BackColor       =   &H00FF8080&
         Caption         =   "Cylinder Booking"
         DownPicture     =   "MDIForm1.frx":52825
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   16.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   3492
      End
      Begin VB.CommandButton cmdEditConsumerDetails 
         BackColor       =   &H0080FF80&
         Caption         =   "Edit Consumer detail"
         DownPicture     =   "MDIForm1.frx":53CB9
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   16.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1080
         Width           =   3492
      End
      Begin VB.CommandButton cmdNewConnection 
         BackColor       =   &H000000FF&
         Caption         =   "New Connection"
         DownPicture     =   "MDIForm1.frx":5514D
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   16.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   3492
      End
      Begin VB.Image Image1 
         Height          =   4332
         Left            =   0
         Picture         =   "MDIForm1.frx":565E1
         Stretch         =   -1  'True
         Top             =   7320
         Width           =   3852
      End
   End
   Begin VB.Menu mnuHome 
      Caption         =   "&Home "
      WindowList      =   -1  'True
      Begin VB.Menu mnuNewConnection 
         Caption         =   "&New Connection"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEmoloyeeEntry 
         Caption         =   "Employee Entry"
      End
      Begin VB.Menu mnuEmployeePayment 
         Caption         =   "Employee Payment"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit "
      Begin VB.Menu mnuEditConsumerDetails 
         Caption         =   "Edit Consumer Details"
      End
      Begin VB.Menu mnuEditEmployeeDetails 
         Caption         =   "Edit Employee Details"
      End
      Begin VB.Menu mnuSetPrice 
         Caption         =   "Set Price"
      End
   End
   Begin VB.Menu mnuBookingAndDelivery 
      Caption         =   "Booking and Delivery"
      Begin VB.Menu mnuBooking 
         Caption         =   "Cylinder Booking"
      End
      Begin VB.Menu mnuDelivery 
         Caption         =   "Cylinder Delivery"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu mnuStockReport 
         Caption         =   "Stock Report"
      End
      Begin VB.Menu mnuBookingReport 
         Caption         =   "Booking Report"
      End
      Begin VB.Menu mnuDeliveryReport 
         Caption         =   "Delivery Report"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About "
      Begin VB.Menu mnuAboutAgency 
         Caption         =   "About Agency"
      End
      Begin VB.Menu mnuAboutProject 
         Caption         =   "About  Project"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frm As Form

Private Sub cmdBookingReport_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmBookingReport.Show
OpenFormName = "frmBookingReport"
End Sub

Private Sub cmdCylinderBooking_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmBooking.Show
OpenFormName = "frmBooking"
End Sub

Private Sub cmdCylinderDelivery_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmDelivery.Show
OpenFormName = "frmDelivery"
End Sub

Private Sub cmdDeliveryReport_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmDeliveryReport.Show
OpenFormName = "frmDeliveryReport"
End Sub

Private Sub cmdEditConsumerDetails_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmEditConsumerDetails.Show
OpenFormName = "frmEditConsumerDetails"
End Sub

Private Sub cmdEditEmployeeDetails_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmEditEmployeeDetails.Show
OpenFormName = "frmEditEmployeeDetails"
End Sub

Private Sub cmdEmployeeEntry_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmEmployee.Show
OpenFormName = "frmEmployeeEntry"
End Sub

Private Sub cmdNewConnection_Click()
    If Len(OpenFormName) > 0 Then Unload Me
    frmNewConnection.Show
    OpenFormName = "frmNewConnection"
End Sub
Private Sub cmdStockReport_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmStock.Show
OpenFormName = "frmStock"
End Sub
Private Sub mnuAboutAgency_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmAboutAgency.Show
OpenFormName = "frmAboutAgency"
End Sub

Private Sub mnuAboutProject_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmAboutProject.Show
OpenFormName = "frmAboutProject"
End Sub

Private Sub mnuBooking_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmBooking.Show
OpenFormName = "frmBooking"
End Sub

Private Sub mnuBookingReport_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmBookingReport.Show
OpenFormName = "frmBookingReport"
End Sub

Private Sub mnuDelivery_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmDelivery.Show
OpenFormName = "frmDelivery"
End Sub

Private Sub mnuDeliveryReport_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmDeliveryReport.Show
OpenFormName = "frmDeliveryReport"
End Sub

Private Sub mnuEditConsumerDetails_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmEditConsumerDetails.Show
OpenFormName = "frmEditConsumerDetails"
End Sub

Private Sub mnuEditEmployeeDetails_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmEditEmployeeDetails.Show
OpenFormName = "frmEditEmployeeDetails"
End Sub

Private Sub mnuEmoloyeeEntry_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmEmployee.Show
OpenFormName = "frmEmployee"
End Sub

Private Sub mnuEmployeePayment_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmEmployeePayment.Show
OpenFormName = "frmEmployeePayment"
End Sub

Private Sub mnuExit_Click()
ans = MsgBox(" Do you really want to Exit ", vbInformation + vbYesNo, "Instruction")
If (ans = vbYes) Then
End
End If
End Sub

Private Sub mnuNewConnection_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmNewConnection.Show
OpenFormName = "frmNewConnection"
End Sub

Private Sub mnuSetPrice_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmSetPrice.Show
OpenFormName = "frmSetPrice"
End Sub

Private Sub mnuStockReport_Click()
If Len(OpenFormName) > 0 Then Unload Me
frmStock.Show
OpenFormName = "frmStock"
End Sub

