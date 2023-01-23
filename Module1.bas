Attribute VB_Name = "Module1"
Dim UserName As String
Dim book As Integer, deli As Integer, avai As Integer
Dim deliver As Integer, available As Integer
Dim ConID As String, ConName As String, DelDate As Date, BookDate As Date, CylNo As Integer
Dim ConNo As String, Conntype As String
Public OpenFormName As String

Public Sub Bill()
ConID = frmDelivery.cmbConsumerID.Text
frmBill.lblConsumerID.Caption = ConID
ConName = frmDelivery.txtConsumerName.Text
frmBill.lblConsumerName.Caption = ConName
DelDate = frmDelivery.DTPDelivery.Value
frmBill.lblDeliveryDate.Caption = DelDate
BookDate = frmDelivery.DTPBooking.Value
frmBill.lblBookingDate.Caption = BookDate
CylNo = frmDelivery.txtCylinderNo.Text
frmBill.lblNoofCylinder.Caption = CylNo
ConNo = frmDelivery.txtPhoneNumber.Text
frmBill.lblContactNo.Caption = ConNo
Conntype = frmDelivery.txtConnectionType.Text
frmBill.lblConnectionType.Caption = Conntype
End Sub


Public Sub lblForgot()
UserName = frmLogin.cmbUser.Text
frmForgot.lblUserName.Caption = UserName
End Sub
Public Sub lblReset()
UserName = frmLogin.cmbUser.Text
frmReset.lblUserName.Caption = UserName
End Sub

Public Sub UserRead()
UserName = frmLogin.cmbUser.Text
End Sub
