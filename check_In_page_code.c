

Private Sub CommandButton1_Click()
Answer = MsgBox("CONFIRM CHECK IN", vbQuestion + vbYesNo, "User Repsonse")
If Answer = vbYes Then
Dim sh As Worksheet
Dim Ls As Long
Set sh = ThisWorkbook.Sheets("CHECK-IN BE")
Ls = Worksheets("CHECK-IN BE").Cells(Rows.Count, 1).End(xlUp).Row
With sh
sh.Cells(Ls + 1, 1) = Ls
sh.Cells(Ls + 1, 2) = TextBox5.value
sh.Cells(Ls + 1, 3) = UserForm2.indatebox.value
sh.Cells(Ls + 1, 4) = UserForm2.intimebox.value
sh.Cells(Ls + 1, 5) = UserForm2.outdatebox.value
sh.Cells(Ls + 1, 7) = UserForm2.namebox.value
sh.Cells(Ls + 1, 8) = UserForm2.contactnobox.value
sh.Cells(Ls + 1, 9) = UserForm2.proofidbox.value
sh.Cells(Ls + 1, 10) = UserForm2.idcardbox.text
sh.Cells(Ls + 1, 11) = UserForm2.TextBox1.text
sh.Cells(Ls + 1, 12) = UserForm2.TextBox2.value
sh.Cells(Ls + 1, 13) = UserForm2.TextBox3.text
sh.Cells(Ls + 1, 16) = UserForm2.depositebox.value
sh.Cells(Ls + 1, 17) = UserForm2.TextBox4.value
sh.Cells(Ls + 1, 19) = UserForm2.paidamountbox.value
sh.Cells(Ls + 1, 21) = "BOOKED"
sh.Cells(Ls + 1, 22) = UserForm2.paymentbox.text
sh.Cells(Ls + 1, 28) = UserForm2.Address.value
End With

Dim h As Worksheet
Dim L As Long
Set h = ThisWorkbook.Sheets("CHECK-IN")
L = Worksheets("CHECK-IN").Cells(Rows.Count, 1).End(xlUp).Row
With h
h.Cells(L + 1, 1) = L
h.Cells(L + 1, 2) = TextBox5.value
h.Cells(L + 1, 3) = UserForm2.indatebox.value
h.Cells(L + 1, 4) = UserForm2.intimebox.value
h.Cells(L + 1, 5) = UserForm2.outdatebox.value
h.Cells(L + 1, 7) = UserForm2.namebox.value
h.Cells(L + 1, 8) = UserForm2.contactnobox.value
h.Cells(L + 1, 9) = UserForm2.TextBox1.text
h.Cells(L + 1, 10) = UserForm2.TextBox2.value
h.Cells(L + 1, 14) = UserForm2.depositebox.value
h.Cells(L + 1, 15) = UserForm2.TextBox4.value
h.Cells(L + 1, 17) = UserForm2.paidamountbox.value
h.Cells(L + 1, 19) = "BOOKED"
h.Cells(L + 1, 20) = UserForm2.paymentbox.value
h.Cells(L + 1, 26) = UserForm2.Address.text
End With


hrow = Worksheets("adv").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To hrow
If contactnobox.value Like Worksheets("adv").Cells(i, 7) Then
Worksheets("adv").Cells(i, 24) = "CLOSED"
End If
Next i


MsgBox "Check In Done"

Unload UserForm2

Else

End If
End Sub

Private Sub TextBox10_Change()

End Sub

Private Sub TextBox8_Change()

End Sub


Private Sub CommandButton2_Click()
On Error Resume Next
hrow = Worksheets("CHECK-IN BE").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To hrow
If contactnobox.value Like Worksheets("CHECK-IN BE").Cells(i, 8) Then
namebox = Worksheets("CHECK-IN BE").Cells(i, 7)
proofidbox = Worksheets("CHECK-IN BE").Cells(i, 9)
idcardbox = Worksheets("CHECK-IN BE").Cells(i, 10)
End If
Next i
End Sub

Private Sub CommandButton3_Click()
Me.outdatebox.value = Calendar.DatePicker
End Sub

Private Sub CommandButton6_Click()
indatebox = Calendar.DatePicker
End Sub

Private Sub CommandButton7_Click()
Unload UserForm2
UserForm2.Show
End Sub

Private Sub CommandButton8_Click()
erow = Worksheets("adv").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To erow
If ((contactnobox.value Like Worksheets("adv").Cells(i, 7) Or idcardbox.value Like Worksheets("adv").Cells(i, 9)) And Worksheets("adv").Cells(i, 24) Like "OPENED") Then
namebox = Worksheets("adv").Cells(i, 5)
indatebox = Worksheets("adv").Cells(i, 3)
proofidbox = Worksheets("adv").Cells(i, 8)
idcardbox = Worksheets("adv").Cells(i, 9)
TextBox2 = Worksheets("adv").Cells(i, 11)
TextBox3 = Worksheets("adv").Cells(i, 12)
depositebox = Worksheets("adv").Cells(i, 15)
paidamountbox = Worksheets("adv").Cells(i, 18)
paymentbox = Worksheets("adv").Cells(i, 20)
cmtbox = Worksheets("adv").Cells(i, 22)
End If
Next i
End Sub

Private Sub CommandButton9_Click()
  
End Sub

Private Sub TextBox1_Change()



arow = Worksheets("ROOM-RENT").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To arow
If TextBox1.value Like Worksheets("ROOM-RENT").Cells(i, 1) Then
TextBox3 = Worksheets("ROOM-RENT").Cells(i, 5)
TextBox2 = Worksheets("ROOM-RENT").Cells(i, 4)

If ComboBox2 Like "CMC" Then
TextBox4 = Worksheets("ROOM-RENT").Cells(i, 2)
Else
TextBox4 = Worksheets("ROOM-RENT").Cells(i, 3)

End If
End If
Next i
End Sub

Private Sub UserForm_Activate()
paymentbox.List = Array("UPI", "CASH", "DEBIT CARD", "UPI+CASH")
proofidbox.List = Array("PASSPORT", "AADHAAR CARD", "LICENCE", "CMC ID CARD", "OTHER GOVT CARD")
TextBox2.List = Array("AC", "NON AC")
TextBox3.List = Array("DOUBLE", "TRIBLE", "FOUR")
ComboBox2.List = Array("CMC", "NON CMC")
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next
UserForm2.indatebox.value = Format(Date, "DD-MMM-YY")
UserForm2.intimebox.value = Time()

End Sub

Private Sub V_Click()

End Sub
