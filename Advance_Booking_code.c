

Private Sub CommandButton1_Click()
On Error Resume Next
erow = Sheet2.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To erow
If ComboBox2.value <> Sheet2.Cells(i, 2) And bedtypebox.value <> Sheet2.Cells(i, 6) Then

ListBox1.ColumnCount = 3

ListBox1.AddItem
ListBox1.List(ListBox1.ListCount - 1, 0) = Sheet2.Cells(i, 2)
ListBox1.List(ListBox1.ListCount - 1, 1) = Sheet2.Cells(i, 7)
ListBox1.List(ListBox1.ListCount - 1, 2) = Sheet2.Cells(i, 6)

End If
Next i
End Sub

Private Sub CommandButton2_Click()
Dim sh As Worksheet
Dim Lr As Long
Set sh = ThisWorkbook.Sheets("adv")
Lr = [Counta(adv!A:A)] + 1
With sh
.Cells(Lr, 1) = Lr - 1
.Cells(Lr, 3) = UserForm7.indatebox.value
.Cells(Lr, 5) = UserForm7.namebox.value
.Cells(Lr, 6) = UserForm7.ComboBox1.text
.Cells(Lr, 7) = UserForm7.Ccont.value
.Cells(Lr, 8) = UserForm7.idcardbox.text
.Cells(Lr, 9) = UserForm7.idnumbox.value
.Cells(Lr, 11) = UserForm7.ComboBox2.text
.Cells(Lr, 12) = UserForm7.bedtypebox.text
.Cells(Lr, 15) = UserForm7.depositbox.value
.Cells(Lr, 20) = UserForm7.ComboBox4.text
.Cells(Lr, 23) = UserForm7.personbox.value
.Cells(Lr, 24) = "OPENED"

End With
Call CommandButton3_Click
 MsgBox "DATA HAS BEEN ADDED IN THE DATABSE", vbInformation
End Sub
Private Sub genderbox_Change()

End Sub

Private Sub CommandButton3_Click()
 Me.namebox.value = ""
 Me.ComboBox1.value = ""
 Me.Ccont.value = ""
 Me.indatebox.value = ""
 Me.idcardbox.value = ""
 Me.idnumbox.value = ""
 Me.ComboBox2.value = ""
 Me.depositbox.value = ""
 Me.personbox.value = ""
 Me.bedtypebox.value = ""
End Sub

Private Sub CommandButton4_Click()
Unload UserForm7
UserForm5.Show
End Sub


Private Sub number_Change()

End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub numberbox_Change()

End Sub

Private Sub roomnobox_Change()

End Sub

Private Sub roomtypebox_Change()

End Sub

Private Sub TabStrip1_Change()

End Sub

Private Sub CommandButton5_Click()
Me.indatebox.value = Calendar.DatePicker
End Sub

Private Sub CommandButton6_Click()
Me.ComboBox3.value = Calendar.DatePicker

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Activate()
ComboBox1.List = Array("MALE", "FEMALE")
ComboBox4.List = Array("CASH", "UPI", "DEBIT CARD", "CASH+UPI")
End Sub

Private Sub UserForm_Initialize()
bedtypebox.AddItem "DOUBLE BED"
bedtypebox.AddItem "TRIBLE BED"
bedtypebox.AddItem "FOUR BED"
ComboBox2.AddItem "AC"
ComboBox2.AddItem "NON AC"
idcardbox.AddItem "AADHAR CARD"
idcardbox.AddItem "PAN CARD"
idcardbox.AddItem "PASSPORT"
idcardbox.AddItem "CMC ID CARD"
End Sub
