Private Sub CommandButton1_Click()
Ay = MsgBox("CONFIRM TO EXTEND THE DATE ", vbQuestion + vbYesNo, "date extend")
On Error Resume Next
Dim value
If Ay Like vbYes Then


erow = Worksheets("CHECK-IN BE").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To erow
If TextBox9.value Like Worksheets("CHECK-IN BE").Cells(i, 2) Then
 Worksheets("CHECK-IN BE").Cells(i, 5) = TextBox8.value
 value = Worksheets("CHECK-IN BE").Cells(i, 19)
 Worksheets("CHECK-IN BE").Cells(i, 19) = value + TextBox11.value
 End If
Next i


trow = Worksheets("CHECK-IN").Cells(Rows.Count, 1).End(xlUp).Row
For x = 2 To trow
If TextBox10.value Like "" Then
If TextBox9.value Like Worksheets("CHECK-IN").Cells(x, 2) Then
 Worksheets("CHECK-IN").Cells(x, 5) = TextBox8.value
 End If
End If
Next x

MsgBox "DATE EXTENDED"
End If

If Ay Like vbNo Then
UserForm8.Show
End If
End Sub

Private Sub CommandButton2_Click()
On Error Resume Next
erow = Worksheets("CHECK-IN BE").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To erow
If ((TextBox9.value Like Worksheets("CHECK-IN BE").Cells(i, 2) Or TextBox6.value Like Worksheets("CHECK-IN BE").Cells(i, 11)) And (Worksheets("CHECK-IN BE").Cells(i, 21) Like "BOOKED")) Then
TextBox9 = Worksheets("CHECK-IN BE").Cells(i, 2)
TextBox3 = Worksheets("CHECK-IN BE").Cells(i, 5)
TextBox5 = Worksheets("CHECK-IN BE").Cells(i, 7)
TextBox1 = Worksheets("CHECK-IN BE").Cells(i, 8)
TextBox2 = Worksheets("CHECK-IN BE").Cells(i, 10)
TextBox6 = Worksheets("CHECK-IN BE").Cells(i, 11)

If Worksheets("CHECK-IN BE").Cells(i, 5) Like "" Then
TextBox7 = "NO DATE GIVE"
Else
TextBox7 = Worksheets("CHECK-IN BE").Cells(i, 5)
End If
End If
Next i
End Sub

Private Sub CommandButton3_Click()

    Dim response As VbMsgBoxResult
    response = MsgBox("Do you want to Exit", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbYes Then
       Unload Me
       UserForm5.Show
    Else
    End If
End Sub

Private Sub CommandButton4_Click()
Unload Me
UserForm8.Show

End Sub

Private Sub CommandButton5_Click()
Me.TextBox8 = Calendar.DatePicker
End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub TextBox11_Change()

End Sub

Private Sub UserForm_Click()

End Sub
