Private Sub CommandButton1_Click()
ListBox1.Clear
Dim trow As Long
Set sh = ThisWorkbook.Sheets("CHECK-IN BE")
trow = sh.Cells(Rows.Count, 1).End(xlUp).Row
For x = 2 To trow
If UserForm4.TextBox1.value Like sh.Cells(x, 8) Or UserForm4.TextBox2.text Like Trim(sh.Cells(x, 7)) Then
ListBox1.ColumnCount = 3
ListBox1.AddItem
ListBox1.List(ListBox1.ListCount - 1, 0) = sh.Cells(x, 7)
ListBox1.List(ListBox1.ListCount - 1, 1) = sh.Cells(x, 11)
ListBox1.List(ListBox1.ListCount - 1, 2) = sh.Cells(x, 8)
End If
Next x

End Sub

Private Sub CommandButton2_Click()
Unload UserForm4
UserForm4.Show
End Sub

Private Sub CommandButton3_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you want to Close", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbYes Then
      Unload UserForm4
      Else
    End If
End Sub

Private Sub CommandButton4_Click()

For x = 0 To ListBox1.ListCount - 1

Customerdetails.TextBox2.value = ListBox1.List(ListBox1.ListIndex, 0)

Next x
Unload Me
Customerdetails.Show
End Sub

Private Sub UserForm_Click()

End Sub
