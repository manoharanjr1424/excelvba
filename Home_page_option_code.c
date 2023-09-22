Private Sub CommandButton1_Click()
Unload UserForm5
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim prevNumber As Long
    Dim newNumber As Long
    
    ' Set the worksheet where the data is stored
    Set ws = ThisWorkbook.Sheets("CHECK-IN BE") ' Change to your sheet name
    
    ' Determine the last used row in the specified column (adjust column as needed)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Get the previous number from the last used row (assuming numbers are in column A)
    If lastRow > 1 Then
        prevNumber = CLng(ws.Cells(lastRow, "B").value)
    Else
        prevNumber = 0 ' Default starting number if no previous numbers
    End If
    
    ' Generate the new number by incrementing the previous number
    newNumber = prevNumber + 1
    
    ' Display the new register number in TextBox5 on UserForm (change UserForm name if different)
    UserForm2.TextBox5.value = newNumber

UserForm2.Show
Unload UserForm2
UserForm5.Show
End Sub

Private Sub CommandButton2_Click()
Unload UserForm5
UserForm4.Show
Unload UserForm4
UserForm5.Show
End Sub

Private Sub CommandButton3_Click()
Application.Quit
End Sub

Private Sub CommandButton4_Click()
Unload UserForm5
UserForm3.Show
Unload UserForm3
UserForm5.Show
End Sub

Private Sub CommandButton5_Click()
Unload UserForm5
UserForm2.Show
Unload UserForm2
UserForm5.Show
End Sub

Private Sub CommandButton6_Click()
Unload UserForm5
UserForm7.Show
Unload UserForm7
UserForm5.Show
End Sub

