Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr



Private Sub CommandButton2_Click()
ListBox1.Clear
End Sub

Private Sub CommandButton3_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you want to Exit", vbYesNo + vbQuestion, "Confirmation")
    
    If response = vbYes Then
       Unload Customerdetails
       UserForm4.Show
    Else
    End If
End Sub

Private Sub CommandButton4_Click()
Dim ws As Worksheet
Dim searchName As String
Dim lastRow As Long
Dim i As Long
Dim colCount As Long

' Set the worksheet reference
Set ws = ThisWorkbook.Sheets("CHECK-IN BE") ' Change the sheet name as needed

' Clear existing items from the ListBox
Customerdetails.ListBox1.Clear

' Get the search name from TextBox2 on the UserForm
searchName = Customerdetails.TextBox2.value
searchnumber = Customerdetails.TextBox3.value

' Find the last row with data in the worksheet
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

' Find the number of columns to consider (B to M)
colCount = 13

' Loop through the rows to search for the name
For i = 2 To lastRow ' Assuming row 1 contains headers
    If ws.Cells(i, 7).value = searchName Or ws.Cells(i, 8) = searchnumber Then
        ' Match found, populate ListBox with data from columns B to M
        Dim item As String
        item = ""

        Dim linkAddress As String
        linkAddress = ws.Cells(i, 13).value ' Get link text from column M

        For j = 2 To colCount - 1 ' Columns B to L
            If j = 4 Or j = 6 Then ' Columns 3 and 5 are time columns
                item = item & Format(ws.Cells(i, j).value, "hh:mm AM/PM") & vbTab
            Else
                item = item & ws.Cells(i, j).value & vbTab
            End If
        Next j

        ' Add the hyperlink text to the end of the item
        item = item & linkAddress

        ' Add the concatenated item to the ListBox
        Customerdetails.ListBox1.AddItem Trim(item)
    End If
Next i

End Sub



Private Sub CommandButton5_Click()
    
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub TextBox2_Change()
    Dim ws As Worksheet
    Dim searchName As String
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim colCount As Long
    
    ' Set the worksheet reference
    Set ws = ThisWorkbook.Sheets("CHECK-IN") ' Change the sheet name as needed
    
    ' Clear existing items from the ListBox
    Customerdetails.ListBox1.Clear
    
    ' Get the search name from TextBox2 on the UserForm
    searchName = Customerdetails.TextBox2.value
    
    ' Find the last row with data in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Find the number of columns to consider (B to M)
    colCount = 9
    
    ' Loop through the rows to search for the name
    For i = 2 To lastRow ' Assuming row 1 contains headers
        If ws.Cells(i, 7).value = searchName Then
            ' Match found, populate ListBox with data from columns B to M
            Dim item As String
            item = ""
            
            For j = 2 To colCount + 1 ' Columns B to M
                If j = 3 Or j = 5 Then ' Columns 3 and 5 are time columns
                    item = item & Format(ws.Cells(i, j).value, "hh:mm AM/PM") & vbTab
                Else
                    item = item & ws.Cells(i, j).value & vbTab
                End If
            Next j
            End If
            Next i
End Sub


Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrHandler
    ThisWorkbook.FollowHyperlink Address:=ListBox1.List(ListBox1.ListIndex, 1)
ExitHere:
    Exit Sub
ErrHandler:
    If Err.Number = -2147221014 Then
        MsgBox "Wrong link!"
    Else
        MsgBox "Error: " & Err.Description
    End If
    Resume ExitHere
End Sub



Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Click()

End Sub
