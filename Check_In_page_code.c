Private Sub Checkoutbutton_Click()
Dim erow As Long
'----------------------------------------------------------------------------------'
'CHNAGE THE STATUS BOOKED TO VACCATED'

'------------------------------------------------------------------------------------------'

 AnswerYes = MsgBox("CONFIRM CHECK OUT ", vbQuestion + vbYesNo, "User Repsonse")

 If AnswerYes = vbYes Then
'======================================================================================================================'
If TextBox24 Like "" Then
TextBox24 = 0
End If

a = Worksheets("CHECK-IN BE").Cells(Rows.Count, 1).End(xlUp).Row
For h = 2 To a
If TextBox2.value Like Worksheets("CHECK-IN BE").Cells(h, 11) And Worksheets("CHECK-IN BE").Cells(h, 21) Like "BOOKED" Then
Worksheets("CHECK-IN BE").Rows(h).Copy
Worksheets("CHECK-OUT").Activate
b = Worksheets("CHECK-OUT").Cells(Rows.Count, 1).End(xlUp).Row
Worksheets("CHECK-OUT").Cells(b + 1, 1).Select
ActiveSheet.Paste
Worksheets("CHECK-IN BE").Activate

'-------------------------------------------------------------------------------------------------------------------------'

Dim sh As Worksheet
Dim Ls As Long
Set sh = ThisWorkbook.Sheets("CHECK-OUT")
Ls = Worksheets("CHECK-OUT").Cells(Rows.Count, 1).End(xlUp).Row
Set ah = ThisWorkbook.Sheets("CHECK-OUT")
Dim trow As Long

For i = 2 To Ls
'========================================================================================================================='
If (TextBox22.value Like Worksheets("CHECK-OUT").Cells(i, 2) Or TextBox1.value Like Worksheets("CHECK-OUT").Cells(i, 8)) And (Worksheets("CHECK-OUT").Cells(i, 21) Like "BOOKED" And TextBox2.value Like Worksheets("CHECK-OUT").Cells(i, 11)) Then

On Error Resume Next
trow = Sheets("CHECK-OUT").Cells(Rows.Count, 1).End(xlUp).Row
For x = 2 To trow
If (TextBox22.value Like Worksheets("CHECK-OUT").Cells(x, 2) Or TextBox1.value Like Worksheets("CHECK-OUT").Cells(x, 8)) And (Worksheets("CHECK-OUT").Cells(x, 21) Like "BOOKED" And TextBox2.value Like Worksheets("CHECK-OUT").Cells(x, 11)) Then
Sheets("CHECK-OUT").Cells(x, 35) = "BILL NOT ISSUED"
Sheets("CHECK-OUT").Cells(x, 36) = TextBox26.text
End If
Next x


' THIS BELOW CODE FOR INPUT THE DATA TO THE CHECK-OUT PAGE'
With sh
sh.Cells(Ls, 1) = Ls - 1
sh.Cells(Ls, 5) = UserForm3.TextBox4.value
sh.Cells(Ls, 6) = UserForm3.TextBox5.value
sh.Cells(Ls, 21) = "VACCATED"
sh.Cells(Ls, 32) = UserForm3.TextBox20.value
sh.Cells(Ls, 33) = UserForm3.ComboBox1.text
sh.Cells(Ls, 29) = UserForm3.TextBox23.text

sh.Cells(Ls, 30) = UserForm3.TextBox24.text
End With

Dim L As Long
Set h = ThisWorkbook.Sheets("CHECK-OUT FE")
L = Worksheets("CHECK-OUT FE").Cells(Rows.Count, 1).End(xlUp).Row
With h
h.Cells(L + 1, 1) = L
h.Cells(L + 1, 2) = TextBox22.value
h.Cells(L + 1, 3) = TextBox3.value
h.Cells(L + 1, 4) = TextBox7.value
h.Cells(L + 1, 5) = TextBox4.value
h.Cells(L + 1, 6) = TextBox5.value
h.Cells(L + 1, 7) = TextBox6.value
h.Cells(L + 1, 8) = TextBox1.value
h.Cells(L + 1, 9) = TextBox2.text
h.Cells(L + 1, 10) = TextBox19.value
h.Cells(L + 1, 12) = Label9.Caption
h.Cells(L + 1, 14) = TextBox27.value
h.Cells(L + 1, 17) = TextBox16.value
h.Cells(L + 1, 19) = "VACCATED"
h.Cells(L + 1, 26) = TextBox13.text
h.Cells(L + 1, 27) = TextBox23.text
h.Cells(L + 1, 30) = TextBox20.value
h.Cells(L + 1, 31) = ComboBox1.text
h.Cells(L + 1, 32) = "BILL NO ISSUED"
h.Cells(L + 1, 28) = TextBox24.value
h.Cells(L + 1, 30) = TextBox20.value
Total = TextBox21.value
Final = Total / Label9.Caption
h.Cells(L + 1, 15) = Final
h.Cells(L + 1, 16) = TextBox21.value
End With

End If
Next i
'=======================================================================================================================''
'================================================================================================================'

'THIS BELOW CODE FOR MAKING THE CHECKOUT CUSTOMER DATA  STATUS ARE MAKE US A VACCATED IN THE THE CHECK-IN PAGE'

erow = Worksheets("CHECK-IN BE").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To erow
If ((TextBox22.value Like Worksheets("CHECK-IN BE").Cells(i, 2) Or TextBox2.value Like Worksheets("CHECK-IN BE").Cells(i, 11)) And (Worksheets("CHECK-IN BE").Cells(i, 21) Like "BOOKED")) Then
Worksheets("CHECK-IN BE").Cells(i, 21) = "VACCATED"
Worksheets("CHECK-IN BE").Cells(i, 5) = UserForm3.TextBox4.value
Worksheets("CHECK-IN BE").Cells(i, 6) = UserForm3.TextBox5.value
Worksheets("CHECK-IN BE").Cells(i, 32) = UserForm3.TextBox20.value
Worksheets("CHECK-IN BE").Cells(i, 33) = UserForm3.ComboBox1.value
Worksheets("CHECK-IN BE").Cells(i, 29) = UserForm3.TextBox23.text
Worksheets("CHECK-IN BE").Cells(i, 34) = "BILL NOT ISSUED"
Worksheets("CHECK-IN BE").Cells(i, 35) = TextBox26.value
Worksheets("CHECK-IN BE").Cells(i, 30) = TextBox24.value

End If
Next i
End If

'=.=.=.=.=.=.=.=.=.=.==.=.=.=.==.=.=.=.=.=.=.=.=.=.=.==.=.=.=.=.=.=.=.=.=.=.=.=.=.=.=.=.=.=.=.=.=.'
erow = Worksheets("CHECK-IN").Cells(Rows.Count, 1).End(xlUp).Row
For j = 2 To erow
If ((TextBox22.value Like Worksheets("CHECK-IN").Cells(j, 2) Or TextBox2.value Like Worksheets("CHECK-IN").Cells(j, 11)) And Worksheets("CHECK-IN").Cells(j, 19) Like "BOOKED") Then
Worksheets("CHECK-IN").Cells(j, 19) = "VACCATED"
Worksheets("CHECK-IN").Cells(j, 5) = UserForm3.TextBox4.value
Worksheets("CHECK-IN").Cells(j, 6) = UserForm3.TextBox5.value
Worksheets("CHECK-IN").Cells(j, 30) = UserForm3.TextBox20.value
Worksheets("CHECK-IN").Cells(j, 31) = UserForm3.ComboBox1.value
Worksheets("CHECK-IN").Cells(j, 27) = UserForm3.TextBox23.text
Worksheets("CHECK-IN").Cells(j, 29) = UserForm3.TextBox24.text

End If
'======================================================================================================================='
Next j
Next h
End If
MsgBox "CHECKOUT DONE"
End Sub

'THIS FOR BILL PRINT'
Private Sub CommandButton3_Click()
Dim ps As Worksheet
Dim lastRow As Long
  Dim billNumber As String
Dim searchText1 As String
Dim searchText2 As String
 Dim foundRow As Long
        
Dim trow As Long
Dim erow As Long
Dim hrow As Long
Dim ws As Long
Dim hs As Worksheet
Dim prevNumber As Long
Dim newNumber As Long
    


Set sh = ThisWorkbook.Sheets("CHECK-OUT")
Set Dsh = ThisWorkbook.Sheets("New Bill")
On Error Resume Next
trow = Sheets("CHECK-OUT").Cells(Rows.Count, 1).End(xlUp).Row
For x = 2 To trow
If ref.Caption = "" Then
If ((TextBox22.value Like Worksheets("CHECK-OUT").Cells(x, 2) Or TextBox2.value Like Worksheets("CHECK-OUT").Cells(x, 11)) And Worksheets("CHECK-OUT").Cells(x, 35) Like "BILL NOT ISSUED") Then

Previous = Worksheets("CHECK-OUT").Cells(1, 60)
If Previous = 1 Then
newNumber = Previous
Previous = Previous + 1
Else
newNumber = Previous
Previous = Previous + 1
End If
Worksheets("CHECK-OUT").Cells(1, 60) = Previous

hrow = Worksheets("CHECK-OUT FE").Cells(Rows.Count, 1).End(xlUp).Row
For j = 2 To hrow
If TextBox22.value Like Worksheets("CHECK-OUT FE").Cells(j, 2) Then
Worksheets("CHECK-OUT FE").Cells(j, 32) = "BILL ISSUED"
Worksheets("CHECK-OUT FE").Cells(j, 33) = newNumber
Worksheets("CHECK-OUT FE").Cells(j, 34) = TextBox26.value
End If
Next j

erow = Worksheets("CHECK-IN BE").Cells(Rows.Count, 1).End(xlUp).Row
For L = 2 To erow
If TextBox22.value Like Worksheets("CHECK-IN BE").Cells(L, 2) Then
Worksheets("CHECK-IN BE").Cells(L, 34) = "BILL ISSUED"
Worksheets("CHECK-IN BE").Cells(L, 35) = newNumber
Worksheets("CHECK-IN BE").Cells(L, 36) = TextBox26.value
End If
Next L

ws = Worksheets("CHECK-IN").Cells(Rows.Count, 1).End(xlUp).Row
For h = 2 To ws
If TextBox22.value Like Worksheets("CHECK-IN").Cells(h, 2) Then
Worksheets("CHECK-IN").Cells(h, 32) = "BILL ISSUED"
Worksheets("CHECK-IN").Cells(h, 33) = newNumber
Worksheets("CHECK-IN").Cells(h, 34) = TextBox26.value
End If
Next h

With sh
sh.Cells(x, 35) = "BILL ISSUED"
sh.Cells(x, 34) = newNumber
Dsh.Range("E4") = sh.Cells(x, 34)
Dsh.Range("E5") = sh.Cells(x, 5)
Dsh.Range("E7") = sh.Cells(x, 3)
Dsh.Range("I7") = sh.Cells(x, 5)
Dsh.Range("E6") = sh.Cells(x, 11)
Dsh.Range("E9") = sh.Cells(x, 7)
Dsh.Range("E10") = sh.Cells(x, 8)
Dsh.Range("E11") = sh.Cells(x, 28)
Dsh.Range("F15") = sh.Cells(x, 12)
Dsh.Range("C15") = Label9.Caption
a = Label9.Caption
value = TextBox21 / a
Dsh.Range("E15") = value
Total = TextBox21
SGST = Total * 0.09
CGST = Total * 0.09
taxt = SGST + CGST
t = Total + CGST + SGST
Final = t + TextBox24
Dsh.Range("L15") = t
Dsh.Range("I15") = SGST
Dsh.Range("J15") = CGST
Dsh.Range("L19") = Total
Dsh.Range("L20") = SGST
Dsh.Range("L21") = CGST
Dsh.Range("E12") = TextBox26.value
Dsh.Range("H15") = Total
Dsh.Range("H22") = sh.Cells(x, 29)
Dsh.Range("L22") = sh.Cells(x, 30)
Dsh.Range("K15") = taxt
Dsh.Range("L23") = Final
End With
End If

Else

If ref.Caption = 1 Then
If TextBox22.value Like Worksheets("CHECK-OUT").Cells(x, 2) Then


hrow = Worksheets("CHECK-OUT FE").Cells(Rows.Count, 1).End(xlUp).Row
For j = 2 To hrow
If TextBox22.value Like Worksheets("CHECK-OUT FE").Cells(j, 2) Then
Worksheets("CHECK-OUT FE").Cells(j, 32) = "BILL ISSUED"
End If
Next j

erow = Worksheets("CHECK-IN BE").Cells(Rows.Count, 1).End(xlUp).Row
For L = 2 To erow
If TextBox22.value Like Worksheets("CHECK-IN BE").Cells(L, 2) Then
Worksheets("CHECK-IN BE").Cells(L, 34) = "BILL ISSUED"
End If
Next L

ws = Worksheets("CHECK-IN").Cells(Rows.Count, 1).End(xlUp).Row
For h = 2 To ws
If TextBox22.value Like Worksheets("CHECK-IN").Cells(h, 2) Then
Worksheets("CHECK-IN").Cells(h, 32) = "BILL ISSUED"
End If
Next h


With sh

Dsh.Range("E4") = sh.Cells(x, 34)
Dsh.Range("E5") = sh.Cells(x, 5)
Dsh.Range("E7") = sh.Cells(x, 3)
Dsh.Range("I7") = sh.Cells(x, 5)
Dsh.Range("E6") = sh.Cells(x, 11)
Dsh.Range("E9") = sh.Cells(x, 7)
Dsh.Range("E10") = sh.Cells(x, 8)
Dsh.Range("E11") = sh.Cells(x, 28)
Dsh.Range("F15") = sh.Cells(x, 12)
Dsh.Range("C15") = Label9.Caption
a = Label9.Caption
value = TextBox21 / a
Dsh.Range("E15") = value
Total = TextBox21
SGST = Total * 0.09
CGST = Total * 0.09
taxt = SGST + CGST
t = Total + CGST + SGST
Final = t + TextBox24
Dsh.Range("L15") = t
Dsh.Range("I15") = SGST
Dsh.Range("J15") = CGST
Dsh.Range("L19") = Total
Dsh.Range("L20") = SGST
Dsh.Range("L21") = CGST
Dsh.Range("E12") = TextBox26.value
Dsh.Range("H15") = Total
Dsh.Range("H22") = sh.Cells(x, 29)
Dsh.Range("L22") = sh.Cells(x, 30)
Dsh.Range("K15") = taxt
Dsh.Range("L23") = Final
End With
End If
End If
End If
Next x
Unload UserForm3
Dsh.PrintPreview
Load UserForm3
End Sub
'=================================================================================================================================='
Private Sub CommandButton4_Click()
Unload UserForm3
UserForm3.Show
End Sub

'SEARCH BUTTON'

Private Sub CommandButton5_Click()
arow = Worksheets("CHECK-IN BE").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To arow
'=============================================================================================================================='
'IF STATEMENT FOR ENTIRE SUB START '
If ((UserForm3.TextBox22.value Like Worksheets("CHECK-IN BE").Cells(i, 2) Or UserForm3.TextBox2.value Like Worksheets("CHECK-IN BE").Cells(i, 11)) And (Worksheets("CHECK-IN BE").Cells(i, 21) Like Trim("BOOKED") And UserForm3.ref.Caption = "")) Then
Label10 = Worksheets("CHECK-IN BE").Cells(i, 2)
TextBox3 = Worksheets("CHECK-IN BE").Cells(i, 3)
TextBox7 = Format(Worksheets("CHECK-IN BE").Cells(i, 4), "hh:mm:ss am/pm")

'==============================================================================================================================

If Worksheets("CHECK-IN BE").Cells(i, 5) Like "" Then
TextBox4 = Format(Date, "DD-MMM-YYYY")

MsgBox "CHECK-OUT DATE ENTERED"
MsgBox "CLICK SEARCH FOR CALCULATE RENT"
Else
TextBox4 = Worksheets("CHECK-IN BE").Cells(i, 5)
End If

'=================================================================================================================================='
If Worksheets("CHECK-IN BE").Cells(i, 19) Like "" Then
MsgBox "NO AMOUNT RECIVIED ON CHECK-IN"
TextBox16.value = 0
Else
TextBox16.value = Worksheets("CHECK-IN BE").Cells(i, 19)
End If
'======================================================================================================================================'
TextBox6 = Worksheets("CHECK-IN BE").Cells(i, 7)
TextBox1 = Worksheets("CHECK-IN BE").Cells(i, 8)
TextBox9 = Worksheets("CHECK-IN BE").Cells(i, 9)
TextBox11 = Worksheets("CHECK-IN BE").Cells(i, 10)
TextBox2 = Worksheets("CHECK-IN BE").Cells(i, 11)
TextBox8 = Worksheets("CHECK-IN BE").Cells(i, 12)
TextBox18 = Worksheets("CHECK-IN BE").Cells(i, 13)
TextBox17 = Worksheets("CHECK-IN BE").Cells(i, 20)
TextBox13 = Worksheets("CHECK-IN BE").Cells(i, 28)
TextBox19 = Worksheets("CHECK-IN BE").Cells(i, 12)
Label9 = Worksheets("CHECK-IN BE").Cells(i, 15)
Worksheets("CHECK-IN BE").Cells(i, 5) = TextBox4.value
TextBox21 = Worksheets("CHECK-IN BE").Cells(i, 18)
TextBox22 = Worksheets("CHECK-IN BE").Cells(i, 2)
TextBox27 = Worksheets("CHECK-IN BE").Cells(i, 16)
'==============================================================================================================================='

'==============================================================================================================================='
If TextBox4 = Worksheets("CHECK-IN BE").Cells(i, 5) Like "" Then
Worksheets("CHECK-IN BE").Cells(i, 5) = TextBox4.value
Else
TextBox4 = Worksheets("CHECK-IN BE").Cells(i, 5)
End If
'==========================================================================================================================='
TextBox21 = Worksheets("CHECK-IN BE").Cells(i, 18)
'================================================================================================================================'
'==========================================================================================================================='
MsgBox "DATA HAS BEEN ENTERED"
'END OF THE IF STATEMENT'
ElseIf ((UserForm3.TextBox22.value Like Worksheets("CHECK-IN BE").Cells(i, 2)) And (Worksheets("CHECK-IN BE").Cells(i, 34) Like "BILL NOT ISSUED" Or Worksheets("CHECK-IN BE").Cells(i, 34) Like "BILL ISSUED") And UserForm3.ref.Caption Like 1) Then

If Worksheets("CHECK-IN BE").Cells(i, 5) Like "" Then
TextBox4 = Format(Date, "DD-MMM-YYYY")
MsgBox "CHECK-OUT DATE ENTERED"
MsgBox "CLICK SEARCH FOR CALCULATE RENT"
Else
TextBox4 = Worksheets("CHECK-IN BE").Cells(i, 5)
End If

'=================================================================================================================================='
If Worksheets("CHECK-IN BE").Cells(i, 19) Like "" Then
MsgBox "NO AMOUNT RECIVIED ON CHECK-IN"
TextBox16.value = 0
Else
TextBox16.value = Worksheets("CHECK-IN BE").Cells(i, 19)
End If
'======================================================================================================================================'
TextBox6 = Worksheets("CHECK-IN BE").Cells(i, 7)
TextBox1 = Worksheets("CHECK-IN BE").Cells(i, 8)
TextBox9 = Worksheets("CHECK-IN BE").Cells(i, 9)
TextBox11 = Worksheets("CHECK-IN BE").Cells(i, 10)
TextBox2 = Worksheets("CHECK-IN BE").Cells(i, 11)
TextBox8 = Worksheets("CHECK-IN BE").Cells(i, 12)
TextBox18 = Worksheets("CHECK-IN BE").Cells(i, 13)
TextBox17 = Worksheets("CHECK-IN BE").Cells(i, 20)
TextBox13 = Worksheets("CHECK-IN BE").Cells(i, 28)
TextBox19 = Worksheets("CHECK-IN BE").Cells(i, 12)
Label9 = Worksheets("CHECK-IN BE").Cells(i, 15)
Worksheets("CHECK-IN BE").Cells(i, 5) = TextBox4.value
TextBox20 = Worksheets("CHECK-IN BE").Cells(i, 32)
ComboBox1 = Worksheets("CHECK-IN BE").Cells(i, 33)
TextBox21 = Worksheets("CHECK-IN BE").Cells(i, 18)
TextBox22 = Worksheets("CHECK-IN BE").Cells(i, 2)
TextBox7 = Format(Worksheets("CHECK-IN BE").Cells(i, 4), "hh:mm:ss am/pm")
TextBox3 = Worksheets("CHECK-IN BE").Cells(i, 3)
TextBox23 = Worksheets("CHECK-IN BE").Cells(i, 29)
TextBox24 = Worksheets("CHECK-IN BE").Cells(i, 30)
TextBox26 = Worksheets("CHECK-IN BE").Cells(i, 36)
'==============================================================================================================================='

'==============================================================================================================================='
If TextBox4 = Worksheets("CHECK-IN BE").Cells(i, 5) Like "" Then
Worksheets("CHECK-IN BE").Cells(i, 5) = TextBox4.value
Else
TextBox4 = Worksheets("CHECK-IN BE").Cells(i, 5)
End If
'==========================================================================================================================='
TextBox21 = Worksheets("CHECK-IN BE").Cells(i, 18)
'================================================================================================================================'
If Worksheets("CHECK-IN BE").Cells(i, 19) Like "" Then '
MsgBox "NO AMOUNT RECIVIED ON CHECK-IN"
Else
TextBox16 = Worksheets("CHECK-IN BE").Cells(i, 19)
End If
'==========================================================================================================================='
MsgBox "DATA HAS BEEN ENTERED"

End If
Next i
'=============================================================================================================================='
End Sub

Private Sub CommandButton6_Click()
Me.TextBox4 = Calendar.DatePicker
End Sub

Private Sub CommandButton7_Click()
If TextBox24 Like "" Then
TextBox24.value = 0
End If

If TextBox25 Like "" Then
TextBox25.value = 0
End If

If TextBox16 Like "" Then
TextBox16 = 0
Else
Totalamount = TextBox21 - TextBox25 - TextBox16
CGST = TextBox21 * 0.09
SGST = TextBox21 * 0.09
t = TextBox24 + Totalnumber
Total = Totalamount + CGST + SGST
TextBox20.value = Total + t
End If
End Sub

Private Sub CommandButton8_Click()

Dim a As Long
Dim t As Long
Dim j As Long
Dim k As Long
 AnswerYes = MsgBox("CONFIRM RETURN ", vbQuestion + vbYesNo, "User Repsonse")
If AnswerYes = vbYes Then

a = Worksheets("CHECK-IN BE").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To a
If UserForm3.TextBox22.value Like Worksheets("CHECK-IN BE").Cells(i, 2) Then
Worksheets("CHECK-IN BE").Cells(i, 23) = TextBox27.value
End If
Next i

t = Worksheets("CHECK-IN").Cells(Rows.Count, 1).End(xlUp).Row
For x = 2 To t
If UserForm3.TextBox22.value Like Worksheets("CHECK-IN").Cells(x, 2) Then
Worksheets("CHECK-IN").Cells(x, 21) = TextBox27.value
End If
Next x

j = Worksheets("CHECK-OUT").Cells(Rows.Count, 1).End(xlUp).Row
For L = 2 To j
If UserForm3.TextBox22.value Like Worksheets("CHECK-OUT").Cells(L, 2) Then
Worksheets("CHECK-OUT").Cells(L, 23) = TextBox27.value
End If
Next L

k = Worksheets("CHECK-OUT FE").Cells(Rows.Count, 1).End(xlUp).Row
For f = 2 To k
If UserForm3.TextBox22.value Like Worksheets("CHECK-OUT FE").Cells(f, 2) Then
Worksheets("CHECK-OUT FE").Cells(f, 21) = TextBox27.value
End If
Next f

ElseIf AnswerYes = vbNo Then
End If
End Sub

Private Sub UserForm_Initialize()
  UserForm3.TextBox10.value = Format(Date, "DD-MMM-YY")
  UserForm3.TextBox4.value = Format(Date, "DD-MMM-YY")
  UserForm3.TextBox5.value = Time()
  UserForm3.ComboBox1.List = Array("UPI", "CASH", "UPI+CASH", "DEBIT CARD")
  End Sub
