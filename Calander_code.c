
Option Explicit

Sub ButtonClick(btn As MSForms.CommandButton)
    If btn.Caption <> "" Then
    Me.TextBox1.value = btn.Caption & "-" & VBA.Left(Me.cmbMonth.value, 3) & "-" & Me.cmbYear
    End If
    Unload Me
End Sub
Private Sub CB1_Click()
        Call ButtonClick(Me.CB1)
End Sub
Private Sub CB2_Click()
        Call ButtonClick(Me.CB2)
End Sub
Private Sub CB3_Click()
        Call ButtonClick(Me.CB3)
End Sub
Private Sub CB4_Click()
        Call ButtonClick(Me.CB4)
End Sub
Private Sub CB5_Click()
        Call ButtonClick(Me.CB5)
End Sub
Private Sub CB6_Click()
        Call ButtonClick(Me.CB6)
End Sub
Private Sub CB7_Click()
        Call ButtonClick(Me.CB7)
End Sub
Private Sub CB8_Click()
        Call ButtonClick(Me.CB8)
End Sub
Private Sub CB9_Click()
        Call ButtonClick(Me.CB9)
End Sub
Private Sub CB10_Click()
        Call ButtonClick(Me.CB10)
End Sub
Private Sub CB11_Click()
        Call ButtonClick(Me.CB11)
End Sub
Private Sub CB12_Click()
        Call ButtonClick(Me.CB12)
End Sub
Private Sub CB13_Click()
        Call ButtonClick(Me.CB13)
End Sub
Private Sub CB14_Click()
        Call ButtonClick(Me.CB14)
End Sub
Private Sub CB15_Click()
        Call ButtonClick(Me.CB15)
End Sub
Private Sub CB16_Click()
        Call ButtonClick(Me.CB16)
End Sub
Private Sub CB17_Click()
        Call ButtonClick(Me.CB17)
End Sub
Private Sub CB18_Click()
        Call ButtonClick(Me.CB18)
End Sub
Private Sub CB19_Click()
        Call ButtonClick(Me.CB19)
End Sub
Private Sub CB20_Click()
        Call ButtonClick(Me.CB20)
End Sub
Private Sub CB21_Click()
        Call ButtonClick(Me.CB21)
End Sub
Private Sub CB22_Click()
        Call ButtonClick(Me.CB22)
End Sub
Private Sub CB23_Click()
        Call ButtonClick(Me.CB23)
End Sub
Private Sub CB24_Click()
        Call ButtonClick(Me.CB24)
End Sub
Private Sub CB25_Click()
        Call ButtonClick(Me.CB25)
End Sub
Private Sub CB26_Click()
        Call ButtonClick(Me.CB26)
End Sub
Private Sub CB27_Click()
        Call ButtonClick(Me.CB27)
End Sub
Private Sub CB28_Click()
        Call ButtonClick(Me.CB28)
End Sub
Private Sub CB29_Click()
        Call ButtonClick(Me.CB29)
End Sub
Private Sub CB30_Click()
        Call ButtonClick(Me.CB30)
End Sub
Private Sub CB31_Click()
        Call ButtonClick(Me.CB31)
End Sub
Private Sub CB32_Click()
        Call ButtonClick(Me.CB32)
End Sub
Private Sub CB33_Click()
        Call ButtonClick(Me.CB33)
End Sub
Private Sub CB34_Click()
        Call ButtonClick(Me.CB34)
End Sub
Private Sub CB35_Click()
        Call ButtonClick(Me.CB35)
End Sub
Private Sub CB36_Click()
        Call ButtonClick(Me.CB36)
End Sub
Private Sub CB37_Click()
        Call ButtonClick(Me.CB37)
End Sub
Private Sub CB38_Click()
        Call ButtonClick(Me.CB38)
End Sub
Private Sub CB39_Click()
        Call ButtonClick(Me.CB39)
End Sub
Private Sub CB40_Click()
        Call ButtonClick(Me.CB40)
End Sub
Private Sub CB41_Click()
        Call ButtonClick(Me.CB41)
End Sub
Private Sub CB42_Click()
        Call ButtonClick(Me.CB42)
End Sub

''''''''''Code to initialize combobox''''''''''''''''''

Private Sub UserForm_Activate()
        Dim i As Integer
        With Me.cmbMonth
             For i = 1 To 12
                 .AddItem VBA.Format(VBA.DateSerial(2020, i, 1), "MMMM")
             Next i
             .value = VBA.Format(VBA.Date, "MMMM")
        End With
        
        With Me.cmbYear
             For i = VBA.Year(Date) - 20 To VBA.Year(Date) + 20
                 .AddItem i
             Next i
             .value = VBA.Format(VBA.Date, "YYYY")
        End With
        
        Call Show_Date
        
        If Me.TextBox1.value <> "" Then
           Call Highlight_Date(CDate(Me.TextBox1.value))
        End If
        
End Sub

Sub Show_Date()
    Dim First_Date As Date
    Dim Last_Date As Date
    
    First_Date = VBA.CDate("1 " & Me.cmbMonth.value & " " & Me.cmbYear.value)
    Last_Date = VBA.Day(VBA.DateSerial(VBA.Year(First_Date), Month(First_Date) + 1, 1) - 1)
    
    Dim i As Integer
    Dim btn As MSForms.CommandButton
    
''''''''''To remove any caption from button'''''''''''''''''

    For i = 1 To 42
        Set btn = Me.Controls("CB" & i)
        btn.Caption = " "
        Next i

'''''''''''''''Set first date of month'''''''''''''
    For i = 1 To 7
    Set btn = Me.Controls("CB" & i)
    If VBA.Weekday(First_Date) = i Then
       btn.Caption = "1"
    End If
    Next i

'''''''''Set all the dates'''''''''''''''
    Dim btn1 As MSForms.CommandButton
    Dim btn2 As MSForms.CommandButton
    
    For i = 1 To 41
        Set btn1 = Me.Controls("CB" & i)
        Set btn2 = Me.Controls("CB" & i + 1)
        If btn1.Caption <> " " Then
           If btn1.Caption < Last_Date Then
              btn2.Caption = btn1.Caption + 1
           End If
        End If
    Next i
    
    Call reset_color
    
End Sub

'''''''''''''Change event of Month & Year'''''''''''''''

Private Sub cmbMonth_Change()
        If Me.cmbMonth.value <> "" & Me.cmbYear.value <> "" Then
           Call Show_Date
           Me.LabelMonth.Caption = Me.cmbMonth & "-" & Me.cmbYear
        End If
End Sub

Private Sub cmbYear_Change()
        If Me.cmbMonth.value <> "" & Me.cmbYear.value <> "" Then
           Call Show_Date
           Me.LabelMonth.Caption = Me.cmbMonth & "-" & Me.cmbYear
        End If
End Sub

''''''''''''''''''Code for next and previous button''''''''''''''''

Private Sub cmdNext_Click()
        If Me.cmbMonth.ListIndex = 11 Then
           Me.cmbMonth.ListIndex = 0
           Me.cmbYear.value = Me.cmbYear.value + 1
        Else
           Me.cmbMonth.ListIndex = Me.cmbMonth.ListIndex + 1
        End If
End Sub

Private Sub cmdPrevious_Click()
        If Me.cmbMonth.ListIndex = 0 Then
           Me.cmbMonth.ListIndex = 11
           Me.cmbYear.value = Me.cmbYear.value - 1
        Else
           Me.cmbMonth.ListIndex = Me.cmbMonth.ListIndex - 1
        End If
End Sub

'''''''''''''''''''Code to give background color to command button''''''''''''''
Sub reset_color()
    Dim i As Integer
    Dim btn As MSForms.CommandButton

    For i = 1 To 42
        Set btn = Me.Controls("CB" & i)
        If btn.Caption = " " Then
           btn.Enabled = False
           btn.BackColor = &H8000000B
        Else
            btn.Enabled = True
            btn.BackColor = VBA.RGB(254, 198, 25)
        End If
        Next i
        
End Sub


'''''''''''''Code for select the date'''''''''''''

Function DatePicker(Optional DateInput As Object) As String

         Dim str As String
         If VBA.TypeName(DateInput) = "Textbox" Or VBA.TypeName(DateInput) = "Range" Then str = DateInput.value
         If VBA.TypeName(DateInput) = "CommandButton" Or VBA.TypeName(DateInput) = "Label" Then str = DateInput.Caption
        '''' End If
         If VBA.IsDate(str) Then
            Me.TextBox1.value = VBA.Format(CDate(str), "D-MMM-YYYY")
            ''''''Call Highlight_Date(CDate(str))
         Else
            Me.TextBox1.value = ""
         End If
    
    Me.Show ' Show Calendar Form
    'Assign the selected date
    
         If VBA.TypeName(DateInput) = "Textbox" Or VBA.TypeName(DateInput) = "Range" Then
            DateInput.value = Me.TextBox1.value
         ElseIf VBA.TypeName(DateInput) = "CommandButton" Or VBA.TypeName(DateInput) = "Label" Then
                DateInput.Caption = Me.TextBox1.value
         Else
            DatePicker = Me.TextBox1.value
         End If
    
End Function


'''''''''''''''''''Code to highlight date''''''''''''''
Sub Highlight_Date(dt As Date)
    Dim i As Integer
    Dim btn As MSForms.CommandButton
    ''''Me.cmbMonth.Value = VBA.Format(dt, "MMMM")
    ''''Me.cmbYear.Value = VBA.Format(dt, "YYYY")

    For i = 1 To 42
        Set btn = Me.Controls("CB" & i)
        If VBA.CStr(VBA.Day(dt)) = btn.Caption Then
           btn.BackColor = vbWhite
        End If
    Next i
End Sub
