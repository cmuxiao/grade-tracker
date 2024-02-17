Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Nicholas Cai
' Date: 2023-07-20
' Program title: Student Grade Tracker
' Description: A program to manage and view student grades
'===========================================================+
Private cancel As Boolean
Public Function ShowInputsDialog(txtstudentID As String, txtfirstname As String, _
            txtlastname As String, txtA1 As Integer, txtA2 As Integer, txtA3 As Integer, txtA4 As Integer, _
            txtMidterm As Integer, txtExam As Integer) As Boolean
    Call UserForm_Initialize
    Me.Show
    ShowInputsDialog = Not cancel
    Unload Me
End Function

Private Sub empid_Click()
End Sub

Private Sub Label10_Click()
End Sub

Private Sub Label3_Click()
End Sub

Private Sub Label4_Click()
End Sub

Private Sub Label9_Click()

End Sub

Private Sub txtstudentID_Change()

End Sub

Private Sub UserForm_Initialize()
   txtstudentID.Value = ""
   txtstudentID.SetFocus
   txtfirstname.Value = ""
   txtlastname.Value = ""
   txtA1.Value = ""
   txtA2.Value = ""
   txtA3.Value = ""
   txtA4.Value = ""
   txtMidterm.Value = ""
   txtExam.Value = ""
End Sub
Private Sub CancelButton_Click()
   Unload Me
End Sub

Private Sub SubmitButton_Click()
   Dim emptyRow As Long
   Worksheets("grades").Activate
   emptyRow = WorksheetFunction.CountA(Range("A:A")) + 1
    
    
    On Error GoTo TextErrorHandler
   Cells(emptyRow, 1).Value = txtstudentID.Value
   Cells(emptyRow, 2).Value = txtfirstname.Value
   Cells(emptyRow, 3).Value = txtlastname.Value
   Cells(emptyRow, 4).Value = txtA1.Value
   Cells(emptyRow, 5).Value = txtA2.Value
   Cells(emptyRow, 6).Value = txtA3.Value
   Cells(emptyRow, 7).Value = txtA4.Value
   Cells(emptyRow, 8).Value = txtMidterm.Value
   Cells(emptyRow, 9).Value = txtExam.Value
Exit Sub

TextErrorHandler:
    MsgBox "Please Enter the Correct Values in the Fields" _
                    & vbCrLf & "[String] Student ID" _
                    & vbCrLf & "[String] FirstName" _
                    & vbCrLf & "[String] LastName" _
                    & vbCrLf & "[Numbers] Assignments" _
                    & vbCrLf & "[Exams] Exams ", vbExclamation, "Field Input Error"
Unload Me
End Sub