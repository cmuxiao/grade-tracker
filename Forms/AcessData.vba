Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Nicholas Cai
' Date: 2023-07-20
' Program title: Student Grade Tracker
' Description: A program to manage and view student grades
'===========================================================+

Private Sub AveragesButton_Click()
Dim averageRange As Range
Dim minimum As Double
Dim maximum As Double

Set averageRange = Worksheets("grades").Range("M10")
minimum = WorksheetFunction.min(Worksheets("grades").Range("I4:I65"))
maximum = WorksheetFunction.max(Worksheets("grades").Range("I4:I65"))
MsgBox "Class average for this course is: " & Round(averageRange, 0) & "%" & vbCrLf & "Highest Final Exam score is " & maximum & "%" & " and the lowest is " & minimum & "%", vbInformation
Unload Me

End Sub

Private Sub ChartButton_Click()
Dim chart As Shape
Dim assignment As String
Dim gradeRng As Range

On Error GoTo ErrorHandler
assignment = InputBox("Please Enter the Assignment you would like as the Histogram: " & vbCrLf _
                    & vbCrLf & "[A1] Assignment 1" _
                    & vbCrLf & "[A2] Assignment 2" _
                    & vbCrLf & "[A3] Assignment 3" _
                    & vbCrLf & "[A4] Assignment 4" _
                    & vbCrLf & "[MidTerm] MidTerm Exam" _
                    & vbCrLf & "[Exam] Final Exam", "Assignment Selection")
If assignment = "A1" Then
    Set gradeRng = Worksheets("grades").Range("D4:D60")
ElseIf assignment = "A2" Then
    Set gradeRng = Worksheets("grades").Range("E4:E60")
ElseIf assignment = "A3" Then
    Set gradeRng = Worksheets("grades").Range("F4:F60")
ElseIf assignment = "A4" Then
    Set gradeRng = Worksheets("grades").Range("G4:G60")
ElseIf assignment = "MidTerm" Then
    Set gradeRng = Worksheets("grades").Range("H4:H60")
ElseIf assignment = "Exam" Then
    Set gradeRng = Worksheets("grades").Range("I4:I60")
End If

For Each chart In Worksheets("grades").Shapes
    If chart.Name = "Chart" Or chart.Type = msoChart Then
        chart.Delete
    End If

Next chart
Worksheets("grades").Select
gradeRng.Select
Worksheets("grades").Shapes.AddChart2(366, xlHistogram).Select
With ActiveChart
    .ChartTitle.Caption = assignment
    .PlotBy = xlColumns
    .HasLegend = False
    With .Parent
        .Name = "Chart"
        .Top = 330
        .Left = 450
        .Width = 300
    End With
    With .Axes(xlCategory, xlPrimary)
       .HasTitle = True
       .AxisTitle.Text = "Grade (%)"
     End With
    With .Axes(xlValue, xlPrimary)
       .HasTitle = True
       .AxisTitle.Text = "Number of Students"
     End With
  ActiveChart.ChartGroups(1).BinsType = xlBinsTypeBinCount
  ActiveChart.ChartGroups(1).BinsCountValue = 10
End With
Exit Sub

ErrorHandler:
    MsgBox "Please Enter the Following Assignment Codes ONLY", vbCritical
Unload Me
End Sub

Private Sub Label2_Click()
End Sub

Private Sub Label3_Click()
End Sub

Private Sub Label4_Click()
End Sub

Private Sub Label5_Click()
End Sub

Private Sub Label6_Click()
End Sub

Private Sub ResetButton_Click()
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "grades" Then
        Application.DisplayAlerts = False
        Sheets("grades").Delete
        Application.DisplayAlerts = True
        End If
        
Next ws
Unload Me

End Sub

Private Sub ReportButton_Click()
Dim wdDoc As word.Document
Dim wdApp As New word.Application
Dim tblRange As Excel.Range
Dim wdTable As word.Table
Dim WordRng As word.Range
Dim wdSel As word.Selection
Dim chrt As ChartObject

wdApp.Visible = True
wdApp.Activate
Set wdDoc = wdApp.Documents.Add
Set wdSel = wdDoc.ActiveWindow.Selection
wdSel.TypeText Text:="The image below is the student grades info, this Word document is made from Excel using VBA"
wdSel.TypeParagraph
On Error GoTo ErrorHandler
Set tblRange = ThisWorkbook.Worksheets("grades").Range("A1:Q60")
tblRange.Copy
Set wdTable = wdDoc.Tables(1)
wdTable.AutoFitBehavior (wdAutoFitWindow)


With wdApp.Selection
    .EndKey Unit:=wdStory
    .TypeParagraph
    .PasteSpecial Link:=False, DataType:=wdPasteBitmap, _
    Placement:=wdInLine, DisplayAsIcon:=False
End With

ErrorHandler:
    If tblRange Is Nothing Then
    MsgBox "Select a  Proper Range and try again! (SELECT COURSE)", vbExclamation, "Export Table To Word"
    Exit Sub
    End If
    Err.Clear
    Resume Next
Set chrt = Worksheets("grades").ChartObjects("Chart")
chrt.chart.ChartArea.Copy
With wdApp.Selection
    .EndKey Unit:=wdStory
    .PasteSpecial Link:=False, DataType:=wdPasteShape
End With

wdDoc.SaveAs ThisWorkbook.Path & "\" & _
"StudentGradeReport" & ".docx"

If ActiveDocument.Saved = False Then
ActiveDocument.Save
End If

Unload Me
End Sub

Private Sub StudentButton_Click()
    Dim txtstudentID As String
    Dim txtfirstname As String
    Dim txtlastname As String
    Dim txtA1 As Integer
    Dim txtA2 As Integer
    Dim txtA3 As Integer
    Dim txtA4 As Integer
    Dim txtMidterm As Integer
    Dim txtExam As Integer
    

    If AddStudent.ShowInputsDialog(txtstudentID, txtfirstname, _
            txtlastname, txtA1, txtA2, txtA3, txtA4, _
            txtMidterm, txtExam) Then
     End If

Unload Me
End Sub

Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Unload Me
End Sub