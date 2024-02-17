Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Nicholas Cai
' Date: 2023-07-20
' Program title: Student Grade Tracker
' Description: A program to manage and view student grades
'===========================================================+

Public messageBox As String
Public Sub ImportFiles()
    Dim fileDialog As fileDialog
    Dim messageBox As String
    Dim selectFile As Variant
    Dim comfirmation As Boolean

    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    messageBox = ""
    fileDialog.InitialFileName = ThisWorkbook.Path
    With fileDialog
    comfirmation = .Show
        If comfirmation Then
            For Each selectFile In .SelectedItems
                messageBox = messageBox & vbCrLf & selectFile
            Next
        End If
    End With
        
End Sub
Sub Main()
Dim connection As New ADODB.connection
Dim courseCode As String
Dim courseName As String
    
    MsgBox "Please Import 'Registrar.mdb'", vbInformation, "Select selectFile"
    Call ImportFiles
    With connection
        .ConnectionString = "Data Source=" & ThisWorkbook.Path & "\Registrar.mdb"
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    
    If CourseSelect.showListBox(connection, courseCode, courseName) Then
        Call GetInfo(connection, courseCode, courseName)
    End If

    connection.Close
    Set connection = Nothing
    
End Sub
Sub GetInfo(connection As ADODB.connection, courseCode As String, courseName As String)

    Dim recordSet As New ADODB.recordSet
    Dim SQL As String
    Dim rowCount As Integer
    Dim startcell As Range
    Dim ws As Worksheet
    Dim i As Integer
    Dim rng As Range
    
    For Each ws In ThisWorkbook.Worksheets
    
        If ws.Name = "grades" Then
        Application.DisplayAlerts = False
        Sheets("grades").Delete
        Application.DisplayAlerts = True
        
        Else
        
          Sheets.Add.Name = "grades"
        
        End If
    Next ws
    
    With Worksheets("grades").Range("A1")
        .Value = "Course Name:"
        .Font.Bold = True
    End With
    
    With Worksheets("grades").Range("B1")
        .Value = courseName
        .Font.Bold = True
    End With
        With Worksheets("grades").Range("A2")
        .Value = "Course Code:"
        .Font.Bold = True
    End With
    With Worksheets("grades").Range("B2")
        .Value = courseCode
        .Font.Bold = True
    End With
    
    With Worksheets("grades")
    .Range("A3").Value = "Student ID"
    .Range("B3").Value = "First Name"
    .Range("C3").Value = "Last Name"
    .Range("D3").Value = "A1"
    .Range("E3").Value = "A2"
    .Range("F3").Value = "A3"
    .Range("G3").Value = "A4"
    .Range("H3").Value = "Midterm"
    .Range("I3").Value = "Final"
    .Range("A3:I3").Font.Bold = True
    End With

    Worksheets("grades").Range("A1:K1").EntireColumn.AutoFit

   Set startcell = Worksheets("grades").Range("A4")

    SQL = "SELECT students.firstName, students.lastName, grades.ID, grades.studentID, grades.course, grades.A1, grades.A2, grades.A3, grades.A4, grades.Midterm, grades.Exam " _
        & "FROM students INNER JOIN (courses INNER JOIN grades ON courses.courseCode = grades.course) ON students.studentID = grades.studentID " _
        & "WHERE grades.course = '" & courseCode & "'"
    
    With recordSet
        .Open SQL, connection
        
        rowCount = 0
      
        Do Until .EOF
            startcell.Offset(rowCount, 0).Value = .Fields("studentID")
            startcell.Offset(rowCount, 1).Value = .Fields("firstName")
            startcell.Offset(rowCount, 2).Value = .Fields("lastName")
            startcell.Offset(rowCount, 3).Value = .Fields("A1")
            startcell.Offset(rowCount, 4).Value = .Fields("A2")
            startcell.Offset(rowCount, 5).Value = .Fields("A3")
            startcell.Offset(rowCount, 6).Value = .Fields("A4")
            startcell.Offset(rowCount, 7).Value = .Fields("Midterm")
            startcell.Offset(rowCount, 8).Value = .Fields("Exam")
            rowCount = rowCount + 1
            
            recordSet.MoveNext

        Loop
        
        .Close
        
    End With

    Set recordSet = Nothing
    

    Application.DisplayAlerts = False
    With Worksheets("grades")
        With .Range("K3:L3")
            .Value = "Averages"
            .MergeCells = True
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        .Range("K4") = "A1"
        .Range("L4").FormulaArray = "=Average(D4:D60)"
        .Range("K5") = "A2"
        .Range("L5").FormulaArray = "=Average(E4:E60)"
        .Range("K6") = "A3"
        .Range("L6").FormulaArray = "=Average(F4:F60)"
        .Range("K7") = "A4"
        .Range("L7").FormulaArray = "=Average(G4:G60)"
        .Range("K8") = "Midterm"
        .Range("L8").FormulaArray = "=Average(H4:H60)"
        .Range("K9") = "Final"
        .Range("L9").FormulaArray = "=Average(I4:I60)"
        .Range("K10") = "Total"
        .Range("K3:M10").Interior.ColorIndex = 24
    End With

    With Worksheets("grades")
        With .Range("K12:L12")
            .Value = "Grades Breakdown"
            .MergeCells = True
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        .Range("K13") = "Assessment"
        .Range("L13") = "Weight"
        .Range("K13:L13").Font.Bold = True
        
        
        .Range("K14") = "A1"
        .Range("K15") = "A2"
        .Range("K16") = "A3"
        .Range("K17") = "A4"
        .Range("L14:L17") = Format("0.05", "Percent")
        
        .Range("K18") = "Midterm"
        .Range("L18") = Format("0.30", "Percent")
        
        .Range("K19") = "Final"
        .Range("L19") = Format("0.50", "Percent")
        
        .Range("K20") = "Total"
        .Range("L20") = Format("1", "Percent")
        
        .Range("K1:L20").EntireColumn.AutoFit
        
    End With
    
    With Worksheets("grades")
        .Range("M3") = "Weighted"
        .Range("M3").Font.Bold = True
        
        .Range("M4").Value = Round(Range("L4") * (Range("L14")), 1)
        .Range("M5").Value = Round(Range("L5") * (Range("L15")), 1)
        .Range("M6").Value = Round(Range("L6") * (Range("L16")), 1)
        .Range("M7").Value = Round(Range("L7") * (Range("L17")), 1)
        .Range("M8").Value = Round(Range("L8") * (Range("L18")), 1)
        .Range("M9").Value = Round(Range("L9") * (Range("L19")), 1)
        
        .Range("M10").Formula = "=SUM(M4:M9)"
    End With
    
End Sub

Sub Reset()
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "grades" Then
        Application.DisplayAlerts = False
        Sheets("grades").Delete
        Application.DisplayAlerts = True
        End If
Next ws
End Sub

Sub Start()
AcessData.Show
End Sub