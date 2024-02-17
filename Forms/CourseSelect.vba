Option Explicit

' ==== CP212 Windows Application Programming ===============+
' Name: Nicholas Cai
' Date: 2023-07-20
' Program title: Student Grade Tracker
' Description: A program to manage and view student grades
'===========================================================+
Public cancel As Boolean
Public Function showListBox(cn As ADODB.connection, courseCode As String, courseName As String) As Boolean
    Call Initialize(cn)
    Call DisplayForm
    
    If Not cancel Then
        courseCode = ListBox1.List(ListBox1.ListIndex, 0)
        courseName = ListBox1.List(ListBox1.ListIndex, 1)
    End If
    
    showListBox = Not cancel
    Unload Me
End Function
Private Sub Initialize(cn As ADODB.connection)
    Dim rs As New ADODB.recordSet
    Dim rowCount As Integer, SQL As String
    Dim coursesArray(15, 2) As Variant

    SQL = "SELECT CourseCode, CourseName FROM courses"
    rowCount = 0
    With rs
        .Open SQL, cn
        Do Until .EOF
            coursesArray(rowCount, 0) = .Fields("CourseCode")
            coursesArray(rowCount, 1) = .Fields("CourseName")
            rowCount = rowCount + 1
            .MoveNext
        Loop
    End With
    
    ListBox1.List = coursesArray
    ListBox1.ListIndex = 0
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Label1_Click()
End Sub

Private Sub OKButton_Click()
    Me.Hide
    cancel = False
End Sub
Private Sub CancelButton_Click()
    Me.Hide
    cancel = True
End Sub

Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then CancelButton_Click
End Sub

Sub DisplayForm()
    CourseSelect.Show
End Sub
