Attribute VB_Name = "Module1"
Option Explicit
Sub Reset_Click()
    Sheets("Applicant Hall Preference raw").Select
    Columns("A:M").Select
    Selection.Copy
    Sheets("Applicant Hall Preference").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
End Sub
Sub AssignHalls_Click()
    'Setting basic values
    Dim number_of_students As Integer
    number_of_students = 500
    Dim number_of_halls As Integer
    number_of_halls = 7
    Dim output_column As Integer
    output_column = 12
    
    'Define vacancies for each hall
    Dim hall_vacancy(1 To 7) As Integer
    hall_vacancy(1) = 18
    hall_vacancy(2) = 7
    hall_vacancy(3) = 9
    hall_vacancy(4) = 19
    hall_vacancy(5) = 17
    hall_vacancy(6) = 15
    hall_vacancy(7) = 14
    
    'Preliminary sorting
    ActiveWorkbook.Worksheets("Applicant Hall Preference").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Applicant Hall Preference").Sort.SortFields.Add2 Key:=Range("K:K") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'Sort by eligible/waitlisted
        ActiveWorkbook.Worksheets("Applicant Hall Preference").Sort.SortFields.Add2 Key:=Range("I:I") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'Sort by number of hall preferences
            ActiveWorkbook.Worksheets("Applicant Hall Preference").Sort.SortFields.Add2 Key:=Range("A:A") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'Sort by matric number
    With ActiveWorkbook.Worksheets("Applicant Hall Preference").Sort
        .SetRange Range("A:L") 'Select range of cells to be sorted
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Generate sorting helper for eligible students
    Range("J2").Select 'Select first cell for eligible students
    ActiveCell.FormulaR1C1 = "=COUNTIF(R2C[-1]:RC[-1], RC[-1]) + RC[-1]/10" 'Change Rxx to row for first eligible student
    Selection.AutoFill Destination:=Range("J2:J207") 'Change to range of eligible students
    
    'Generate sorting helper for waitlisted students
    Range("J208").Select 'Select first cell for waitlisted students
    ActiveCell.FormulaR1C1 = "=COUNTIF(R208C[-1]:RC[-1], RC[-1]) + RC[-1]/10" 'Change Rxx to row for first waitlisted student
    Selection.AutoFill Destination:=Range("J208:J501") 'Change to range of waitlisted students
    
    '234 pattern
    ActiveWorkbook.Worksheets("Applicant Hall Preference").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Applicant Hall Preference").Sort.SortFields.Add2 Key:=Range("K:K") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'Ensure that pattern sort is within eligible/waitlisted subgroups
    ActiveWorkbook.Worksheets("Applicant Hall Preference").Sort.SortFields.Add2 Key:=Range("J:J") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'Pattern sort
    With ActiveWorkbook.Worksheets("Applicant Hall Preference").Sort
        .SetRange Range("A:L") 'Select range of cells to be sorted
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Sort halls by alphabetical (numerical) order
    Range("B:H").Select
    ActiveWorkbook.Worksheets("Applicant Hall Preference").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Applicant Hall Preference").Sort.SortFields.Add2 Key:=Range("B1:H1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Applicant Hall Preference").Sort
        .SetRange Range("B:H")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlLeftToRight
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Declaration of variables
    Dim current_hall As Integer 'track the current hall
    Dim current_student As Integer 'track the current student
    Dim hall_assigned As Integer 'track if hall has been allocated to current student
    
    'Command to output names
    For current_student = 1 To number_of_students
        hall_assigned = 0
    
    'Actual loop to assign halls
    For current_hall = 1 To number_of_halls
        If Cells(current_student + 1, current_hall + 1).Value = "Preferred Choice" And hall_vacancy(current_hall) > 0 And hall_assigned < 1 Then
            hall_vacancy(current_hall) = hall_vacancy(current_hall) - 1 'Remove one vacancy from current hall
            hall_assigned = hall_assigned + 1 'Add one to hall counter for current student
            Cells(current_student + 1, output_column).Value = current_hall
        End If
    
    'Select next hall
    Next current_hall
    
    'Select next student
    Next current_student
    
    Range("A1").Select
End Sub
