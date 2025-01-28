Attribute VB_Name = "Setup_Schedule_Sheet_Module"
Option Explicit


Public Function Make_ScheduleTabel(sh As Worksheet, LabSetting As LabSettings, LanguageSetting As LanguageSettings, Abort As Boolean)
    Dim i As Integer, j As Integer, sh3 As Worksheet
    Dim RowIndex As Integer, ColumnIndex As Integer, NumOfRowsPerTabel As Integer
    Dim NameRowColIndex As Variant, StudentName As String, StudentGroup As Integer
    Dim dict As Scripting.Dictionary, Day As String
    
    Set dict = New Scripting.Dictionary
    dict.Add "PON", RGB(255, 95, 31)
    dict.Add "UTO", RGB(31, 255, 15)
    dict.Add "SRI", RGB(148, 10, 206)
    dict.Add ChrW$(&H10C) & "ET", RGB(207, 255, 4)
    dict.Add "PET", RGB(63, 63, 255)
    dict.Add "", RGB(200, 200, 200)
    
    If Abort Then
        Exit Function
    End If
    
    If ActiveWorkbook.Sheets.Count = 4 Then
        ActiveWorkbook.Sheets(4).Name = ActiveWorkbook.Sheets(4).Name & "-Backup"
    End If
        Set sh3 = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
    
    sh3.Name = LanguageSetting.ScheduleLabel
    sh3.Cells.RowHeight = 18
    
    With sh3.Cells(3, 5)
        .Value = LabSetting.SubjectName
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    RowIndex = 5
    ColumnIndex = 1
    NumOfRowsPerTabel = 3 + LabSetting.NumberOfStudentsPerGroup + 4
    ReDim NameRowColIndex(LabSetting.NumberOfGroups)
    
    NameRowColIndex(0) = Array(5, 12)
    
    For i = 1 To LabSetting.NumberOfGroups
        sh3.Cells(RowIndex, ColumnIndex + 1).Value = "G" & i
        sh3.Cells(RowIndex + 1, ColumnIndex + 1).Value = "Lab " & LabSetting.get_GroupRoom(i)
        sh3.Cells(RowIndex + 2, ColumnIndex + 1).Value = LabSetting.get_GroupDate(i)
        Day = Left(LabSetting.get_GroupDate(i), 4)
        Day = Trim(Day)
        
        NameRowColIndex(i) = Array(RowIndex + 5, ColumnIndex + 1)
        
        For j = 1 To LabSetting.NumberOfStudentsPerGroup
            sh3.Cells(RowIndex + 4 + j, ColumnIndex).Value = j
        Next j
        
        With sh3.Range(sh3.Cells(RowIndex, ColumnIndex), sh3.Cells(RowIndex + 2, ColumnIndex + 2))
                .Borders(xlInsideHorizontal).Weight = xlThick
                .Borders(xlBottom).Weight = xlThick
                .HorizontalAlignment = xlCenter
                .Interior.Color = dict(Day)
        End With
        
        sh3.Range(sh3.Cells(RowIndex + 5, ColumnIndex), sh3.Cells(RowIndex + NumOfRowsPerTabel - 1 - 2, ColumnIndex)) _
                .HorizontalAlignment = xlCenter
        
        sh3.Range(sh3.Cells(RowIndex, ColumnIndex), sh3.Cells(RowIndex + NumOfRowsPerTabel - 1, ColumnIndex + 2)) _
                .BorderAround Weight:=xlThick
        
        With sh3.Range(sh3.Cells(RowIndex + 5, ColumnIndex + 1), sh3.Cells(RowIndex + NumOfRowsPerTabel - 1 - 2, ColumnIndex + 1))
                .Borders.Weight = xlThick
                .HorizontalAlignment = xlLeft
        End With
                
        ColumnIndex = ColumnIndex + 3
        
        If ColumnIndex = 10 Then
            ColumnIndex = 1
            RowIndex = RowIndex + 7 + LabSetting.NumberOfStudentsPerGroup + 3
        End If
    Next i
    
    For i = 2 To LabSetting.NumberOfStudents + 1
        StudentName = sh.Cells(i, 1).Value
        StudentGroup = sh.Cells(i, LabSetting.GroupColumnIndex).Value
        sh3.Cells(NameRowColIndex(StudentGroup)(0), NameRowColIndex(StudentGroup)(1)).Value = StudentName
        NameRowColIndex(StudentGroup)(0) = NameRowColIndex(StudentGroup)(0) + 1
    Next i
    
    sh3.Columns(2).AutoFit
    sh3.Columns(5).AutoFit
    sh3.Columns(8).AutoFit
End Function
