Attribute VB_Name = "Setup_GroupTables_Sheet_Module"
Option Explicit

'Private functions
'------------------------------------------------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------


Private Function Make_GroupTables_Sheet(sh1 As Worksheet, LanguageSetting As LanguageSettings) As Worksheet
    Dim sh2 As Worksheet, shOld As Worksheet
    'Check if old table sheet exists and make backup
    'Work in progress
    If ActiveWorkbook.Sheets.Count = 2 Then
        Set shOld = ActiveWorkbook.Sheets(2)
        shOld.Name = shOld.Name & "-Backup"
    ElseIf ActiveWorkbook.Sheets.Count > 2 Then
        MsgBox "Please delete old backup sheet and restart macro."
        Exit Function
    End If
           
    Set sh2 = ActiveWorkbook.Sheets.Add(, sh1)
    sh2.Name = LanguageSetting.Sheet2Name
    sh2.Cells.RowHeight = 18
    
    Set Make_GroupTables_Sheet = sh2
End Function


Private Function Set_GroupTables_Dict(LabSetting As LabSettings) As Scripting.Dictionary
    Dim y As Long, flip As Boolean, x As Integer, TotalNumberOfExercises As Long
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    ''Set dict = CreateObject("Scripting.Dictionary")
    
    TotalNumberOfExercises = LabSetting.Get_TotalNumberOfExercises
    
    'Group 0 is for all students that didnt select a group and is separated
    dict.Add 0, Array(3, (TotalNumberOfExercises + 2) * 2 + 4)
    
    y = 8
    flip = True
    
    'dict.Add 1, Array(3, 1)    add LabGroup, arr(GroupRowIndexWrite, GroupColIndexWrite)
    'dict.Add 2, Array(3, 12)
    'dict.Add 3, Array(18, 1)
    'dict.Add 4, Array(18, 12)
    'dict.Add 5, Array(33, 1)
    '...
    For x = 1 To LabSetting.NumberOfGroups + 1
        If flip Then
            flip = Not flip             'Creates two columns of group tables (row1 g1/2, row2 g3/4, ...)
            dict.Add x, Array(y, 1)
        Else
            flip = Not flip
            dict.Add x, Array(y, TotalNumberOfExercises + 5)    'TotalNumberOfExercises + 5 set the number of column between tables 5 -> 2 Columns
            y = y + LabSetting.NumberOfStudentsPerGroup + 7     'Sets number of rows between tables 7 -> 4 rows
        End If
    Next x
    
    Set Set_GroupTables_Dict = dict
End Function

Private Function Make_GroupTables(sh1 As Worksheet, sh2 As Worksheet, LabSetting As LabSettings, LanguageSetting As LanguageSettings, GroupTables_Dict As Scripting.Dictionary)
    Dim k As Integer
    Dim GroupRowIndexWrite As Long, GroupColIndexWrite As Long, TotalNumberOfExercises As Long
    Dim LastGroup0Row As Long
    TotalNumberOfExercises = LabSetting.Get_TotalNumberOfExercises
    
    For k = 1 To LabSetting.NumberOfGroups
        'Get position for table
        GroupRowIndexWrite = GroupTables_Dict(k)(0)
        GroupColIndexWrite = GroupTables_Dict(k)(1)
        
        'Set labels
        sh2.Cells(GroupRowIndexWrite - 2, GroupColIndexWrite).Value = "G" & k
        sh2.Cells(GroupRowIndexWrite - 2, GroupColIndexWrite + 1).Value = LabSetting.get_GroupDate(k)
        sh2.Cells(GroupRowIndexWrite - 2, GroupColIndexWrite + 2).Value = LabSetting.get_GroupRoom(k)
        sh2.Cells(GroupRowIndexWrite - 1, GroupColIndexWrite + 1).Value = LanguageSetting.NameLabel
        
        'Copy lab labels from main sheet
        sh1.Range(sh1.Cells(1, 2), sh1.Cells(1, TotalNumberOfExercises + 1)).Copy _
        Destination:=sh2.Range(sh2.Cells(GroupRowIndexWrite - 1, GroupColIndexWrite + 2), sh2.Cells(GroupRowIndexWrite - 1, GroupColIndexWrite + 1 + TotalNumberOfExercises))
        
        'Set row indexeing for students
        sh2.Cells(GroupRowIndexWrite, GroupColIndexWrite).Value = 1
        sh2.Cells(GroupRowIndexWrite + 1, GroupColIndexWrite).Value = 2
        sh2.Range(sh2.Cells(GroupRowIndexWrite, GroupColIndexWrite), sh2.Cells(GroupRowIndexWrite + 1, GroupColIndexWrite)).AutoFill Destination:=sh2.Range(sh2.Cells(GroupRowIndexWrite, GroupColIndexWrite), sh2.Cells(GroupRowIndexWrite + LabSetting.NumberOfStudentsPerGroup, GroupColIndexWrite))
        
        'Format table
        sh2.Range(sh2.Cells(GroupRowIndexWrite - 1, GroupColIndexWrite), sh2.Cells(GroupRowIndexWrite + LabSetting.NumberOfStudentsPerGroup, GroupColIndexWrite + TotalNumberOfExercises + 1)).Borders _
        .LineStyle = xlContinuous   'Table with label row 2
        
        sh2.Range(sh2.Cells(GroupRowIndexWrite - 2, GroupColIndexWrite), sh2.Cells(GroupRowIndexWrite - 2, GroupColIndexWrite + 2)).Borders _
        .Weight = xlThick   'label row 1
        
        sh2.Range(sh2.Cells(GroupRowIndexWrite - 2, GroupColIndexWrite), sh2.Cells(GroupRowIndexWrite - 2, GroupColIndexWrite + 1)).EntireColumn _
        .HorizontalAlignment = xlLeft   'Column 1 and 2
        
        'sh2.Range(sh2.Cells(GroupRowIndexWrite - 2, GroupColIndexWrite), sh2.Cells(GroupRowIndexWrite + NumOfStudentsPerGroup, GroupColIndexWrite)).BorderAround _
        'Weight:=xlThick     'Column 1
        
        sh2.Range(sh2.Cells(GroupRowIndexWrite - 1, GroupColIndexWrite), sh2.Cells(GroupRowIndexWrite - 1, GroupColIndexWrite + TotalNumberOfExercises + 1)).BorderAround _
        Weight:=xlThick     'Label row 2
        
        sh2.Range(sh2.Cells(GroupRowIndexWrite, GroupColIndexWrite), sh2.Cells(GroupRowIndexWrite + LabSetting.NumberOfStudentsPerGroup, GroupColIndexWrite + TotalNumberOfExercises + 1)).BorderAround _
        Weight:=xlThick     'Table without labels
        
        sh2.Range(sh2.Cells(GroupRowIndexWrite, GroupColIndexWrite + 2), sh2.Cells(GroupRowIndexWrite + LabSetting.NumberOfStudentsPerGroup, GroupColIndexWrite + TotalNumberOfExercises + 1)) _
        .HorizontalAlignment = xlCenter      'Table without labels and names
    Next k
    
    'Group 0 table:
    GroupRowIndexWrite = GroupTables_Dict(0)(0)
    GroupColIndexWrite = GroupTables_Dict(0)(1)
    LastGroup0Row = LabSetting.NumberOfGroup0Students + 2
    sh1.Range(sh1.Cells(1, 2), sh1.Cells(1, TotalNumberOfExercises + 1)).Copy _
        Destination:=sh2.Range(sh2.Cells(GroupRowIndexWrite - 1, GroupColIndexWrite + 2), sh2.Cells(GroupRowIndexWrite - 1, GroupColIndexWrite + 1 + TotalNumberOfExercises))
    With sh2.Range(sh2.Cells(GroupRowIndexWrite - 1, GroupColIndexWrite), sh2.Cells(LastGroup0Row, GroupColIndexWrite + TotalNumberOfExercises + 1))
        .Borders.Weight = xlThin
        .BorderAround Weight:=xlThick
    End With
    sh2.Range(sh2.Cells(GroupRowIndexWrite - 1, GroupColIndexWrite), sh2.Cells(GroupRowIndexWrite - 1, GroupColIndexWrite + 1 + TotalNumberOfExercises)) _
        .BorderAround Weight:=xlThick
    
    sh2.Range(sh2.Cells(GroupRowIndexWrite, GroupColIndexWrite), sh2.Cells(LastGroup0Row, GroupColIndexWrite)).HorizontalAlignment = xlLeft
    sh2.Range(sh2.Cells(GroupRowIndexWrite, GroupColIndexWrite + 2), sh2.Cells(LastGroup0Row, GroupColIndexWrite + 1 + TotalNumberOfExercises)).HorizontalAlignment = xlCenter
    
    For k = 3 To LastGroup0Row
        sh2.Cells(k, GroupColIndexWrite).Value = k - 2
    Next k
End Function

'Fill all group tables and link with main table on sh1
Private Function Fill_GroupTables(sh1 As Worksheet, sh2 As Worksheet, LabSetting As LabSettings, GroupTables_Dict As Scripting.Dictionary)
    Dim GroupRowIndexWrite As Long, GroupColIndexWrite As Long, TotalNumberOfExercises As Long
    Dim StudentName As String
    Dim LabGroup As Integer
    Dim i As Long, j As Long
    
    TotalNumberOfExercises = LabSetting.Get_TotalNumberOfExercises
    
    For i = 2 To LabSetting.NumberOfStudents + 1
        StudentName = sh1.Cells(i, "A").Value      'Get student name
        LabGroup = sh1.Cells(i, LabSetting.GroupColumnIndex).Value      'Get lab group for student
        GroupRowIndexWrite = GroupTables_Dict(LabGroup)(0)     'Get first free row index in group table
        GroupColIndexWrite = GroupTables_Dict(LabGroup)(1) + 1      'Get column for names in group table
        sh2.Cells(GroupRowIndexWrite, GroupColIndexWrite).Value = StudentName   'Paste student name in group table
        For j = 1 To TotalNumberOfExercises     'Link student row from main table to row in group table
            sh1.Cells(i, j + 1).Formula = "=INDEX(" & sh2.Name & "!" & Range(Cells(GroupRowIndexWrite, GroupColIndexWrite), Cells(GroupRowIndexWrite, GroupColIndexWrite + TotalNumberOfExercises)).Address & ",," & j + 1 & ")"
                                          '=INDEX(    TabliceGrupa!    RangeAddress                                                                                                                            ,,    column   )
        Next j
        GroupTables_Dict.Item(LabGroup) = Array(GroupRowIndexWrite + 1, GroupColIndexWrite - 1)  'Set new first free row index in group
    Next i
End Function


Private Function Format_Sheet2(sh1 As Worksheet, sh2 As Worksheet, LabSetting As LabSettings)
    Dim TotalNumberOfExercises As Long
    TotalNumberOfExercises = LabSetting.Get_TotalNumberOfExercises
    
    'Autofit name columns in sh2
    sh2.Columns(2).AutoFit
    sh2.Columns(TotalNumberOfExercises + 6).AutoFit
    sh2.Columns((TotalNumberOfExercises + 2) * 2 + 5).AutoFit
    'Fit first column in tables of sh2
    sh2.Columns(1).ColumnWidth = sh2.Columns(1).ColumnWidth * 0.6
    sh2.Columns(TotalNumberOfExercises + 5).ColumnWidth = sh2.Columns(TotalNumberOfExercises + 5).ColumnWidth * 0.6
    
    sh1.Cells(1, TotalNumberOfExercises + 2).EntireColumn.Hidden = True
    
    sh1.Cells.Locked = False
    sh1.Range(sh1.Cells(1, 1), sh1.Cells(LabSetting.NumberOfStudents + 1, TotalNumberOfExercises + 5)).Locked = True
    sh1.Protect
End Function



'Public functions
'------------------------------------------------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------

'
Public Function Make_Group_Tables(sh1 As Worksheet, LabSetting As LabSettings, LanguageSetting As LanguageSettings)
    Dim sh2 As Worksheet
    Dim TotalNumberOfExercises As Long
    Dim GroupTables_Dict As Scripting.Dictionary
    
    TotalNumberOfExercises = LabSetting.Get_TotalNumberOfExercises
    
    'Make sheet and backup old
    Set sh2 = Make_GroupTables_Sheet(sh1, LanguageSetting)
    If sh2 Is Nothing Then
        Exit Function
    End If

    'Set GroupTables_Dict
    Set GroupTables_Dict = Set_GroupTables_Dict(LabSetting)
    
    
    With sh2.Range(sh2.Cells(2, 5), sh2.Cells(2, 5))
        .Value = LabSetting.SubjectName
        .Font.Bold = True
    End With
    With sh2.Range(sh2.Cells(3, 4), sh2.Cells(3, 4))
        .Value = "LABORATORIJSKE VJEŽBE FESB"
        .Font.Bold = True
    End With
    
    sh2.Cells(2, ((TotalNumberOfExercises + 2) * 2 + 4)).Value = "G0"
    sh2.Cells(2, ((TotalNumberOfExercises + 2) * 2 + 4 + 1)).Value = LanguageSetting.NameLabel
    sh2.Cells(2, ((TotalNumberOfExercises + 2) * 2 + 4 + 1)).Interior.ColorIndex = 15
    
    
    Call Make_GroupTables(sh1, sh2, LabSetting, LanguageSetting, GroupTables_Dict)
    
    Call Fill_GroupTables(sh1, sh2, LabSetting, GroupTables_Dict)
    
    Call Format_Sheet2(sh1, sh2, LabSetting)
End Function









