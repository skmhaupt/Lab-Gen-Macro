Attribute VB_Name = "Setup_Main_Sheet_Module"
Option Explicit

'Private functions
'------------------------------------------------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------

'Fill small table with group room and date data.
'First cell is in row 2 and column GroupColumnIndex + 2
Private Function Make_SmallGroupDataTable(sh As Worksheet, LabSetting As LabSettings)
    Dim i As Long
    For i = 3 To LabSetting.NumberOfGroups + 2
        sh.Cells(i, LabSetting.GroupColumnIndex + 4).Value = i - 2
        sh.Cells(i, LabSetting.GroupColumnIndex + 5).Value = LabSetting.get_GroupDate(i - 2)
        sh.Cells(i, LabSetting.GroupColumnIndex + 6).Value = LabSetting.get_GroupRoom(i - 2)
    Next i
End Function


'Set text for all labels
Private Function Set_Labels(sh As Worksheet, LabSetting As LabSettings, LanguageSetting As LanguageSettings)
    Dim currentCell_y As Long, LAB0_offset As Long, i As Long
    sh.Range("A1").EntireRow.Insert
    sh.Range("A1").Value = LanguageSetting.NameLabel
    
    'Move group data to GRUPA column
    sh.Range("B2", sh.Cells(LabSetting.NumberOfStudents + 1, 2)).Cut sh.Range(sh.Cells(2, LabSetting.GroupColumnIndex), sh.Cells(LabSetting.NumberOfStudents + 1, LabSetting.GroupColumnIndex))
    
    'Check if first lab label is 0 or 1
    If LabSetting.Lab0 Then
        sh.Range("B1").Value = "LAB0"
        LAB0_offset = 1
    Else
        LAB0_offset = 0
    End If
    
    'Set exercises labels
    For i = 1 To LabSetting.NumberOfLabExercises
        currentCell_y = 1 + i + LAB0_offset
        If Not LabSetting.UsingCustomExcerciseLabels Then
            sh.Cells(1, currentCell_y).Value = "LAB" & i
        Else
            sh.Cells(1, currentCell_y).Value = LabSetting.CustomExcerciseLabels(i - 1)
        End If
    Next i
    
    'Set text for done, average and group cells
    sh.Cells(1, LabSetting.GroupColumnIndex - 2).Value = LanguageSetting.DoneLabel
    sh.Cells(1, LabSetting.GroupColumnIndex - 1).Value = LanguageSetting.AverageLabel
    sh.Cells(1, LabSetting.GroupColumnIndex).Value = LanguageSetting.GroupLabel
    sh.Cells(1, LabSetting.GroupColumnIndex + 1).Value = LanguageSetting.AlreadyDoneLabel
    
    'Set text for small tabel
    'First cell is in row 2 and column GroupColumnIndex + 2
    sh.Cells(2, LabSetting.GroupColumnIndex + 4).Value = LanguageSetting.GroupLabel
    sh.Cells(2, LabSetting.GroupColumnIndex + 5).Value = LanguageSetting.ScheduleLabel
    sh.Cells(2, LabSetting.GroupColumnIndex + 6).Value = LanguageSetting.RoomLabel
End Function


'Formating for main sheet. Main and small table
Private Function Set_TabelFormating(sh As Worksheet, LabSetting As LabSettings, LanguageSetting As LanguageSettings)
    'Set row heights and column widths
    sh.Cells.RowHeight = 18
    sh.Cells.ColumnWidth = 8.11
    sh.Rows(1).RowHeight = sh.Rows(1).RowHeight * 2
    
    'Wrap already done label
    sh.Cells(1, LabSetting.GroupColumnIndex + 1).WrapText = True
    
    'Center all except for col 1
    sh.Range(sh.Cells(1, 2), sh.Cells(LabSetting.NumberOfStudents + 1, LabSetting.GroupColumnIndex + 1)).HorizontalAlignment = xlCenter
    'Autofit
    sh.Columns(1).AutoFit
    sh.Range(sh.Columns(LabSetting.GroupColumnIndex - 3), sh.Columns(LabSetting.GroupColumnIndex + 1)).AutoFit
    sh.Columns(LabSetting.GroupColumnIndex + 5).AutoFit
    
    'Set label row color
    sh.Range("A1", sh.Cells(1, LabSetting.GroupColumnIndex + 1)).Interior.ColorIndex = 15
    sh.Range(sh.Cells(2, LabSetting.GroupColumnIndex + 4), sh.Cells(2, LabSetting.GroupColumnIndex + 6)).Interior.ColorIndex = 15
    
    'Set border:
        'Border around table
    With sh.Range("A1", sh.Cells(LabSetting.NumberOfStudents + 1, LabSetting.GroupColumnIndex - 1))
        .BorderAround _
            Weight:=xlThick
    End With
        'Border around labels
    sh.Range("A1", sh.Cells(1, LabSetting.GroupColumnIndex - 1)).BorderAround _
        Weight:=xlThick
        'Border next to names and align names left
    With sh.Range("A1", sh.Cells(LabSetting.NumberOfStudents + 1, 1))
        .Borders(xlEdgeRight).Weight = xlThick
        .HorizontalAlignment = xlLeft
    End With
        'Border on small table and center text
    With sh.Range(sh.Cells(2, LabSetting.GroupColumnIndex + 4), sh.Cells(LabSetting.NumberOfGroups + 2, LabSetting.GroupColumnIndex + 6))
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .BorderAround Weight:=xlThick
    End With
    sh.Range(sh.Cells(2, LabSetting.GroupColumnIndex + 4), sh.Cells(2, LabSetting.GroupColumnIndex + 6)).BorderAround Weight:=xlThick
    
    'Set percentage formating for average column (LabSetting.GroupColumnIndex - 1)
    sh.Range(sh.Cells(2, LabSetting.GroupColumnIndex - 1), sh.Cells(LabSetting.NumberOfStudents + 1, LabSetting.GroupColumnIndex - 1)).NumberFormat = "0.00%"
End Function


'Conditional formating for cell shading
Private Function Set_ConditionalFormating(sh As Worksheet, LabSetting As LabSettings, LanguageSetting As LanguageSettings)
    Dim NoEvalFirstLab_offset As Long
    NoEvalFirstLab_offset = 0
    If LabSetting.NoEvalFirstLab Then
        NoEvalFirstLab_offset = 1
        With sh.Range("B2", sh.Cells(LabSetting.NumberOfStudents + 1, 2))
            .FormatConditions.Delete
            'Add first rule: shade green if x = 1
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                Formula1:="=1"
            .FormatConditions(1).Interior.Color = RGB(166, 240, 80)
            'Add second rule: shade red if x > 1
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                Formula1:="=1"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0)
        End With
    End If
    With sh.Range(sh.Cells(2, 2 + NoEvalFirstLab_offset), sh.Cells(LabSetting.NumberOfStudents + 1, LabSetting.GroupColumnIndex - 4))
        .FormatConditions.Delete
        'Add first rule: shade green if 5 <= x <= 10
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
            Formula1:="=5", Formula2:="=10"
        .FormatConditions(1).Interior.Color = RGB(146, 208, 80)
        'Add second rule: shade red if x > 10
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=10"
        .FormatConditions(2).Interior.Color = RGB(255, 0, 0)
    End With
    With sh.Range(sh.Cells(2, LabSetting.GroupColumnIndex - 2), sh.Cells(LabSetting.NumberOfStudents + 1, LabSetting.GroupColumnIndex - 2))
        .FormatConditions.Delete
        'Add first rule: shade green if x = yes
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
            Formula1:=LanguageSetting.YesLabel
        .FormatConditions(1).Interior.Color = RGB(146, 250, 80)
        'Add second rule: shade red if x = no
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
            Formula1:=LanguageSetting.NoLabel
        .FormatConditions(2).Interior.Color = RGB(200, 80, 80)
    End With
End Function


'Set formulas for columns "odradio-1", odradio and prosjek depending on lab settings
Private Function Set_Formulas(sh As Worksheet, LabSetting As LabSettings, LanguageSetting As LanguageSettings)
    Dim i As Long
    If LabSetting.Lab0 Then
        For i = 2 To LabSetting.NumberOfStudents + 1
            'x=NumberOfLabExercises
            '=AND(COUNTIF(range{3:x+2},">=5")=x,B2=1)
            sh.Cells(i, LabSetting.GroupColumnIndex - 3).Formula = "=AND(COUNTIF(" & sh.Range(sh.Cells(i, 3), sh.Cells(i, LabSetting.NumberOfLabExercises + 2)).Address & "," & Chr(34) & ">=5" & Chr(34) & ")=" & LabSetting.NumberOfLabExercises & ",B" & i & "=1)"
            '=IF(odradeno=TRUE,"DA","NE")
            sh.Cells(i, LabSetting.GroupColumnIndex - 2).Formula = "=IF(" & sh.Cells(i, LabSetting.GroupColumnIndex - 3).Address & "=TRUE," & Chr(34) & LanguageSetting.YesLabel & Chr(34) & "," & Chr(34) & LanguageSetting.NoLabel & Chr(34) & ")"
            '=IF(odradeno="DA",SUM(range{3:x+1})/(10*x),"0.00%")
            sh.Cells(i, LabSetting.GroupColumnIndex - 1).Formula = "=IF(" & sh.Cells(i, LabSetting.GroupColumnIndex - 2).Address & "=" & Chr(34) & LanguageSetting.YesLabel & Chr(34) & ",SUM(" & sh.Range(sh.Cells(i, 3), sh.Cells(i, LabSetting.NumberOfLabExercises + 2)).Address & ")/(10*" & LabSetting.NumberOfLabExercises & "), 0%)"
        Next i
    ElseIf LabSetting.NoEvalFirstLab Then
        For i = 2 To LabSetting.NumberOfStudents + 1
            'x=NumberOfLabExercises
            '=AND(COUNTIF(range{3:x+1},">=5")=x-1,B2=1)
            sh.Cells(i, LabSetting.GroupColumnIndex - 3).Formula = "=AND(COUNTIF(" & sh.Range(sh.Cells(i, 3), sh.Cells(i, LabSetting.NumberOfLabExercises + 1)).Address & "," & Chr(34) & ">=5" & Chr(34) & ")=" & LabSetting.NumberOfLabExercises - 1 & ",B" & i & "=1)"
            '=IF(odradeno=TRUE,"DA","NE")
            sh.Cells(i, LabSetting.GroupColumnIndex - 2).Formula = "=IF(" & sh.Cells(i, LabSetting.GroupColumnIndex - 3).Address & "=TRUE," & Chr(34) & LanguageSetting.YesLabel & Chr(34) & "," & Chr(34) & LanguageSetting.NoLabel & Chr(34) & ")"
            '=IF(odradeno="DA",SUM(range{3:x+2})/(10*x-1),"0.00%")
            sh.Cells(i, LabSetting.GroupColumnIndex - 1).Formula = "=IF(" & sh.Cells(i, LabSetting.GroupColumnIndex - 2).Address & "=" & Chr(34) & LanguageSetting.YesLabel & Chr(34) & ",SUM(" & sh.Range(sh.Cells(i, 3), sh.Cells(i, LabSetting.NumberOfLabExercises + 1)).Address & ")/(10*" & LabSetting.NumberOfLabExercises - 1 & "), 0%)"
        Next i
    Else
        For i = 2 To LabSetting.NumberOfStudents + 1
            'x=NumberOfLabExercises
            '=COUNTIF(range{2:x+1},">=5")=x
            sh.Cells(i, LabSetting.GroupColumnIndex - 3).Formula = "=COUNTIF(" & sh.Range(sh.Cells(i, 2), sh.Cells(i, LabSetting.NumberOfLabExercises + 1)).Address & "," & Chr(34) & ">=5" & Chr(34) & ")=" & LabSetting.NumberOfLabExercises
            '=IF(odradeno=TRUE,"DA","NE")
            sh.Cells(i, LabSetting.GroupColumnIndex - 2).Formula = "=IF(" & sh.Cells(i, LabSetting.GroupColumnIndex - 3).Address & "=TRUE," & Chr(34) & LanguageSetting.YesLabel & Chr(34) & "," & Chr(34) & LanguageSetting.NoLabel & Chr(34) & ")"
            '=IF(odradeno="DA",SUM(range{2:x+1})/(10*x),"0.00%")
            sh.Cells(i, LabSetting.GroupColumnIndex - 1).Formula = "=IF(" & sh.Cells(i, LabSetting.GroupColumnIndex - 2).Address & "=" & Chr(34) & LanguageSetting.YesLabel & Chr(34) & ",SUM(" & sh.Range(sh.Cells(i, 2), sh.Cells(i, LabSetting.NumberOfLabExercises + 1)).Address & ")/(10*" & LabSetting.NumberOfLabExercises & "), 0%)"
        Next i
    End If
End Function

Private Function FreezLabels(sh As Worksheet)
    sh.Activate
    With ActiveWindow
    If .FreezePanes Then .FreezePanes = False
    .SplitColumn = 1
    .SplitRow = 1
    .FreezePanes = True
End With
End Function

'Button to get failed students
Private Function Set_Button(sh As Worksheet, Column As Long)
    Dim btn As Button
    Dim t As Range
    Set t = sh.Range(sh.Cells(2, Column), sh.Cells(2, Column))
    sh.Columns(Column).ColumnWidth = sh.Columns(Column).ColumnWidth * 2.5
    t.Interior.Color = RGB(200, 80, 80)
    Set btn = sh.Buttons.Add(t.Left + 1.5, t.Top + 1.5, t.Width - 2, t.Height - 2)
    With btn
      .OnAction = "Get_FailedStudents"
      .Caption = "Get Failed Students "
      .Name = "Get_FailedStudents_Btn"
    End With
End Function

Private Sub Get_FailedStudents()
    Dim sh As Worksheet
    Dim AverageColumn As Long, LastStudentRow As Long, i As Integer, j As Integer
    Set sh = ActiveSheet
    LastStudentRow = sh.Cells(1, 1).End(xlDown).Row
    AverageColumn = sh.Cells(1, 1).End(xlToRight).Column - 2
    sh.Range(sh.Cells(3, AverageColumn + 10), sh.Cells(3, AverageColumn + 10).End(xlDown)).Value = ""
    j = 3
    For i = 2 To LastStudentRow
        If sh.Cells(i, AverageColumn).Value < 0.5 And IsEmpty(sh.Cells(i, AverageColumn + 2)) Then
            sh.Cells(j, AverageColumn + 10).Value = sh.Cells(i, 1).Value
            j = j + 1
        End If
    Next i
End Sub


'Public functions
'------------------------------------------------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------

'Calls all private functions to set main sheet
Public Function Setup_Main_Sheet(sh As Worksheet, LabSetting As LabSettings, LanguageSetting As LanguageSettings)
    Call Set_Labels(sh, LabSetting, LanguageSetting)
    Call Make_SmallGroupDataTable(sh, LabSetting)
    Call Set_TabelFormating(sh, LabSetting, LanguageSetting)
    Call Set_ConditionalFormating(sh, LabSetting, LanguageSetting)
    Call Set_Formulas(sh, LabSetting, LanguageSetting)
    Call FreezLabels(sh)
    Call Set_Button(sh, LabSetting.GroupColumnIndex + 9)
End Function
