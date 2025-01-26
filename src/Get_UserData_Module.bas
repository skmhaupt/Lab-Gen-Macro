Attribute VB_Name = "Get_UserData_Module"
Option Explicit


'Public functions
'------------------------------------------------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------

Private Function Clear_Old_workbook(sh As Worksheet, NumberOfStudents, GroupColumnIndex)
    sh.Range(sh.Cells(2, GroupColumnIndex), sh.Cells(NumberOfStudents + 1, GroupColumnIndex)).Cut sh.Range("B2", sh.Cells(NumberOfStudents + 1, 2))
    sh.Range("A1").EntireRow.Delete
    sh.Columns(GroupColumnIndex + 9).Delete
    sh.Range("A1", sh.Cells(NumberOfStudents, GroupColumnIndex + 11)).ClearFormats
    sh.Range("C1", sh.Cells(NumberOfStudents, GroupColumnIndex + 11)).Clear
    sh.Range("A1", sh.Cells(NumberOfStudents, GroupColumnIndex + 11)).Locked = False
    sh.Range(sh.Cells(1, 1), sh.Cells(NumberOfStudents, 2)).Sort Key1:=Range(sh.Cells(1, 1), sh.Cells(NumberOfStudents, 1)), Order1:=xlAscending
End Function

Private Function Clear_MerlinData(sh As Worksheet, NumberOfStudents)
    sh.Columns(6).Delete Shift:=xlShiftToLeft
    sh.Columns(4).Delete Shift:=xlShiftToLeft
    sh.Columns(3).Delete Shift:=xlShiftToLeft
    sh.Columns(2).Delete Shift:=xlShiftToLeft
    sh.Rows(1).Delete Shift:=xlShiftUp
    sh.Range("A1", sh.Cells(NumberOfStudents, 2)).Sort Key1:=Range(sh.Cells(1, 1), sh.Cells(NumberOfStudents, 1)), Order1:=xlAscending
End Function


'Public functions
'------------------------------------------------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------

'Get user data with form
'If canceled stops macro.
'Stores data in collection of LabSettings and LanguageSetting
Public Function Get_UserData(sh As Worksheet, LabSetting As LabSettings, LanguageSetting As LanguageSettings) As Collection
    Dim ReturnData As New Collection
    Dim UserForm As LabDetailesUserForm
    Set UserForm = New LabDetailesUserForm
    
    'If using old workbook setup form with old data
    If Not LabSetting.UsingNewWorkbook Then
        UserForm.SetUserFormOldData LabSetting, LanguageSetting
    End If
    
    'Show form and get user data
    UserForm.Show
    If UserForm.Cancelled Then
        Exit Function
    Else
        'Clear old workbook
        If LabSetting.UsingNewMerlinWorkbook Then
            Call Clear_MerlinData(sh, LabSetting.NumberOfStudents)
        ElseIf Not LabSetting.UsingNewWorkbook Then
            Call Clear_Old_workbook(sh, LabSetting.NumberOfStudents, LabSetting.GroupColumnIndex)
        End If
        'Set new data
        LabSetting.set_LabSettings UserForm.SubjectName, UserForm.NumberOfLabExcercises, UserForm.Lab0, UserForm.NoEvalFirstLab, UserForm.CustomLabels, UserForm.UsingCustomLabels
        LanguageSetting.SetLanguage (UserForm.LanguageIndex)
        sh.Name = UserForm.SubjectName & LanguageSetting.Sheet1Name
        Unload UserForm
    End If
    
    'Setup return data
    ReturnData.Add Item:=LabSetting, Key:="labSetting"
    ReturnData.Add Item:=LanguageSetting, Key:="language"
    
    Set Get_UserData = ReturnData
End Function
