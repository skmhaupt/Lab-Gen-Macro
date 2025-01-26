Attribute VB_Name = "Setup_Lab_Workbook_Module"
Option Explicit

Sub Make_Labs_Workbook()
    Dim var As String
    
    Dim sh As Worksheet
    Set sh = ActiveWorkbook.Sheets(1)
    Dim LabSetting As New LabSettings
    Dim LanguageSetting As New LanguageSettings
    Dim StartingData As Collection, NewData As Collection
    
    sh.Unprotect
    
    'Get number of students
    'And if workbook is old get old data
    Set StartingData = Get_StartingData(sh)
    If StartingData Is Nothing Then
        Exit Sub
    End If
    Set LabSetting = StartingData.Item("labSetting")
    Set LanguageSetting = StartingData.Item("language")
    
    'Get new data from user
    Set NewData = Get_UserData(sh, LabSetting, LanguageSetting)
    If NewData Is Nothing Then
        Exit Sub
    End If
    Set LabSetting = NewData.Item("labSetting")
    Set LanguageSetting = NewData.Item("language")
    
    'Setup workbook
    Call Setup_Main_Sheet(sh, LabSetting, LanguageSetting)
    
    Call Make_Group_Tables(sh, LabSetting, LanguageSetting)
    
    sh.Activate
    ActiveWindow.ScrollRow = 20
    MsgBox "Completed macro!"
    ActiveWindow.ScrollRow = 2
    sh.Cells(2, 1).Select
End Sub

