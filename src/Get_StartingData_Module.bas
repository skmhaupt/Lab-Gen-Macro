Attribute VB_Name = "Get_StartingData_Module"
Option Explicit

'Private functions
'------------------------------------------------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------

'Get input data from merlin.
'Sets number of students, number of groups, date for every group, room for every group, max number of students per group
Private Function Get_DataFromMerlinTable(sh As Worksheet, LabSetting As LabSettings, NumberOfCol1Entries As Long) As LabSettings
    Dim GroupIndex As Integer, i As Integer, NumOfStudentsPerGroup As Integer
    Dim GroupDate As String, GroupLabRoom As String, temp As String, temp2 As String
    
    LabSetting.NumberOfStudents = NumberOfCol1Entries - 1
    
    For i = 2 To NumberOfCol1Entries
        
        temp = sh.Cells(i, 6).Value
        'Check if student selected a group
        If temp Like "*-*" Then     'Student selected a group
            temp2 = Split(temp, "-", 2)(0)
            GroupIndex = CInt(WorksheetFunction.Substitute(temp2, "G", ""))
            sh.Cells(i, 5).Value = GroupIndex   'Write group index i  group column
            
            'Check if new group
            If Not LabSetting.GroupExists(GroupIndex) Then
                temp = Split(temp, "-", 2)(1)
                GroupDate = Split(temp, "(", 2)(0)
                GroupLabRoom = Split(temp, "(", 2)(1)
                GroupLabRoom = WorksheetFunction.Substitute(GroupLabRoom, ")", "")
                Call LabSetting.add_NewGroupData(GroupIndex, GroupDate, GroupLabRoom)   'set group date and room for new group
            End If
            
        Else    'Student didnt selec a group so set to 0
            sh.Cells(i, 5).Value = 0
        End If
    Next i
    
    'Get max number of students per group
    For GroupIndex = 1 To LabSetting.NumberOfGroups
        i = WorksheetFunction.CountIf(sh.Range(sh.Cells(2, 5), sh.Cells(NumberOfCol1Entries, 5)), GroupIndex)
        If i > LabSetting.NumberOfStudentsPerGroup Then
            LabSetting.NumberOfStudentsPerGroup = i
        End If
    Next GroupIndex
    
    LabSetting.NumberOfGroup0Students = WorksheetFunction.CountIf(sh.Range(sh.Cells(2, 5), sh.Cells(NumberOfCol1Entries, 5)), 0)
    
    Set Get_DataFromMerlinTable = LabSetting
End Function

'Function to get data from workbook that was already setup using this macro.
'It gets the language used, the number of excercises, group column index,
'whether the first lab is evaled, whether lab0 is used
Private Function Get_OldData(sh As Worksheet, LabSetting As LabSettings, LanguageSetting As LanguageSettings, NumberOfCol1Entries As Long) As Collection
    Dim TotalNumberOfExercises As Long, GroupIndex As Integer, i As Integer, NumOfGroups As Integer
    Dim ErrorIndex As Integer, ErrorMsg As String, Error1 As String, Error2 As String, Error3 As String
    Dim FormulaString As String, FirstLabLabel As String
    Dim GroupDate As String, GroupLabRoom As String
    Dim CustomLabels As Variant
    Dim ReturnData As New Collection
    
    ErrorMsg = ""
    Error1 = "Couldn't get language from old workbook" 'Index 1
    Error2 = "Couldn't get formula string" 'Index 2
    Error3 = "Couldn't get exercises data" 'Index 4
    
    'If old workbook was made using this macro it will have a hidden column
    sh.Cells(1, sh.Range("A1").End(xlToRight).Column - 4).EntireColumn.Hidden = False
    
    'Get total number of lab exercises including lab0 if used and label of first lab
    TotalNumberOfExercises = sh.Range("A1").End(xlToRight).Column - 1
    FirstLabLabel = sh.Cells(1, 2).Value
    'sh.Range(sh.Cells(1, 1), sh.Cells(NumberOfCol1Entries, TotalNumberOfExercises + 5)).Locked = False
    
    
    'Check what language was used
    If sh.Cells(1, 1).Value = "Prezime i Ime" Then    'Using cro workbook
        LanguageSetting.SetLanguage (1)
    ElseIf sh.Cells(1, 1).Value = "Full Name" Then     'Using eng workbook
        LanguageSetting.SetLanguage (2)
    ElseIf sh.Cells(1, 1).Value = "Nachname und Vorname" Then  'Using ger workbook
        LanguageSetting.SetLanguage (3)
    ElseIf sh.Cells(1, 1).Value = "Nom de famille et Nom" Then   'Using fr workbook
        LanguageSetting.SetLanguage (4)
    Else                                                'Sheet 1 name not recognized set language to cro
        LanguageSetting.SetLanguage (1)
        ErrorIndex = ErrorIndex + 1 'Set Error1
    End If
    
    'Get old subject name
    LabSetting.SubjectName = Split(sh.Name, "-")(0)
    
    'Get formula string
    If sh.Cells(2, TotalNumberOfExercises + 2).Formula Like "*(*" Then  'Make sure split will not error
        FormulaString = Split(sh.Cells(2, TotalNumberOfExercises + 2).Formula, "(")(0)
    Else    'Formula string not found
        FormulaString = ""
        ErrorIndex = ErrorIndex + 2     'Set Error2
    End If
        
    'Check if first lab is LAB0
    If FirstLabLabel = "LAB0" Then
        LabSetting.Lab0 = True
        LabSetting.NoEvalFirstLab = True
        LabSetting.NumberOfLabExercises = TotalNumberOfExercises - 1
        LabSetting.set_GroupColumnIndex
    'Check if first lab is not evaled
    ElseIf FormulaString = "=AND" Then
        LabSetting.Lab0 = False
        LabSetting.NoEvalFirstLab = True
        LabSetting.NumberOfLabExercises = TotalNumberOfExercises
        LabSetting.set_GroupColumnIndex
    'Check if first lab is evaled
    ElseIf FormulaString = "=COUNTIF" Then
        LabSetting.Lab0 = False
        LabSetting.NoEvalFirstLab = False
        LabSetting.NumberOfLabExercises = TotalNumberOfExercises
        LabSetting.set_GroupColumnIndex
    'Error! Workbook formula for ODRADENO not found
    Else
        ErrorIndex = ErrorIndex + 4     'Set Error3
    End If
        
    'Set number of students
    LabSetting.NumberOfStudents = NumberOfCol1Entries - 1
    
    
    'Check for custom labels
    If (Not ((FirstLabLabel = "LAB0") Or (FirstLabLabel = "LAB1")) And (Not (LabSetting.NumberOfLabExercises = 0))) Then
        ReDim CustomLabels(LabSetting.NumberOfLabExercises - 1)
        LabSetting.UsingCustomExcerciseLabels = True
        CustomLabels(0) = FirstLabLabel
        For i = 1 To LabSetting.NumberOfLabExercises - 1
            CustomLabels(i) = sh.Cells(1, i + 2).Value
        Next i
        LabSetting.CustomExcerciseLabels = CustomLabels
    End If
    
    
    'Check if small group table exists
    If Not IsEmpty(sh.Cells(3, LabSetting.GroupColumnIndex + 4)) Then   'Small table found, now getting number of groups, date and room data
        NumOfGroups = sh.Cells(2, LabSetting.GroupColumnIndex + 4).End(xlDown).Row - 2
        For i = 3 To NumOfGroups + 2
            GroupDate = sh.Cells(i, LabSetting.GroupColumnIndex + 5).Value
            GroupLabRoom = sh.Cells(i, LabSetting.GroupColumnIndex + 6).Value
            Call LabSetting.add_NewGroupData(i - 2, GroupDate, GroupLabRoom)
        Next i

    Else    'Small table not found, now getting number of groups
        MsgBox LabSetting.GroupColumnIndex & "Didnt finde small table"
        LabSetting.NumberOfGroups = WorksheetFunction.Max(sh.Range(sh.Cells(2, LabSetting.GroupColumnIndex), sh.Cells(NumberOfCol1Entries, LabSetting.GroupColumnIndex)))
    End If
    
    'Get max number of students per group
    For GroupIndex = 1 To LabSetting.NumberOfGroups
        i = WorksheetFunction.CountIf(sh.Range(sh.Cells(2, LabSetting.GroupColumnIndex), sh.Cells(NumberOfCol1Entries, LabSetting.GroupColumnIndex)), GroupIndex)
        If i > LabSetting.NumberOfStudentsPerGroup Then
            LabSetting.NumberOfStudentsPerGroup = i
        End If
    Next GroupIndex
    
    LabSetting.NumberOfGroup0Students = WorksheetFunction.CountIf(sh.Range(sh.Cells(2, LabSetting.GroupColumnIndex), sh.Cells(NumberOfCol1Entries, LabSetting.GroupColumnIndex)), 0)
    
    'Setup error msg
    While ErrorIndex > 0
        If ErrorIndex >= 4 Then
            ErrorMsg = Error3
            ErrorIndex = ErrorIndex - 4
        ElseIf ErrorIndex >= 2 Then
            ErrorMsg = ErrorMsg & vbCrLf & Error2
            ErrorIndex = ErrorIndex - 2
        Else
            ErrorMsg = ErrorMsg & vbCrLf & Error1
            ErrorIndex = ErrorIndex - 1
        End If
    Wend
    
    'Setup return data
    ReturnData.Add Item:=LabSetting, Key:="labSetting"
    ReturnData.Add Item:=LanguageSetting, Key:="language"
    ReturnData.Add Item:=ErrorMsg, Key:="error"
    
    Set Get_OldData = ReturnData
    
End Function

'Public functions
'------------------------------------------------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------

'Function to get number of students, number of groups, max number of students per group and old data if it exists
Public Function Get_StartingData(sh As Worksheet) As Collection
    Dim NumberOfCol1Entries As Long
    Dim LabSetting As New LabSettings
    Dim LanguageSetting As New LanguageSettings
    Dim ReturnData As New Collection, OldData As Collection
    
    'Get the number of students
    'Students have to be in column 1 starting with row 1
    'Group data on every student has to be in column 2
    'if A1 is empty macro will stop
    NumberOfCol1Entries = WorksheetFunction.CountA(sh.Range("A1", sh.Cells(1, 1).End(xlDown)))
    If IsEmpty(sh.Cells(1, 1)) Then
        MsgBox "Missing Students!" & vbCrLf & "Make sure you are using data from merlin OR that all Students are in column A starting with row 1 and the group data is in column B"
        Exit Function
    End If
    
    'Check if using new or used workboook
    If (sh.Range("A1").End(xlToRight).Column = 6 Or sh.Range("A1").End(xlToRight).Column = 2) Then
        'Using new workbook
        LanguageSetting.SetLanguage (1)
        LabSetting.UsingNewWorkbook = True
        
        'Check if data is from merlin
        If sh.Range("A1").End(xlToRight).Column = 6 Then
            LabSetting.UsingNewMerlinWorkbook = True
            Set LabSetting = Get_DataFromMerlinTable(sh, LabSetting, NumberOfCol1Entries)
            
        'If data is not from merlin check if group data exists
        ElseIf Not (sh.Cells(1, 2).End(xlDown).Row = NumberOfCol1Entries) Then
            MsgBox "Missing group data!" & vbCrLf & "Make sure you are using data from merlin OR that all Students are in column A starting with row 1 and the group data is in column B"
            Exit Function
        
        'Using new workbook that isnt from merlin and found all data (students and groups)
        'Now setting number of students, number of groups and max number of students per group
        Else
            LabSetting.NumberOfStudents = NumberOfCol1Entries
            LabSetting.NumberOfGroups = WorksheetFunction.Max(sh.Range(sh.Cells(1, 2), sh.Cells(NumberOfCol1Entries, 2)))
            Dim GroupIndex As Integer, i As Integer
            For GroupIndex = 1 To LabSetting.NumberOfGroups
                i = WorksheetFunction.CountIf(sh.Range(sh.Cells(1, 2), sh.Cells(NumberOfCol1Entries, 2)), GroupIndex)
                If i > LabSetting.NumberOfStudentsPerGroup Then
                    LabSetting.NumberOfStudentsPerGroup = i
                End If
            Next GroupIndex
            
            LabSetting.NumberOfGroup0Students = WorksheetFunction.CountIf(sh.Range(sh.Cells(1, 2), sh.Cells(NumberOfCol1Entries, 2)), 0)
        End If
      
    'Using old workbook
    Else
        'Get old data
        Set OldData = Get_OldData(sh, LabSetting, LanguageSetting, NumberOfCol1Entries)
        Set LabSetting = OldData.Item("labSetting")
        Set LanguageSetting = OldData.Item("language")
        If Not OldData.Item("error") = "" Then
            MsgBox OldData.Item("error")
        End If
    End If
    
    
    'Setup return data
    ReturnData.Add Item:=LabSetting, Key:="labSetting"
    ReturnData.Add Item:=LanguageSetting, Key:="language"
    
    Set Get_StartingData = ReturnData
End Function
