VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SubjectName As String
Public NumberOfLabExercises As Long
Public NumberOfStudents As Long
Public Lab0 As Boolean
Public NoEvalFirstLab As Boolean
Public UsingNewWorkbook As Boolean
Public UsingNewMerlinWorkbook As Boolean
Public GroupColumnIndex As Long
Private dictGroupDateAndRoom As Object
Public NumberOfGroups As Long
Public NumberOfStudentsPerGroup As Long
Public NumberOfGroup0Students As Long
Public CustomExcerciseLabels As Variant
Public UsingCustomExcerciseLabels As Boolean


Private Sub Class_Initialize()
    SubjectName = ""
    NumberOfLabExercises = 0
    NumberOfStudents = 0
    Lab0 = False
    NoEvalFirstLab = False
    UsingNewWorkbook = False
    UsingNewMerlinWorkbook = False
    Set dictGroupDateAndRoom = CreateObject("Scripting.Dictionary")
    NumberOfGroups = 0
    NumberOfStudentsPerGroup = 0
    UsingCustomExcerciseLabels = False
End Sub

Public Sub set_LabSettings(arg_SubjectName As String, arg_NumberOfLabExercises As Long, arg_lab0 As Boolean, arg_noEvalFirstLab As Boolean, arg_CustomExcerciseLabels As Variant, arg_UsingCustomExcerciseLabels As Boolean)
    SubjectName = arg_SubjectName
    NumberOfLabExercises = arg_NumberOfLabExercises
    Lab0 = arg_lab0
    NoEvalFirstLab = arg_noEvalFirstLab
    If Lab0 Then
        GroupColumnIndex = 2 + NumberOfLabExercises + 4
    Else
        GroupColumnIndex = 1 + NumberOfLabExercises + 4
    End If
    CustomExcerciseLabels = arg_CustomExcerciseLabels
    UsingCustomExcerciseLabels = arg_UsingCustomExcerciseLabels
End Sub

Public Sub set_GroupColumnIndex()
    If Lab0 Then
        GroupColumnIndex = 2 + NumberOfLabExercises + 4
    Else
        GroupColumnIndex = 1 + NumberOfLabExercises + 4
    End If
End Sub

Public Function add_NewGroupData(arg_groupIndex As Integer, arg_date As String, arg_room As String)
    If dictGroupDateAndRoom.Exists(arg_groupIndex) Then
        MsgBox "Group already exists"
    Else
        dictGroupDateAndRoom.Add arg_groupIndex, Array(arg_date, arg_room)
        NumberOfGroups = NumberOfGroups + 1
    End If
End Function

Public Function get_GroupDate(GroupIndex As Integer) As String
    If Not dictGroupDateAndRoom.Exists(GroupIndex) Then
        get_GroupDate = ""
    Else
        get_GroupDate = dictGroupDateAndRoom(GroupIndex)(0)
    End If
End Function

Public Function get_GroupRoom(GroupIndex As Integer) As String
    If Not dictGroupDateAndRoom.Exists(GroupIndex) Then
        get_GroupRoom = ""
    Else
        get_GroupRoom = dictGroupDateAndRoom(GroupIndex)(1)
    End If
End Function

Public Function GroupExists(GroupIndex As Integer) As Boolean
    GroupExists = dictGroupDateAndRoom.Exists(GroupIndex)
End Function

Public Function Get_TotalNumberOfExercises() As Long
    Dim TotalNumberOfExercises As Long
    
    If Lab0 Then
        TotalNumberOfExercises = NumberOfLabExercises + 1
    Else
        TotalNumberOfExercises = NumberOfLabExercises
    End If
    
    Get_TotalNumberOfExercises = TotalNumberOfExercises
End Function




