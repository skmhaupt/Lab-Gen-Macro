VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LabDetailesUserForm 
   Caption         =   "Lab Details"
   ClientHeight    =   6444
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5472
   OleObjectBlob   =   "LabDetailesUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LabDetailesUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public NumberOfLabExcercises As Long, Lab0 As Boolean, NoEvalFirstLab As Boolean, SubjectName As String, LanguageIndex As Integer
Public CustomLabels As Variant, i As Integer, NumberOfCustomLabels As Integer, UsingCustomLabels As Boolean
Private IsCancelled As Boolean, Sheet1Names As Variant

Public Property Get Cancelled() As Boolean
    Cancelled = IsCancelled
End Property

Private Sub CancelButton_Click()
    OnCancel
End Sub

Private Sub CustomLabelsCheckBox_Click()
    CustomLabelsLabel.Visible = CustomLabelsCheckBox.Value
    CustomLabelsTextBox.Visible = CustomLabelsCheckBox.Value
    If FirstLabEvalCheckBox.Value Then
        LAB0CheckBox.Visible = Not CustomLabelsCheckBox.Value
        LAB0CheckBox.Value = False
    End If
End Sub

Private Sub ExercisesNumberSpinButton_Change()
    ExercisesNumberTextBox.Text = ExercisesNumberSpinButton.Value
End Sub

Private Sub FirstLabEvalCheckBox_Click()
    If Not CustomLabelsCheckBox.Value Then
        LAB0CheckBox.Visible = FirstLabEvalCheckBox.Value
    End If
    
    If FirstLabEvalCheckBox.Value = False Then
        LAB0CheckBox.Value = False
    End If
End Sub

Private Sub LanguageComboBox_Change()
    LanguageIndex = LanguageComboBox.ListIndex + 1
    PrisutnostStudentaLabel.Caption = Sheet1Names(LanguageIndex)
End Sub

Private Sub OKButton_Click()
    If (IsNumeric(ExercisesNumberTextBox.Text) And (Not SubjectTextBox.Text = "")) Then
        SubjectName = SubjectTextBox.Text
        NumberOfLabExcercises = CLng(ExercisesNumberTextBox.Text)
        Lab0 = LAB0CheckBox.Value
        NoEvalFirstLab = FirstLabEvalCheckBox.Value
        
        UsingCustomLabels = CustomLabelsCheckBox.Value
        CustomLabels = Split(CustomLabelsTextBox.Value, ",")
        NumberOfCustomLabels = UBound(CustomLabels) - LBound(CustomLabels) + 1
        For i = 0 To NumberOfCustomLabels - 1
            CustomLabels(i) = WorksheetFunction.Trim(CustomLabels(i))
        Next i
        
        If Not CustomLabelsCheckBox.Value Then
            Me.Hide
        ElseIf NumberOfCustomLabels = NumberOfLabExcercises Then
            Me.Hide
        Else
            MsgBox NumberOfLabExcercises & NumberOfCustomLabels & "Number of custom labels and excercises is not equal. Pleas fix or cancel."
        End If
    Else
        MsgBox "Input the number of excercises and subject name!"
    End If
End Sub

Private Sub UserForm_Initialize()
    Sheet1Names = Array("-Prisutnost-studenta", "-Prisutnost-studenta", _
                   "-Student-attendance", _
                   "-Studenten-anwesenheit", _
                   "-présence-d'etudiants")

    SubjectTextBox.Text = ""
    FirstLabEvalCheckBox.Value = False
    ExercisesNumberTextBox.Value = ""
    LAB0CheckBox.Value = False
    LAB0CheckBox.Visible = False
    CustomLabelsCheckBox.Value = False
    CustomLabelsLabel.Visible = False
    CustomLabelsTextBox.Visible = False
    
    CustomLabelsTextBox.MultiLine = True
    MultiPage.Value = 0
    LanguageIndex = 1
    PrisutnostStudentaLabel.Caption = Sheet1Names(LanguageIndex)
    SubjectTextBox.SetFocus
    
    'Font settings
    'SubjectTextBox.Font.Size = 12
    'SubjectLabel.Font.Size = 12
    'PrisutnostStudentaLabel.Font.Size = 12
    'MultiPage.Font.Size = 12
    'CustomLabelsTextBox.Font.Size = 12
    
    
    With Me.LanguageComboBox
    .Clear      'Clear previous items
    .AddItem "Hr"
    .AddItem "Eng"
    .AddItem "Ger"
    .AddItem "Fr"
    .ListIndex = 0
    End With

End Sub

Public Sub SetUserFormOldData(arg_LabSetting As LabSettings, arg_languageSetting As LanguageSettings)
    
    Lab0 = arg_LabSetting.Lab0
    NoEvalFirstLab = arg_LabSetting.NoEvalFirstLab
    SubjectName = arg_LabSetting.SubjectName
    NumberOfLabExcercises = arg_LabSetting.NumberOfLabExercises
    
    If arg_LabSetting.UsingCustomExcerciseLabels Then
        CustomLabelsCheckBox.Value = True
        CustomLabelsTextBox.Value = arg_LabSetting.CustomExcerciseLabels(0)
        For i = 1 To arg_LabSetting.NumberOfLabExercises - 1
            CustomLabelsTextBox.Value = CustomLabelsTextBox.Value & ", " & arg_LabSetting.CustomExcerciseLabels(i)
        Next i
    End If
    
    LanguageIndex = arg_languageSetting.LanguageIndex
    PrisutnostStudentaLabel.Caption = Sheet1Names(LanguageIndex)
    LanguageComboBox.ListIndex = LanguageIndex - 1
    
    ExercisesNumberTextBox.Value = NumberOfLabExcercises
    ExercisesNumberSpinButton.Value = NumberOfLabExcercises
    SubjectTextBox.Text = SubjectName
    
    'Check if first lab is LAB0
    If Lab0 Then
        FirstLabEvalCheckBox.Value = True
        LAB0CheckBox.Value = True
        LAB0CheckBox.Visible = True
    'Check if first lab is not evaled
    ElseIf NoEvalFirstLab Then
        FirstLabEvalCheckBox.Value = True
        LAB0CheckBox.Value = False
        LAB0CheckBox.Visible = True
    Else
        FirstLabEvalCheckBox.Value = False
        LAB0CheckBox.Value = False
        LAB0CheckBox.Visible = False
    'Error! Workbook formula for ODRADENO not found
    End If
    
    If arg_LabSetting.UsingCustomExcerciseLabels Then
        CustomLabelsCheckBox.Value = True
        LAB0CheckBox.Visible = False
        CustomLabelsTextBox.Value = arg_LabSetting.CustomExcerciseLabels(0)
        For i = 1 To arg_LabSetting.NumberOfLabExercises - 1
            CustomLabelsTextBox.Value = CustomLabelsTextBox.Value & ", " & arg_LabSetting.CustomExcerciseLabels(i)
        Next i
    End If
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    IsCancelled = True
    Me.Hide
End Sub
