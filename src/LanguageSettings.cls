VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LanguageSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Hr_txt As Variant, Eng_txt As Variant
Private Ger_txt As Variant, Fr_txt As Variant
Private dict As Object, mirror As Object
Private ChosenLang As Variant

Public NameLabel As String, DoneLabel As String
Public AverageLabel As String, GroupLabel As String
Public YesLabel As String, NoLabel As String
Public Sheet1Name As String, Sheet2Name As String
Public AlreadyDoneLabel As String, ScheduleLabel As String
Public RoomLabel As String

Private Sub Class_Initialize()
    Hr_txt = Array("Prezime i Ime", _
                   "ODRAÐENO", _
                   "PROSJEK", _
                   "GRUPA", _
                   "DA", _
                   "NE", _
                   "-Prisutnost-studenta", _
                   "TabliceGrupa", _
                   "Hr_txt", _
                   "PRETHODNO" & vbLf & "POLOŽILI", _
                   "RASPORED")
                   
    Eng_txt = Array("Full Name", _
                    "DONE", _
                    "AVERAGE", _
                    "GROUP", _
                    "YES", _
                    "NO", _
                    "-Student-attendance", _
                    "GroupTables", _
                    "Eng_txt", _
                    "ALREADY" & vbLf & "DONE", _
                    "SCHEDULE")
                   
    Ger_txt = Array("Nachname und Vorname", _
                    "ABGESCHLOSSEN", _
                    "DURCHSCHNITT", _
                    "GRUPE", _
                    "JA", _
                    "NEIN", _
                    "-Studenten-anwesenheit", _
                    "-Grupentabellen", _
                    "Ger_txt", _
                    "BEREITS" & vbLf & "ABGESCHLOSSEN", _
                    "ZEITPLAN")
                   
    Fr_txt = Array("Nom de famille et Nom", _
                   "REALISE", _
                   "MOYENNE", _
                   "GROUPE", _
                   "OUI", _
                   "NON", _
                   "-présence-d'etudiants", _
                   "TabliceGrupa", _
                   "Fr_txt", _
                   "DÉJÀ" & vbLf & "REALISE", _
                   "L'HORAIRE")
                                     
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add 1, Hr_txt       'LanguageIndex - x_txt
    dict.Add 2, Eng_txt
    dict.Add 3, Ger_txt
    dict.Add 4, Fr_txt
    
    Set mirror = CreateObject("Scripting.Dictionary")
    mirror.Add "Hr_txt", 1       'x_txt - LanguageIndex
    mirror.Add "Eng_txt", 2
    mirror.Add "Ger_txt", 3
    mirror.Add "Fr_txt", 4
    
    ChosenLang = Hr_txt
End Sub

Public Sub SetLanguage(LanguageIndex As Integer)
    If LanguageIndex > 0 And LanguageIndex < 5 Then
        ChosenLang = dict(LanguageIndex)
    Else
        MsgBox "Error in language setup! Using default setting (Cro)."
        ChosenLang = Hr_txt
    End If
    
    NameLabel = ChosenLang(0)
    DoneLabel = ChosenLang(1)
    AverageLabel = ChosenLang(2)
    GroupLabel = ChosenLang(3)
    YesLabel = ChosenLang(4)
    NoLabel = ChosenLang(5)
    Sheet1Name = ChosenLang(6)
    Sheet2Name = ChosenLang(7)
    AlreadyDoneLabel = ChosenLang(9)
    ScheduleLabel = ChosenLang(10)
    RoomLabel = "LAB"
End Sub

Public Property Get LanguageIndex() As Integer
    LanguageIndex = mirror(ChosenLang(8))
End Property
