VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStart 
   Caption         =   "Health and quality of life"
   ClientHeight    =   8640
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   14910
   OleObjectBlob   =   "frmStart-2022-4-0-2.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------
' GNU GENERAL PUBLIC LICENSE
' Version 3, 29 June 2007
'
' Copyright (C) 2007 Free Software Foundation, Inc. <https://fsf.org/>
' Everyone is permitted to copy and distribute verbatim copies
' of this license document, but changing it is not allowed.
'
' Updates and modifications can be requested at:
' https://github.com/Veronika-Lindberg/SF-36-registration-program
'--------------------------------------------------------------------------------------------------
' Measuring Health by VAS and HRQoL (SF-36 Health and Quality of Life)
'
' frmStart.frm
'
' 2015-04-13 1.0.0 Veronika Lindberg    Created.
'                                       Main picture.
' 2019-08-27 2.0.0 Veronika Lindberg    New dataset for normative data with 2 more age groups.
'                                       Removed scale OH - overall health, the mean of PCS and MCS.
'                                       Swapped rows and columns in myDataArrayByRow.
' 2019-09-11 3.0.0 Veronika Lindberg    Added normative data for Rand Bodily Pain and Rand General Health.
'                                       and Correlated PCS and MCS.
' 2021-12-01 4.0.0 Veronika Lindberg    Windows 10, Office 365, Get local path for microsoft onedrive.
' 2022-02-01 4.0.1 Veronika Lindberg    Use local path if Onedrive path is not found.
'                                       Skip Fain and Fatigue and other fields specific for old projects.
'                                       Survey date is set to Now as default value.
' 2022-02-09 4.0.2 Veronika Lindberg    Allow unknown gender and unknown birth year.
'                                       Set to age 20 if age is unknown.
'--------------------------------------------------------------------------------------------------

Option Explicit

Private Sub cboUsers_Change()
On Error GoTo Errhandler
    SelectedUser = cboUsers.List(cboUsers.ListIndex, 0)
    SelectedBirthYear = cboUsers.List(cboUsers.ListIndex, 1)
    Call select_birthYear_text
    SelectedGender = cboUsers.List(cboUsers.ListIndex, 2)
    SelectedGenderCode = cboUsers.List(cboUsers.ListIndex, 3)
    Call select_gender_text
    Call populate_surveys
Exit Sub
Errhandler:
      ErrorHandling "frmStart. Sub cboUsers_Change", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub cmdSurvey2_Click()
On Error GoTo Errhandler
    If ComboBoxSurvey.ListIndex = -1 Then
        SelectedSheet = ""
    Else
        SelectedSheet = ComboBoxSurvey.List(ComboBoxSurvey.ListIndex, 0)
    End If
    frmNewSurvey.Show
Exit Sub
Errhandler:
      ErrorHandling "frmStart. Sub cmdSurvey2_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub ComboBoxSurvey_Change()
On Error GoTo Errhandler
    If ComboBoxSurvey.ListIndex = -1 Then
        SelectedSheet = ""
    Else
        SelectedSheet = ComboBoxSurvey.List(ComboBoxSurvey.ListIndex, 0)
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmStart. Sub cmdNewSurvey_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub cmdNewSurvey_Click()
On Error GoTo Errhandler
    SelectedSheet = "" 'Shall be empty here
    frmNewSurvey.Show
Exit Sub
Errhandler:
      ErrorHandling "frmStart. Sub cmdNewSurvey_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub cmdNewUser_Click()
On Error GoTo Errhandler
    frmRegisterNewUser.Show
Exit Sub
Errhandler:
      ErrorHandling "frmStart. Sub cmdNewUser_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub




Private Sub initCaptions()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        frmStart.Caption = "Health and quality of life"
        lblHealthAndQOL.Caption = "Health and quality of life"
        FrameUser.Caption = "Person" 'User"
        cmdNewUser.Caption = "Register a new person" 'user"
        lblUser.Caption = "Select an existing person:" 'user:"
        frameLanguage.Caption = "Language"
        FrameSurvey.Caption = "Health survey"
        cmdNewSurvey.Caption = "New survey for selected person" 'user"
        cmdSurvey2.Caption = "New survey for selected person" 'user" '"Create a new survey by editing an old one:" '"Select survey to display:"
        CommandButtonDelete.Caption = "Delete selected survey:"
        FrameGraphs.Caption = "Graphs"
        CommandButtonGraphs.Caption = "View health problems: Visual Analogue Scale" '"Look at graphs for selected user"
        CommandButtonGraphsAll.Caption = "View general health condition: RAND SF-36" '"LoOk at graphs for all users"
    Else
        frmStart.Caption = "Helse og livskvalitet"
        lblHealthAndQOL.Caption = "Helse og livskvalitet"
        FrameUser.Caption = "Person" 'Bruker"
        cmdNewUser.Caption = "Registrer en ny person" 'bruker"
        lblUser.Caption = "Velg en eksisterende person" 'bruker:"
        frameLanguage.Caption = "Språk"
        FrameSurvey.Caption = "Helse spørreundersøkelse"
        cmdSurvey2.Caption = "Ny spørreundersøkelse for valgt person" 'bruker" '"Lag ny spørreundersøkelse av en eksisterende:"
        cmdNewSurvey.Caption = "Ny spørreundersøkelse for valgt person" 'bruker"
        CommandButtonDelete.Caption = "Slett valgte spørreundersøkelse:"
        FrameGraphs.Caption = "Grafikk"
        CommandButtonGraphs.Caption = "Følg helseproblemer: Visual Analogue Scale" '"Se på grafikk for valgt bruker"
        CommandButtonGraphsAll.Caption = "Følg almenntilstand: RAND SF-36" '"Se på grafikk for alle brukere"
    End If
    Me.Repaint
Exit Sub
Errhandler:
      ErrorHandling "frmStart. Sub initCaptions", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub





Private Sub CommandButtonDelete_Click()
On Error GoTo Errhandler
Dim answer As Integer
Dim strPromt As String
Dim strTitle As String
Dim rowCount As Integer
Dim myRange As Variant
Dim c As Variant
Dim vt2 As Variant

    If ComboBoxSurvey.ListIndex = -1 Then
        SelectedSheet = ""
    Else
        SelectedSheet = ComboBoxSurvey.List(ComboBoxSurvey.ListIndex, 0)
    End If
    If SelectedSheet <> "" Then
        If SelectedLanguage = "UK" Then
            strPromt = "Do you want to delete " & SelectedSheet & "?"
            strTitle = "Delete survey"
        Else
            strPromt = "Vil du slette '" & SelectedSheet & "'?"
            strTitle = "Slett spørreundersøkelse"
        End If
        answer = VBA.MsgBox(strPromt, vbYesNo, strTitle)
        If answer = vbYes Then
            Application.DisplayAlerts = False
            Sheets(SelectedSheet).Delete
            Worksheets("SurveySummary").Activate
            
            rowCount = Range("A1").CurrentRegion.Rows.Count
            myRange = "A1:" & "A" & Trim(Str(rowCount))
            With Worksheets("SurveySummary").Range(myRange)
                Set c = .Find(SelectedSheet)
            End With
            
             If Not c Is Nothing Then
                vt2 = VarType(c)
                If Not vt2 = vbString Then
                    c = CStr(c)
                End If
                If UCase(c) = UCase(SelectedSheet) Then
                    Cells(c.Row, 1).EntireRow.Delete
                End If
            End If
    
            Application.DisplayAlerts = True
            
            ThisWorkbook.Save
            Call populate_surveys
        End If
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmStart. Sub CommandButtonDelete_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub CommandButtonGraphs_Click()
On Error GoTo Errhandler
    frmUserFormVAS.Show False
    frmUserFormVAS.initCaptions
    frmUserFormVAS.chartVASScores
Exit Sub
Errhandler:
      ErrorHandling "frmStart. Sub cmdNewUser_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub CommandButtonGraphsAll_Click()
On Error GoTo Errhandler
    frmUserFormSF36.Show False
    frmUserFormSF36.initCaptions
    frmUserFormSF36.chartScaleScores
Exit Sub
Errhandler:
      ErrorHandling "frmStart. Sub cmdNewUser_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub ImageGraphs_Click()
    Call CommandButtonGraphsAll_Click
    Call CommandButtonGraphs_Click
End Sub

Private Sub ImageSurvey_Click()
    Call cmdSurvey2_Click
End Sub

Private Sub imgCircle_Click()
    Call cmdNewUser_Click
End Sub

Private Sub imgNO_Click()
On Error GoTo Errhandler
    SelectedLanguage = "NO"
    imgNO.SpecialEffect = fmSpecialEffectRaised
    imgUK.SpecialEffect = fmSpecialEffectFlat
    Call initCaptions
Exit Sub
Errhandler:
      ErrorHandling "frmStart. Sub imgNO_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub imgUK_Click()
On Error GoTo Errhandler
    SelectedLanguage = "UK"
    imgUK.SpecialEffect = fmSpecialEffectRaised
    imgNO.SpecialEffect = fmSpecialEffectFlat
    Call initCaptions
Exit Sub
Errhandler:
      ErrorHandling "frmStart. Sub imgUK_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub UserForm_Initialize()
On Error GoTo Errhandler
    Call initCaptions
    'MsgBox "With= " & Me.Width & " height " & Me.Height
Exit Sub
Errhandler:
      ErrorHandling "frmStart. Sub UserForm_Initialize", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub
