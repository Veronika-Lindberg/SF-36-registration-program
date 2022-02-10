VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNewSurvey 
   Caption         =   "Health survey"
   ClientHeight    =   11020
   ClientLeft      =   80
   ClientTop       =   500
   ClientWidth     =   16240
   OleObjectBlob   =   "frmNewSurvey-2022-4-0-2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNewSurvey"
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
' frmNewSurvey.frm
'
' 2015-04-13 1.0.0 Veronika Lindberg    Created.
'                                       Calculate SF-36 scores.
'                                       Each question is given equal weight in the summary scores.
' 2015-06-28 1.1.0 Veronika Lindberg    (Option removed again. Changed calculation of summary scores:
'                                       Each category is given equal weight in the summary scores.)
' 2019-08-27 2.0.0 Veronika Lindberg    New dataset for normative data with 2 more age groups.
'                                       Removed scale OH - overall health, the mean of PCS and MCS
'                                       Swapped rows and columns in myDataArrayByRow
'                                       add z value to summary sheet
'                                       Changed calculation of PCS and MCS - the mean is already norm based
' 2019-09-11 2.0.0 Veronika Lindberg    Added normative data for Rand Bodily Pain and Rand General Health
'                                       and Correlated PCS and MCS
'                                       Each category is given equal weight in summary scores
'                                       PCS = 0.25(PF + RP + BP + GH) SCALE 1,2,7,8
'                                       MCS = 0.25(RE + VT + MH + SF) SCALE 3,4,5,6
'                                       Global Health Composite = 0.125 (PF + RP + BP + GH + RE + VT + MH + SF) SCALE 1-8
' 2019-09-11 3.0.0 Veronika Lindberg    Added normative data for Rand Bodily Pain and Rand General Health.
'                                       and Correlated PCS and MCS.
' 2019-10-18 3.0.0 Veronika Lindberg    Compare version 2 and 3
''2019-10-22                            Uses correlated scores instead
'                                       PF + RP + BP + GH + VT + SF + RE + MH
'                                       PCS 0,20    0,31    0,23    0,20    0,13    0,11    0,03    -0,03
'                                       MCS -0,02   0,03    0,04    0,10    0,29    0,14    0,20    0,35
'                                       Global Health Composite = 0.125 (PF + RP + BP + GH + RE + VT + MH + SF) SCALE 1-8
'                                       Added check on number_of_health_problems = 0
' 2021-12-01 4.0.0 Veronika Lindberg    Windows 10, Office 365, Get local path for microsoft onedrive.
' 2022-02-01 4.0.1 Veronika Lindberg    Use local path if Onedrive path is not found.
'                                       Skip Fain and Fatigue and other fields specific for old projects.
'                                       Survey date is set to Now as default value.
' 2022-02-09 4.0.2 Veronika Lindberg    Allow unknown gender and unknown birth year.
'                                       Set to age 20 if age is unknown.
'--------------------------------------------------------------------------------------------------


Option Explicit

'user data shall be stored in an excel sheet
Const maxUserDataRows = 1000
Const maxUserDataCols = 117
Dim UserData(1 To maxUserDataRows, 1 To maxUserDataCols) As Variant
'Dim UserDataTransposed(1 To maxUserDataRows, 1 To maxUserDataCols) As Variant
'survey data shall be stored in an excel sheet
Const nRows = 100
Const nCols = 15
Dim SurveyData(nRows, nCols) As Variant
Dim myOldSurveyData As Variant
Dim normativeData As Variant
'values controls must be stored in a variable
'because the values of the user controls behave a bit strange
'when the tab is disabled
'
Dim userControlValuesPage0(5) As Variant
Dim userTextboxValuesPage1(19) As Variant
Dim userControlValuesPage1(19) As Variant
Dim userControlValuesPage2(nRows) As Variant
Dim userControlValuesPage11(1 To 3) As Variant


Dim sf_36_questions(1 To 39, 1 To 5) As Variant 'question number, response, calculated value, question text, answer text
                                      'sf 36 items + 3 additional items = 39

Dim SurveyRegisteredDate As String
Dim surveyWSName As String 'work sheet name
Dim oldSurveyWSName As String 'old work sheet name
Public surveyDate As Variant

'CalcSF36scores
Dim scale_group_label(1 To 11) As String
Dim scale_number_in_group(1 To 11) As Double
Dim scale_sum_group(1 To 11) As Double
Dim scale_average_group(1 To 11) As Double
Dim new_sorted_scale_average_group(1 To 11) As Double
Dim z_scale_average_group(1 To 11) As Double
Dim normalized_scale_average_group(1 To 11) As Double
Dim countMissingItems As Integer
'
Const n_summary_values = 77 '67 '22 VAS, 11 norm based scores, 11 scores, 11 norms, 11 sd, + 11  z scores
'18.05.2015 added q1 and q2 to summary_q1q2
'30.08.2019 added number of missing items to summary
Dim summary_values(1 To n_summary_values) As Double
Dim summary_values_ix As Integer
Dim summary_vas_text(1 To 22) As String
Dim summary_vas_text_ix As Integer
Dim summary_text(1 To 4) As String
Dim summary_q1q2(1 To 6) As String
Dim summary_extra_values(1 To 6) As String
Dim hasSaved As Boolean




Private Sub calc_all_SF36scores()
On Error GoTo Errhandler
    Dim ix As Integer
    Dim groupIX As Integer
    Dim calculatedValue As Integer
    Dim myNormScore As Double
    Dim myNormSD As Double
    
    'Call get_normdata(SelectedGenderCode, selectedAge, myNormScore, myNormSD)
    
    countMissingItems = 0
    For groupIX = 1 To 11
        scale_group_label(groupIX) = ""
        scale_number_in_group(groupIX) = 0
        scale_sum_group(groupIX) = 0
        scale_average_group(groupIX) = 0
        new_sorted_scale_average_group(groupIX) = 0
    Next groupIX
    
    Call get_scale_label_text
  
    For ix = 1 To 36
        If sf_36_questions(ix, 5) = "Missing" Then 'answerText
            countMissingItems = countMissingItems + 1
        Else
            If sf_36_questions(ix, 1) = "Q2" Then
                calculatedValue = sf_36_questions(ix, 3)
                'item 2 does not belong to a group
                'scale_number_in_group(11) = scale_number_in_group(11) + 1 'overall group
                'scale_sum_group(11) = scale_sum_group(11) + calculatedValue 'overall health
            End If
        
            If (sf_36_questions(ix, 1) = "Q3" _
                Or sf_36_questions(ix, 1) = "Q4" _
                Or sf_36_questions(ix, 1) = "Q5" _
                Or sf_36_questions(ix, 1) = "Q6" _
                Or sf_36_questions(ix, 1) = "Q7" _
                Or sf_36_questions(ix, 1) = "Q8" _
                Or sf_36_questions(ix, 1) = "Q9" _
                Or sf_36_questions(ix, 1) = "Q10" _
                Or sf_36_questions(ix, 1) = "Q11" _
                Or sf_36_questions(ix, 1) = "Q12") Then
                calculatedValue = sf_36_questions(ix, 3)
                scale_number_in_group(1) = scale_number_in_group(1) + 1 'belongs to group 1
                scale_sum_group(1) = scale_sum_group(1) + calculatedValue 'sum group
                scale_number_in_group(11) = scale_number_in_group(11) + 1 'overall group
                scale_sum_group(11) = scale_sum_group(11) + calculatedValue 'overall health
                scale_number_in_group(9) = scale_number_in_group(9) + 1 '
                scale_sum_group(9) = scale_sum_group(9) + calculatedValue
            End If
        
            If (sf_36_questions(ix, 1) = "Q13" _
                Or sf_36_questions(ix, 1) = "Q14" _
                Or sf_36_questions(ix, 1) = "Q15" _
                Or sf_36_questions(ix, 1) = "Q16") Then
                calculatedValue = sf_36_questions(ix, 3)
                scale_number_in_group(2) = scale_number_in_group(2) + 1 'belongs to group 2
                scale_sum_group(2) = scale_sum_group(2) + calculatedValue 'sum group
                scale_number_in_group(11) = scale_number_in_group(11) + 1 'overall group
                scale_sum_group(11) = scale_sum_group(11) + calculatedValue 'overall health
                scale_number_in_group(9) = scale_number_in_group(9) + 1 '
                scale_sum_group(9) = scale_sum_group(9) + calculatedValue
            End If
        
            If (sf_36_questions(ix, 1) = "Q17" _
                Or sf_36_questions(ix, 1) = "Q18" _
                Or sf_36_questions(ix, 1) = "Q19") Then
                calculatedValue = sf_36_questions(ix, 3)
                scale_number_in_group(3) = scale_number_in_group(3) + 1 'belongs to group 3
                scale_sum_group(3) = scale_sum_group(3) + calculatedValue 'sum group
                scale_number_in_group(11) = scale_number_in_group(11) + 1 'overall group
                scale_sum_group(11) = scale_sum_group(11) + calculatedValue 'overall health
                scale_number_in_group(10) = scale_number_in_group(10) + 1 '
                scale_sum_group(10) = scale_sum_group(10) + calculatedValue
            End If

        
            If (sf_36_questions(ix, 1) = "Q23" _
                Or sf_36_questions(ix, 1) = "Q27" _
                Or sf_36_questions(ix, 1) = "Q29" _
                Or sf_36_questions(ix, 1) = "Q31") Then
                calculatedValue = sf_36_questions(ix, 3)
                scale_number_in_group(4) = scale_number_in_group(4) + 1 'belongs to group 4
                scale_sum_group(4) = scale_sum_group(4) + calculatedValue 'sum group
                scale_number_in_group(11) = scale_number_in_group(11) + 1 'overall group
                scale_sum_group(11) = scale_sum_group(11) + calculatedValue 'overall health
                scale_number_in_group(10) = scale_number_in_group(10) + 1 '
                scale_sum_group(10) = scale_sum_group(10) + calculatedValue
            End If

        
            If (sf_36_questions(ix, 1) = "Q24" _
                Or sf_36_questions(ix, 1) = "Q25" _
                Or sf_36_questions(ix, 1) = "Q26" _
                Or sf_36_questions(ix, 1) = "Q28" _
                Or sf_36_questions(ix, 1) = "Q30") Then
                calculatedValue = sf_36_questions(ix, 3)
                scale_number_in_group(5) = scale_number_in_group(5) + 1 'belongs to group 5
                scale_sum_group(5) = scale_sum_group(5) + calculatedValue 'sum group
                scale_number_in_group(11) = scale_number_in_group(11) + 1 'overall group
                scale_sum_group(11) = scale_sum_group(11) + calculatedValue 'overall health
                scale_number_in_group(10) = scale_number_in_group(10) + 1 '
                scale_sum_group(10) = scale_sum_group(10) + calculatedValue
            End If
        
            If (sf_36_questions(ix, 1) = "Q20" _
                Or sf_36_questions(ix, 1) = "Q32") Then
                calculatedValue = sf_36_questions(ix, 3)
                scale_number_in_group(6) = scale_number_in_group(6) + 1 'belongs to group 6
                scale_sum_group(6) = scale_sum_group(6) + calculatedValue 'sum group
                scale_number_in_group(11) = scale_number_in_group(11) + 1 'overall group
                scale_sum_group(11) = scale_sum_group(11) + calculatedValue 'overall health
                scale_number_in_group(10) = scale_number_in_group(10) + 1 '
                scale_sum_group(10) = scale_sum_group(10) + calculatedValue
            End If
        
        
            If (sf_36_questions(ix, 1) = "Q21" _
                Or sf_36_questions(ix, 1) = "Q22") Then
                calculatedValue = sf_36_questions(ix, 3)
                
                scale_number_in_group(7) = scale_number_in_group(7) + 1 'belongs to group 7
                scale_sum_group(7) = scale_sum_group(7) + calculatedValue 'sum group
                scale_number_in_group(11) = scale_number_in_group(11) + 1 'overall group
                scale_sum_group(11) = scale_sum_group(11) + calculatedValue 'overall health
                scale_number_in_group(9) = scale_number_in_group(9) + 1 '
                scale_sum_group(9) = scale_sum_group(9) + calculatedValue
            End If
        
            If (sf_36_questions(ix, 1) = "Q1" _
                Or sf_36_questions(ix, 1) = "Q33" _
                Or sf_36_questions(ix, 1) = "Q34" _
                Or sf_36_questions(ix, 1) = "Q35" _
                Or sf_36_questions(ix, 1) = "Q36") Then
                calculatedValue = sf_36_questions(ix, 3)
                scale_number_in_group(8) = scale_number_in_group(8) + 1 'belongs to group 8
                scale_sum_group(8) = scale_sum_group(8) + calculatedValue 'sum group
                scale_number_in_group(11) = scale_number_in_group(11) + 1 'overall group
                scale_sum_group(11) = scale_sum_group(11) + calculatedValue 'overall health
                scale_number_in_group(9) = scale_number_in_group(9) + 1 '
                scale_sum_group(9) = scale_sum_group(9) + calculatedValue

            End If
        
        End If
        
    Next ix
    'Each question is given equal weiht in summary scores
     
    'PCS
    scale_number_in_group(9) = scale_number_in_group(1) _
                                + scale_number_in_group(2) _
                                + scale_number_in_group(7) _
                                + scale_number_in_group(8)
    scale_sum_group(9) = scale_sum_group(1) _
                                + scale_sum_group(2) _
                                + scale_sum_group(7) _
                                + scale_sum_group(8)
                                
    'MCS
    scale_number_in_group(10) = scale_number_in_group(3) _
                                + scale_number_in_group(4) _
                                + scale_number_in_group(5) _
                                + scale_number_in_group(6)
    scale_sum_group(10) = scale_sum_group(3) _
                                + scale_sum_group(4) _
                                + scale_sum_group(5) _
                                + scale_sum_group(6)
    'Global Health Composite
    scale_number_in_group(11) = scale_number_in_group(9) + scale_number_in_group(10)
    scale_sum_group(11) = scale_sum_group(9) + scale_sum_group(10)
    
    'Calculate average for each group, 1-8, This is the sf-36 raw score
    For ix = 1 To 11
        If scale_number_in_group(ix) > 0 Then
            scale_average_group(ix) = scale_sum_group(ix) / scale_number_in_group(ix)
        End If
        scale_group_label(ix) = "Scale" & Str(ix)
    Next ix
    
    '2015-06-28: Each category is given equal weight in summary scores
    '2019-09-11: Below, back to original RAND-36 computation: Each category is given equal weight in summary scores
    ' PCS = 0.25(PF + RP + BP + GH) SCALE 1,2,7,8
    ' MCS = 0.25(RE + VT + MH + SF) SCALE 3,4,5,6
    ' Global Health Composite = 0.125 (PF + RP + BP + GH + RE + VT + MH + SF) SCALE 1-8

    scale_average_group(9) = (scale_average_group(1) + scale_average_group(2) + scale_average_group(7) + scale_average_group(8)) / 4
    scale_average_group(10) = (scale_average_group(3) + scale_average_group(4) + scale_average_group(5) + scale_average_group(6)) / 4
    scale_average_group(11) = (scale_average_group(1) + scale_average_group(2) + scale_average_group(7) + scale_average_group(8) + scale_average_group(3) + scale_average_group(4) + scale_average_group(5) + scale_average_group(6)) / 8
    
    '2019-09-13 Uses correlated scores instead
    ' PF + RP + BP + GH + VT + SF + RE + MH
    'PCS 0,20    0,31    0,23    0,20    0,13    0,11    0,03    -0,03
    'MCS -0,02   0,03    0,04    0,10    0,29    0,14    0,20    0,35
    'Global Health Composite = 0.125 (PF + RP + BP + GH + RE + VT + MH + SF) SCALE 1-8

    scale_average_group(9) = (0.2 * scale_average_group(1) + 0.31 * scale_average_group(2) + 0.23 * scale_average_group(7) + 0.2 * scale_average_group(8) + 0.13 * scale_average_group(4) + 0.11 * scale_average_group(6) + 0.03 * scale_average_group(3) + (-0.03) * scale_average_group(5))
    scale_average_group(10) = ((-0.02) * scale_average_group(1) + 0.03 * scale_average_group(2) + 0.04 * scale_average_group(7) + 0.1 * scale_average_group(8) + 0.29 * scale_average_group(4) + 0.14 * scale_average_group(6) + 0.2 * scale_average_group(3) + 0.35 * scale_average_group(5))
    scale_average_group(11) = (scale_average_group(1) + scale_average_group(2) + scale_average_group(7) + scale_average_group(8) + scale_average_group(3) + scale_average_group(4) + scale_average_group(5) + scale_average_group(6)) / 8
    
Exit Sub
Errhandler:
      ErrorHandling "ModuleCalcSF36scores. Sub calc_all_SF36scores", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub



Private Sub FetchUserControlValuesBeforeSave()
On Error GoTo Errhandler
    Dim yearNow As Integer
    Dim i As Integer
    Dim ix As Integer
    Dim surveyAge As Integer
    Dim colIx As Integer
    Dim sortIX As Integer
    
    
    If Me.MultiPage1.Value = 0 Then
        Call keepUserControlValuesFromPage0
    ElseIf Me.MultiPage1.Value = 1 Then
        Call keepUserControlValuesFromPage1
    ElseIf Me.MultiPage1.Value = 2 Then
        Call keepUserControlValuesFromPage2
    ElseIf Me.MultiPage1.Value = 3 Then
        Call keepUserControlValuesFromPage3
    ElseIf Me.MultiPage1.Value = 4 Then
        Call keepUserControlValuesFromPage4
    ElseIf Me.MultiPage1.Value = 5 Then
        Call keepUserControlValuesFromPage5
    ElseIf Me.MultiPage1.Value = 6 Then
        Call keepUserControlValuesFromPage6
    ElseIf Me.MultiPage1.Value = 7 Then
        Call keepUserControlValuesFromPage7
    ElseIf Me.MultiPage1.Value = 8 Then
        Call keepUserControlValuesFromPage8
    ElseIf Me.MultiPage1.Value = 9 Then
        Call keepUserControlValuesFromPage9
    ElseIf Me.MultiPage1.Value = 10 Then
        Call keepUserControlValuesFromPage10
    ElseIf Me.MultiPage1.Value = 11 Then
        Call keepUserControlValuesFromPage11
    End If
    
    yearNow = Format(surveyDate, "yyyy")
    'MsgBox "Fetchusercontrolvaluesbeforsave  yearnow " & yearNow
    If IsNumeric(SelectedBirthYear) Then
        surveyAge = yearNow - SelectedBirthYear
    Else
        surveyAge = -1
    End If
    For ix = 0 To nRows - 1
        SurveyData(ix, 0) = surveyWSName
        SurveyData(ix, 1) = SurveyRegisteredDate
        SurveyData(ix, 2) = SelectedUser
        SurveyData(ix, 3) = SelectedGender
        SurveyData(ix, 4) = SelectedGenderCode
        SurveyData(ix, 5) = surveyAge
    Next ix
    i = 0
    'page0
    SurveyData(i, 6) = "Survey date"
    SurveyData(i, 7) = userControlValuesPage0(0) 'Format(DTPickerSurveyDate, "yyyy-mm-dd")
    'MsgBox "Fetchusercontrolvaluesbeforsave  SurveyData(i, 7)" & SurveyData(i, 7)
    i = i + 1
    SurveyData(i, 6) = "Survey group"
    SurveyData(i, 7) = userControlValuesPage0(1) 'SliderSurveyGroup.Value '0=none,1=retrospective,2=prospective
    i = i + 1
    SurveyData(i, 6) = "Form number"
    SurveyData(i, 7) = userControlValuesPage0(2) 'TextBoxFormNumber.Text 'default 1
    summary_extra_values(2) = userControlValuesPage0(2)
    i = i + 1
    SurveyData(i, 6) = "Form language"
    SurveyData(i, 7) = userControlValuesPage0(3) 'SliderSurveyLanguage.Value '0=english,1=norwegian,2=german,3=dutch
    i = i + 1
    SurveyData(i, 6) = "Months after baseline"
    SurveyData(i, 7) = userControlValuesPage0(4) 'TextBoxWeeksAfterBaseline.Text  'default =0
    summary_extra_values(4) = userControlValuesPage0(4)
    i = i + 1
    SurveyData(i, 6) = "Notes"
    SurveyData(i, 7) = userControlValuesPage0(5) 'TextBoxNotes.Text  'default =""
    summary_text(1) = userControlValuesPage0(5)

    'page1
    summary_values_ix = 0
    
    Dim number_of_health_problems As Integer
    Dim mean_of_health_problems As Double
    
    number_of_health_problems = 0
    mean_of_health_problems = 0
    
    For ix = 0 To 19
        i = i + 1
        SurveyData(i, 6) = "Health problem " & Str(ix + 1)
        SurveyData(i, 7) = userControlValuesPage1(ix) 'SliderHealthProblem1.Value   'default =0
        SurveyData(i, 8) = userTextboxValuesPage1(ix) 'TextBoxHealthProblem1.Text   'default =""
        summary_values_ix = summary_values_ix + 1
        If IsNumeric(userControlValuesPage1(ix)) Then
            summary_values(summary_values_ix) = userControlValuesPage1(ix) 'save health problem values 1-20
        Else
            summary_values(summary_values_ix) = 0
        End If
        summary_vas_text(summary_values_ix) = userTextboxValuesPage1(ix)
        If summary_vas_text(summary_values_ix) <> "" Then
            number_of_health_problems = number_of_health_problems + 1 ' count number of problems
            mean_of_health_problems = mean_of_health_problems + userControlValuesPage1(ix) ' add values
        End If
    Next ix
    summary_extra_values(5) = number_of_health_problems
    If number_of_health_problems = 0 Then
        summary_extra_values(6) = 0
    Else
        summary_extra_values(6) = mean_of_health_problems / number_of_health_problems ' calculate mean of problems
    End If
    
    'page 2
    i = i + 1
    Dim strPain As String
    Dim strLackOfEnergy As String
    
    strPain = get_pain_text
    strLackOfEnergy = get_lack_of_energy_text


    summary_values_ix = summary_values_ix + 1
    SurveyData(i, 6) = strPain '"Pain"
    SurveyData(i, 7) = userControlValuesPage2(0) 'SliderPain.Value   'default =0
    If IsNumeric(userControlValuesPage2(0)) Then
        summary_values(summary_values_ix) = userControlValuesPage2(0)
    Else
        summary_values(summary_values_ix) = 0
    End If
    summary_vas_text(summary_values_ix) = strPain
    
    i = i + 1
    summary_values_ix = summary_values_ix + 1
    SurveyData(i, 6) = strLackOfEnergy '"Lack of energy"
    SurveyData(i, 7) = userControlValuesPage2(1) 'SliderEnergy.Value   'default =0
    If IsNumeric(userControlValuesPage2(1)) Then
        summary_values(summary_values_ix) = userControlValuesPage2(1)
    Else
        summary_values(summary_values_ix) = 0
    End If
    summary_vas_text(summary_values_ix) = strLackOfEnergy '"Lack of energy"
    
    summary_q1q2(1) = sf_36_questions(1, 5) 'Q1 answer
    summary_q1q2(2) = sf_36_questions(1, 2) 'Q1 slider value
    summary_q1q2(3) = sf_36_questions(1, 3) 'Q1 calculatedValue
    summary_q1q2(4) = sf_36_questions(2, 5) 'Q2 answer
    summary_q1q2(5) = sf_36_questions(2, 2) 'Q2 slider value
    summary_q1q2(6) = sf_36_questions(2, 3) 'Q2 calculatedValue
    'page 3+, q1-q36
    For ix = 1 To 36
        i = i + 1
        SurveyData(i, 6) = sf_36_questions(ix, 1) '"Q"
        SurveyData(i, 7) = sf_36_questions(ix, 2) 'SliderQ1.Value 'default =0
        SurveyData(i, 8) = sf_36_questions(ix, 3) 'calculatedValue
        SurveyData(i, 9) = sf_36_questions(ix, 4) 'LabelQ.Caption
        SurveyData(i, 10) = sf_36_questions(ix, 5) 'answerText
    Next ix
    
    For ix = 1 To 3
        i = i + 1
        SurveyData(i, 6) = sf_36_questions(36 + ix, 1) '"Q"
        SurveyData(i, 7) = sf_36_questions(36 + ix, 2) 'SliderQ1.Value 'default =0
        SurveyData(i, 8) = sf_36_questions(36 + ix, 3) 'calculatedValue
        SurveyData(i, 9) = sf_36_questions(36 + ix, 4) 'LabelQ.Caption
        SurveyData(i, 10) = sf_36_questions(36 + ix, 5) 'answerText
        If ix = 1 Then
            SurveyData(i, 11) = userControlValuesPage11(ix) '= TextBoxq37.Text
            summary_text(2) = userControlValuesPage11(ix)
        ElseIf ix = 2 Then
            SurveyData(i, 11) = userControlValuesPage11(ix) ' = TextBoxq38.Text
            summary_text(3) = userControlValuesPage11(ix)
        Else
            SurveyData(i, 11) = userControlValuesPage11(ix) ' = TextBoxq39.Text
            summary_text(4) = userControlValuesPage11(ix)
        End If
    Next ix
   
    Call calc_all_SF36scores
    i = i + 1
    SurveyData(i, 6) = "Number of missing items"
    SurveyData(i, 7) = Str(countMissingItems)
    summary_extra_values(1) = Str(countMissingItems)
        
    For ix = 1 To 11 '10 ' 11
        i = i + 1
        SurveyData(i, 6) = scale_group_label(ix)
        SurveyData(i, 7) = scale_average_group(ix)
        SurveyData(i, 8) = scale_number_in_group(ix)
        SurveyData(i, 9) = scale_sum_group(ix)
        SurveyData(i, 10) = scale_group_label_text(ix)
    Next ix
    
    
    
    'get group to compare to
    'If (SelectedGenderCode = "-1") Or (SelectedGenderCode = "") Then
    '    colIx = 2 '15 'gender is unknown. Suppose age is 20
    'Else
    '    If SelectedGenderCode = 1 Then
    '        colIx = 0 'male
    '    Else
    '        colIx = 1 'female
    '    End If
    
        Select Case SelectedGenderCode
            Case Is = 1
                colIx = 0 'male
            Case Is = 0
                colIx = 1 'female
            Case Else
                colIx = 2 'unknown
        End Select
        
        Select Case surveyAge
            Case Is < 0
                colIx = colIx + 4 'set to age 20 if unknown
            Case 0 To 19
                colIx = colIx + 1
            Case 20 To 29
                colIx = colIx + 4
            Case 30 To 39
                colIx = colIx + 7
            Case 40 To 49
                colIx = colIx + 10
            Case 50 To 59
                colIx = colIx + 13
            Case 60 To 69
                colIx = colIx + 16
            Case 70 To 79
                colIx = colIx + 19
            Case Is > 79
                colIx = colIx + 22
            Case Else
                colIx = colIx + 4 'set to age 20 if unknown
        End Select
    'End If
    
    'need to rearrange data to get correct order according to new norm dataset
    
    new_sorted_scale_average_group(1) = scale_average_group(1)
    new_sorted_scale_average_group(2) = scale_average_group(2)
    new_sorted_scale_average_group(7) = scale_average_group(3)
    new_sorted_scale_average_group(5) = scale_average_group(4)
    new_sorted_scale_average_group(8) = scale_average_group(5)
    new_sorted_scale_average_group(6) = scale_average_group(6)
    new_sorted_scale_average_group(3) = scale_average_group(7)
    new_sorted_scale_average_group(4) = scale_average_group(8)
    new_sorted_scale_average_group(9) = scale_average_group(9)
    new_sorted_scale_average_group(10) = scale_average_group(10)
    new_sorted_scale_average_group(11) = scale_average_group(11)
    
    
    'zPF = (PFNN – PFRef) / |SD PFRef |
    'calculate normalized scores
    For ix = 1 To 8 '11
        z_scale_average_group(ix) = (new_sorted_scale_average_group(ix) - myDataArrayByRow(colIx, ix).d_mean) / _
            (Abs(myDataArrayByRow(colIx, ix).d_SD))
        normalized_scale_average_group(ix) = 50 + (10 * z_scale_average_group(ix))
    Next ix
    'Dim z_from_norm_based_mean As Double
    'Dim z_from_4_scales As Double
    
    'ix = 9
    'z_from_norm_based_mean = (myDataArrayByRow(colIx, ix).d_mean - 50) / 10 ' from norm based t-value to z-value
    'z_from_4_scales = z_scale_average_group(1) + z_scale_average_group(2) + z_scale_average_group(3) + z_scale_average_group(4)
    'z_scale_average_group(ix) = (z_from_4_scales / 4) - z_from_norm_based_mean
    'normalized_scale_average_group(ix) = 50 + (10 * z_scale_average_group(ix))
    
    'ix = 10
    'z_from_norm_based_mean = (myDataArrayByRow(colIx, ix).d_mean - 50) / 10 ' from norm based t-value to z-value
    'z_from_4_scales = z_scale_average_group(5) + z_scale_average_group(6) + z_scale_average_group(7) + z_scale_average_group(8)
    'z_scale_average_group(ix) = (z_from_4_scales / 4) - z_from_norm_based_mean
    'normalized_scale_average_group(ix) = 50 + (10 * z_scale_average_group(ix))
    
    'PCS equal weights 0.25(PF+RP+BP+GH)
    'z_scale_average_group(9) = (z_scale_average_group(1) + z_scale_average_group(2) + z_scale_average_group(3) + z_scale_average_group(4)) / 4
    ' PF + RP + BP + GH + VT + SF + RE + MH
    'PCS 0,20    0,31    0,23    0,20    0,13    0,11    0,03    -0,03
    z_scale_average_group(9) = (0.2 * z_scale_average_group(1) + 0.31 * z_scale_average_group(2) + 0.23 * z_scale_average_group(3) + 0.2 * z_scale_average_group(4) + 0.13 * z_scale_average_group(5) + 0.11 * z_scale_average_group(6) + 0.03 * z_scale_average_group(7) + (-0.03) * z_scale_average_group(8))
    normalized_scale_average_group(9) = 50 + (10 * z_scale_average_group(9))
    
    'MCS equal ewights 0.25(VT+SF+RE+MH)
    'z_scale_average_group(10) = (z_scale_average_group(5) + z_scale_average_group(6) + z_scale_average_group(7) + z_scale_average_group(8)) / 4
    ' PF + RP + BP + GH + VT + SF + RE + MH
    'MCS -0,02   0,03    0,04    0,10    0,29    0,14    0,20    0,35
    z_scale_average_group(10) = ((-0.02) * z_scale_average_group(1) + 0.03 * z_scale_average_group(2) + 0.04 * z_scale_average_group(3) + 0.1 * z_scale_average_group(4) + 0.29 * z_scale_average_group(5) + 0.14 * z_scale_average_group(6) + 0.2 * z_scale_average_group(7) + 0.35 * z_scale_average_group(8))
    normalized_scale_average_group(10) = 50 + (10 * z_scale_average_group(10))
    
    'GHC equal weights for all scales 0.125(SUM OF 8 SCALES)
    z_scale_average_group(11) = (z_scale_average_group(1) + z_scale_average_group(2) + z_scale_average_group(3) + z_scale_average_group(4) + z_scale_average_group(5) + z_scale_average_group(6) + z_scale_average_group(7) + z_scale_average_group(8)) / 8
    normalized_scale_average_group(11) = 50 + (10 * z_scale_average_group(11))
      
    
    'compare to normative data
    Dim copyValues As myDataTypeByRow
    'Public Type myDataTypeByRow
    'd_mean As Double
    'd_SD As Double
    'd_N As Double
    'd_group_ix As Integer
    'd_group As String
    'd_group_name As String
    'd_gender As String
    'd_gender_code As String
    'd_age_group As String
    'End Type

    summary_values_ix = 22
    For ix = 1 To 11 '10 '11
        If ix = 11 Then
            copyValues.d_group = "GHC"
            copyValues.d_mean = 50
            copyValues.d_SD = 25
            copyValues.d_group_name = "General Health Composite"
           Else
            copyValues.d_group = myDataArrayByRow(colIx, ix).d_group
            copyValues.d_mean = myDataArrayByRow(colIx, ix).d_mean
            copyValues.d_SD = myDataArrayByRow(colIx, ix).d_SD
            copyValues.d_group_name = myDataArrayByRow(colIx, ix).d_group_name
            copyValues.d_age_group = myDataArrayByRow(colIx, ix).d_age_group
        End If
        i = i + 1
        SurveyData(i, 6) = copyValues.d_group
        SurveyData(i, 7) = z_scale_average_group(ix)
        SurveyData(i, 8) = normalized_scale_average_group(ix)
        SurveyData(i, 9) = 50
        SurveyData(i, 10) = new_sorted_scale_average_group(ix) 'scale_average_group(ix)
        SurveyData(i, 11) = copyValues.d_mean
        SurveyData(i, 12) = copyValues.d_SD
        SurveyData(i, 13) = copyValues.d_group_name
        SurveyData(i, 14) = copyValues.d_age_group
        summary_extra_values(3) = Str(copyValues.d_age_group)
        Select Case ix
            Case Is = 1
                sortIX = 1
            Case Is = 2
                sortIX = 2
            Case Is = 3
                sortIX = 3 '7
            Case Is = 4
                sortIX = 4 '5
            Case Is = 5
                sortIX = 5 '8
            Case Is = 6
                sortIX = 6
            Case Is = 7
                sortIX = 7 '3
            Case Is = 8
                sortIX = 8 '4
            Case Is > 8
                sortIX = ix
        End Select
        SurveyData(i, 15) = sortIX
        'change from scale_average_group to new_sorted_scale_average_group
        If IsNumeric(normalized_scale_average_group(ix)) Then
            summary_values(summary_values_ix + sortIX) = normalized_scale_average_group(ix)
        Else
            summary_values(summary_values_ix + sortIX) = 0
        End If
        If IsNumeric(new_sorted_scale_average_group(ix)) Then
            summary_values(summary_values_ix + sortIX + 11) = new_sorted_scale_average_group(ix)
        Else
            summary_values(summary_values_ix + sortIX + 11) = 0
        End If
        'If IsNumeric(myDataArrayByRow(colIx, ix).d_mean) Then
        If IsNumeric(copyValues.d_mean) Then
            summary_values(summary_values_ix + sortIX + 22) = copyValues.d_mean ' myDataArrayByRow(colIx, ix).d_mean
        Else
            summary_values(summary_values_ix + sortIX + 22) = 0
        End If
        'If IsNumeric(myDataArrayByRow(colIx, ix).d_SD) Then
        If IsNumeric(copyValues.d_SD) Then
            summary_values(summary_values_ix + sortIX + 33) = copyValues.d_SD ' myDataArrayByRow(colIx, ix).d_SD
        Else
            summary_values(summary_values_ix + sortIX + 33) = 0
        End If
        If IsNumeric(z_scale_average_group(ix)) Then
            summary_values(summary_values_ix + sortIX + 44) = z_scale_average_group(ix) 'add z value to summary sheet
        Else
            summary_values(summary_values_ix + sortIX + 44) = 0
        End If
        
    Next ix
Exit Sub
Errhandler:
      Me.MousePointer = fmMousePointerDefault
      ErrorHandling "frmNewSurvey. Sub FetchUserControlValuesBeforeSave", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub keepUserControlValuesFromPage0()
On Error GoTo Errhandler
    'userControlValuesPage0(0) = Format(DTPickerSurveyDate.Value, "yyyy-mm-dd")
    If surveyDate = "" Then
        surveyDate = Now
    End If
    userControlValuesPage0(0) = Format(surveyDate, "yyyy-mm-dd")
    'MsgBox "keepUserControlValuesFromPage0 userControlValuesPage0(0) " & userControlValuesPage0(0)
    userControlValuesPage0(1) = SliderSurveyGroup.Value '0=none,1=retrospective,2=prospective
    userControlValuesPage0(2) = TextBoxFormNumber.Text 'default 1
    userControlValuesPage0(3) = SliderSurveyLanguage.Value '0=english,1=norwegian,2=german,3=dutch
    userControlValuesPage0(4) = TextBoxWeeksAfterBaseline.Text 'default =0
    userControlValuesPage0(5) = Left(TextBoxNotes.Text, 32766) 'default =""
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub keepUserControlValuesFromPage0", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub oldPage0()
On Error GoTo Errhandler
    Dim rowPos As Integer
    Dim colPos As Integer
    
    colPos = 8
    rowPos = 0
    'DTPickerSurveyDate.Value = myOldSurveyData(rowPos + 1, colPos)
    'surveyDate = DTPickerSurveyDate.Value
    'Me.TextBoxSurveyDate.Text = myOldSurveyData(rowPos + 1, colPos)
    'surveyDate = Me.TextBoxSurveyDate.Text
    'MsgBox "oldPage0 DTPickerSurveyDate.Value " & DTPickerSurveyDate.Value & " " & surveyDate & " " & myOldSurveyData(rowPos + 1, colPos)
    TextBoxSurveyDate_Init 'Change new servey date to now
    'TextBoxSurveyDate_Exit
    SliderSurveyGroup.Value = myOldSurveyData(rowPos + 2, colPos) '0=none,1=retrospective,2=prospective
    TextBoxFormNumber.Text = myOldSurveyData(rowPos + 3, colPos) 'default 1
    SliderSurveyLanguage.Value = myOldSurveyData(rowPos + 4, colPos) '0=english,1=norwegian,2=german,3=dutch
    TextBoxWeeksAfterBaseline.Text = myOldSurveyData(rowPos + 5, colPos) 'default =0
    TextBoxNotes.Text = myOldSurveyData(rowPos + 6, colPos) 'default =""
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub oldPage0", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub keepUserControlValuesFromPage1()
On Error GoTo Errhandler
    userTextboxValuesPage1(0) = TextBoxHealthProblem1.Text   'default =""
    userTextboxValuesPage1(1) = TextBoxHealthProblem2.Text   'default =""
    userTextboxValuesPage1(2) = TextBoxHealthProblem3.Text   'default =""
    userTextboxValuesPage1(3) = TextBoxHealthProblem4.Text   'default =""
    userTextboxValuesPage1(4) = TextBoxHealthProblem5.Text   'default =""
    userTextboxValuesPage1(5) = TextBoxHealthProblem6.Text   'default =""
    userTextboxValuesPage1(6) = TextBoxHealthProblem7.Text   'default =""
    userTextboxValuesPage1(7) = TextBoxHealthProblem8.Text   'default =""
    userTextboxValuesPage1(8) = TextBoxHealthProblem9.Text   'default =""
    userTextboxValuesPage1(9) = TextBoxHealthProblem10.Text   'default =""
    userTextboxValuesPage1(10) = TextBoxHealthProblem11.Value   'default =0
    userTextboxValuesPage1(11) = TextBoxHealthProblem12.Text   'default =""
    userTextboxValuesPage1(12) = TextBoxHealthProblem13.Text   'default =""
    userTextboxValuesPage1(13) = TextBoxHealthProblem14.Text   'default =""
    userTextboxValuesPage1(14) = TextBoxHealthProblem15.Text   'default =""
    userTextboxValuesPage1(15) = TextBoxHealthProblem16.Text   'default =""
    userTextboxValuesPage1(16) = TextBoxHealthProblem17.Text   'default =""
    userTextboxValuesPage1(17) = TextBoxHealthProblem18.Text   'default =""
    userTextboxValuesPage1(18) = TextBoxHealthProblem19.Text   'default =""
    userTextboxValuesPage1(19) = TextBoxHealthProblem20.Text   'default =""
     
    userControlValuesPage1(0) = SliderHealthProblem1.Value   'default =""
    userControlValuesPage1(1) = SliderHealthProblem2.Value   'default =""
    userControlValuesPage1(2) = SliderHealthProblem3.Value   'default =""
    userControlValuesPage1(3) = SliderHealthProblem4.Value   'default =""
    userControlValuesPage1(4) = SliderHealthProblem5.Value   'default =""
    userControlValuesPage1(5) = SliderHealthProblem6.Value   'default =""
    userControlValuesPage1(6) = SliderHealthProblem7.Value   'default =""
    userControlValuesPage1(7) = SliderHealthProblem8.Value   'default =""
    userControlValuesPage1(8) = SliderHealthProblem9.Value   'default =""
    userControlValuesPage1(9) = SliderHealthProblem10.Value   'default =""
    userControlValuesPage1(10) = SliderHealthProblem11.Value   'default =0
    userControlValuesPage1(11) = SliderHealthProblem12.Value   'default =0
    userControlValuesPage1(12) = SliderHealthProblem13.Value   'default =0
    userControlValuesPage1(13) = SliderHealthProblem14.Value   'default =0
    userControlValuesPage1(14) = SliderHealthProblem15.Value   'default =0
    userControlValuesPage1(15) = SliderHealthProblem16.Value   'default =0
    userControlValuesPage1(16) = SliderHealthProblem17.Value   'default =0
    userControlValuesPage1(17) = SliderHealthProblem18.Value   'default =0
    userControlValuesPage1(18) = SliderHealthProblem19.Value   'default =0
    userControlValuesPage1(19) = SliderHealthProblem20.Value   'default =0
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub keepUserControlValuesFromPage1", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub oldPage1()
On Error GoTo Errhandler
    Dim rowPos As Integer
    Dim colPos As Integer
    
    colPos = 8
    rowPos = 6
    TextBoxHealthProblem1.Text = myOldSurveyData(rowPos + 1, colPos + 1)
    TextBoxHealthProblem2.Text = myOldSurveyData(rowPos + 2, colPos + 1)
    TextBoxHealthProblem3.Text = myOldSurveyData(rowPos + 3, colPos + 1)
    TextBoxHealthProblem4.Text = myOldSurveyData(rowPos + 4, colPos + 1)
    TextBoxHealthProblem5.Text = myOldSurveyData(rowPos + 5, colPos + 1)
    TextBoxHealthProblem6.Text = myOldSurveyData(rowPos + 6, colPos + 1)
    TextBoxHealthProblem7.Text = myOldSurveyData(rowPos + 7, colPos + 1)
    TextBoxHealthProblem8.Text = myOldSurveyData(rowPos + 8, colPos + 1)
    TextBoxHealthProblem9.Text = myOldSurveyData(rowPos + 9, colPos + 1)
    TextBoxHealthProblem10.Text = myOldSurveyData(rowPos + 10, colPos + 1)
    TextBoxHealthProblem11.Value = myOldSurveyData(rowPos + 11, colPos + 1)
    TextBoxHealthProblem12.Text = myOldSurveyData(rowPos + 12, colPos + 1)
    TextBoxHealthProblem13.Text = myOldSurveyData(rowPos + 13, colPos + 1)
    TextBoxHealthProblem14.Text = myOldSurveyData(rowPos + 14, colPos + 1)
    TextBoxHealthProblem15.Text = myOldSurveyData(rowPos + 15, colPos + 1)
    TextBoxHealthProblem16.Text = myOldSurveyData(rowPos + 16, colPos + 1)
    TextBoxHealthProblem17.Text = myOldSurveyData(rowPos + 17, colPos + 1)
    TextBoxHealthProblem18.Text = myOldSurveyData(rowPos + 18, colPos + 1)
    TextBoxHealthProblem19.Text = myOldSurveyData(rowPos + 19, colPos + 1)
    TextBoxHealthProblem20.Text = myOldSurveyData(rowPos + 20, colPos + 1)
     
    SliderHealthProblem1.Value = myOldSurveyData(rowPos + 1, colPos)
    SliderHealthProblem2.Value = myOldSurveyData(rowPos + 2, colPos)
    SliderHealthProblem3.Value = myOldSurveyData(rowPos + 3, colPos)
    SliderHealthProblem4.Value = myOldSurveyData(rowPos + 4, colPos)
    SliderHealthProblem5.Value = myOldSurveyData(rowPos + 5, colPos)
    SliderHealthProblem6.Value = myOldSurveyData(rowPos + 6, colPos)
    SliderHealthProblem7.Value = myOldSurveyData(rowPos + 7, colPos)
    SliderHealthProblem8.Value = myOldSurveyData(rowPos + 8, colPos)
    SliderHealthProblem9.Value = myOldSurveyData(rowPos + 9, colPos)
    SliderHealthProblem10.Value = myOldSurveyData(rowPos + 10, colPos)
    SliderHealthProblem11.Value = myOldSurveyData(rowPos + 11, colPos)
    SliderHealthProblem12.Value = myOldSurveyData(rowPos + 12, colPos)
    SliderHealthProblem13.Value = myOldSurveyData(rowPos + 13, colPos)
    SliderHealthProblem14.Value = myOldSurveyData(rowPos + 14, colPos)
    SliderHealthProblem15.Value = myOldSurveyData(rowPos + 15, colPos)
    SliderHealthProblem16.Value = myOldSurveyData(rowPos + 16, colPos)
    SliderHealthProblem17.Value = myOldSurveyData(rowPos + 17, colPos)
    SliderHealthProblem18.Value = myOldSurveyData(rowPos + 18, colPos)
    SliderHealthProblem19.Value = myOldSurveyData(rowPos + 19, colPos)
    SliderHealthProblem20.Value = myOldSurveyData(rowPos + 20, colPos)
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub oldPage1", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub keepUserControlValuesFromPage2()
On Error GoTo Errhandler
    userControlValuesPage2(0) = SliderPain.Value   'default =0
    userControlValuesPage2(1) = SliderEnergy.Value   'default =0
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub keepUserControlValuesFromPage2", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub oldPage2()
On Error GoTo Errhandler
    Dim rowPos As Integer
    Dim colPos As Integer
    
    colPos = 8
    rowPos = 26
    SliderPain.Value = myOldSurveyData(rowPos + 1, colPos)
    SliderEnergy.Value = myOldSurveyData(rowPos + 2, colPos)
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub oldPage2", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub keepUserControlValuesFromPage3()
On Error GoTo Errhandler
    Dim sliderValue As Integer
    Dim calculatedValue As Integer
    Dim answerText As String
    
'Dim sf_36_questions(36, 4) As Variant 'question number, response, calculated value, question text, answer text

    sf_36_questions(1, 1) = "Q1"
    sf_36_questions(1, 2) = SliderQ1.Value 'default =0
    sliderValue = SliderQ1.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelq1_1.Caption
        calculatedValue = 100
    ElseIf sliderValue = 2 Then
        answerText = Labelq1_2.Caption
        calculatedValue = 75
    ElseIf sliderValue = 3 Then
        answerText = Labelq1_3.Caption
        calculatedValue = 50
    ElseIf sliderValue = 4 Then
        answerText = Labelq1_4.Caption
        calculatedValue = 25
    ElseIf sliderValue = 5 Then
        answerText = Labelq1_5.Caption
        calculatedValue = 0
    End If
    sf_36_questions(1, 3) = calculatedValue
    sf_36_questions(1, 4) = LabelQ1.Caption
    sf_36_questions(1, 5) = answerText

    sf_36_questions(2, 1) = "Q2"
    sf_36_questions(2, 2) = SliderQ2.Value 'default =0
    sliderValue = SliderQ2.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelQ2_1.Caption
        calculatedValue = 100
    ElseIf sliderValue = 2 Then
        answerText = LabelQ2_2.Caption
        calculatedValue = 75
    ElseIf sliderValue = 3 Then
        answerText = LabelQ2_3.Caption
        calculatedValue = 50
    ElseIf sliderValue = 4 Then
        answerText = LabelQ2_4.Caption
        calculatedValue = 25
    ElseIf sliderValue = 5 Then
        answerText = LabelQ2_5.Caption
        calculatedValue = 0
    End If
    sf_36_questions(2, 3) = calculatedValue
    sf_36_questions(2, 4) = LabelQ2.Caption
    sf_36_questions(2, 5) = answerText
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub keepUserControlValuesFromPage3", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub oldPage3()
On Error GoTo Errhandler
    Dim rowPos As Integer
    Dim colPos As Integer
    
    colPos = 8
    rowPos = 28
    SliderQ1.Value = myOldSurveyData(rowPos + 1, colPos)
    SliderQ2.Value = myOldSurveyData(rowPos + 2, colPos)

Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub oldPage3", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub keepUserControlValuesFromPage4()
On Error GoTo Errhandler
Dim sliderValue As Integer
Dim answerText As String
Dim calculatedValue As Double
    sf_36_questions(3, 1) = "Q3"
    sf_36_questions(3, 2) = SliderQ3.Value 'default =0
    sliderValue = SliderQ3.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelPage4_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = LabelPage4_2.Caption
        calculatedValue = 50
    ElseIf sliderValue = 3 Then
        answerText = LabelPage4_3.Caption
        calculatedValue = 100
    End If
    sf_36_questions(3, 3) = calculatedValue
    sf_36_questions(3, 4) = Labelqu3.Caption
    sf_36_questions(3, 5) = answerText
    
    sf_36_questions(4, 1) = "Q4"
    sf_36_questions(4, 2) = SliderQ4.Value 'default =0
    sliderValue = SliderQ4.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelPage4_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = LabelPage4_2.Caption
        calculatedValue = 50
    ElseIf sliderValue = 3 Then
        answerText = LabelPage4_3.Caption
        calculatedValue = 100
    End If
    sf_36_questions(4, 3) = calculatedValue
    sf_36_questions(4, 4) = Labelqu4.Caption
    sf_36_questions(4, 5) = answerText
    
    sf_36_questions(5, 1) = "Q5"
    sf_36_questions(5, 2) = SliderQ5.Value 'default =0
    sliderValue = SliderQ5.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelPage4_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = LabelPage4_2.Caption
        calculatedValue = 50
    ElseIf sliderValue = 3 Then
        answerText = LabelPage4_3.Caption
        calculatedValue = 100
    End If
    sf_36_questions(5, 3) = calculatedValue
    sf_36_questions(5, 4) = Labelqu3.Caption
    sf_36_questions(5, 5) = answerText
    
    sf_36_questions(6, 1) = "Q6"
    sf_36_questions(6, 2) = SliderQ6.Value 'default =0
    sliderValue = SliderQ6.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelPage4_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = LabelPage4_2.Caption
        calculatedValue = 50
    ElseIf sliderValue = 3 Then
        answerText = LabelPage4_3.Caption
        calculatedValue = 100
    End If
    sf_36_questions(6, 3) = calculatedValue
    sf_36_questions(6, 4) = Labelqu6.Caption
    sf_36_questions(6, 5) = answerText
    
    sf_36_questions(7, 1) = "Q7"
    sf_36_questions(7, 2) = SliderQ7.Value 'default =0
    sliderValue = SliderQ7.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelPage4_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = LabelPage4_2.Caption
        calculatedValue = 50
    ElseIf sliderValue = 3 Then
        answerText = LabelPage4_3.Caption
        calculatedValue = 100
    End If
    sf_36_questions(7, 3) = calculatedValue
    sf_36_questions(7, 4) = Labelqu7.Caption
    sf_36_questions(7, 5) = answerText
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub keepUserControlValuesFromPage4", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub oldPage4()
On Error GoTo Errhandler
    Dim rowPos As Integer
    Dim colPos As Integer
    
    colPos = 8
    rowPos = 30
    SliderQ3.Value = myOldSurveyData(rowPos + 1, colPos)
    SliderQ4.Value = myOldSurveyData(rowPos + 2, colPos)
    SliderQ5.Value = myOldSurveyData(rowPos + 3, colPos)
    SliderQ6.Value = myOldSurveyData(rowPos + 4, colPos)
    SliderQ7.Value = myOldSurveyData(rowPos + 5, colPos)
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub oldPage4", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub keepUserControlValuesFromPage5()
On Error GoTo Errhandler
Dim sliderValue As Integer
Dim answerText As String
Dim calculatedValue As Double
    sf_36_questions(8, 1) = "Q8"
    sf_36_questions(8, 2) = SliderQ8.Value 'default =0
    sliderValue = SliderQ8.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelPage5_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = LabelPage5_2.Caption
        calculatedValue = 50
    ElseIf sliderValue = 3 Then
        answerText = LabelPage5_3.Caption
        calculatedValue = 100
    End If
    sf_36_questions(8, 3) = calculatedValue
    sf_36_questions(8, 4) = LabelQu8.Caption
    sf_36_questions(8, 5) = answerText
    
    sf_36_questions(9, 1) = "Q9"
    sf_36_questions(9, 2) = SliderQ9.Value 'default =0
    sliderValue = SliderQ9.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelPage5_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = LabelPage5_2.Caption
        calculatedValue = 50
    ElseIf sliderValue = 3 Then
        answerText = LabelPage5_3.Caption
        calculatedValue = 100
    End If
    sf_36_questions(9, 3) = calculatedValue
    sf_36_questions(9, 4) = LabelQu9.Caption
    sf_36_questions(9, 5) = answerText
    
    sf_36_questions(10, 1) = "Q10"
    sf_36_questions(10, 2) = SliderQ10.Value 'default =0
    sliderValue = SliderQ10.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelPage5_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = LabelPage5_2.Caption
        calculatedValue = 50
    ElseIf sliderValue = 3 Then
        answerText = LabelPage5_3.Caption
        calculatedValue = 100
    End If
    sf_36_questions(10, 3) = calculatedValue
    sf_36_questions(10, 4) = LabelQu10.Caption
    sf_36_questions(10, 5) = answerText
    
    sf_36_questions(11, 1) = "Q11"
    sf_36_questions(11, 2) = SliderQ11.Value 'default =0
    sliderValue = SliderQ11.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelPage5_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = LabelPage5_2.Caption
        calculatedValue = 50
    ElseIf sliderValue = 3 Then
        answerText = LabelPage5_3.Caption
        calculatedValue = 100
    End If
    sf_36_questions(11, 3) = calculatedValue
    sf_36_questions(11, 4) = LabelQu11.Caption
    sf_36_questions(11, 5) = answerText
    
    sf_36_questions(12, 1) = "Q12"
    sf_36_questions(12, 2) = SliderQ12.Value 'default =0
    sliderValue = SliderQ12.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelPage5_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = LabelPage5_2.Caption
        calculatedValue = 50
    ElseIf sliderValue = 3 Then
        answerText = LabelPage5_3.Caption
        calculatedValue = 100
    End If
    sf_36_questions(12, 3) = calculatedValue
    sf_36_questions(12, 4) = LabelQu12.Caption
    sf_36_questions(12, 5) = answerText
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub keepUserControlValuesFromPage5", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub oldPage5()
On Error GoTo Errhandler
    Dim rowPos As Integer
    Dim colPos As Integer
    
    colPos = 8
    rowPos = 35
    SliderQ8.Value = myOldSurveyData(rowPos + 1, colPos)
    SliderQ9.Value = myOldSurveyData(rowPos + 2, colPos)
    SliderQ10.Value = myOldSurveyData(rowPos + 3, colPos)
    SliderQ11.Value = myOldSurveyData(rowPos + 4, colPos)
    SliderQ12.Value = myOldSurveyData(rowPos + 5, colPos)
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub oldPage5", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub
Private Sub keepUserControlValuesFromPage6()
On Error GoTo Errhandler
Dim sliderValue As Integer
Dim answerText As String
Dim calculatedValue As Double
    sf_36_questions(13, 1) = "Q13"
    sf_36_questions(13, 2) = SliderQ13.Value 'default =0
    sliderValue = SliderQ13.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp6yes1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelp6no1.Caption
        calculatedValue = 100
    End If
    sf_36_questions(13, 3) = calculatedValue
    sf_36_questions(13, 4) = LabelQ13.Caption
    sf_36_questions(13, 5) = answerText
    
    sf_36_questions(14, 1) = "Q14"
    sf_36_questions(14, 2) = Sliderq14.Value 'default =0
    sliderValue = Sliderq14.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp6yes1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelp6no1.Caption
        calculatedValue = 100
    End If
    sf_36_questions(14, 3) = calculatedValue
    sf_36_questions(14, 4) = LabelQ14.Caption
    sf_36_questions(14, 5) = answerText
    
    sf_36_questions(15, 1) = "Q15"
    sf_36_questions(15, 2) = SliderQ15.Value 'default =0
    sliderValue = SliderQ15.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp6yes1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelp6no1.Caption
        calculatedValue = 100
    End If
    sf_36_questions(15, 3) = calculatedValue
    sf_36_questions(15, 4) = LabelQ15.Caption
    sf_36_questions(15, 5) = answerText
    
    sf_36_questions(16, 1) = "Q16"
    sf_36_questions(16, 2) = SliderQ16.Value 'default =0
    sliderValue = SliderQ16.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp6yes1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelp6no1.Caption
        calculatedValue = 100
    End If
    sf_36_questions(16, 3) = calculatedValue
    sf_36_questions(16, 4) = LabelQ16.Caption
    sf_36_questions(16, 5) = answerText
    
    sf_36_questions(17, 1) = "Q17"
    sf_36_questions(17, 2) = SliderQ17.Value 'default =0
    sliderValue = SliderQ17.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp6yes1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelp6no1.Caption
        calculatedValue = 100
    End If
    sf_36_questions(17, 3) = calculatedValue
    sf_36_questions(17, 4) = LabelQ17.Caption
    sf_36_questions(17, 5) = answerText
    
    sf_36_questions(18, 1) = "Q18"
    sf_36_questions(18, 2) = Sliderq18.Value 'default =0
    sliderValue = Sliderq18.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp6yes1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelp6no1.Caption
        calculatedValue = 100
    End If
    sf_36_questions(18, 3) = calculatedValue
    sf_36_questions(18, 4) = LabelQ18.Caption
    sf_36_questions(18, 5) = answerText
    
    sf_36_questions(19, 1) = "Q19"
    sf_36_questions(19, 2) = SliderQ19.Value 'default =0
    sliderValue = SliderQ19.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp6yes1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelp6no1.Caption
        calculatedValue = 100
    End If
    sf_36_questions(19, 3) = calculatedValue
    sf_36_questions(19, 4) = LabelQ19.Caption
    sf_36_questions(19, 5) = answerText
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub keepUserControlValuesFromPage6", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub oldPage6()
On Error GoTo Errhandler
    Dim rowPos As Integer
    Dim colPos As Integer
    
    colPos = 8
    rowPos = 40
    SliderQ13.Value = myOldSurveyData(rowPos + 1, colPos)
    Sliderq14.Value = myOldSurveyData(rowPos + 2, colPos)
    SliderQ15.Value = myOldSurveyData(rowPos + 3, colPos)
    SliderQ16.Value = myOldSurveyData(rowPos + 4, colPos)
    SliderQ17.Value = myOldSurveyData(rowPos + 5, colPos)
    Sliderq18.Value = myOldSurveyData(rowPos + 6, colPos)
    SliderQ19.Value = myOldSurveyData(rowPos + 7, colPos)
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub oldPage6", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub keepUserControlValuesFromPage7()
On Error GoTo Errhandler
Dim sliderValue As Integer
Dim answerText As String
Dim calculatedValue As Double
    sf_36_questions(20, 1) = "Q20"
    sf_36_questions(20, 2) = SliderQ20.Value 'default =0
    sliderValue = SliderQ20.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelq20_1.Caption
        calculatedValue = 100
    ElseIf sliderValue = 2 Then
        answerText = Labelq20_2.Caption
        calculatedValue = 75
    ElseIf sliderValue = 3 Then
        answerText = Labelq20_3.Caption
        calculatedValue = 50
    ElseIf sliderValue = 4 Then
        answerText = Labelq20_4.Caption
        calculatedValue = 25
    ElseIf sliderValue = 5 Then
        answerText = Labelq20_5.Caption
        calculatedValue = 0
    End If
    sf_36_questions(20, 3) = calculatedValue
    sf_36_questions(20, 4) = TextBoxQ20.Text
    sf_36_questions(20, 5) = answerText
    
    sf_36_questions(21, 1) = "Q21"
    sf_36_questions(21, 2) = SliderQ21.Value 'default =0
    sliderValue = SliderQ21.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelq21_1.Caption
        calculatedValue = 100
    ElseIf sliderValue = 2 Then
        answerText = Labelq21_2.Caption
        calculatedValue = 80
    ElseIf sliderValue = 3 Then
        answerText = Labelq21_3.Caption
        calculatedValue = 60
    ElseIf sliderValue = 4 Then
        answerText = Labelq21_4.Caption
        calculatedValue = 40
    ElseIf sliderValue = 5 Then
        answerText = Labelq21_5.Caption
        calculatedValue = 20
    ElseIf sliderValue = 6 Then
        answerText = Labelq21_6.Caption
        calculatedValue = 0
    End If
    sf_36_questions(21, 3) = calculatedValue
    sf_36_questions(21, 4) = TextBoxQ21.Text
    sf_36_questions(21, 5) = answerText
    
    sf_36_questions(22, 1) = "Q22"
    sf_36_questions(22, 2) = SliderQ22.Value 'default =0
    sliderValue = SliderQ22.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelq22_1.Caption
        calculatedValue = 100
    ElseIf sliderValue = 2 Then
        answerText = Labelq22_2.Caption
        calculatedValue = 75
    ElseIf sliderValue = 3 Then
        answerText = Labelq22_3.Caption
        calculatedValue = 50
    ElseIf sliderValue = 4 Then
        answerText = Labelq22_4.Caption
        calculatedValue = 25
    ElseIf sliderValue = 5 Then
        answerText = Labelq22_5.Caption
        calculatedValue = 0
    End If
    sf_36_questions(22, 3) = calculatedValue
    sf_36_questions(22, 4) = TextBoxQ22.Text
    sf_36_questions(22, 5) = answerText
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub keepUserControlValuesFromPage7", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub oldPage7()
On Error GoTo Errhandler
    Dim rowPos As Integer
    Dim colPos As Integer
    
    colPos = 8
    rowPos = 47
    SliderQ20.Value = myOldSurveyData(rowPos + 1, colPos)
    SliderQ21.Value = myOldSurveyData(rowPos + 2, colPos)
    SliderQ22.Value = myOldSurveyData(rowPos + 3, colPos)
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub oldPage7", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub keepUserControlValuesFromPage8()
On Error GoTo Errhandler
Dim sliderValue As Integer
Dim answerText As String
Dim calculatedValue As Double
    sf_36_questions(23, 1) = "Q23"
    sf_36_questions(23, 2) = Sliderq23.Value 'default =0
    sliderValue = Sliderq23.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelP8_1.Caption
        calculatedValue = 100
    ElseIf sliderValue = 2 Then
        answerText = LabelP8_2.Caption
        calculatedValue = 80
    ElseIf sliderValue = 3 Then
        answerText = LabelP8_3.Caption
        calculatedValue = 60
    ElseIf sliderValue = 4 Then
        answerText = LabelP8_4.Caption
        calculatedValue = 40
    ElseIf sliderValue = 5 Then
        answerText = LabelP8_5.Caption
        calculatedValue = 20
    ElseIf sliderValue = 6 Then
        answerText = LabelP8_6.Caption
        calculatedValue = 0
    End If
    sf_36_questions(23, 3) = calculatedValue
    sf_36_questions(23, 4) = Labelq23.Caption
    sf_36_questions(23, 5) = answerText
    
    sf_36_questions(24, 1) = "Q24"
    sf_36_questions(24, 2) = Sliderq24.Value 'default =0
    sliderValue = Sliderq24.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelP8_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = LabelP8_2.Caption
        calculatedValue = 20
    ElseIf sliderValue = 3 Then
        answerText = LabelP8_3.Caption
        calculatedValue = 40
    ElseIf sliderValue = 4 Then
        answerText = LabelP8_4.Caption
        calculatedValue = 60
    ElseIf sliderValue = 5 Then
        answerText = LabelP8_5.Caption
        calculatedValue = 80
    ElseIf sliderValue = 6 Then
        answerText = LabelP8_6.Caption
        calculatedValue = 100
    End If
    sf_36_questions(24, 3) = calculatedValue
    sf_36_questions(24, 4) = Labelq24.Caption
    sf_36_questions(24, 5) = answerText
    
    sf_36_questions(25, 1) = "Q25"
    sf_36_questions(25, 2) = Sliderq25.Value 'default =0
    sliderValue = Sliderq25.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelP8_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = LabelP8_2.Caption
        calculatedValue = 20
    ElseIf sliderValue = 3 Then
        answerText = LabelP8_3.Caption
        calculatedValue = 40
    ElseIf sliderValue = 4 Then
        answerText = LabelP8_4.Caption
        calculatedValue = 60
    ElseIf sliderValue = 5 Then
        answerText = LabelP8_5.Caption
        calculatedValue = 80
    ElseIf sliderValue = 6 Then
        answerText = LabelP8_6.Caption
        calculatedValue = 100
    End If
    sf_36_questions(25, 3) = calculatedValue
    sf_36_questions(25, 4) = Labelq25.Caption
    sf_36_questions(25, 5) = answerText
    
    
    sf_36_questions(26, 1) = "Q26"
    sf_36_questions(26, 2) = Sliderq26.Value 'default =0
    sliderValue = Sliderq26.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = LabelP8_1.Caption
        calculatedValue = 100
    ElseIf sliderValue = 2 Then
        answerText = LabelP8_2.Caption
        calculatedValue = 80
    ElseIf sliderValue = 3 Then
        answerText = LabelP8_3.Caption
        calculatedValue = 60
    ElseIf sliderValue = 4 Then
        answerText = LabelP8_4.Caption
        calculatedValue = 40
    ElseIf sliderValue = 5 Then
        answerText = LabelP8_5.Caption
        calculatedValue = 20
    ElseIf sliderValue = 6 Then
        answerText = LabelP8_6.Caption
        calculatedValue = 0
    End If
    sf_36_questions(26, 3) = calculatedValue
    sf_36_questions(26, 4) = Labelq26.Caption
    sf_36_questions(26, 5) = answerText
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub keepUserControlValuesFromPage8", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub oldPage8()
On Error GoTo Errhandler
    Dim rowPos As Integer
    Dim colPos As Integer
    
    colPos = 8
    rowPos = 50
    Sliderq23.Value = myOldSurveyData(rowPos + 1, colPos)
    Sliderq24.Value = myOldSurveyData(rowPos + 2, colPos)
    Sliderq25.Value = myOldSurveyData(rowPos + 3, colPos)
    Sliderq26.Value = myOldSurveyData(rowPos + 4, colPos)
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub oldPage8", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub keepUserControlValuesFromPage9()
On Error GoTo Errhandler
Dim sliderValue As Integer
Dim answerText As String
Dim calculatedValue As Double
    sf_36_questions(27, 1) = "Q27"
    sf_36_questions(27, 2) = Sliderq27.Value 'default =0
    sliderValue = Sliderq27.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp9_1.Caption
        calculatedValue = 100
    ElseIf sliderValue = 2 Then
        answerText = Labelp9_2.Caption
        calculatedValue = 80
    ElseIf sliderValue = 3 Then
        answerText = Labelp9_3.Caption
        calculatedValue = 60
    ElseIf sliderValue = 4 Then
        answerText = Labelp9_4.Caption
        calculatedValue = 40
    ElseIf sliderValue = 5 Then
        answerText = Labelp9_5.Caption
        calculatedValue = 20
    ElseIf sliderValue = 6 Then
        answerText = Labelp9_6.Caption
        calculatedValue = 0
    End If
    sf_36_questions(27, 3) = calculatedValue
    sf_36_questions(27, 4) = Labelq27.Caption
    sf_36_questions(27, 5) = answerText
    
    sf_36_questions(28, 1) = "Q28"
    sf_36_questions(28, 2) = Sliderq28.Value 'default =0
    sliderValue = Sliderq28.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp9_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelp9_2.Caption
        calculatedValue = 20
    ElseIf sliderValue = 3 Then
        answerText = Labelp9_3.Caption
        calculatedValue = 40
    ElseIf sliderValue = 4 Then
        answerText = Labelp9_4.Caption
        calculatedValue = 60
    ElseIf sliderValue = 5 Then
        answerText = Labelp9_5.Caption
        calculatedValue = 80
    ElseIf sliderValue = 6 Then
        answerText = Labelp9_6.Caption
        calculatedValue = 100
    End If
    sf_36_questions(28, 3) = calculatedValue
    sf_36_questions(28, 4) = Labelq28.Caption
    sf_36_questions(28, 5) = answerText
    
    sf_36_questions(29, 1) = "Q29"
    sf_36_questions(29, 2) = Sliderq29.Value 'default =0
    sliderValue = Sliderq29.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp9_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelp9_2.Caption
        calculatedValue = 20
    ElseIf sliderValue = 3 Then
        answerText = Labelp9_3.Caption
        calculatedValue = 40
    ElseIf sliderValue = 4 Then
        answerText = Labelp9_4.Caption
        calculatedValue = 60
    ElseIf sliderValue = 5 Then
        answerText = Labelp9_5.Caption
        calculatedValue = 80
    ElseIf sliderValue = 6 Then
        answerText = Labelp9_6.Caption
        calculatedValue = 100
    End If
    sf_36_questions(29, 3) = calculatedValue
    sf_36_questions(29, 4) = Labelq29.Caption
    sf_36_questions(29, 5) = answerText
    
    sf_36_questions(30, 1) = "Q30"
    sf_36_questions(30, 2) = Sliderq30.Value 'default =0
    sliderValue = Sliderq30.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp9_1.Caption
        calculatedValue = 100
    ElseIf sliderValue = 2 Then
        answerText = Labelp9_2.Caption
        calculatedValue = 80
    ElseIf sliderValue = 3 Then
        answerText = Labelp9_3.Caption
        calculatedValue = 60
    ElseIf sliderValue = 4 Then
        answerText = Labelp9_4.Caption
        calculatedValue = 40
    ElseIf sliderValue = 5 Then
        answerText = Labelp9_5.Caption
        calculatedValue = 20
    ElseIf sliderValue = 6 Then
        answerText = Labelp9_6.Caption
        calculatedValue = 0
    End If
    sf_36_questions(30, 3) = calculatedValue
    sf_36_questions(30, 4) = Labelq30.Caption
    sf_36_questions(30, 5) = answerText
    
    sf_36_questions(31, 1) = "Q31"
    sf_36_questions(31, 2) = Sliderq31.Value 'default =0
    sliderValue = Sliderq31.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp9_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelp9_2.Caption
        calculatedValue = 20
    ElseIf sliderValue = 3 Then
        answerText = Labelp9_3.Caption
        calculatedValue = 40
    ElseIf sliderValue = 4 Then
        answerText = Labelp9_4.Caption
        calculatedValue = 60
    ElseIf sliderValue = 5 Then
        answerText = Labelp9_5.Caption
        calculatedValue = 80
    ElseIf sliderValue = 6 Then
        answerText = Labelp9_6.Caption
        calculatedValue = 100
    End If
    sf_36_questions(31, 3) = calculatedValue
    sf_36_questions(31, 4) = Labelq31.Caption
    sf_36_questions(31, 5) = answerText
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub keepUserControlValuesFromPage9", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub oldPage9()
On Error GoTo Errhandler
    Dim rowPos As Integer
    Dim colPos As Integer
    
    colPos = 8
    rowPos = 54
    Sliderq27.Value = myOldSurveyData(rowPos + 1, colPos)
    Sliderq28.Value = myOldSurveyData(rowPos + 2, colPos)
    Sliderq29.Value = myOldSurveyData(rowPos + 3, colPos)
    Sliderq30.Value = myOldSurveyData(rowPos + 4, colPos)
    Sliderq31.Value = myOldSurveyData(rowPos + 5, colPos)
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub oldPage9", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub keepUserControlValuesFromPage10()
On Error GoTo Errhandler
Dim sliderValue As Integer
Dim answerText As String
Dim calculatedValue As Double

    sf_36_questions(32, 1) = "Q32"
    sf_36_questions(32, 2) = Sliderq32.Value 'default =0
    sliderValue = Sliderq32.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelq32_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelq32_2.Caption
        calculatedValue = 25
    ElseIf sliderValue = 3 Then
        answerText = Labelq32_3.Caption
        calculatedValue = 50
    ElseIf sliderValue = 4 Then
        answerText = Labelq32_4.Caption
        calculatedValue = 75
    ElseIf sliderValue = 5 Then
        answerText = Labelq32_5.Caption
        calculatedValue = 100
    End If
    sf_36_questions(32, 3) = calculatedValue
    sf_36_questions(32, 4) = TextBoxPage10Header1.Text
    sf_36_questions(32, 5) = answerText

    sf_36_questions(33, 1) = "Q33"
    sf_36_questions(33, 2) = Sliderq33.Value 'default =0
    sliderValue = Sliderq33.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp10_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelp10_2.Caption
        calculatedValue = 25
    ElseIf sliderValue = 3 Then
        answerText = Labelp10_3.Caption
        calculatedValue = 50
    ElseIf sliderValue = 4 Then
        answerText = Labelp10_4.Caption
        calculatedValue = 75
    ElseIf sliderValue = 5 Then
        answerText = Labelp10_5.Caption
        calculatedValue = 100
    End If
    sf_36_questions(33, 3) = calculatedValue
    sf_36_questions(33, 4) = Labelq33.Caption
    sf_36_questions(33, 5) = answerText
    
    sf_36_questions(34, 1) = "Q34"
    sf_36_questions(34, 2) = Sliderq34.Value 'default =0
    sliderValue = Sliderq34.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp10_1.Caption
        calculatedValue = 100
    ElseIf sliderValue = 2 Then
        answerText = Labelp10_2.Caption
        calculatedValue = 75
    ElseIf sliderValue = 3 Then
        answerText = Labelp10_3.Caption
        calculatedValue = 50
    ElseIf sliderValue = 4 Then
        answerText = Labelp10_4.Caption
        calculatedValue = 25
    ElseIf sliderValue = 5 Then
        answerText = Labelp10_5.Caption
        calculatedValue = 0
    End If
    sf_36_questions(34, 3) = calculatedValue
    sf_36_questions(34, 4) = Labelq34.Caption
    sf_36_questions(34, 5) = answerText
    
    sf_36_questions(35, 1) = "Q35"
    sf_36_questions(35, 2) = Sliderq35.Value 'default =0
    sliderValue = Sliderq35.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp10_1.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelp10_2.Caption
        calculatedValue = 25
    ElseIf sliderValue = 3 Then
        answerText = Labelp10_3.Caption
        calculatedValue = 50
    ElseIf sliderValue = 4 Then
        answerText = Labelp10_4.Caption
        calculatedValue = 75
    ElseIf sliderValue = 5 Then
        answerText = Labelp10_5.Caption
        calculatedValue = 100
    End If
    sf_36_questions(35, 3) = calculatedValue
    sf_36_questions(35, 4) = Labelq35.Caption
    sf_36_questions(35, 5) = answerText
    
    
    sf_36_questions(36, 1) = "Q36"
    sf_36_questions(36, 2) = Sliderq36.Value 'default =0
    sliderValue = Sliderq36.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelp10_1.Caption
        calculatedValue = 100
    ElseIf sliderValue = 2 Then
        answerText = Labelp10_2.Caption
        calculatedValue = 75
    ElseIf sliderValue = 3 Then
        answerText = Labelp10_3.Caption
        calculatedValue = 50
    ElseIf sliderValue = 4 Then
        answerText = Labelp10_4.Caption
        calculatedValue = 25
    ElseIf sliderValue = 5 Then
        answerText = Labelp10_5.Caption
        calculatedValue = 0
    End If
    sf_36_questions(36, 3) = calculatedValue
    sf_36_questions(36, 4) = Labelq36.Caption
    sf_36_questions(36, 5) = answerText
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub keepUserControlValuesFromPage10", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub oldPage10()
On Error GoTo Errhandler
    Dim rowPos As Integer
    Dim colPos As Integer
    
    colPos = 8
    rowPos = 59
    Sliderq32.Value = myOldSurveyData(rowPos + 1, colPos)
    Sliderq33.Value = myOldSurveyData(rowPos + 2, colPos)
    Sliderq34.Value = myOldSurveyData(rowPos + 3, colPos)
    Sliderq35.Value = myOldSurveyData(rowPos + 4, colPos)
    Sliderq36.Value = myOldSurveyData(rowPos + 5, colPos)
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub oldPage10", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub keepUserControlValuesFromPage11()
On Error GoTo Errhandler
Dim sliderValue As Integer
Dim answerText As String
Dim calculatedValue As Double
    sf_36_questions(37, 1) = "Q37"
    sf_36_questions(37, 2) = Sliderq37.Value 'default =0
    sliderValue = Sliderq37.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelq37_no.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelq37_yes.Caption
        calculatedValue = 1
    End If
    sf_36_questions(37, 3) = calculatedValue
    sf_36_questions(37, 4) = Labelq37.Caption
    sf_36_questions(37, 5) = answerText
    userControlValuesPage11(1) = TextBoxq37.Text
    
    sf_36_questions(38, 1) = "Q38"
    sf_36_questions(38, 2) = Sliderq38.Value 'default =0
    sliderValue = Sliderq38.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelq38_no.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelq38_yes.Caption
        calculatedValue = 1
    End If
    sf_36_questions(38, 3) = calculatedValue
    sf_36_questions(38, 4) = Labelq38.Caption
    sf_36_questions(38, 5) = answerText
    userControlValuesPage11(2) = TextBoxq38.Text
    
    sf_36_questions(39, 1) = "Q39"
    sf_36_questions(39, 2) = Sliderq39.Value 'default =0
    sliderValue = Sliderq39.Value
    If sliderValue = 0 Then
        answerText = "Missing"
        calculatedValue = 0
    ElseIf sliderValue = 1 Then
        answerText = Labelq39_no.Caption
        calculatedValue = 0
    ElseIf sliderValue = 2 Then
        answerText = Labelq39_yes.Caption
        calculatedValue = 1
    End If
    sf_36_questions(39, 3) = calculatedValue
    sf_36_questions(39, 4) = Labelq39.Caption
    sf_36_questions(39, 5) = answerText
    userControlValuesPage11(3) = TextBoxq39.Text
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub keepUserControlValuesFromPage11", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub oldPage11()
On Error GoTo Errhandler
    Dim rowPos As Integer
    Dim colPos As Integer
    
    colPos = 8
    rowPos = 64
    Sliderq37.Value = myOldSurveyData(rowPos + 1, colPos)
    Sliderq38.Value = myOldSurveyData(rowPos + 2, colPos)
    Sliderq39.Value = myOldSurveyData(rowPos + 3, colPos)
    TextBoxq37.Text = myOldSurveyData(rowPos + 1, colPos + 4)
    TextBoxq38.Text = myOldSurveyData(rowPos + 2, colPos + 4)
    TextBoxq39.Text = myOldSurveyData(rowPos + 3, colPos + 4)
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub oldPage11", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub



Private Sub init_page0()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        LabelHeaderPage0.Caption = "This survey asks for your views about your health. " & vbCrLf & _
            "This information will help you keep track of how you feel " & vbCrLf & _
            "and how well you are able to do your usual activities."
        LabelDate.Caption = "Choose date for answering the survey:"
        Me.FrameResearch.Caption = "Research group info"
        Me.LabelResarchGroupNone = "Not in use" '"No group"
        Me.LabelResearchGroupRetrospective = "Not in use" 'Retrospective group"
        Me.LabelResearchGroupProspective = "Not in use" ' "Prospective group"
        Me.LabelFormNumber.Caption = "Form number:"
        LabelLangEng.Caption = "Not in use" '"English"
        LabelLangNo.Caption = "Not in use" '"Norwegian"
        LabelLangGerman.Caption = "Not in use" '"German"
        LabelLangHolland.Caption = "Not in use" '"Dutch"
        LabelNotes.Caption = "Your notes:"
        LabelWeeksAfterBaseline.Caption = "Months after baseline:"
    Else
        LabelHeaderPage0.Caption = "Denne spørreundersøkelsen handler om hvordan du ser på din egne helse. " & vbCrLf & _
            "Denne informasjonene vil hjelpe deg til å holde oversikt over hvordan du føler deg " & vbCrLf & _
            "og hvor godt du er i stand til å utføre dine vanlige aktiviteter."
        LabelDate.Caption = "Velg dato for utfylling av spørreskjemaet:"
        Me.FrameResearch.Caption = "Informasjon om gruppe"
        Me.LabelResarchGroupNone = "Ikke i bruk" ' "Ingen gruppe"
        Me.LabelResearchGroupRetrospective = "Ikke i bruk" ' "Retrospektiv gruppe"
        Me.LabelResearchGroupProspective = "Ikke i bruk" ' "Prospektiv gruppe"
        Me.LabelFormNumber.Caption = "Spørreskjema nummer:"
        LabelLangEng.Caption = "Ikke i bruk" ' "Engelsk"
        LabelLangNo.Caption = "Ikke i bruk" ' "Norsk"
        LabelLangGerman.Caption = "Ikke i bruk" ' "Tysk"
        LabelLangHolland.Caption = "Ikke i bruk" ' "Nederlandsk"
        LabelLangGerman.Caption = "Ikke i bruk" ' "Tysk"
        LabelNotes.Caption = "Dine notater:"
        LabelWeeksAfterBaseline.Caption = "Måneder etter baseline:"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub init_page0", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub init_page1()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        TextBoxPage1Header.Text = "Please complete the table below with" & _
            UCase(" the diseases or health problems you are in treatment for") & vbCrLf & _
            " and mark" & _
            UCase(" how the condition is now") & _
            " by selecting a value for each medical condition or ailment."
        LabelBestPossibleCondition.Caption = "Best possible condition, " & vbCrLf & "no problem at all."
        LabelWorstPossibleCondition.Caption = "Worst possible condition, " & vbCrLf & "a really big problem."
        FrameHealthProblems.Caption = "Diseases or health problems:"
        LabelConditionNow.Caption = "Condition NOW:"
    Else
        TextBoxPage1Header.Text = "Vennligst fyll ut tabellen nedenfor" & _
            UCase(" med de sykdommene eller plagene du er til behandling for, ") & vbCrLf & _
            "og marker" & _
            UCase(" hvordan tilstanden er nå") & _
            " ved å velge en verdi for hver sykdom eller plage."
        LabelBestPossibleCondition.Caption = "Best mulige tilstand, " & vbCrLf & "ikke noe problem i det hele tatt."
        LabelWorstPossibleCondition.Caption = "Verst tenkelige tilstand, " & vbCrLf & "et veldig stort problem."
        FrameHealthProblems.Caption = "Sykdommer eller helseplager:"
        LabelConditionNow.Caption = "Tilstand NÅ:"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub init_page1", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub init_page2()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        TextBoxPage2header.Text = "Please mark in the table below" & _
            UCase(" how severe the pain") & vbCrLf & _
            " you have felt during" & _
            UCase(" the last four weeks") & _
            " by selecting a value from 0 to 10."
        TextBoxPage2header2.Text = "Please mark in the table below how much you have experienced" & _
            UCase(" a lack of energy") & vbCrLf & _
            " during" & _
            UCase(" the last four weeks") & _
            " by selecting a value from 0 to 10."
        FramePain.Caption = "Pain:"
        LabelPainLast4weeks.Caption = "Pain the last 4 weeks:"
        LabelNoPain.Caption = "No pain"
        LabelWorstPossiblePain.Caption = "Worst possible pain"
        FrameLackOfEnergy.Caption = "Lack of energy:"
        LabelLackOfEnergy.Caption = "Lack of energy the last 4 weeks:"
        LabelNoLackOfEnergy.Caption = "No lack of energy"
        LabelWorstPossibleLackOfEnergy.Caption = "Worst possible lack of energy"
    Else
        TextBoxPage2header.Text = "Vennligst marker i tabellen nedenfor" & _
            UCase(" hvor sterke smerter") & vbCrLf & _
            " du har følt i løpet av" & _
            UCase(" de siste fire ukene") & _
            " ved å velge en verdi fra 0 til 10."
        TextBoxPage2header2.Text = "Vennligst marker i tabellen nedenfor hvor mye du har opplevd" & vbCrLf & _
            UCase(" mangel på energi") & _
            " i løpet av" & _
            UCase(" de siste fire ukene") & _
            " ved å velge en verdi fra 0 til 10."
        FramePain.Caption = "Smerte:"
        LabelPainLast4weeks.Caption = "Smerter des siste 4 uker:"
        LabelNoPain.Caption = "Ingen smerter"
        LabelWorstPossiblePain.Caption = "Verst tenkelige smerter"
        FrameLackOfEnergy.Caption = "Mangel på energi:"
        LabelLackOfEnergy.Caption = "Mangel på energi de siste 4 uker:"
        LabelNoLackOfEnergy.Caption = "Ingen mangel på energi"
        LabelWorstPossibleLackOfEnergy.Caption = "Verst tenkelige mangel på energi"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub init_page2", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub init_page3()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        TextBoxPage3header1.Text = "The following questions are about how you feel about your own health. " & _
            vbCrLf & _
            "Each question should be answered by selecting the answer that best suits you." & _
            vbCrLf & _
            "If you are unsure how to answer, please answer as best you can."
        LabelQ1.Caption = "In general, would you say your health is:"
        Labelq1_1.Caption = "Excellent"
        Labelq1_2.Caption = "Very good"
        Labelq1_3.Caption = "Good"
        Labelq1_4.Caption = "Fair"
        Labelq1_5.Caption = "Poor"
        LabelQ2.Caption = UCase("Compared to 1 year ago,") & _
            " how would you rate your health in general" & _
            UCase(" now?")
        LabelQ2_1.Caption = "Much better now than 1 year ago"
        LabelQ2_2.Caption = "Somewhat better now than 1 year ago"
        LabelQ2_3.Caption = "About the same as 1 year ago"
        LabelQ2_4.Caption = "Somewhat worse now than 1 year ago"
        LabelQ2_5.Caption = "Much worse now than 1 year ago"
    Else
        TextBoxPage3header1.Text = "De følgende spørsmålene handler om hvordan du ser på din egen helse." & _
            vbCrLf & _
            "Hvert spørsmål skal besvares ved å velge det svaret som passer best for deg." & _
            vbCrLf & _
            "Hvis du er usikker på hva du skal svare, vennligst svar så godt du kan."
        LabelQ1.Caption = "Stort sett, hvordan vil du si din helse er:"
        Labelq1_1.Caption = "Utmerket"
        Labelq1_2.Caption = "Meget god"
        Labelq1_3.Caption = "God"
        Labelq1_4.Caption = "Nokså god"
        Labelq1_5.Caption = "Dårlig"
        LabelQ2.Caption = UCase("Sammenliknet med for ett år siden,") & _
            " hvordan vil du si at din helse stort sett er" & _
            UCase(" nå?")
        LabelQ2_1.Caption = "Mye bedre nå enn for ett år siden"
        LabelQ2_2.Caption = "Litt bedre nå enn for ett år siden"
        LabelQ2_3.Caption = "Omtrent den samme som for ett år siden"
        LabelQ2_4.Caption = "Litt dårligere nå enn for ett år siden"
        LabelQ2_5.Caption = "Mye dårligere nå enn for ett år siden"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub init_page3", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub init_page4()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        TextBoxPage4header1.Text = "The following items are about activities you might do during a typical day." & _
            vbCrLf & UCase(" Does your health now limit you") & _
            " in these activities?  If so, how much?"
        LabelPage4_1.Caption = "Yes " & vbCrLf & "Limited " & vbCrLf & "a lot"
        LabelPage4_2.Caption = "Yes " & vbCrLf & "Limited " & vbCrLf & "a little"
        LabelPage4_3.Caption = "No " & vbCrLf & "not limited " & vbCrLf & "at all"
        Labelqu3.Caption = UCase("Vigorous activities,") & _
            " such as running, lifting heavy objects, participating in strenuous sports"
        Labelqu4.Caption = UCase("Moderate activities,") & _
            " such as moving a table, pushing a vacuum cleaner, bowling or playing golf"
        Labelqu5.Caption = "Lifting or carrying groceries"
        Labelqu6.Caption = "Climbing " & UCase("several") & " flights of stairs"
        Labelqu7.Caption = "Climbing " & UCase("one") & " flight of stairs"
    Else
        TextBoxPage4header1.Text = "De neste spørsmålene handler om aktiviteter som du kanskje utfører i løpet av en vanlig dag." & _
            vbCrLf & UCase("Er din helse slik at den begrenser deg") & _
            " i utførelsen av disse aktivitetene NÅ? Hvis ja, hvor mye?"
        LabelPage4_1.Caption = "Ja, " & vbCrLf & "begrenser " & vbCrLf & "meg mye"
        LabelPage4_2.Caption = "Ja, " & vbCrLf & "begrenser " & vbCrLf & "meg litt"
        LabelPage4_3.Caption = "Nei," & vbCrLf & "begrenser meg ikke " & vbCrLf & "i det hele tatt"
        Labelqu3.Caption = UCase("Anstrengende aktiviteter,") & _
            " som å løpe, løfte tunge gjenstander, delta i anstrengende idrett"
        Labelqu4.Caption = UCase("Moderate aktiviteter,") & _
            " som å flytte et bord, støvsuge, gå en tur eller drive med hagearbeid"
        Labelqu5.Caption = "Løfte eller bære en handlekurv"
        Labelqu6.Caption = "Gå opp trappen " & UCase("flere") & " etasjer"
        Labelqu7.Caption = "Gå opp trappen " & UCase("en") & " etasje"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub init_page4", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub init_page5()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        TextBoxPage5header1.Text = "The following items are about activities you might do during a typical day." & _
            vbCrLf & UCase(" Does your health now limit you") & _
            " in these activities?  If so, how much?"
        LabelPage5_1.Caption = "Yes " & vbCrLf & "Limited " & vbCrLf & "a lot"
        LabelPage5_2.Caption = "Yes " & vbCrLf & "Limited " & vbCrLf & "a little"
        LabelPage5_3.Caption = "No " & vbCrLf & "not limited " & vbCrLf & "at all"
        LabelQu8.Caption = "Bending , kneeling Or stooping"
        LabelQu9.Caption = "Walking" & UCase(" more than a mile")
        LabelQu10.Caption = "Walking" & UCase(" several blocks")
        LabelQu11.Caption = "Walking" & UCase(" one block")
        LabelQu12.Caption = "Bathing or dressing yourself"
    Else
        TextBoxPage5header1.Text = "De neste spørsmålene handler om aktiviteter som du kanskje utfører i løpet av en vanlig dag." & _
            vbCrLf & UCase("Er din helse slik at den begrenser deg") & _
            " i utførelsen av disse aktivitetene NÅ? Hvis ja, hvor mye?"
        LabelPage5_1.Caption = "Ja, " & vbCrLf & "begrenser " & vbCrLf & "meg mye"
        LabelPage5_2.Caption = "Ja, " & vbCrLf & "begrenser " & vbCrLf & "meg litt"
        LabelPage5_3.Caption = "Nei," & vbCrLf & "begrenser meg ikke " & vbCrLf & "i det hele tatt"
        LabelQu8.Caption = "Bøye deg eller sitte på huk"
        LabelQu9.Caption = "Gå" & UCase(" mer enn to kilometer")
        LabelQu10.Caption = "Gå" & UCase(" noen hundre meter")
        LabelQu11.Caption = "Gå" & UCase(" hundre meter")
        LabelQu12.Caption = "Vaske deg eller kle på deg"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub init_page5", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub init_page6()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        TextBoxPage6Header1.Text = "During the " & UCase("past 4 weeks,") & _
            " have you had any of the following problems with your work or other regular daily activities" & _
            vbCrLf & UCase(" as a result of your physical health?")
        Labelp6yes1.Caption = "YES"
        Labelp6no1.Caption = "NO"
        Labelp6yes2.Caption = "YES"
        Labelp6no2.Caption = "NO"
        LabelQ13.Caption = "Cut down the " & UCase("amount of time") & " you spend on work or other activities"
        LabelQ14.Caption = UCase("Accomplished less") & " than you would like"
        LabelQ15.Caption = "Were limited in the " & UCase("kind of") & " work or other activities"
        LabelQ16.Caption = "Had " & UCase("difficulty") & " performing the work or other activities" & _
            " (for example it took extra effort)"
        TextBoxPage6Header2.Text = "During the " & UCase("past 4 weeks,") & _
            " have you had any of the following problems with your work or other regular daily activities" & _
            vbCrLf & UCase(" as a result of any emotional problems") & _
            " (such as feeling depressed or anxious)?"
        LabelQ17.Caption = "Cut down the " & UCase("amount of time") & " you spend on work or other activities"
        LabelQ18.Caption = UCase("Accomplished less") & " than you would like"
        LabelQ19.Caption = "Didn't do work or other activities as " & UCase("carefully") & " as usual"
    Else
        TextBoxPage6Header1.Text = "I løpet av " & UCase("de siste 4 ukene,") & _
            " har du hatt noen av de  følgende problemer i ditt arbeid eller i andre av  dine daglige gjøremål" & _
            vbCrLf & UCase(" på grunn av din fysiske helse?")
        Labelp6yes1.Caption = "JA"
        Labelp6no1.Caption = "NEI"
        Labelp6yes2.Caption = "JA"
        Labelp6no2.Caption = "NEI"
        LabelQ13.Caption = "Du har måttet " & UCase("redusere tiden") & " du har brukt på arbeid eller på andre gjøremål"
        LabelQ14.Caption = "Du har " & UCase("utrettet mindre") & " enn du hadde ønsket"
        LabelQ15.Caption = "Du har vært hindret i å utføre " & UCase("visse typer") & " arbeid eller gjøremål"
        LabelQ16.Caption = "Du har hatt  " & UCase("problemer") & " med å gjennomføre arbeidet eller andre gjøremål" & _
            " (for eksempel fordi det krevde ekstra anstrengelser)"
        TextBoxPage6Header2.Text = "I løpet av " & UCase("de siste 4 ukene,") & _
            " har du hatt noen av de  følgende problemer i ditt arbeid eller i andre av  dine daglige gjøremål" & _
            vbCrLf & UCase(" på grunn av følelsesmessige problemer") & _
            vbCrLf & " (som for eksempel å være deprimert eller engstelig)?"
        LabelQ17.Caption = "Du har måttet " & UCase("redusere tiden") & " du har brukt på arbeid eller på andre gjøremål"
        LabelQ18.Caption = "Du har " & UCase("utrettet mindre") & " enn du hadde ønsket"
        LabelQ19.Caption = "Du har utført arbeidet eller andre gjøremål " & UCase("mindre grundig") & " enn vanlig"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub init_page6", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub init_page7()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        TextBoxQ20.Text = "During the " & UCase("past 4 weeks") & ", to what extent has your physical health or emotional problems interfered with your normal social activities with family, friends, neighbors or groups?"
        Labelq20_1.Caption = "Not at all"
        Labelq20_2.Caption = "Slightly"
        Labelq20_3.Caption = "Moderately"
        Labelq20_4.Caption = "Quite a bit"
        Labelq20_5.Caption = "Extremely"
        
        TextBoxQ21.Text = "How much " & UCase("bodily") & " pain have you had in the " & UCase("past 4 weeks?")
        Labelq21_1.Caption = "None"
        Labelq21_2.Caption = "Very Mild"
        Labelq21_3.Caption = "Mild"
        Labelq21_4.Caption = "Moderate"
        Labelq21_5.Caption = "Severe"
        Labelq21_6.Caption = "Very severe"
 
        TextBoxQ22.Text = "During the " & UCase("past 4 weeks,") & " how much did " & UCase("pain") & _
            " interfere with your normal work (Including work outside the house " & UCase("and") & " housework)"
        Labelq22_1.Caption = "Not at all"
        Labelq22_2.Caption = "Slightly"
        Labelq22_3.Caption = "Moderately"
        Labelq22_4.Caption = "Quite a bit"
        Labelq22_5.Caption = "Extremely"
    Else
        TextBoxQ20.Text = "I løpet av " & UCase("de siste 4 ukene,") & _
            " i hvilken grad har din fysiske helse eller følelsesmessige problemer hatt innvirkning på din vanlige sosiale omgang med familie, venner, naboer eller foreninger?"
        Labelq20_1.Caption = "Ikke i det hele tatt"
        Labelq20_2.Caption = "Litt"
        Labelq20_3.Caption = "En del"
        Labelq20_4.Caption = "Mye"
        Labelq20_5.Caption = "Svært mye"
        
        TextBoxQ21.Text = "Hvor sterke " & UCase("kroppslige") & " smerter har du hatt i løpet av " & _
            UCase("de siste 4 ukene?")
        Labelq21_1.Caption = "Ingen"
        Labelq21_2.Caption = "Meget svake"
        Labelq21_3.Caption = "Svake"
        Labelq21_4.Caption = "Moderate"
        Labelq21_5.Caption = "Sterke"
        Labelq21_6.Caption = "Meget sterke"
        
        TextBoxQ22.Text = "I løpet av " & UCase("de siste 4 ukene,") & _
            " hvor mye har " & UCase("smerter") & _
            " påvirket ditt vanlige arbeid (gjelder både arbeid utenfor hjemmet " & UCase("og") & " husarbeid)?"
        Labelq22_1.Caption = "Ikke i det hele tatt"
        Labelq22_2.Caption = "Litt"
        Labelq22_3.Caption = "En del"
        Labelq22_4.Caption = "Mye"
        Labelq22_5.Caption = "Svært mye"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub init_page7", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub init_page8()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        TextBoxPage8Header1.Text = "These questions are about how you feel and how things have been with you " & _
            UCase("during the last 4 weeks.") & vbCrLf & _
            "For each question, please give the 1 answer that comes closest to the way you have been feeling. " & vbCrLf & _
            UCase("How much of the time during the last 4 weeks...")
        LabelP8_1.Caption = "All" & vbCrLf & "of the" & vbCrLf & "time"
        LabelP8_2.Caption = "Most" & vbCrLf & "of the" & vbCrLf & "time"
        LabelP8_3.Caption = "A good bit" & vbCrLf & "of the" & vbCrLf & "time"
        LabelP8_4.Caption = "Some" & vbCrLf & "of the" & vbCrLf & "time"
        LabelP8_5.Caption = "A little" & vbCrLf & "of the" & vbCrLf & "time"
        LabelP8_6.Caption = "None" & vbCrLf & "of the" & vbCrLf & "time"
        Labelq23.Caption = "Did you feel full of pep?"
        Labelq24.Caption = "Have you been a very nervous person?"
        Labelq25.Caption = "Have you felt so down in the dumps that nothing could cheer you up?"
        Labelq26.Caption = "Have you felt calm and peaceful?"
    Else
        TextBoxPage8Header1.Text = "De neste spørsmålene handler om hvordan du har følt deg og hvordan du har hatt det " & _
            UCase("de siste 4 ukene.") & vbCrLf & _
            "For hvert spørsmål, vennligst velg det svaralternativet som best beskriver hvordan du har hatt det. " & vbCrLf & _
            "Hvor ofte i løpet av " & UCase("de siste 4 ukene") & " har du:"
        LabelP8_1.Caption = "Hele" & vbCrLf & "tiden"
        LabelP8_2.Caption = "Nesten hele" & vbCrLf & "tiden"
        LabelP8_3.Caption = "Mye av" & vbCrLf & "tiden"
        LabelP8_4.Caption = "En del av" & vbCrLf & "tiden"
        LabelP8_5.Caption = "Litt av" & vbCrLf & "tiden"
        LabelP8_6.Caption = "Ikke" & vbCrLf & "i det hele tatt"
        Labelq23.Caption = "Følt deg full av tiltakslyst?"
        Labelq24.Caption = "Følt deg veldig nervøs?"
        Labelq25.Caption = "Vært så langt nede at ingenting har kunnet muntre deg opp?"
        Labelq26.Caption = "Følt deg rolig og harmonisk?"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub init_page8", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub init_page9()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        TextBoxPage9Header1.Text = "These questions are about how you feel and how things have been with you " & _
            UCase("during the last 4 weeks.") & vbCrLf & _
            "For each question, please give the 1 answer that comes closest to the way you have been feeling. " & vbCrLf & _
            UCase("How much of the time during the last 4 weeks...")
        Labelp9_1.Caption = "All" & vbCrLf & "of the" & vbCrLf & "time"
        Labelp9_2.Caption = "Most" & vbCrLf & "of the" & vbCrLf & "time"
        Labelp9_3.Caption = "A good bit" & vbCrLf & "of the" & vbCrLf & "time"
        Labelp9_4.Caption = "Some" & vbCrLf & "of the" & vbCrLf & "time"
        Labelp9_5.Caption = "A little" & vbCrLf & "of the" & vbCrLf & "time"
        Labelp9_6.Caption = "None" & vbCrLf & "of the" & vbCrLf & "time"
        Labelq27.Caption = "Did you have a lot of energy?"
        Labelq28.Caption = "Have you felt downhearted and blue?"
        Labelq29.Caption = "Did you feel worn out?"
        Labelq30.Caption = "Have you been a happy person?"
        Labelq31.Caption = "Did you feel tired?"
    Else
        TextBoxPage9Header1.Text = "De neste spørsmålene handler om hvordan du har følt deg og hvordan du har hatt det " & _
            UCase("de siste 4 ukene.") & vbCrLf & _
            "For hvert spørsmål, vennligst velg det svaralternativet som best beskriver hvordan du har hatt det. " & vbCrLf & _
            "Hvor ofte i løpet av " & UCase("de siste 4 ukene") & " har du:"
        Labelp9_1.Caption = "Hele" & vbCrLf & "tiden"
        Labelp9_2.Caption = "Nesten hele" & vbCrLf & "tiden"
        Labelp9_3.Caption = "Mye av" & vbCrLf & "tiden"
        Labelp9_4.Caption = "En del av" & vbCrLf & "tiden"
        Labelp9_5.Caption = "Litt av" & vbCrLf & "tiden"
        Labelp9_6.Caption = "Ikke" & vbCrLf & "i det hele tatt"
        Labelq27.Caption = "Hatt mye overskudd?"
        Labelq28.Caption = "Følt deg nedfor og trist?"
        Labelq29.Caption = "Følt deg sliten?"
        Labelq30.Caption = "Følt deg glad?"
        Labelq31.Caption = "Følt deg trett?"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub init_page9", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub init_page10()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        TextBoxPage10Header1.Text = "During the " & UCase("past 4 weeks,") & " how much of the time has your " & _
            UCase("physical health or emotional problems") & " interfered with your social activities" & _
            vbCrLf & "(like visiting friends, relatives, etc.)?"
        Labelq32_1.Caption = "All" & vbCrLf & "of the" & vbCrLf & "time"
        Labelq32_2.Caption = "Most" & vbCrLf & "of the" & vbCrLf & "time"
        Labelq32_3.Caption = "Some" & vbCrLf & "of the" & vbCrLf & "time"
        Labelq32_4.Caption = "A little" & vbCrLf & "of the" & vbCrLf & "time"
        Labelq32_5.Caption = "None" & vbCrLf & "of the" & vbCrLf & "time"
        
        TextBoxPage10Header2.Text = "How TRUE or FALSE is " & UCase("each") & " of the following statements for you?"
        Labelp10_1.Caption = "Definitely true"
        Labelp10_2.Caption = "Mostly true"
        Labelp10_3.Caption = "Don't know"
        Labelp10_4.Caption = "Mostly false"
        Labelp10_5.Caption = "Definitely false"
        
        Labelq33.Caption = "I seem to get sick a lot easier than other people"
        Labelq34.Caption = "I am as healthy as anybody I know"
        Labelq35.Caption = "I expect my health to get worse"
        Labelq36.Caption = "My health is excellent"
    Else
        TextBoxPage10Header1.Text = "I løpet av " & UCase("de siste 4 ukene,") & _
            " hvor mye av tiden har din " & UCase("fysiske helse eller følelsesmessige problemer") & _
            " påvirket din sosiale omgang " & _
            vbCrLf & "(som det å besøke venner, slektninger osv.)?"
        Labelq32_1.Caption = "Hele" & vbCrLf & "tiden"
        Labelq32_2.Caption = "Nesten hele" & vbCrLf & "tiden"
        Labelq32_3.Caption = "En del av" & vbCrLf & "tiden"
        Labelq32_4.Caption = "Litt av" & vbCrLf & "tiden"
        Labelq32_5.Caption = "Ikke" & vbCrLf & "i det hele tatt"
        
        TextBoxPage10Header2.Text = "Hvor RIKTIG eller GAL er " & UCase("hver") & " av de følgende påstander for deg?"
        Labelp10_1.Caption = "Helt riktig"
        Labelp10_2.Caption = "Delvis riktig"
        Labelp10_3.Caption = "Vet ikke"
        Labelp10_4.Caption = "Delvis gal"
        Labelp10_5.Caption = "Helt gal"
        
        Labelq33.Caption = "Det virker som jeg blir syk litt lettere enn andre"
        Labelq34.Caption = "Jeg er like frisk som de fleste jeg kjenner"
        Labelq35.Caption = "Jeg tror at helsen min vil forverres"
        Labelq36.Caption = "Jeg har utmerket helse"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub init_page10", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub init_page11()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        TextBoxPage11header.Text = "Additional questions"
        Labelq37.Caption = "Have you had any injuries or other important changes in your health since you completed the LAST survey?"
        Labelq37_no.Caption = "No"
        Labelq37_yes.Caption = "Yes"
        Labelq37_2.Caption = "If YES, please list whatever injuries or health changes you have had:"
        
        Labelq38.Caption = "Have you had any surgeries or hospitalizations since you completed the LAST survey?"
        Labelq38_no.Caption = "No"
        Labelq38_yes.Caption = "Yes"
        Labelq38_2.Caption = "If YES, please list whatever surgeries or hospitalizations you have had:"

        Labelq39.Caption = "Have you changed your medication since you completed the LAST survey?"
        Labelq39_no.Caption = "No"
        Labelq39_yes.Caption = "Yes"
        Labelq39_2.Caption = "If YES, please list changes you have done to your medication:"
    Else
        TextBoxPage11header.Text = "Tilleggs spørsmål"
        Labelq37.Caption = "Har du hatt noen skader eller andre viktige helsemessige endringer siden du fullførte den forrige spørreundersøkelsen?"
        Labelq37_no.Caption = "Nei"
        Labelq37_yes.Caption = "Ja"
        Labelq37_2.Caption = "Hvis JA, vennligst liste opp hvilke skader eller helsemessige endringer du har hatt:"
        
        Labelq38.Caption = "Har du hatt noen operasjoner eller sykehusinnleggelser siden du fullførte den forrige spørreundersøkelsen?"
        Labelq38_no.Caption = "Nei"
        Labelq38_yes.Caption = "Ja"
        Labelq38_2.Caption = "Hvis JA, vennligst nevn operasjoner eller sykehusinnleggelser du har hatt:"

        Labelq39.Caption = "Har du endret medisinering siden du fullførte den forrige spørreundersøkelsen?"
        Labelq39_no.Caption = "Nei"
        Labelq39_yes.Caption = "Ja"
        Labelq39_2.Caption = "Hvis JA, vennligst nevn endringene du har gjort på medisineringen din:"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub init_page11", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub CommandButtonPrev_Click()
On Error GoTo Errhandler

    Select Case MultiPage1.Value
        Case 1
            Call keepUserControlValuesFromPage1
            MultiPage1.Pages(1).Enabled = False
            MultiPage1.Pages(0).Enabled = True
            MultiPage1.Value = 0
            Call init_page0
            CommandButtonPrev.Enabled = False
            CommandButtonNext.Enabled = True
        Case 2
            Call keepUserControlValuesFromPage2
            MultiPage1.Pages(2).Enabled = False
            MultiPage1.Pages(1).Enabled = True
            MultiPage1.Value = 1
            Call init_page1
            CommandButtonPrev.Enabled = True
            CommandButtonNext.Enabled = True
        Case 3
            Call keepUserControlValuesFromPage3
            MultiPage1.Pages(3).Enabled = False
            'Skip page 2
            'MultiPage1.Pages(2).Enabled = True
            'MultiPage1.Value = 2
            'Call init_page2
            '
            ' Go to page 1
            MultiPage1.Pages(1).Enabled = True
            MultiPage1.Value = 1
            Call init_page1
            CommandButtonPrev.Enabled = True
            CommandButtonNext.Enabled = True
        Case 4
            Call keepUserControlValuesFromPage4
            MultiPage1.Pages(4).Enabled = False
            MultiPage1.Pages(3).Enabled = True
            MultiPage1.Value = 3
            Call init_page3
            CommandButtonPrev.Enabled = True
            CommandButtonNext.Enabled = True
        Case 5
            Call keepUserControlValuesFromPage5
            MultiPage1.Pages(5).Enabled = False
            MultiPage1.Pages(4).Enabled = True
            MultiPage1.Value = 4
            Call init_page4
            CommandButtonPrev.Enabled = True
            CommandButtonNext.Enabled = True
        Case 6
            Call keepUserControlValuesFromPage6
            MultiPage1.Pages(6).Enabled = False
            MultiPage1.Pages(5).Enabled = True
            MultiPage1.Value = 5
            Call init_page5
            CommandButtonPrev.Enabled = True
            CommandButtonNext.Enabled = True
        Case 7
            Call keepUserControlValuesFromPage7
            MultiPage1.Pages(7).Enabled = False
            MultiPage1.Pages(6).Enabled = True
            MultiPage1.Value = 6
            Call init_page6
            CommandButtonPrev.Enabled = True
            CommandButtonNext.Enabled = True
        Case 8
            Call keepUserControlValuesFromPage8
            MultiPage1.Pages(8).Enabled = False
            MultiPage1.Pages(7).Enabled = True
            MultiPage1.Value = 7
            Call init_page7
            CommandButtonPrev.Enabled = True
            CommandButtonNext.Enabled = True
        Case 9
            Call keepUserControlValuesFromPage9
            MultiPage1.Pages(9).Enabled = False
            MultiPage1.Pages(8).Enabled = True
            MultiPage1.Value = 8
            Call init_page8
            CommandButtonPrev.Enabled = True
            CommandButtonNext.Enabled = True
        Case 10
            Call keepUserControlValuesFromPage10
            MultiPage1.Pages(10).Enabled = False
            MultiPage1.Pages(9).Enabled = True
            MultiPage1.Value = 9
            Call init_page9
            CommandButtonPrev.Enabled = True
            CommandButtonNext.Enabled = True
        Case 11
            Call keepUserControlValuesFromPage11
            MultiPage1.Pages(11).Enabled = False
            MultiPage1.Pages(10).Enabled = True
            MultiPage1.Value = 10
            Call init_page10
            CommandButtonPrev.Enabled = True
            CommandButtonNext.Enabled = True
        Case Else
            Err.Raise (9) 'Out of range
    End Select
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub CommandButtonPrev_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub CommandButtonNext_Click()
On Error GoTo Errhandler

    Select Case MultiPage1.Value
        Case 0
            Call keepUserControlValuesFromPage0
            MultiPage1.Pages(0).Enabled = False
            MultiPage1.Pages(1).Enabled = True
            MultiPage1.Value = 1
            Call init_page1
            CommandButtonNext.Enabled = True
            CommandButtonPrev.Enabled = True
        Case 1
            Call keepUserControlValuesFromPage1
            MultiPage1.Pages(1).Enabled = False
            'Skip page 2
            'MultiPage1.Pages(2).Enabled = True
            'MultiPage1.Value = 2
            'Call init_page2
            '
            'Go to page 3
            MultiPage1.Pages(3).Enabled = True
            MultiPage1.Value = 3
            Call init_page3
            CommandButtonNext.Enabled = True
            CommandButtonPrev.Enabled = True
        Case 2
            Call keepUserControlValuesFromPage2
            MultiPage1.Pages(2).Enabled = False
            MultiPage1.Pages(3).Enabled = True
            MultiPage1.Value = 3
            Call init_page3
            CommandButtonNext.Enabled = True
            CommandButtonPrev.Enabled = True
        Case 3
            Call keepUserControlValuesFromPage3
            MultiPage1.Pages(3).Enabled = False
            MultiPage1.Pages(4).Enabled = True
            MultiPage1.Value = 4
            Call init_page4
            CommandButtonNext.Enabled = True
            CommandButtonPrev.Enabled = True
        Case 4
            Call keepUserControlValuesFromPage4
            MultiPage1.Pages(4).Enabled = False
            MultiPage1.Pages(5).Enabled = True
            MultiPage1.Value = 5
            Call init_page5
            CommandButtonNext.Enabled = True
            CommandButtonPrev.Enabled = True
        Case 5
            Call keepUserControlValuesFromPage5
            MultiPage1.Pages(5).Enabled = False
            MultiPage1.Pages(6).Enabled = True
            MultiPage1.Value = 6
            Call init_page6
            CommandButtonNext.Enabled = True
            CommandButtonPrev.Enabled = True
        Case 6
            Call keepUserControlValuesFromPage6
            MultiPage1.Pages(6).Enabled = False
            MultiPage1.Pages(7).Enabled = True
            MultiPage1.Value = 7
            Call init_page7
            CommandButtonNext.Enabled = True
            CommandButtonPrev.Enabled = True
        Case 7
            Call keepUserControlValuesFromPage7
            MultiPage1.Pages(7).Enabled = False
            MultiPage1.Pages(8).Enabled = True
            MultiPage1.Value = 8
            Call init_page8
            CommandButtonNext.Enabled = True
            CommandButtonPrev.Enabled = True
        Case 8
            Call keepUserControlValuesFromPage8
            MultiPage1.Pages(8).Enabled = False
            MultiPage1.Pages(9).Enabled = True
            MultiPage1.Value = 9
            Call init_page9
            CommandButtonNext.Enabled = True
            CommandButtonPrev.Enabled = True
        Case 9
            Call keepUserControlValuesFromPage9
            MultiPage1.Pages(9).Enabled = False
            MultiPage1.Pages(10).Enabled = True
            MultiPage1.Value = 10
            Call init_page10
            CommandButtonNext.Enabled = True
            CommandButtonPrev.Enabled = True
        Case 10
            Call keepUserControlValuesFromPage10
            MultiPage1.Pages(10).Enabled = False
            MultiPage1.Pages(11).Enabled = True
            MultiPage1.Value = 11
            Call init_page11
            CommandButtonNext.Enabled = False
            CommandButtonPrev.Enabled = True
        Case Else
            Err.Raise (9) 'Out of range
    End Select
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub CommandButtonNext_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub



Public Sub FetchUserData()
On Error GoTo Errhandler
Dim oldUserData As Variant
Dim rowCount As Integer
Dim colCount As Integer
Dim cIX As Integer
Dim rIX As Integer
Dim mItem As String
Dim mValue As String
Dim iValue As Integer
Dim mAge As Double
Dim ix As Integer
Dim colNr As Integer
Dim UserDataTransposed As Variant


    Worksheets(SelectedUser).Activate
    rowCount = Range("A1").CurrentRegion.Rows.Count
    colCount = Range("A1").CurrentRegion.Columns.Count
    
    If rowCount < 2 Then
        ReDim UserDataTransposed(1 To 2, 1 To maxUserDataCols) As Variant
        'headers in row 1
        rIX = 1
        UserDataTransposed(rIX, 1) = "Row id"
        UserDataTransposed(rIX, 2) = "Form name"
        UserDataTransposed(rIX, 3) = "User"
        UserDataTransposed(rIX, 4) = "Gender"
        UserDataTransposed(rIX, 5) = "Gender code"
        UserDataTransposed(rIX, 6) = "Age"
        UserDataTransposed(rIX, 7) = "Age group"
        cIX = 7
        rowCount = 0 'loop through surveydata
        For ix = 0 To 5
            cIX = cIX + 1
            mItem = SurveyData(rowCount + ix, 6)
            UserDataTransposed(rIX, cIX) = CStr(mItem) 'item name
            'survey date'survey group''form number''form language'Months after baseline'note
        Next ix
        
        rowCount = 6
        For ix = 0 To 19 'health problems, value, row 7 to 26
            cIX = cIX + 1
            mItem = SurveyData(rowCount + ix, 6) & " value" 'health problem value
            UserDataTransposed(rIX, cIX) = CStr(mItem)
        Next ix
        
        For ix = 0 To 19 'health problems, value, row 7 to 26
            cIX = cIX + 1
            mItem = SurveyData(rowCount + ix, 6) & " description" 'health problem description
            UserDataTransposed(rIX, cIX) = CStr(mItem)
        Next ix
     
        rowCount = 26
        cIX = cIX + 1
        mItem = SurveyData(rowCount, 6) 'pain
        UserDataTransposed(rIX, cIX) = CStr(mItem)
            
        rowCount = 27
        cIX = cIX + 1
        mItem = SurveyData(rowCount, 6) 'Lack of energy
        UserDataTransposed(rIX, cIX) = CStr(mItem)
        
        rowCount = 64
        cIX = cIX + 1
        mItem = SurveyData(rowCount, 9) 'q37
        UserDataTransposed(rIX, cIX) = CStr(mItem)
        
        cIX = cIX + 1
        mItem = "q37 text"
        UserDataTransposed(rIX, cIX) = CStr(mItem)
        
        rowCount = 65
        cIX = cIX + 1
        mItem = SurveyData(rowCount, 9) 'q38
        UserDataTransposed(rIX, cIX) = CStr(mItem)
        
        cIX = cIX + 1
        mItem = "q38 text"
        UserDataTransposed(rIX, cIX) = CStr(mItem)
        
        rowCount = 66
        cIX = cIX + 1
        mItem = SurveyData(rowCount, 9) 'q39
        UserDataTransposed(rIX, cIX) = CStr(mItem)
        
        cIX = cIX + 1
        mItem = "q39 text"
        UserDataTransposed(rIX, cIX) = CStr(mItem)
        
        rowCount = 67
        cIX = cIX + 1
        mItem = SurveyData(rowCount, 6) 'Number of missing items
        UserDataTransposed(rIX, cIX) = CStr(mItem)
        
        rowCount = 68
        For ix = 0 To 11 '10 'scales
            cIX = cIX + 1
            mItem = SurveyData(rowCount + ix, 10)
            UserDataTransposed(rIX, cIX) = CStr(mItem)
        Next ix
        
        rowCount = 79
        For ix = 0 To 11 '10 'z scores
            cIX = cIX + 1
            mItem = SurveyData(rowCount + ix, 6) & " z"
            UserDataTransposed(rIX, cIX) = CStr(mItem)
        Next ix
        
        For ix = 0 To 11 '10 'normalized scores
            cIX = cIX + 1
            mItem = SurveyData(rowCount + ix, 6)
            UserDataTransposed(rIX, cIX) = CStr(mItem)
        Next ix
        
        For ix = 0 To 11 '10 'normdata
            cIX = cIX + 1
            mItem = SurveyData(rowCount + ix, 6) & " norm"
            UserDataTransposed(rIX, cIX) = CStr(mItem)
        Next ix
        
        For ix = 0 To 11 '10 'sd
            cIX = cIX + 1
            mItem = SurveyData(rowCount + ix, 6) & " SD"
            UserDataTransposed(rIX, cIX) = CStr(mItem)
        Next ix
        rIX = 2
    Else
        iValue = rowCount + 1
        ReDim UserDataTransposed(1 To iValue, 1 To maxUserDataCols) As Variant
        UserDataTransposed = Range(Cells(1, 1), Cells(iValue, colCount)).Value
        rIX = iValue
    End If
                
    'add a row with data values
    mItem = rIX - 1
    UserDataTransposed(rIX, 1) = CStr(mItem)
    mItem = SurveyData(0, 0)
    UserDataTransposed(rIX, 2) = CStr(mItem) '"Form name"
    mItem = SurveyData(0, 2)
    UserDataTransposed(rIX, 3) = CStr(mItem) '"User"
    mItem = SurveyData(0, 3)
    UserDataTransposed(rIX, 4) = CStr(mItem) '"Gender"
    mItem = SurveyData(0, 4)
    UserDataTransposed(rIX, 5) = CStr(mItem) '"Gender code"
    mItem = SurveyData(0, 5)
    UserDataTransposed(rIX, 6) = CStr(mItem) '"Age"
    
    If (mItem = "-1") Or (mItem = "") Then
        mValue = "Unknown"
    Else
        If IsNumeric(mItem) Then
            mValue = CStr(mItem)
            Select Case mAge
                Case Is < 0
                    mValue = "Unknown"
                Case 0 To 19
                    mValue = "<19"
                Case 20 To 29
                    mValue = "20-29"
                Case 30 To 39
                    mValue = "30-39"
                Case 40 To 49
                    mValue = "40-49"
                Case 50 To 59
                    mValue = "50-59"
                Case 60 To 69
                    mValue = "60-69"
                Case 70 To 79
                    mValue = "70-79"
                Case Is > 79
                    mValue = "80+"
            End Select
        Else
            mValue = "Unknown"
        End If
    End If
    
    cIX = cIX + 1
    UserDataTransposed(rIX, 7) = CStr(mValue) '"Age group"
    
    cIX = 7
    rowCount = 0 'loop through surveydata
    For ix = 0 To 5
        cIX = cIX + 1
        mItem = SurveyData(rowCount + ix, 7)
        UserDataTransposed(rIX, cIX) = CStr(mItem) 'item value
        'survey date'survey group''form number''form language'Months after baseline'note
    Next ix
    
    rowCount = 6
    For ix = 0 To 19 'health problems, value, row 7 to 26
        cIX = cIX + 1
        mItem = SurveyData(rowCount + ix, 7)  'health problem value
        UserDataTransposed(rIX, cIX) = CStr(mItem)
    Next ix
    
    For ix = 0 To 19 'health problems, description, row 7 to 26
        cIX = cIX + 1
        mItem = SurveyData(rowCount + ix, 8) 'health problem description
        UserDataTransposed(rIX, cIX) = CStr(mItem)
    Next ix
    
    rowCount = 26
    cIX = cIX + 1
    mItem = SurveyData(rowCount, 7) 'pain
    UserDataTransposed(rIX, cIX) = CStr(mItem)
        
    rowCount = 27
    cIX = cIX + 1
    mItem = SurveyData(rowCount, 7) 'Lack of energy
    UserDataTransposed(rIX, cIX) = CStr(mItem)
    
    rowCount = 64
    cIX = cIX + 1
    mItem = SurveyData(rowCount, 10) 'q37
    UserDataTransposed(rIX, cIX) = CStr(mItem)
    
    cIX = cIX + 1
    mItem = SurveyData(rowCount, 11) 'q37
    UserDataTransposed(rIX, cIX) = CStr(mItem)
    
    rowCount = 65
    cIX = cIX + 1
    mItem = SurveyData(rowCount, 10) 'q38
    UserDataTransposed(rIX, cIX) = CStr(mItem)
    
    cIX = cIX + 1
    mItem = SurveyData(rowCount, 11) 'q38
    UserDataTransposed(rIX, cIX) = CStr(mItem)
    
    rowCount = 66
    cIX = cIX + 1
    mItem = SurveyData(rowCount, 10) 'q39
    UserDataTransposed(rIX, cIX) = CStr(mItem)
    
    cIX = cIX + 1
    mItem = SurveyData(rowCount, 11) 'q39
    UserDataTransposed(rIX, cIX) = CStr(mItem)
    
    rowCount = 67
    cIX = cIX + 1
    mItem = SurveyData(rowCount, 7) 'Number of missing items
    UserDataTransposed(rIX, cIX) = CStr(mItem)
    
    rowCount = 68
    For ix = 0 To 11 '10 'scales
        cIX = cIX + 1
        mItem = SurveyData(rowCount + ix, 7)
        UserDataTransposed(rIX, cIX) = CStr(mItem)
    Next ix
        
    rowCount = 79
    
    For ix = 0 To 11 '10 'z scores
        cIX = cIX + 1
        mItem = SurveyData(rowCount + ix, 7) & " z"
        UserDataTransposed(rIX, cIX) = CStr(mItem)
    Next ix
        
    For ix = 0 To 11 '10 'normalized scores
        cIX = cIX + 1
        mItem = SurveyData(rowCount + ix, 8)
        UserDataTransposed(rIX, cIX) = CStr(mItem)
    Next ix
        
    For ix = 0 To 11 '10 'normdata
        cIX = cIX + 1
        mItem = SurveyData(rowCount + ix, 9)
        UserDataTransposed(rIX, cIX) = CStr(mItem)
    Next ix
    
    For ix = 0 To 11 '10 'sd
        cIX = cIX + 1
        mItem = SurveyData(rowCount + ix, 10)
        UserDataTransposed(rIX, cIX) = CStr(mItem)
    Next ix
    
    rowCount = rIX '=85
    colCount = cIX
    For rIX = 1 To rowCount
        For cIX = 1 To colCount
            UserData(rIX, cIX) = CStr(UserDataTransposed(rIX, cIX))
        Next cIX
    Next rIX
        
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub FetchUserData", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub CommandButtonSave_Click()
On Error GoTo Errhandler
    Dim rowCount As Integer
    Dim vt2 As Variant
    Dim myRange As Variant
    Dim c As Variant
    Dim lastIx As Integer
    
    
    Me.MousePointer = fmMousePointerHourGlass
    Call FetchUserControlValuesBeforeSave
    Worksheets(surveyWSName).Activate
    ' Save new survey
    Range(Cells(1, 1), Cells(nRows, nCols)) = SurveyData
    
    Worksheets("SurveySummary").Activate
    'Add summary for selected sheet
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
            'lastIx = 6 + n_summary_values
            'Range(Cells(c.Row, 7), Cells(c.Row, lastIx)) = summary_values
            Cells(c.Row, 7) = Format(surveyDate, "yyyy-mm-dd")
            'MsgBox "commandbuttonsave click " & surveyDate
            'lastIx = 7 + n_summary_values
            lastIx = 8 + n_summary_values
            Range(Cells(c.Row, 8), Cells(c.Row, lastIx)) = summary_values
            Range(Cells(c.Row, lastIx), Cells(c.Row, lastIx + 21)) = summary_vas_text
            Range(Cells(c.Row, lastIx + 22), Cells(c.Row, lastIx + 21 + 4)) = summary_text
            Range(Cells(c.Row, lastIx + 26), Cells(c.Row, lastIx + 25 + 6)) = summary_q1q2
            Range(Cells(c.Row, lastIx + 32), Cells(c.Row, lastIx + 31 + 6)) = summary_extra_values
        End If
    End If
    
    
    ' Add to user summary
    'Call FetchUserData
    'Worksheets(SelectedUser).Activate
    'Range(Cells(1, 1), Cells(maxUserDataRows, maxUserDataCols)) = UserData
    ThisWorkbook.Save
    
    Me.MousePointer = fmMousePointerDefault
Exit Sub
Errhandler:
      Me.MousePointer = fmMousePointerDefault
      ErrorHandling "frmNewSurvey. Sub CommandButtonSave_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub CommandButtonSaveAndQuit_Click()
On Error GoTo Errhandler
    Call CommandButtonSave_Click
    hasSaved = True
    'Call populate_users
    Call populate_surveys
    Unload Me
Exit Sub
Errhandler:
      Me.MousePointer = fmMousePointerDefault
      ErrorHandling "frmNewSurvey. Sub CommandButtonSaveAndQuit_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub CommandButtonSF36_Click()
On Error GoTo Errhandler
    ActiveWorkbook.FollowHyperlink "http://www.rand.org/health/surveys_tools/mos/mos_core_36item.html"
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub CommandButtonSF36_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub






'Private Sub DTPickerSurveyDate_Change()
'On Error GoTo Errhandler
'    surveyDate = DTPickerSurveyDate.Value
'    'MsgBox "DTPickerSurveyDate_Change " & surveyDate
'Exit Sub
'Errhandler:
'      ErrorHandling "frmNewSurvey. Sub DTPickerSurveyDate_Change", Err, Action
'      If Action = Err_Exit Then
'         Exit Sub
'      ElseIf Action = Err_Resume Then
'         Resume
'      Else
'         Resume Next
'      End If
'End Sub

Private Sub TextBoxSurveyDate_Init()
On Error GoTo Errhandler
    surveyDate = Now
    TextBoxSurveyDate.Text = Format(surveyDate, "yyyy-mm-dd")
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub TextBoxSurveyDate_Init", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub SliderHealthProblem1_Click()
Dim x As Integer
x = 1
End Sub

Private Sub TextBoxSurveyDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'Private Sub TextBoxSurveyDate_Change()
On Error GoTo Errhandler
    If TextBoxSurveyDate.Text = "" Then
        TextBoxSurveyDate.Text = Format(Now, "yyyy-mm-dd")
    End If
    If IsDate(TextBoxSurveyDate.Text) Then
        TextBoxSurveyDate.Text = Format(TextBoxSurveyDate.Text, "yyyy-mm-dd")
    Else
        TextBoxSurveyDate.Text = Format(Now, "yyyy-mm-dd")
    End If
    surveyDate = TextBoxSurveyDate.Text
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub TextBoxSurveyDate_Exit", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub refresh_selected_page()
On Error GoTo Errhandler

    If Me.MultiPage1.Value = 0 Then
        Call init_page0
    ElseIf Me.MultiPage1.Value = 1 Then
        Call init_page1
    ElseIf Me.MultiPage1.Value = 2 Then
        Call init_page2
    ElseIf Me.MultiPage1.Value = 3 Then
        Call init_page3
    ElseIf Me.MultiPage1.Value = 4 Then
        Call init_page4
    ElseIf Me.MultiPage1.Value = 5 Then
        Call init_page5
    ElseIf Me.MultiPage1.Value = 6 Then
        Call init_page6
    ElseIf Me.MultiPage1.Value = 7 Then
        Call init_page7
    ElseIf Me.MultiPage1.Value = 8 Then
        Call init_page8
    ElseIf Me.MultiPage1.Value = 9 Then
        Call init_page9
    ElseIf Me.MultiPage1.Value = 10 Then
        Call init_page10
    ElseIf Me.MultiPage1.Value = 11 Then
        Call init_page11
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub imgNO_Click", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub imgNO_Click()
On Error GoTo Errhandler

    SelectedLanguage = "NO"
    imgNO.SpecialEffect = fmSpecialEffectRaised
    imgUK.SpecialEffect = fmSpecialEffectFlat
    Call select_gender_text
    Call select_birthYear_text
    Call initCaptions
    Call refresh_selected_page
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub imgNO_Click", Err, Action
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
    Call select_gender_text
    Call select_birthYear_text
    Call initCaptions
    Call refresh_selected_page
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub imgUK_Click", Err, Action
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
    Dim ix As Integer
    For ix = 1 To 39
        sf_36_questions(ix, 1) = "Q" & Str(ix)
        sf_36_questions(ix, 2) = 0
        sf_36_questions(ix, 3) = 0
        sf_36_questions(ix, 4) = ""
        sf_36_questions(ix, 5) = "Missing"
    Next ix
    If SelectedLanguage = "UK" Then
        frmNewSurvey.Caption = "Health survey"
        Label1.Caption = "Health survey"
        frameLanguage.Caption = "Language"
        frmUser.Caption = "User"
        LabelName = "Name or code:"
        LabelMyName = SelectedUser
        LabelGender = "Gender:"
        LabelMyGender = SelectedGender
        LabelYear = "Year of birth:"
        LabelMyYear = SelectedBirthYear
        MultiPage1.Pages(0).Caption = "Questionnaire"
        MultiPage1.Pages(1).Caption = "Health problems"
        MultiPage1.Pages(2).Caption = "Not in use" 'Pain and fatigue"
        MultiPage1.Pages(3).Caption = "Q1-Q2"
        MultiPage1.Pages(4).Caption = "Q3-Q7"
        MultiPage1.Pages(5).Caption = "Q8-Q12"
        MultiPage1.Pages(6).Caption = "Q13-Q19"
        MultiPage1.Pages(7).Caption = "Q20-22"
        MultiPage1.Pages(8).Caption = "Q23-Q26"
        MultiPage1.Pages(9).Caption = "Q27-Q31"
        MultiPage1.Pages(10).Caption = "Q32-Q36"
        MultiPage1.Pages(11).Caption = "Q37-Q39"
        'Me.CommandButtonSave.Caption = "Save"
        Me.CommandButtonSaveAndQuit.Caption = "Save and quit"
        LabelBottom.Caption = "Please fill out the questions on this and the next pages," & vbCrLf & "and hit 'Save' when you are finished."
        'LabelPage0.Caption = "Please fill out the questions on the following pages, and hit 'Save' when you are finished."
        CommandButtonSF36.Caption = "Read more about the SF-36 Health Survey"
        frmInfo.Caption = "Information"
    Else
        frmNewSurvey.Caption = "Spørreskjema om helse"
        Label1.Caption = "Spørreskjema om helse"
        frameLanguage.Caption = "Språk"
        frmUser.Caption = "Bruker"
        LabelName = "Navn eller kode:"
        LabelMyName = SelectedUser
        LabelGender = "Kjønn:"
        LabelMyGender = SelectedGender
        LabelYear = "Fødselsår:"
        LabelMyYear = SelectedBirthYear
        MultiPage1.Pages(0).Caption = "Spørreskjema"
        MultiPage1.Pages(1).Caption = "Helse plager"
        MultiPage1.Pages(2).Caption = "Ikke i bruk" 'Smerte og utmattelse"
        MultiPage1.Pages(3).Caption = "SP1-SP2"
        MultiPage1.Pages(4).Caption = "SP3-SP7"
        MultiPage1.Pages(5).Caption = "SP8-SP12"
        MultiPage1.Pages(6).Caption = "SP13-SP19"
        MultiPage1.Pages(7).Caption = "SP20-SP22"
        MultiPage1.Pages(8).Caption = "SP23-SP26"
        MultiPage1.Pages(9).Caption = "SP27-SP31"
        MultiPage1.Pages(10).Caption = "SP32-SP36"
        MultiPage1.Pages(11).Caption = "SP37-SP39"
        'Me.CommandButtonSave.Caption = "Lagre"
        Me.CommandButtonSaveAndQuit.Caption = "Lagre og avslutt"
        LabelBottom.Caption = "Vennligst fyll ut spørsmålene på denne og de følgende sidene," & vbCrLf & "og velg 'Lagre' når du er ferdig."
        'LabelPage0.Caption = "Vennligst fyll ut spørreskjemaet på de følgende sidene, og velg 'Lagre' når du er ferdig."
        CommandButtonSF36.Caption = "Les mer om SF-36 Spørreundersøkelse om helse"
        frmInfo.Caption = "Informasjon"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub initCaptions", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub saveNewSurveySummaryEntry()
On Error GoTo Errhandler
    Dim mydate As String
    Dim myFormName As String
    Dim rowCount As Integer
    Dim myRange As Variant

    SurveyRegisteredDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
    mydate = Format(SurveyRegisteredDate, "yyyy-mm-dd hh-nn-ss")
    myFormName = SelectedUser & " " & mydate
    surveyWSName = Application.WorksheetFunction.Clean(myFormName)
    
    Worksheets("SurveySummary").Activate
    rowCount = Range("A1").CurrentRegion.Rows.Count
    myRange = "A1:" & "A" & Trim(Str(rowCount))
    ' Save new survey entry
     With Range("A1")
        .Offset(rowCount, 0).Value = surveyWSName
        .Offset(rowCount, 1).Value = SurveyRegisteredDate
        .Offset(rowCount, 2).Value = SelectedUser
        .Offset(rowCount, 3).Value = SelectedBirthYear
        .Offset(rowCount, 4).Value = SelectedGender
        .Offset(rowCount, 5).Value = SelectedGenderCode
    End With
    
    Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = surveyWSName
    Worksheets(surveyWSName).Visible = False
    SelectedSheet = surveyWSName
    
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub saveNewSurveySummaryEntry", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub getOldSurveySummaryEntry()
On Error GoTo Errhandler
    Dim mydate As String
    Dim myFormName As String
    
    oldSurveyWSName = frmStart.ComboBoxSurvey.List(frmStart.ComboBoxSurvey.ListIndex, 0)
    'myDate = Format(frmStart.ComboBoxSurvey.List(frmStart.ComboBoxSurvey.ListIndex, 1), "yyyymmddhhnnss")
    SurveyRegisteredDate = Format(Now, "yyyy-mm-dd hh:nn:ss") 'get new date for a new form
    mydate = Format(SurveyRegisteredDate, "yyyymmddhhnnss")
    SelectedUser = frmStart.ComboBoxSurvey.List(frmStart.ComboBoxSurvey.ListIndex, 2)
    SelectedBirthYear = frmStart.ComboBoxSurvey.List(frmStart.ComboBoxSurvey.ListIndex, 3)
    SelectedGender = frmStart.ComboBoxSurvey.List(frmStart.ComboBoxSurvey.ListIndex, 4)
    SelectedGenderCode = frmStart.ComboBoxSurvey.List(frmStart.ComboBoxSurvey.ListIndex, 5)
    myFormName = SelectedUser & mydate
    surveyWSName = Application.WorksheetFunction.Clean(myFormName)
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub getOldSurveySummaryEntry", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub fetchOldSurveyData()
On Error GoTo Errhandler
    Dim mydate As String
    Dim myFormName As String
    
    
    Me.MousePointer = fmMousePointerHourGlass
    Worksheets(oldSurveyWSName).Activate
    myOldSurveyData = Range(Cells(1, 1), Cells(nRows, nCols)).Value
    Call init_page0
    Call init_page1
    Call init_page2
    Call init_page3
    Call init_page4
    Call init_page5
    Call init_page6
    Call init_page7
    Call init_page8
    Call init_page9
    Call init_page10
    Call init_page11
    Call oldPage0
    Call keepUserControlValuesFromPage0
    Call oldPage1
    Call keepUserControlValuesFromPage1
    'Call oldPage2
    'Call keepUserControlValuesFromPage2
    Call oldPage3
    Call init_page3
    Call keepUserControlValuesFromPage3
    Call oldPage4
    Call keepUserControlValuesFromPage4
    Call oldPage5
    Call keepUserControlValuesFromPage5
    Call oldPage6
    Call keepUserControlValuesFromPage6
    Call oldPage7
    Call keepUserControlValuesFromPage7
    Call oldPage8
    Call keepUserControlValuesFromPage8
    Call oldPage9
    Call keepUserControlValuesFromPage9
    Call oldPage10
    Call keepUserControlValuesFromPage10
    Call oldPage11
    Call keepUserControlValuesFromPage11
    Me.MousePointer = fmMousePointerDefault
Exit Sub
Errhandler:
      Me.MousePointer = fmMousePointerDefault
      ErrorHandling "frmNewSurvey. Sub fetchOldSurveyData", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub UserForm_Activate()
   ' Call UserForm_Initialize
    hasSaved = False
    Call initCaptions
    If SelectedLanguage = "UK" Then
        imgUK.SpecialEffect = fmSpecialEffectRaised
        imgNO.SpecialEffect = fmSpecialEffectFlat
    Else
        imgUK.SpecialEffect = fmSpecialEffectFlat
        imgNO.SpecialEffect = fmSpecialEffectRaised
    End If
    
    MultiPage1.Pages(0).Enabled = True
    MultiPage1.Pages(1).Enabled = False
    MultiPage1.Pages(2).Enabled = False
    MultiPage1.Pages(3).Enabled = False
    MultiPage1.Pages(4).Enabled = False
    MultiPage1.Pages(5).Enabled = False
    MultiPage1.Pages(6).Enabled = False
    MultiPage1.Pages(7).Enabled = False
    MultiPage1.Pages(8).Enabled = False
    MultiPage1.Pages(9).Enabled = False
    MultiPage1.Pages(10).Enabled = False
    MultiPage1.Pages(11).Enabled = False
    'select first page:
    MultiPage1.Value = 0
    'Me.DTPickerSurveyDate.Value = Now
    Me.TextBoxSurveyDate = Format(Now, "yyyy-mm-dd")
    surveyDate = Me.TextBoxSurveyDate
    'surveyDate = Me.DTPickerSurveyDate.Value
    'MsgBox "Init " & DTPickerSurveyDate.Value
    CommandButtonPrev.Enabled = False
    'CommandButtonSave.Enabled = True
    CommandButtonNext.Enabled = True
    CommandButtonNext.SetFocus
    
    If SelectedSheet = "" Then
        Call saveNewSurveySummaryEntry
    Else
        Call getOldSurveySummaryEntry
        Call fetchOldSurveyData
        Call saveNewSurveySummaryEntry
    End If
    Call init_page0
    
    'Me.Repaint
End Sub

Private Sub UserForm_Initialize()
On Error GoTo Errhandler
    surveyDate = Now
    TextBoxFormNumber.Visible = True
    Me.Repaint
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub UserForm_Initialize", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Private Sub UserForm_Terminate()
Call CommandButtonSaveAndQuit_Click
    If hasSaved = False Then
        Call populate_users
        Call populate_surveys
        Unload Me
    End If
End Sub
