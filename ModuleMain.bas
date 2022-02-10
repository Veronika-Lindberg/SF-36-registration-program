Attribute VB_Name = "ModuleMain"
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
' ModuleMain.bas
'
' 2015-04-13 1.0.0 Veronika Lindberg    Created.
'                                       Excel startup macro.
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

'language settings
Global SelectedLanguage As String
'selected user
Global SelectedUser As Variant
Global SelectedGender As String 'Unknown, Male, Female
Global SelectedGenderCode As String '-1,1,0
Global SelectedBirthYear As Variant
Global SelectedSheet As Variant
'Global scale_group_label_text(1 To 10) As String 'was 11
'Global scale_group_label_text_sorted(1 To 10) As String 'was 11
Global scale_group_label_text(1 To 11) As String 'was 11
Global scale_group_label_text_sorted(1 To 11) As String 'was 11



'error handling
Global Action As Integer
Public Const Err_Exit = 0
Public Const Err_Resume = 1
Public Const Err_Resume_Next = 2

Public Type myDataType
    d_mean As Double
    d_SD As Double
    d_N As Double
    d_mean_times_n As Double
    d_N_minus_1 As Double
    d_VARIANCE As Double
    d_VARIANCE_times_N_minus_1 As Double
    d_displayN As Double
    'd_VARIANCE_times_N As Double
    'd_VARIANCE2 As Double
    'd_SD2 As Double
End Type

Public Type myDataTypeByRow
    d_mean As Double
    d_SD As Double
    d_N As Double
    d_group_ix As Integer
    d_group As String
    d_group_name As String
    d_gender As String
    d_gender_code As String
    d_age_group As String
End Type

'Public myDataArray(1 To 11, 1 To 15) As myDataType ' 11 scales, 15 groups: 2 gender x 6 age groups = 12 + 3 overall groups
'Public myDataArrayByRow(1 To 11, 1 To 15) As myDataTypeByRow
'Public NormativeDataByRows(1 To 166, 1 To 9) As Variant ' 11 x 15 = 165 + header, 9 values per variable,
'Group ix    Group   Group name  Gender  Gender code Age group   Mean    SD  N

' More data to fetch in new normative data set
Public myDataArray(1 To 24, 1 To 10) As myDataType ' 10 scales (removed OH), 2 gender + 1 general x 8 age groups = 24
Public myDataArrayByRow(1 To 24, 1 To 10) As myDataTypeByRow
Public NormativeDataByRows(1 To 241, 1 To 10) As Variant '10 scales * 3 gender * 8 age groups = 240 + heading = 241
'Public myDataArray(1 To 24, 1 To 11) As myDataType ' 10 scales (removed OH), 2 gender + 1 general x 8 age groups = 24
'Public myDataArrayByRow(1 To 24, 1 To 11) As myDataTypeByRow
'Public NormativeDataByRows(1 To 241, 1 To 11) As Variant '10 scales * 3 gender * 8 age groups = 240 + heading = 241



Public Const maxAllUserDataRows = 10000
Public Const maxAllUserDataCols = 119
Public AllUserData(1 To maxAllUserDataRows, 1 To maxAllUserDataCols) As Variant

Public Function strOneDriveLocalFilePath() As String
On Error Resume Next 'invalid or non existin registry keys check would evaluate error
'On Error GoTo Errhandler ' do not use error handler here
    Dim ShellScript As Object
    Dim strOneDriveLocalPath As String
    Dim strFileURL As String
    Dim iTryCount As Integer
    Dim strRegKeyName As String
    Dim strFileEndPath As String
    Dim iDocumentsPosition As Integer
    Dim i4thSlashPosition As Integer
    Dim iSlashCount As Integer
    Dim blnFileExist As Boolean
    Dim objFSO As Object
    
    'strFileURL = ThisWorkbook.FullName
    strFileURL = ThisWorkbook.Path
    
    'get OneDrive local path from registry
    Set ShellScript = CreateObject("WScript.Shell")
    '3 possible registry keys to be checked
    For iTryCount = 1 To 3
        Select Case (iTryCount)
            Case 1:
                strRegKeyName = "OneDriveCommercial"
            Case 2:
                strRegKeyName = "OneDriveConsumer"
            Case 3:
                strRegKeyName = "OneDrive"
        End Select
        strOneDriveLocalPath = ShellScript.RegRead("HKEY_CURRENT_USER\Environment\" & strRegKeyName)
        'check if OneDrive location found
        If strOneDriveLocalPath <> vbNullString Then
            'for commercial OneDrive file path seems to be like "https://companyName-my.sharepoint.com/personal/userName_domain_com/Documents" & file.FullName)
            If InStr(1, strFileURL, "my.sharepoint.com") <> 0 Then
                'find "/Documents" in string and replace everything before the end with OneDrive local path
                iDocumentsPosition = InStr(1, strFileURL, "/Documents") + Len("/Documents") 'find "/Documents" position in file URL
                strFileEndPath = Mid(strFileURL, iDocumentsPosition, Len(strFileURL) - iDocumentsPosition + 1)  'get the ending file path without pointer in OneDrive
            Else
                'do nothing
            End If
            'for personal onedrive it looks like "https://d.docs.live.net/d7bbaa#######1/" & file.FullName, _
            '   by replacing "https.." with OneDrive local path obtained from registry we can get local file path
            If InStr(1, strFileURL, "d.docs.live.net") <> 0 Then
                iSlashCount = 1
                i4thSlashPosition = 1
                Do Until iSlashCount > 4
                    i4thSlashPosition = InStr(i4thSlashPosition + 1, strFileURL, "/")   'loop 4 times, looking for "/" after last found
                    iSlashCount = iSlashCount + 1
                Loop
                strFileEndPath = Mid(strFileURL, i4thSlashPosition, Len(strFileURL) - i4thSlashPosition + 1)  'get the ending file path without pointer in OneDrive
            Else
                'do nothing
            End If
        Else
            'continue to check next registry key
        End If
        If Len(strFileEndPath) > 0 Then 'check if path found
            strFileEndPath = Replace(strFileEndPath, "/", "\")  'flip slashes from URL type to File path type
            strOneDriveLocalFilePath = strOneDriveLocalPath & strFileEndPath    'this is the final file path on Local drive
            'verify if file exist in this location and exit for loop if True
            If objFSO Is Nothing Then Set objFSO = CreateObject("Scripting.FileSystemObject")
            If objFSO.FileExist(strOneDriveLocalFilePath) Then
                blnFileExist = True     'that is it - WE GOT IT
                Exit For                'terminate for loop
            Else
                blnFileExist = False    'not there try another OneDrive type (personal/business)
            End If
        Else
            'continue to check next registry key
        End If
    Next iTryCount
    'display message if file could not be located in any OneDrive folders
    If Not blnFileExist Then
        'MsgBox "File could not be found in any OneDrive folders"
        'Use local path if onedrive path is not found
        strOneDriveLocalFilePath = ThisWorkbook.Path
    End If
    
    'clean up
    Set ShellScript = Nothing
    Set objFSO = Nothing
    Exit Function
Errhandler:
      ErrorHandling "Module1. Sub populate_surveys", Err, Action
      If Action = Err_Exit Then
         Exit Function
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Function
    


Public Sub health_and_quality_of_life_macro()
On Error GoTo Errhandler
   'MsgBox "health_and_quality_of_life_macro"
   'test = 1 / 0
   SelectedUser = ""
   SelectedBirthYear = ""
   SelectedGender = ""
   SelectedGenderCode = ""
   'SelectedLanguage = "NO"
   'frmStart.imgNO.SpecialEffect = fmSpecialEffectRaised
   'frmStart.imgUK.SpecialEffect = fmSpecialEffectFlat
    
    
   SelectedLanguage = "UK"
   frmStart.imgUK.SpecialEffect = fmSpecialEffectRaised
   frmStart.imgNO.SpecialEffect = fmSpecialEffectFlat
    
   Call populate_users
   Call populate_surveys
   Call get_normdata
   frmStart.Show False
Exit Sub
Errhandler:
      ErrorHandling "Module1. Sub health_and_quality_of_life_macro", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Public Sub get_scale_label_text()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        scale_group_label_text(1) = "Physical Functioning (PF)"
        scale_group_label_text(2) = "Role limitations due to physical health (RP)"
        scale_group_label_text(3) = "Role limitations due to emotional problems (RE)"
        scale_group_label_text(4) = "Vitality (VT)"
        scale_group_label_text(5) = "Mental health (MH)"
        scale_group_label_text(6) = "Social functioning (SF)"
        scale_group_label_text(7) = "Bodily pain (BP)"
        scale_group_label_text(8) = "General health (GH)"
        scale_group_label_text(9) = "Sum Physical Component Summary (PCS)" 'Physical Component Summary
        scale_group_label_text(10) = "Sum Mental Component Summary (MCS)"
        'scale_group_label_text(11) = "Sum General health condition (OH)"
        scale_group_label_text(11) = "Global Health Composite (GHC)"
    Else
        scale_group_label_text(1) = "Fysisk funksjon (PF)"
        scale_group_label_text(2) = "Rolle begrensninger av fysiske årsaker (RP)"
        scale_group_label_text(3) = "Rolle begrensninger av emosjonelle årsaker (RE)"
        scale_group_label_text(4) = "Vitalitet (VT)"
        scale_group_label_text(5) = "Mental helse (MH)"
        scale_group_label_text(6) = "Sosial funksjon (SF)"
        scale_group_label_text(7) = "Fysiske smerter (BP)"
        scale_group_label_text(8) = "Generell oppfatning av egen helse (GH)"
        scale_group_label_text(9) = "Sum Fysisk Helse (PCS)"
        scale_group_label_text(10) = "Sum Mental Helse (MCS)"
        'scale_group_label_text(11) = "Sum Allmenntilstand (OH)"
        scale_group_label_text(11) = "Global Health Composite (GHC)"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub get_scale_label_text", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Public Sub get_scale_label_text_sorted()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        scale_group_label_text_sorted(1) = "Physical Functioning (PF)"
        scale_group_label_text_sorted(2) = "Role limitations due to physical health (RP)"
        scale_group_label_text_sorted(7) = "Role limitations due to emotional problems (RE)"
        scale_group_label_text_sorted(5) = "Vitality (VT)"
        scale_group_label_text_sorted(8) = "Mental health (MH)"
        scale_group_label_text_sorted(6) = "Social functioning (SF)"
        scale_group_label_text_sorted(3) = "Bodily pain (BP)"
        scale_group_label_text_sorted(4) = "General health (GH)"
        scale_group_label_text_sorted(9) = "Sum Physical Component Summary (PCS)" 'Physical Component Summary
        scale_group_label_text_sorted(10) = "Sum Mental Component Summary (MCS)"
        'scale_group_label_text_sorted(11) = "Sum General health condition (OH)"
        scale_group_label_text_sorted(11) = "Global Health Composite (GHC)"
    Else
        scale_group_label_text_sorted(1) = "Fysisk funksjon (PF)"
        scale_group_label_text_sorted(2) = "Rolle begrensninger av fysiske årsaker (RP)"
        scale_group_label_text_sorted(7) = "Rolle begrensninger av emosjonelle årsaker (RE)"
        scale_group_label_text_sorted(5) = "Vitalitet (VT)"
        scale_group_label_text_sorted(8) = "Mental helse (MH)"
        scale_group_label_text_sorted(6) = "Sosial funksjon (SF)"
        scale_group_label_text_sorted(3) = "Fysiske smerter (BP)"
        scale_group_label_text_sorted(4) = "Generell oppfatning av egen helse (GH)"
        scale_group_label_text_sorted(9) = "Sum Fysisk Helse (PCS)"
        scale_group_label_text_sorted(10) = "Sum Mental Helse (MCS)"
        'scale_group_label_text_sorted(11) = "Sum Allmenntilstand (OH)"
        scale_group_label_text_sorted(11) = "Global Health Composite(GHC)"
    End If
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub get_scale_label_text", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Public Sub fetch_user_data_retro_group()
On Error GoTo Errhandler
    Dim rowCount As Integer, rowCount2 As Integer
    Dim userRowCount As Integer
    Dim userRowIx As Integer
    Dim userColumnCount As Integer
    Dim all_user_data As Variant
    Dim one_user_data As Variant
    Dim myData As Variant, myData2 As Variant
    Dim startIx As Integer
    Dim userIx As Integer
    Dim rowIx As Integer
    Dim colIx As Integer
    Dim myUser As String
    Dim rowIxAllUserData As Integer
    
    userIx = 0
    rowIxAllUserData = 0
    Worksheets("Users").Activate
    rowCount = Trim(Str(Range("A1").CurrentRegion.Rows.Count))
    If rowCount > 1 Then 'at least one user is registered
        myData = Range(Cells(1, 1), Cells(rowCount, 5)).Value 'users
        For userRowIx = 2 To rowCount
        'loop through users
            myUser = myData(userRowIx, 1)
            Worksheets("SurveySummaryRetroGroup").Activate
            'rowCount2 = Trim(Str(Range("A1").CurrentRegion.Rows.Count))
            userRowCount = Trim(Str(Range("A1").CurrentRegion.Rows.Count))
            userColumnCount = Trim(Str(Range("A1").CurrentRegion.Columns.Count))
            If userRowCount > 1 Then 'user has data
                one_user_data = Range(Cells(1, 1), Cells(userRowCount, userColumnCount)).Value
                userIx = userIx + 1
                If userIx = 1 Then
                    startIx = 1
                Else
                    startIx = 2
                End If
                For rowIx = startIx To userRowCount
                    If rowIxAllUserData = 0 Then
                        rowIxAllUserData = rowIxAllUserData + 1
                        AllUserData(rowIxAllUserData, 1) = one_user_data(rowIx, 1)
                    Else
                        rowIxAllUserData = rowIxAllUserData + 1
                        AllUserData(rowIxAllUserData, 1) = rowIxAllUserData - 1
                    End If
                    
                    For colIx = 2 To userColumnCount
                        AllUserData(rowIxAllUserData, colIx) = one_user_data(rowIx, colIx)
                    Next colIx
                Next rowIx
            End If
        Next userRowIx
        If userIx > 0 Then
            Worksheets("Graphs").Activate
            Range(Cells(1, 1), Cells(maxAllUserDataRows, maxAllUserDataCols)) = AllUserData
            ThisWorkbook.Save
        End If
    End If
Exit Sub
Errhandler:
      ErrorHandling "Module1. Sub fetch_user_data_retro_group", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Public Sub fetch_user_data()
On Error GoTo Errhandler
    Dim rowCount As Integer, rowCount2 As Integer
    Dim userRowCount As Integer
    Dim userRowIx As Integer
    Dim userColumnCount As Integer
    Dim all_user_data As Variant
    Dim one_user_data As Variant
    Dim myData As Variant, myData2 As Variant
    Dim startIx As Integer
    Dim userIx As Integer
    Dim rowIx As Integer
    Dim colIx As Integer
    Dim myUser As String
    Dim rowIxAllUserData As Integer
    
    userIx = 0
    rowIxAllUserData = 0
    'Worksheets("Users").Activate
    Worksheets("SurveySummary").Activate
    rowCount = Trim(Str(Range("A1").CurrentRegion.Rows.Count))
    If rowCount > 1 Then 'at least one user is registered
        myData = Range(Cells(1, 1), Cells(rowCount, 5)).Value 'users
        For userRowIx = 2 To rowCount 'loop through users
            myUser = myData(userRowIx, 1)
            'Worksheets("SurveySummary").Activate
            'rowCount2 = Trim(Str(Range("A1").CurrentRegion.Rows.Count))
            
            Worksheets(myUser).Activate
            userRowCount = Trim(Str(Range("A1").CurrentRegion.Rows.Count))
            userColumnCount = Trim(Str(Range("A1").CurrentRegion.Columns.Count))
            If userRowCount > 1 Then 'user has data
                one_user_data = Range(Cells(1, 1), Cells(userRowCount, userColumnCount)).Value
                userIx = userIx + 1
                If userIx = 1 Then
                    startIx = 1
                Else
                    startIx = 2
                End If
                For rowIx = startIx To userRowCount
                    If rowIxAllUserData = 0 Then
                        rowIxAllUserData = rowIxAllUserData + 1
                        AllUserData(rowIxAllUserData, 1) = one_user_data(rowIx, 1)
                    Else
                        rowIxAllUserData = rowIxAllUserData + 1
                        AllUserData(rowIxAllUserData, 1) = rowIxAllUserData - 1
                    End If
                    
                    For colIx = 2 To userColumnCount
                        AllUserData(rowIxAllUserData, colIx) = one_user_data(rowIx, colIx)
                    Next colIx
                Next rowIx
            End If
        Next userRowIx
        If userIx > 0 Then
            Worksheets("Graphs").Activate
            Range(Cells(1, 1), Cells(maxAllUserDataRows, maxAllUserDataCols)) = AllUserData
            ThisWorkbook.Save
        End If
    End If
Exit Sub
Errhandler:
      ErrorHandling "Module1. Sub fetch_user_data", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Public Sub fetch_user_data_by_row() 'populate the graphs sheet
On Error GoTo Errhandler
    Dim formCount As Integer
    Dim formRowCount As Integer
    Dim formColumnCount As Integer
    Dim all_forms As Variant
    
    Dim result As Variant
    Dim numberOfParameters As Integer
    Dim newRowCount As Integer
    Dim totalRowCount As Integer
    Dim parameterIX As Integer
    Dim colIx As Integer
    
    Worksheets("Graphs").Activate
    formCount = Trim(Str(Range("A1").CurrentRegion.Rows.Count))
    formColumnCount = Trim(Str(Range("A1").CurrentRegion.Columns.Count))
    numberOfParameters = formColumnCount - 13
    all_forms = Range(Cells(1, 1), Cells(formCount, formColumnCount)).Value
    
    totalRowCount = 0
    If formCount > 1 Then 'at least one form is registered
        For formRowCount = 2 To formCount 'loop through surveys
            For parameterIX = 1 To numberOfParameters 'loop through parameters
                totalRowCount = totalRowCount + 1
                For colIx = 1 To 13
                    AllUserData(totalRowCount, colIx) = all_forms(formRowCount, colIx)
                Next colIx
                colIx = parameterIX + 13
                AllUserData(totalRowCount, 14) = all_forms(1, colIx)
                AllUserData(totalRowCount, 15) = all_forms(formRowCount, colIx)
            Next parameterIX
        Next formRowCount
        
        Worksheets("Graphs2").Activate
        Range(Cells(1, 1), Cells(maxAllUserDataRows, maxAllUserDataCols)) = AllUserData
        ThisWorkbook.Save
        
    End If
Exit Sub
Errhandler:
      ErrorHandling "Module1. Sub fetch_user_data_by_row", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Public Sub populate_users()
On Error GoTo Errhandler
    Dim myValue As Variant
    Dim rowCount As String
    
    Worksheets("Users").Activate
    rowCount = Trim(Str(Range("A1").CurrentRegion.Rows.Count))
    If rowCount > 1 Then
        With frmStart.cboUsers
            .columnCount = 5
            .ColumnWidths = "60;60;60;60;60"
            .ColumnHeads = False ' True
            .RowSource = "Users!A1:E" & rowCount
            .ListIndex = rowCount - 1
            .SetFocus
        End With
    End If
Exit Sub
Errhandler:
      ErrorHandling "Module1. Sub populate_users", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Public Sub populate_surveys()
On Error GoTo Errhandler
    Dim myValue As Variant
    Dim rowCount As String
    Dim myData As Variant
    Dim rowIx As Integer
    Dim UserData As Variant
    Dim userFormCount As Integer
    Dim mIX As Integer
    
    With frmStart.ComboBoxSurvey
        .columnCount = 6
        .ColumnWidths = "200;60;60;60;60;60"
        .ColumnHeads = False ' True
    End With
        
    Worksheets("SurveySummary").Activate
    rowCount = Trim(Str(Range("A1").CurrentRegion.Rows.Count))
    myData = Range(Cells(1, 1), Cells(rowCount, 6)).Value
    userFormCount = 0
    For rowIx = 2 To rowCount
        If myData(rowIx, 3) = SelectedUser Then
            userFormCount = userFormCount + 1
        End If
    Next rowIx
    
    If userFormCount > 0 Then
        ReDim UserData(1 To userFormCount, 1 To 6)
        mIX = 0
        For rowIx = 2 To rowCount
            If myData(rowIx, 3) = SelectedUser Then
                mIX = mIX + 1
                UserData(mIX, 1) = myData(rowIx, 1)
                UserData(mIX, 2) = myData(rowIx, 2)
                UserData(mIX, 3) = myData(rowIx, 3)
                UserData(mIX, 4) = myData(rowIx, 4)
                UserData(mIX, 5) = myData(rowIx, 5)
                UserData(mIX, 6) = myData(rowIx, 6)
            End If
        Next rowIx
        
        frmStart.ComboBoxSurvey.List = UserData
        frmStart.ComboBoxSurvey.ListIndex = userFormCount - 1
        frmStart.ComboBoxSurvey.SetFocus
        
    Else
        frmStart.ComboBoxSurvey.Clear
    End If
    
    
Exit Sub
Errhandler:
      ErrorHandling "Module1. Sub populate_surveys", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Public Sub select_gender_text()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        If SelectedGenderCode = 0 Then
            SelectedGender = "Female"
        ElseIf SelectedGenderCode = 1 Then
            SelectedGender = "Male"
        Else
            SelectedGender = "Unknown"
        End If
    Else
        If SelectedGenderCode = 0 Then
            SelectedGender = "Kvinne"
        ElseIf SelectedGenderCode = 1 Then
            SelectedGender = "Mann"
        Else
            SelectedGender = "Ukjent"
        End If
    End If
Exit Sub
Errhandler:
      ErrorHandling "Module1. Sub select_gender_text", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Public Sub select_birthYear_text()
On Error GoTo Errhandler
    If Not IsNumeric(SelectedBirthYear) Then
        If SelectedLanguage = "UK" Then
            SelectedBirthYear = "Unknown"
        Else
            SelectedBirthYear = "Ukjent"
        End If
    End If
Exit Sub
Errhandler:
      ErrorHandling "Module1. Sub select_birthYear_text", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Public Function get_pain_text()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        get_pain_text = "Not in use" ' "Pain"
    Else
        get_pain_text = "Ikke i bruk" '"Smerter"
    End If

Exit Function
Errhandler:
      ErrorHandling "Module1. Function get_pain_text", Err, Action
      If Action = Err_Exit Then
         Exit Function
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Function

Public Function get_lack_of_energy_text()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        get_lack_of_energy_text = "Not in use" ' "Lack of energy"
    Else
        get_lack_of_energy_text = "Ikke i bruk" ' "Mangel på energi"
    End If

Exit Function
Errhandler:
      ErrorHandling "Module1. Function get_lack_of_energy_text", Err, Action
      If Action = Err_Exit Then
         Exit Function
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Function


Public Sub populate_NormativeDataByRowsOld()
On Error GoTo Errhandler

Dim normativeData As Variant
Dim myDataArrayByRow(1 To 11, 1 To 15) As myDataTypeByRow


Dim rowIx As Integer
Dim colIx As Integer
Dim fetchRowIX As Integer
Dim fetchColIX As Integer
Dim nRows As Integer
Dim nCols As Integer
Dim myStr As String
Dim rIX As Integer
Dim cIX As Integer

    For rowIx = 1 To 11
        For colIx = 1 To 15
            myDataArrayByRow(rowIx, colIx).d_mean = 0
            myDataArrayByRow(rowIx, colIx).d_SD = 0
            myDataArrayByRow(rowIx, colIx).d_N = 0
            myDataArrayByRow(rowIx, colIx).d_group_ix = 0
            myDataArrayByRow(rowIx, colIx).d_group = ""
            myDataArrayByRow(rowIx, colIx).d_group_name = ""
            myDataArrayByRow(rowIx, colIx).d_gender = ""
            myDataArrayByRow(rowIx, colIx).d_gender_code = ""
            myDataArrayByRow(rowIx, colIx).d_age_group = ""
        Next colIx
    Next rowIx

    'fetch overall mean,sd and n for a row (scale group)
    Worksheets("NormativeDataOld").Activate
    nRows = 34
    nCols = 18
    normativeData = Range(Cells(1, 1), Cells(nRows, nCols)).Value
    
    'Group label Scale   M29 F29 M30 F30 M40 F40 M50 F50 M60 F60 M70 F70 ALL_M   ALL_F   ALL
    'PF  physical_functioning_mean   94,7
    'PF  physical_functioning_SD 12,4
    'PF  physical_functioning_N  231

    'get mean, sd and n
    fetchRowIX = 2
    For rowIx = 1 To 11
        fetchColIX = 4
        'Group label Scale   M29 F29 M30 F30 M40 F40 M50 F50 M60 F60 M70 F70 ALL_M   ALL_F   ALL
        'PF  physical_functioning_mean   94,7
        'PF  physical_functioning_SD 12,4
        'PF  physical_functioning_N  231
        For colIx = 1 To 15
            myDataArrayByRow(rowIx, colIx).d_mean = normativeData(fetchRowIX, fetchColIX)
            myDataArrayByRow(rowIx, colIx).d_SD = normativeData(fetchRowIX + 1, fetchColIX)
            myDataArrayByRow(rowIx, colIx).d_N = normativeData(fetchRowIX + 2, fetchColIX)
            
            myDataArrayByRow(rowIx, colIx).d_group_ix = normativeData(fetchRowIX, 1)
            myDataArrayByRow(rowIx, colIx).d_group = normativeData(fetchRowIX, 2)
            myDataArrayByRow(rowIx, colIx).d_group_name = normativeData(fetchRowIX, 3)
            myStr = normativeData(1, fetchColIX) 'gender and age
            Select Case myStr
                Case "M29"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Male"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "1"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "29"
                Case "F29"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Female"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "0"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "29"
                Case "M30"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Male"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "1"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "30"
                Case "F30"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Female"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "0"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "30"
                Case "M40"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Male"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "1"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "40"
                Case "F40"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Female"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "0"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "40"
                Case "M50"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Male"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "1"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "50"
                Case "F50"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Female"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "0"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "50"
                Case "M60"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Male"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "1"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "60"
                Case "F60"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Female"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "0"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "60"
                Case "M70"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Male"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "1"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "70"
                Case "F70"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Female"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "0"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "70"
               Case "ALL_M"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Male"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "1"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "Unknown"
                Case "ALL_F"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Female"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "0"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "Unknown"
                Case "ALL"
                    myDataArrayByRow(rowIx, colIx).d_gender = "Unknown"
                    myDataArrayByRow(rowIx, colIx).d_gender_code = "-1"
                    myDataArrayByRow(rowIx, colIx).d_age_group = "Unknown"
            End Select
            fetchColIX = fetchColIX + 1
        Next colIx
        fetchRowIX = fetchRowIX + 3
    Next rowIx
    
    rIX = 1
    NormativeDataByRows(rIX, 1) = "Group ix"
    NormativeDataByRows(rIX, 2) = "Group"
    NormativeDataByRows(rIX, 3) = "Group name"
    NormativeDataByRows(rIX, 4) = "Gender"
    NormativeDataByRows(rIX, 5) = "Gender code"
    NormativeDataByRows(rIX, 6) = "Age group"
    NormativeDataByRows(rIX, 7) = "Mean"
    NormativeDataByRows(rIX, 8) = "SD"
    NormativeDataByRows(rIX, 9) = "N"
    For rowIx = 1 To 11
         For colIx = 1 To 15
            rIX = rIX + 1
            NormativeDataByRows(rIX, 1) = myDataArrayByRow(rowIx, colIx).d_group_ix
            NormativeDataByRows(rIX, 2) = myDataArrayByRow(rowIx, colIx).d_group
            NormativeDataByRows(rIX, 3) = Left(myDataArrayByRow(rowIx, colIx).d_group_name, Len(myDataArrayByRow(rowIx, colIx).d_group_name) - 5)
            NormativeDataByRows(rIX, 4) = myDataArrayByRow(rowIx, colIx).d_gender
            NormativeDataByRows(rIX, 5) = myDataArrayByRow(rowIx, colIx).d_gender_code
            NormativeDataByRows(rIX, 6) = myDataArrayByRow(rowIx, colIx).d_age_group
            NormativeDataByRows(rIX, 7) = myDataArrayByRow(rowIx, colIx).d_mean
            NormativeDataByRows(rIX, 8) = myDataArrayByRow(rowIx, colIx).d_SD
            NormativeDataByRows(rIX, 9) = myDataArrayByRow(rowIx, colIx).d_N
        Next colIx
    Next rowIx
    
    
    Worksheets("NormativeDataByRows").Activate
    Range(Cells(1, 1), Cells(166, 9)) = NormativeDataByRows
    ThisWorkbook.Save
    

Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub populate_NormativeDataByRowsOld", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Public Sub populate_NormativeDataByRows()
On Error GoTo Errhandler

Dim normativeData As Variant


Dim rowIx As Integer
Dim colIx As Integer
Dim groupIX As Integer
Dim scaleIX As Integer
Dim fetchScaleIX As Integer
Dim fetchGroupIX As Integer
Dim nRows As Integer
Dim nCols As Integer
Dim myStr As String
Dim rIX As Integer
Dim cIX As Integer

    For rowIx = 1 To 24 '10 ' was 11, removed OH
        For colIx = 1 To 10 '24 ' was 15 2x6=12 +3=15, 2 more age groups for 2 genders + 1 general group for each age, 8x3=24
            myDataArrayByRow(rowIx, colIx).d_mean = 0
            myDataArrayByRow(rowIx, colIx).d_SD = 0
            myDataArrayByRow(rowIx, colIx).d_N = 0
            myDataArrayByRow(rowIx, colIx).d_group_ix = 0
            myDataArrayByRow(rowIx, colIx).d_group = ""
            myDataArrayByRow(rowIx, colIx).d_group_name = ""
            myDataArrayByRow(rowIx, colIx).d_gender = ""
            myDataArrayByRow(rowIx, colIx).d_gender_code = ""
            myDataArrayByRow(rowIx, colIx).d_age_group = ""
        Next colIx
    Next rowIx

    'fetch overall mean,sd and n for a row (scale group)
    Worksheets("NormativeDataNew").Activate
    nRows = 75 '34
    nCols = 14 '18
    normativeData = Range(Cells(1, 1), Cells(nRows, nCols)).Value
    
    'Group label Scale   M29 F29 M30 F30 M40 F40 M50 F50 M60 F60 M70 F70 ALL_M   ALL_F   ALL
    'PF  physical_functioning_mean   94,7
    'PF  physical_functioning_SD 12,4
    'PF  physical_functioning_N  231

    'get mean, sd and n
    fetchScaleIX = 4 'values start by row 4
    
    For scaleIX = 1 To 24 '24 age - gender - groups
        'Age Group code  Gender  Variable    PF  RP  BP  GH  VT  SF  RE  MH  PCS MCS
        'Group label Scale   M29 F29 M30 F30 M40 F40 M50 F50 M60 F60 M70 F70 ALL_M   ALL_F   ALL
        'PF  physical_functioning_mean   94,7
        'PF  physical_functioning_SD 12,4
        'PF  physical_functioning_N  231
        fetchGroupIX = 5 'values start by column 5
        For groupIX = 1 To 10 'loop through 10 sf-36 variables
            
            myDataArrayByRow(scaleIX, groupIX).d_N = normativeData(fetchScaleIX, fetchGroupIX)
            myDataArrayByRow(scaleIX, groupIX).d_mean = normativeData(fetchScaleIX + 1, fetchGroupIX)
            myDataArrayByRow(scaleIX, groupIX).d_SD = normativeData(fetchScaleIX + 2, fetchGroupIX)
            
            'myDataArrayByRow(scaleIX, groupIX).d_mean = normativeData(fetchscaleIX, fetchgroupIX)
            'myDataArrayByRow(scaleIX, groupIX).d_SD = normativeData(fetchscaleIX + 1, fetchgroupIX)
            'myDataArrayByRow(scaleIX, groupIX).d_N = normativeData(fetchscaleIX + 2, fetchgroupIX)
            
            myDataArrayByRow(scaleIX, groupIX).d_group_ix = normativeData(2, fetchGroupIX)
            myDataArrayByRow(scaleIX, groupIX).d_group = normativeData(1, fetchGroupIX)
            myDataArrayByRow(scaleIX, groupIX).d_group_name = normativeData(3, fetchGroupIX) ' + normativeData(fetchscaleIX, 4)
            myStr = normativeData(fetchScaleIX, 2) 'gender and age
            Select Case myStr
                Case "M19"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Male"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "19"
                Case "F19"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Female"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "0"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "19"
                Case "A19"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Unknown"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "-1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "19"
                Case "M20"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Male"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "20"
                Case "F20"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Female"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "0"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "20"
                Case "A20"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Unknown"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "-1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "20"
                Case "M30"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Male"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "30"
                Case "F30"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Female"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "0"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "30"
                Case "A30"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Unknown"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "-1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "30"
                Case "M40"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Male"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "40"
                Case "F40"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Female"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "0"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "40"
                Case "A40"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Unknown"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "-1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "40"
                Case "M50"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Male"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "50"
                Case "F50"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Female"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "0"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "50"
                Case "A50"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Unknown"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "-1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "50"
                Case "M60"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Male"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "60"
                Case "F60"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Female"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "0"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "60"
                Case "A60"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Unknown"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "-1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "60"
                Case "M70"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Male"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "70"
                Case "F70"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Female"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "0"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "70"
                Case "A70"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Unknown"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "-1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "70"
                Case "M80"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Male"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "80"
                Case "F80"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Female"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "0"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "80"
                Case "A80"
                    myDataArrayByRow(scaleIX, groupIX).d_gender = "Unknown"
                    myDataArrayByRow(scaleIX, groupIX).d_gender_code = "-1"
                    myDataArrayByRow(scaleIX, groupIX).d_age_group = "80"
               'Case "ALL_M"
               '     myDataArrayByRow(scaleIX, groupIX).d_gender = "Male"
               '     myDataArrayByRow(scaleIX, groupIX).d_gender_code = "1"
               '     myDataArrayByRow(scaleIX, groupIX).d_age_group = "Unknown"
               ' Case "ALL_F"
               '     myDataArrayByRow(scaleIX, groupIX).d_gender = "Female"
               '     myDataArrayByRow(scaleIX, groupIX).d_gender_code = "0"
               '     myDataArrayByRow(scaleIX, groupIX).d_age_group = "Unknown"
               ' Case "ALL"
               '     myDataArrayByRow(scaleIX, groupIX).d_gender = "Unknown"
               '     myDataArrayByRow(scaleIX, groupIX).d_gender_code = "-1"
               '     myDataArrayByRow(scaleIX, groupIX).d_age_group = "Unknown"
            End Select
            fetchGroupIX = fetchGroupIX + 1
        Next groupIX
        fetchScaleIX = fetchScaleIX + 3
    Next scaleIX
    
    rIX = 1
    NormativeDataByRows(rIX, 1) = "Group ix"
    NormativeDataByRows(rIX, 2) = "Group"
    NormativeDataByRows(rIX, 3) = "Group name"
    NormativeDataByRows(rIX, 4) = "Gender"
    NormativeDataByRows(rIX, 5) = "Gender code"
    NormativeDataByRows(rIX, 6) = "Age group"
    NormativeDataByRows(rIX, 7) = "Mean"
    NormativeDataByRows(rIX, 8) = "SD"
    NormativeDataByRows(rIX, 9) = "N"
    NormativeDataByRows(rIX, 10) = "Sort ix"
    For rowIx = 1 To 24 '10
         For colIx = 1 To 10 '24 '15
            rIX = rIX + 1
            NormativeDataByRows(rIX, 1) = myDataArrayByRow(rowIx, colIx).d_group_ix
            NormativeDataByRows(rIX, 2) = myDataArrayByRow(rowIx, colIx).d_group
            NormativeDataByRows(rIX, 3) = myDataArrayByRow(rowIx, colIx).d_group_name
            NormativeDataByRows(rIX, 4) = myDataArrayByRow(rowIx, colIx).d_gender
            NormativeDataByRows(rIX, 5) = myDataArrayByRow(rowIx, colIx).d_gender_code
            NormativeDataByRows(rIX, 6) = myDataArrayByRow(rowIx, colIx).d_age_group
            NormativeDataByRows(rIX, 7) = myDataArrayByRow(rowIx, colIx).d_mean
            NormativeDataByRows(rIX, 8) = myDataArrayByRow(rowIx, colIx).d_SD
            NormativeDataByRows(rIX, 9) = myDataArrayByRow(rowIx, colIx).d_N
            NormativeDataByRows(rIX, 10) = rowIx
        Next colIx
    Next rowIx
    
    
    Worksheets("NormativeDataByRows").Activate
    Range(Cells(1, 1), Cells(241, 10)) = NormativeDataByRows
    ThisWorkbook.Save
    

Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub populate_NormativeDataByRows", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub



Public Sub get_normdata()
On Error GoTo Errhandler
Dim mnRows As Integer
Dim mnCols As Integer
Dim ix As Integer
Dim normativeData As Variant
Dim rIX As Integer
Dim cIX As Integer
Dim fetchValues As Variant
Dim rowIx As Integer, colIx As Integer
    

    Worksheets("NormativeDataByRows").Activate
    mnRows = 241
    mnCols = 10
    fetchValues = Range(Cells(1, 1), Cells(mnRows, mnCols)).Value
    rIX = 1
    For rowIx = 1 To 24
        For colIx = 1 To 10
            rIX = rIX + 1
            'Group ix    Group   Group name  Gender  Gender code Age group   Mean    SD  N
            myDataArrayByRow(rowIx, colIx).d_group_ix = fetchValues(rIX, 1)
            myDataArrayByRow(rowIx, colIx).d_group = fetchValues(rIX, 2)
            myDataArrayByRow(rowIx, colIx).d_group_name = fetchValues(rIX, 3)
            myDataArrayByRow(rowIx, colIx).d_gender = fetchValues(rIX, 4)
            myDataArrayByRow(rowIx, colIx).d_gender_code = fetchValues(rIX, 5)
            myDataArrayByRow(rowIx, colIx).d_age_group = fetchValues(rIX, 6)
            myDataArrayByRow(rowIx, colIx).d_mean = fetchValues(rIX, 7)
            myDataArrayByRow(rowIx, colIx).d_SD = fetchValues(rIX, 8)
            myDataArrayByRow(rowIx, colIx).d_N = fetchValues(rIX, 9)
        Next colIx
    Next rowIx
    
Exit Sub
Errhandler:
      ErrorHandling "frmNewSurvey. Sub get_normdata", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub



Public Sub ErrorHandling(ErrorSource As String, ErrorValue As Integer, ReturnValue As Integer)
      Dim result As Integer
      Dim ErrMsg As String
      Dim Choices As Integer

      Select Case ErrorValue
         Case 68:     ' Device  not available.
            ErrMsg = "The device you are trying to access is either " & _
               "not online or does not exist. Retry?"
            Choices = vbOKCancel
         Case 75:     ' Path/File access error.
            ErrMsg = "There is an error accessing the path and/or " & _
                 "file specified. Retry?"
            Choices = vbOKCancel
         Case 76:     ' Path not found.
            ErrMsg = "The path and/or file specified was not found. Retry?"
            Choices = vbOKCancel
         Case Else:   'An error other than 68, 75 or 76 has occurred
            ErrMsg = "An unrecognized error has occurred ( " & _
               Error(Err) & " )."
            MsgBox ErrMsg, vbOKOnly, ErrorSource
            ReturnValue = Err_Exit
            Exit Sub
      End Select
      ' Display the error message.
      result = MsgBox(ErrMsg, Choices, ErrorSource)
      ' Determine the ReturnValue based on the user's choice from MsgBox.
      If result = vbOK Then
         ReturnValue = Err_Resume
      Else
         ReturnValue = Err_Exit
      End If
End Sub

Public Sub ShowSheet()
' For test only - make all worksheets visible
' Application.Worksheets("u1210 2015-05-18 10-21-58").Visible = xlSheetVisible
 Dim mIX As Integer
 Dim ix As Integer
 
 mIX = Application.Worksheets.Count
 
For ix = 1 To mIX
Application.Worksheets(ix).Visible = True
Next ix

 End Sub
