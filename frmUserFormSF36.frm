VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUserFormSF36 
   Caption         =   "SF-36"
   ClientHeight    =   7980
   ClientLeft      =   30
   ClientTop       =   310
   ClientWidth     =   15750
   OleObjectBlob   =   "frmUserFormSF36-2022-4-0-2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUserFormSF36"
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
' frmUserFormSF36.frm
'
' 2015-04-13 1.0.0 Veronika Lindberg    Created.
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

Public Sub initCaptions()
On Error GoTo Errhandler
    If SelectedLanguage = "UK" Then
        frmUserFormSF36.Caption = "RAND SF-36" '"User graphs"
        CheckBox1.Caption = "Show Min, Max, SD"
        CheckBox2.Caption = "Show norm based values"
        'TextBox1.Text = "" 'SelectedUser
    Else
        frmUserFormSF36.Caption = "RAND SF-36" '"Grafikk for bruker"
        CheckBox1.Caption = "Vis Min, Max, SD"
        CheckBox2.Caption = "Vis norm baserte verdier"
        'TextBox1.Text = "" 'SelectedUser
    End If
    Me.Repaint
Exit Sub
Errhandler:
      ErrorHandling "frmUserFormVAS. Sub initCaptions", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub

Public Sub chartScaleScores()
On Error GoTo Errhandler
Dim objSheet As Object
Dim objChart As Object
Dim x As Long
Dim myPic As Variant
Dim vt2 As Variant
Dim strChartTitle As String, strNormdata As String, strXaxis As String, strYaxis As String
Dim r1 As Range
Dim r2 As Range
Dim myMultipleRange As Range
Dim xaxis As Range
Dim yaxis As Range
Dim s As Series
Dim strMax As String, strMin As String, strPlussSD As String, strMinusSD As String


    Me.MousePointer = fmMousePointerHourGlass
    If SelectedLanguage = "UK" Then
        strChartTitle = SelectedUser & ": General Health Condition by category"
        'strChartTitle = "General Health Condition for " & SelectedUser
        strNormdata = "Mean for general population"
        strXaxis = "RAND SF-36 Categories"
        strYaxis = "RAND SF-36 Scale Scores, 100 = Best"
        strMax = "Best possible value"
        strMin = "Worst possible value"
        strPlussSD = "+1 Standard Deviation"
        strMinusSD = "-1 Standard Deviation"
    Else
        strChartTitle = SelectedUser & ": Allmenntilstand kategorisert"
        strNormdata = "Gjennomsnitt for befolkningen"
        strXaxis = "RAND SF-36 Kategorier"
        strYaxis = "RAND SF-36 Verdier, 100 = Best"
        strMax = "Høyest mulige verdi"
        strMin = "Lavest mulige verdi"
        strPlussSD = "+1 Standardavvik"
        strMinusSD = "-1 Standardavvik"
    End If
    
    Call get_scale_label_text_sorted

    Set objSheet = Worksheets("SurveySummary")
    'clean up, delete old charts
    x = objSheet.ChartObjects.Count
    If x > 0 Then
        objSheet.ChartObjects.Delete
    End If
    x = objSheet.Pictures.Count
    If x > 0 Then
        objSheet.Pictures.Delete
    End If
    
    'create empty chart
    'Set objChart = objSheet.ChartObjects.Add(0, 0, frmUserFormSF36.Width, frmUserFormSF36.Height).chart
    Set objChart = objSheet.ChartObjects.Add(0, 0, Image2.Width, Image2.Height).chart
    objSheet.ChartObjects(1).Activate
    With ActiveChart
     .ChartType = xlColumnClustered
     .HasTitle = True
     .ChartTitle.Text = strChartTitle
     
     .Axes(xlCategory, xlPrimary).HasTitle = True
     .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = strXaxis
    
     .Axes(xlValue, xlPrimary).HasTitle = True
     .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = strYaxis
     
     .Axes(xlValue).MinimumScale = 0
     .Axes(xlValue).MaximumScale = 100
     '.Axes(xlCategory).TickLabels.Orientation = xlTickLabelOrientationHorizontal
     .Refresh
    End With
    
    Dim rowIx As Long
    Dim rowCount As Long
    Dim strUser As String
    Dim strRange As String
    Dim serieCount As Integer
    Dim rrowCount As Long
    Dim dValues As Range, sdValues As Range, sumValues As Range
    Dim dMin As Double, dMax As Double
    'Dim max_values(1 To 10) As Double, min_values(1 To 10) As Double
    Dim max_values(1 To 11) As Double, min_values(1 To 11) As Double
    
    dMin = 0
    dMax = 100
    rowCount = Range("C1").CurrentRegion.Rows.Count
    serieCount = 0
    For rowIx = 2 To rowCount
        strUser = Cells(rowIx, 3)
        If UCase(strUser) = UCase(SelectedUser) Then
        ' adding series
            serieCount = serieCount + 1
            rrowCount = rowIx
            objChart.SeriesCollection.NewSeries
            objChart.SeriesCollection(serieCount).Name = Cells(rowIx, 7) 'Survey date
            'objChart.SeriesCollection(serieCount).XValues = objSheet.Range("AD1:AN1") 'Group labels
            objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
            'strRange = "AO" & rowIx & ":AY" & rowIx
            strRange = "AO" & rowIx & ":AY" & rowIx
            objChart.SeriesCollection(serieCount).Values = objSheet.Range(strRange)
        End If
    Next rowIx
    If serieCount > 0 Then
    ' adding reference values
        serieCount = serieCount + 1
        objChart.SeriesCollection.NewSeries
        objChart.SeriesCollection(serieCount).Name = strNormdata
        'objChart.SeriesCollection(serieCount).XValues = objSheet.Range("AD1:AN1")
        objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
        'strRange = "AZ" & rrowCount & ":BJ" & rrowCount
        strRange = "AZ" & rrowCount & ":BJ" & rrowCount
        objChart.SeriesCollection(serieCount).Values = objSheet.Range(strRange)
        objChart.SeriesCollection(serieCount).Type = xlLine
        objChart.SeriesCollection(serieCount).MarkerStyle = xlNone
        If CheckBox1.Value = True Then
        
            'strRange = "AZ" & rrowCount & ":BJ" & rrowCount
            strRange = "AZ" & rrowCount & ":BJ" & rrowCount
            Set dValues = objSheet.Range(strRange)
            'strRange = "BK" & rrowCount & ":BU" & rrowCount
            strRange = "BK" & rrowCount & ":BU" & rrowCount
            Set sdValues = objSheet.Range(strRange)
            
            Dim i As Integer
            For i = 1 To 11 '10 '11
                max_values(i) = dValues.Cells(i).Value + sdValues.Cells(i).Value
                If max_values(i) > dMax Then
                    dMax = max_values(i)
                End If
                min_values(i) = dValues.Cells(i).Value - sdValues.Cells(i).Value
                If min_values(i) < dMin Then
                    dMin = min_values(i)
                End If
            Next i
            
            serieCount = serieCount + 1
            objChart.SeriesCollection.NewSeries
            objChart.SeriesCollection(serieCount).Name = strPlussSD
            objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
            objChart.SeriesCollection(serieCount).Values = max_values
            objChart.SeriesCollection(serieCount).Type = xlLine
            objChart.SeriesCollection(serieCount).MarkerStyle = xlNone
            objChart.SeriesCollection(serieCount).Format.Line.DashStyle = msoLineRoundDot

            
            serieCount = serieCount + 1
            objChart.SeriesCollection.NewSeries
            objChart.SeriesCollection(serieCount).Name = strMinusSD
            objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
            objChart.SeriesCollection(serieCount).Values = min_values
            objChart.SeriesCollection(serieCount).Type = xlLine
            objChart.SeriesCollection(serieCount).MarkerStyle = xlNone
            objChart.SeriesCollection(serieCount).Format.Line.DashStyle = msoLineRoundDot
            
            For i = 1 To 11 '10 '11
                max_values(i) = 100
                min_values(i) = 0
            Next i
            
            serieCount = serieCount + 1
            objChart.SeriesCollection.NewSeries
            objChart.SeriesCollection(serieCount).Name = strMax
            objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
            objChart.SeriesCollection(serieCount).Values = max_values
            objChart.SeriesCollection(serieCount).Type = xlLine
            objChart.SeriesCollection(serieCount).MarkerStyle = xlNone
            objChart.SeriesCollection(serieCount).Format.Line.DashStyle = msoLineDash
            
            serieCount = serieCount + 1
            objChart.SeriesCollection.NewSeries
            objChart.SeriesCollection(serieCount).Name = strMin
            objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
            objChart.SeriesCollection(serieCount).Values = min_values
            objChart.SeriesCollection(serieCount).Type = xlLine
            objChart.SeriesCollection(serieCount).MarkerStyle = xlNone
            objChart.SeriesCollection(serieCount).Format.Line.DashStyle = msoLineDash
            
            objChart.Axes(xlValue).MinimumScale = dMin
            objChart.Axes(xlValue).MaximumScale = dMax
            'objChart.Axes(xlValue).MinorUnit = 10
            'objChart.Axes(xlValue).MajorUnit = 100
        
        End If
    End If
    ActiveChart.Refresh
    
    Worksheets("SurveySummary").ChartObjects(1).CopyPicture
    ActiveSheet.Paste
    objSheet.ChartObjects(1).Activate
    
    Dim strPath As String
    x = objSheet.Pictures.Count
    If x > 0 Then
        'objSheet.Pictures.Activate
        'Set myPic = objSheet.Pictures(x)
        strPath = strOneDriveLocalFilePath
        strPath = strPath & "\" & "tmp.bmp"
        'strPath = ActiveWorkbook.Path & "\" & "tmp.bmp"
        ActiveChart.Export strPath
        'frmUserFormSF36.Picture = LoadPicture(strPath)
        Image2.Picture = LoadPicture(strPath)
     End If
 
    'clean up, delete old charts, but keep picture of graph
    x = objSheet.ChartObjects.Count
    If x > 0 Then
        objSheet.ChartObjects.Delete
    End If
    Me.MousePointer = fmMousePointerDefault
    
Exit Sub
Errhandler:
    Me.MousePointer = fmMousePointerDefault
      ErrorHandling "frmUserFormsSF36. Sub chartScaleScores", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Public Sub chartNormScores()
On Error GoTo Errhandler
Dim objSheet As Object
Dim objChart As Object
Dim x As Long
Dim myPic As Variant
Dim vt2 As Variant
Dim strChartTitle As String, strNormdata As String, strXaxis As String, strYaxis As String
Dim r1 As Range
Dim r2 As Range
Dim myMultipleRange As Range
Dim xaxis As Range
Dim yaxis As Range
Dim s As Series
Dim strMax As String, strMin As String, strPlussSD As String, strMinusSD As String, strBelowMean As String


    Me.MousePointer = fmMousePointerHourGlass
    
    If SelectedLanguage = "UK" Then
        strChartTitle = SelectedUser & ": General Health Condition by category"
        strNormdata = "Mean for general population"
        strBelowMean = "Health below average"
        strXaxis = "RAND SF-36 Categories"
        strYaxis = "Norm based RAND SF-36 Scale Scores, 50 = Norm"
        strMax = "Best possible value"
        strMin = "Worst possible value"
        strPlussSD = "+1 Standard Deviation"
        strMinusSD = "-1 Standard Deviation"
    Else
        strChartTitle = SelectedUser & ": Allmenntilstand kategorisert"
        strNormdata = "Gjennomsnitt for befolkningen"
        strBelowMean = "Helse under gjennomsnittet"
        strXaxis = "RAND SF-36 Kategorier"
        strYaxis = "Norm baserte RAND SF-36 Verdier, 50 = Norm"
        strMax = "Høyest mulige verdi"
        strMin = "Lavest mulige verdi"
        strPlussSD = "+1 Standardavvik"
        strMinusSD = "-1 Standardavvik"
    End If
    
    Call get_scale_label_text_sorted

    Set objSheet = Worksheets("SurveySummary")
    'clean up, delete old charts
    x = objSheet.ChartObjects.Count
    If x > 0 Then
        objSheet.ChartObjects.Delete
    End If
    x = objSheet.Pictures.Count
    If x > 0 Then
        objSheet.Pictures.Delete
    End If
    
    'create empty chart
    'Set objChart = objSheet.ChartObjects.Add(0, 0, frmUserFormSF36.Width, frmUserFormSF36.Height).chart
    Set objChart = objSheet.ChartObjects.Add(0, 0, Image2.Width, Image2.Height).chart
    objSheet.ChartObjects(1).Activate
    With ActiveChart
     .ChartType = xlColumnClustered
     .HasTitle = True
     .ChartTitle.Text = strChartTitle
     
     .Axes(xlCategory, xlPrimary).HasTitle = True
     .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = strXaxis
    
     .Axes(xlValue, xlPrimary).HasTitle = True
     .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = strYaxis
     
     .Axes(xlValue).MinimumScale = 0
     .Axes(xlValue).MaximumScale = 100
     '.Axes(xlCategory).TickLabels.Orientation = xlTickLabelOrientationHorizontal
     .Refresh
    End With
    
    Dim rowIx As Long
    Dim rowCount As Long
    Dim strUser As String
    Dim strRange As String
    Dim serieCount As Integer
    Dim rrowCount As Long
    Dim dValues As Range, sdValues As Range, sumValues As Range
    Dim dMin As Double, dMax As Double
    Dim max_values(1 To 11) As Double, min_values(1 To 11) As Double
    Dim mean_values(1 To 11) As Double, sd_values(1 To 11) As Double
    'Dim max_values(1 To 10) As Double, min_values(1 To 10) As Double
    'Dim mean_values(1 To 10) As Double, sd_values(1 To 10) As Double
   

    Dim i As Integer
    
    Dim recalculated_summary_mean As Double
    Dim z_from_norm_based_mean As Double
    Dim raw_mean_score As Double
    
    dMin = 0
    dMax = 100
    rowCount = Range("C1").CurrentRegion.Rows.Count
    serieCount = 0
    For rowIx = 2 To rowCount
        strUser = Cells(rowIx, 3)
        If UCase(strUser) = UCase(SelectedUser) Then
        ' adding series
            serieCount = serieCount + 1
            rrowCount = rowIx
            objChart.SeriesCollection.NewSeries
            objChart.SeriesCollection(serieCount).Name = Cells(rowIx, 7) 'Survey date
            'objChart.SeriesCollection(serieCount).XValues = objSheet.Range("AD1:AN1") 'Group labels
            objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
            strRange = "AD" & rowIx & ":AN" & rowIx
            'strRange = "AD" & rowIx & ":AM" & rowIx
            objChart.SeriesCollection(serieCount).Values = objSheet.Range(strRange)
        End If
    Next rowIx
    If serieCount > 0 Then
    ' adding reference values
        serieCount = serieCount + 1
        objChart.SeriesCollection.NewSeries
        objChart.SeriesCollection(serieCount).Name = strNormdata
        objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
        ' add mean=50
        For i = 1 To 11 '10 '11
            mean_values(i) = 50
            sd_values(i) = 10
        Next i
        objChart.SeriesCollection(serieCount).Values = mean_values
        objChart.SeriesCollection(serieCount).Type = xlLine
        objChart.SeriesCollection(serieCount).MarkerStyle = xlNone
        
        serieCount = serieCount + 1
        objChart.SeriesCollection.NewSeries
        objChart.SeriesCollection(serieCount).Name = strBelowMean
        objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
        ' add below mean=45
        Dim below_mean(1 To 11) As Double
        'Dim below_mean(1 To 10) As Double
        For i = 1 To 11 '10 '11
            below_mean(i) = 45
        Next i
        objChart.SeriesCollection(serieCount).Format.Line.DashStyle = msoLineLongDash
        objChart.SeriesCollection(serieCount).Values = below_mean
        objChart.SeriesCollection(serieCount).Type = xlLine
        objChart.SeriesCollection(serieCount).MarkerStyle = xlNone
        objChart.SeriesCollection(serieCount).Border.Color = RGB(255, 0, 0) ' pure red
        'rgb(0, 0, 0) 'black
        'RGB(255, 0, 0) ' pure red
        'RGB(0, 0, 255) ' pure blue
        'RGB(128, 128, 128) ' middle gray
        

        
        If CheckBox1.Value = True Then
        
            For i = 1 To 11 '10 '11
                max_values(i) = mean_values(i) + sd_values(i)
                min_values(i) = mean_values(i) - sd_values(i)
            Next i
            'For i = 9 To 11
            '    max_values(i) = 50
            '    min_values(i) = 50
            'Next i
            
            serieCount = serieCount + 1
            objChart.SeriesCollection.NewSeries
            objChart.SeriesCollection(serieCount).Name = strPlussSD
            objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
            objChart.SeriesCollection(serieCount).Values = max_values
            objChart.SeriesCollection(serieCount).Type = xlLine
            objChart.SeriesCollection(serieCount).MarkerStyle = xlNone
            objChart.SeriesCollection(serieCount).Format.Line.DashStyle = msoLineRoundDot
            
            
            serieCount = serieCount + 1
            objChart.SeriesCollection.NewSeries
            objChart.SeriesCollection(serieCount).Name = strMinusSD
            objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
            objChart.SeriesCollection(serieCount).Values = min_values
            objChart.SeriesCollection(serieCount).Type = xlLine
            objChart.SeriesCollection(serieCount).MarkerStyle = xlNone
            objChart.SeriesCollection(serieCount).Format.Line.DashStyle = msoLineRoundDot
            
            
            strRange = "AZ" & rrowCount & ":BJ" & rrowCount
            'strRange = "AZ" & rrowCount & ":BI" & rrowCount
            Set dValues = objSheet.Range(strRange)
            strRange = "BK" & rrowCount & ":BU" & rrowCount
            'strRange = "BK" & rrowCount & ":BT" & rrowCount
            Set sdValues = objSheet.Range(strRange)
            For i = 1 To 11 '10 '11
                mean_values(i) = dValues.Cells(i).Value
                sd_values(i) = sdValues.Cells(i).Value
            Next i
            
            
            
            For i = 1 To 8 '11 '8 '11
                max_values(i) = 50 + Round((10 * _
                                (100 - mean_values(i)) / sd_values(i)), 2)
                If max_values(i) > dMax Then
                    dMax = max_values(i)
                End If
                min_values(i) = 50 + Round((10 * _
                                (0 - mean_values(i)) / sd_values(i)), 2)
                If min_values(i) < dMin Then
                    dMin = min_values(i)
                End If
            Next i
            'For i = 9 To 10 ' ?
            '    z_from_norm_based_mean = (mean_values(i) - 50) / 10 ' from norm based t-value to z-value
            '    max_values(i) = 50 + Round((((100 / sd_values(i)) - (10 * z_from_norm_based_mean))), 2)
            '    If max_values(i) > dMax Then
            '        dMax = max_values(i)
            '    End If
            '    min_values(i) = 50 + Round((10 * (0 - z_from_norm_based_mean)), 2)
            '    If min_values(i) < dMin Then
            '        dMin = min_values(i)
            '    End If
            'Next i
            For i = 9 To 11
                max_values(i) = 50
                min_values(i) = 50
            Next i
            
   
            
            serieCount = serieCount + 1
            objChart.SeriesCollection.NewSeries
            objChart.SeriesCollection(serieCount).Name = strMax
            objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
            objChart.SeriesCollection(serieCount).Values = max_values
            objChart.SeriesCollection(serieCount).Type = xlLine
            objChart.SeriesCollection(serieCount).MarkerStyle = xlNone
            objChart.SeriesCollection(serieCount).Format.Line.DashStyle = msoLineDash
            
            serieCount = serieCount + 1
            objChart.SeriesCollection.NewSeries
            objChart.SeriesCollection(serieCount).Name = strMin
            objChart.SeriesCollection(serieCount).XValues = scale_group_label_text_sorted
            objChart.SeriesCollection(serieCount).Values = min_values
            objChart.SeriesCollection(serieCount).Type = xlLine
            objChart.SeriesCollection(serieCount).MarkerStyle = xlNone
            objChart.SeriesCollection(serieCount).Format.Line.DashStyle = msoLineDash
            
            Dim dMinTick As Double
            dMinTick = (dMin - 1) / 10
            dMinTick = Round(dMinTick, 0)
            dMinTick = dMinTick * 10
            objChart.Axes(xlValue).MinimumScale = dMinTick
            objChart.Axes(xlValue).MaximumScale = dMax
            'objChart.Axes(xlValue).MinorUnit = 10
            'objChart.Axes(xlValue).MajorUnit = 100

            
            
       
        
        End If
    End If
    ActiveChart.Refresh
    
    Worksheets("SurveySummary").ChartObjects(1).CopyPicture
    ActiveSheet.Paste
    objSheet.ChartObjects(1).Activate
    
    Dim strPath As String
    x = objSheet.Pictures.Count
    If x > 0 Then
        'objSheet.Pictures.Activate
        'Set myPic = objSheet.Pictures(x)
        strPath = strOneDriveLocalFilePath
        strPath = strPath & "\" & "tmp.bmp"
        'strPath = ActiveWorkbook.Path & "\" & "tmp.bmp"
        ActiveChart.Export strPath
        'frmUserFormSF36.Picture = LoadPicture(strPath)
        Image2.Picture = LoadPicture(strPath)
     End If
 
    'clean up, delete old charts, but keep picture of graph
    x = objSheet.ChartObjects.Count
    If x > 0 Then
        objSheet.ChartObjects.Delete
    End If
    Me.MousePointer = fmMousePointerDefault

Exit Sub
Errhandler:
      Me.MousePointer = fmMousePointerDefault
      ErrorHandling "frmUserFormsSF36. Sub chartNormScores", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


Private Sub CheckBox1_Change()
    If CheckBox2.Value = False Then
        Call chartScaleScores
    Else
        Call chartNormScores
    End If
End Sub

Private Sub CheckBox2_Change()
    If CheckBox2.Value = False Then
        Call chartScaleScores
    Else
        Call chartNormScores
    End If
End Sub

Private Sub UserForm_Initialize()
On Error GoTo Errhandler
    
    Call initCaptions
    'Call chartScaleScores
   
Exit Sub
Errhandler:
      ErrorHandling "frmUserFormsSF36. Sub UserForm_Initialize", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub


