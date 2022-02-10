VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegisterNewUser 
   Caption         =   "Register new user"
   ClientHeight    =   3090
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   10060
   OleObjectBlob   =   "frmRegisterNewUser-2022-4-0-2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRegisterNewUser"
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
' frmRegisterNewUser.frm
'
' 2015-04-13 1.0.0 Veronika Lindberg    Created.
'                                       Register a new user.
' 2015-06-28 1.0.1 Veronika Lindberg    User code can not consist of numbers only.
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

    Dim strHeader3 As String
    Dim strUserExist1 As String
    Dim strUserExist2 As String
    Dim strUser1 As String
    Dim strUser2 As String
    Dim strUser3 As String
    Dim strNotNumbers As String


Private Sub cmdSaveNewUser_Click()
On Error GoTo Errhandler
    Dim myValue As Variant
    Dim myYear As Variant
    Dim myGender As String
    Dim myGenderCode As String
    Dim rowCount As Long
    Dim c As Variant
    Dim myRange As Variant
    Dim vt1 As Variant
    Dim vt2 As Variant
    
    'myValue = VBA.InputBox(strEnter1, strHeader1) 'Enter user code or name
    ' Check if user hit the cancel button
    'If StrPtr(myValue) = 0 Then
        'MsgBox "User hit cancel"
    '    Exit Sub
    'End If
    'myValue = Trim(myValue)
    'If myValue = "" Then
        'MsgBox "No value entered"
    '    Exit Sub
    'End If
    
    myValue = TextBox1.Value
    
    
    vt1 = VarType(myValue)
    If IsNumeric(myValue) Then
        MsgBox strUser3, vbExclamation, strUser1
        TextBox1.SetFocus
        Exit Sub
    End If
    
    If Not vt1 = vbString Then
        myValue = CStr(myValue)
    End If
    
    If myValue = "" Then
        MsgBox strUser1, vbExclamation, strUser2
        TextBox1.SetFocus
        Exit Sub
    End If
    myValue = Trim(myValue)
    myValue = Replace(myValue, " ", "")
    
    Worksheets("Users").Activate
    rowCount = Range("A1").CurrentRegion.Rows.Count
    myRange = "A1:" & "A" & Trim(Str(rowCount))
    With Worksheets("Users").Range(myRange)
        Set c = .Find(myValue)
    End With
    If Not c Is Nothing Then
        vt2 = VarType(c)
        If Not vt2 = vbString Then
            c = CStr(c)
        End If
        If UCase(c) = UCase(myValue) Then
            MsgBox strUserExist1 & myValue & strUserExist2, , strHeader3
            '"Det finnes allerede en bruker med navn eller kode " & Str(myValue) & "."
            TextBox1.SetFocus
            Exit Sub
        End If
    End If
    
    'myYear = VBA.InputBox(strEnter2, strHeader2) 'Skriv inn brukerens fødselsår
    ' Check if user hit the cancel button
    'If StrPtr(myYear) = 0 Then
        'MsgBox "User hit cancel"
    '    Exit Sub
    'End If
    'myYear = Trim(myYear)
    'If myYear = "" Then
        'MsgBox "No value entered"
    '    Exit Sub
    'End If
    
    myGender = ComboBox1.Value
    vt1 = VarType(myGender)
    If Not vt1 = vbString Then
        myGender = CStr(myGender)
    End If
    myGender = Trim(myGender)
    ' Unknown gender is allowed
    If ComboBox1.ListIndex = 0 Then
        myGenderCode = -1
    ElseIf ComboBox1.ListIndex = 1 Then
        myGenderCode = 1 'Female
    ElseIf ComboBox1.ListIndex = 2 Then
        myGenderCode = 0 'Male
    End If
    ' Removed unknown
    'If ComboBox1.ListIndex = 0 Then
    '    myGenderCode = 1
    'ElseIf ComboBox1.ListIndex = 1 Then
    '    myGenderCode = 0
    'End If
    
    myYear = ComboBox2.Value
    vt1 = VarType(myYear)
    If Not vt1 = vbString Then
        myYear = CStr(myYear)
    End If
    myYear = Trim(myYear)
    
    ' Save new user
     With Range("A1")
        .Offset(rowCount, 0).Value = CStr(myValue)
        .Offset(rowCount, 1).Value = CStr(myYear)
        .Offset(rowCount, 2).Value = CStr(myGender)
        .Offset(rowCount, 3).Value = CStr(myGenderCode)
        .Offset(rowCount, 4).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    End With
    'MsgBox "En ny bruker med navn eller kode " & Str(myValue) & " er registrert."
    SelectedUser = myValue
    SelectedBirthYear = myYear
    SelectedGender = myGender
    SelectedGenderCode = myGenderCode
    
    Call populate_users
    Call populate_surveys
    
    'Worksheets.Add(after:=Worksheets(Worksheets.Count)).Name = SelectedUser
    ThisWorkbook.Save
    
    Unload Me
Exit Sub
Errhandler:
      ErrorHandling "frmRegisterNewUser. Sub cmdSaveNewUser_Click", Err, Action
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
    Dim vnow As Integer
    Dim vStart As Integer
    Dim i As Integer
    
    vnow = Format(Now, "yyyy")
    vStart = vnow - 150
    'MsgBox CStr(vnow) & " " & CStr(vStart)
    For i = vStart To vnow
        ComboBox2.AddItem i
    Next i
    If SelectedLanguage = "UK" Then
        frmRegisterNewUser.Caption = "Register a new person"
        Label1.Caption = "Enter name:" '"Enter user code or name:"
        Label2.Caption = "Gender (at birth):"
        Label3.Caption = "Year of birth:"
        ComboBox1.AddItem "Prefer not to say"
        ComboBox1.AddItem "Male"
        ComboBox1.AddItem "Female"
        ComboBox1.ListIndex = 0
        ComboBox2.AddItem "Unknown"
        ComboBox2.ListIndex = ComboBox2.ListCount - 1
        strUserExist1 = "An person with name or code '"
        strUserExist2 = "' already exist."
        strHeader3 = "Name already exist"
        strUser1 = "Please enter name" '"Please enter user code or name."
        strUser2 = "Name is missing" '"User code or name is missing"
        strUser3 = "Name cannot be a number"
        cmdSaveNewUser.Caption = "Save"
    Else
        frmRegisterNewUser.Caption = "Registrer en ny person"
        Label1.Caption = "Skriv inn navn:" '"Bruker kode eller navn:"
        Label2.Caption = "Kjønn (ved fødselen):"
        Label3.Caption = "Fødselsår:"
        ComboBox1.AddItem "Ønsker ikke å oppgi"
        ComboBox1.AddItem "Mann"
        ComboBox1.AddItem "Kvinne"
        ComboBox1.ListIndex = 0
        ComboBox2.AddItem "Ukjent"
        ComboBox2.ListIndex = ComboBox2.ListCount - 1
        strUserExist1 = "Det finnes allerede en person registrert med navn eller kode '"
        strUserExist2 = "'."
        strHeader3 = "Navnet finnes allerede"
        strUser1 = "Vennligst fyll in navn" '"Vennligst fyll inn brukerkode eller navn."
        strUser2 = "Navn mangeler" '"Brukerkode eller navn mangler"
        strUser3 = "Navn kan ikke være et nummer"
        cmdSaveNewUser.Caption = "Lagre"
    End If
    TextBox1.SetFocus
    Me.Repaint
Exit Sub
Errhandler:
      ErrorHandling "frmRegisterNewUser. Sub UserForm_Initialize", Err, Action
      If Action = Err_Exit Then
         Exit Sub
      ElseIf Action = Err_Resume Then
         Resume
      Else
         Resume Next
      End If
End Sub
