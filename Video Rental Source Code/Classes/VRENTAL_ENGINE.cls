VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VRENTAL_ENGINE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Option Base 1
Function CheckPermission(Level As Integer, Fcode As Integer) As Boolean
Dim loop1, loop2, counter As Integer
Dim str As String
Dim db As Database
Dim rec As Recordset
Set db = OpenDatabase(App.Path & "\Permission.mdb" _
             , False, False, ";pwd=AdmiN")
Set rec = db.OpenRecordset("PermissionTable", dbOpenTable)
counter = 0
    rec.MoveFirst
For loop1 = 1 To Level
    str = Trim(rec.Fields("Permissions"))
    For loop2 = 1 To 7
        If loop2 = Fcode And loop1 = Level Then
           If Int(Val(Mid(str, loop2, 1))) = 1 Then
              CheckPermission = True
              Exit Function
           Else
              MsgBox "You don't have the permission to use this feature.  ", vbInformation, "Access Denied"
              CheckPermission = False
              Exit Function
           End If
        End If
        counter = counter + 1
    Next loop2
    If rec.EOF = False Then rec.MoveNext
Next loop1

Set db = Nothing
Set rec = Nothing
End Function
Function Date_GETDATE(NumberOfDaysFromJanYear1 As Long) As String
    Dim YearCount As Integer
    Dim Days, i, iMonth(12) As Integer
    Dim GET_Day As Long
    Dim strDay As String
    'Get Day
     GET_Day = NumberOfDaysFromJanYear1 Mod 7
     
   Select Case GET_Day
    Case 1: strDay = "Sunday"
    Case 2: strDay = "Monday"
    Case 3: strDay = "Tuesday"
    Case 4: strDay = "Wednesday"
    Case 5: strDay = "Thursday"
    Case 6: strDay = "Friday"
    Case 0: strDay = "Saturday"
   End Select
    
    iMonth(1) = 31 'Jan
    iMonth(2) = 28 'Feb
    iMonth(3) = 31 'Mar
    iMonth(4) = 30 'Apr
    iMonth(5) = 31 'May
    iMonth(6) = 30 'Jun
    iMonth(7) = 31 'Jul
    iMonth(8) = 31 'Aug
    iMonth(9) = 30 'Sep
    iMonth(10) = 31 'Oct
    iMonth(11) = 30 'Nov
    iMonth(12) = 31 'Dec
    
    'Get Year
    YearCount = 1
    Days = 365
    Do While NumberOfDaysFromJanYear1 > Days
       NumberOfDaysFromJanYear1 = NumberOfDaysFromJanYear1 - Days
       If NumberOfDaysFromJanYear1 > 0 Then YearCount = YearCount + 1
       If YearCount Mod 4 = 0 Then
          Days = 366
       Else
          Days = 365
       End If
    Loop
    
    'Get Month and Day
    If YearCount Mod 4 = 0 Then iMonth(2) = 29
    If NumberOfDaysFromJanYear1 <> 0 Then
      i = 1
      Do While NumberOfDaysFromJanYear1 > iMonth(i)
         NumberOfDaysFromJanYear1 = NumberOfDaysFromJanYear1 - iMonth(i)
         i = i + 1
      Loop
    Else
      i = 12
      NumberOfDaysFromJanYear1 = 31
    End If
    
   Date_GETDATE = str(i) & "/" & Trim(str(NumberOfDaysFromJanYear1)) & "/" & Trim(str(YearCount))
    
End Function
Function Date_CountNumberOfDaysFromJan1Year1ToDec31YearEntered(YearEntered As Long) As Long
    Dim TotalDays, Year As Long
    Dim Days, DaysInAYear, counter As Integer

    TotalDays = 730500 ' No. Of days from 1/1/1 to 12/31/2000
    counter = 1
    DaysInAYear = 365
'730499
    For Year = 2001 To YearEntered
        For Days = 1 To DaysInAYear
            TotalDays = TotalDays + 1
        Next Days
        counter = counter + 1
        If counter = 4 Then  ' Leap Year
            counter = 0
            DaysInAYear = 366
        Else
            DaysInAYear = 365
        End If
    Next Year

    Date_CountNumberOfDaysFromJan1Year1ToDec31YearEntered = TotalDays

End Function
Function Date_CountDaysInAYear(DateEntered As String) As Integer
Dim DateDay, DateMonth, DateYear, iMonth(12), _
    loop1, TotalDays As Integer

iMonth(1) = 31 'Jan
iMonth(2) = 28 'Feb
iMonth(3) = 31 'Mar
iMonth(4) = 30 'Apr
iMonth(5) = 31 'May
iMonth(6) = 30 'Jun
iMonth(7) = 31 'Jul
iMonth(8) = 31 'Aug
iMonth(9) = 30 'Sep
iMonth(10) = 31 'Oct
iMonth(11) = 30 'Nov
iMonth(12) = 31 'Dec

   DateYear = Year(DateEntered)
   DateMonth = Month(DateEntered)
   DateDay = Day(DateEntered)
'Check if year is leapyear
If DateYear Mod 4 = 0 Then iMonth(2) = 29
TotalDays = 0
For loop1 = 1 To DateMonth
   If loop1 = DateMonth Then TotalDays = TotalDays + DateDay
   If loop1 < DateMonth Then TotalDays = TotalDays + iMonth(loop1)
Next loop1

Date_CountDaysInAYear = TotalDays

End Function
Sub CopyFlexDataToExcel(Flex As MSFlexGrid)
On Error GoTo ErrHandler
Dim EXCELApp As Excel.Application
Dim EXCELWorkBook As Excel.Workbook
Dim Rows, Cols As Integer
Dim iRow, hCol, iCol As Integer
Dim New_Col As Boolean


    If Flex.Rows <= 1 Then
        MsgBox "No Data to extract", vbInformation, App.Title
        Exit Sub
    End If
    
    Set EXCELApp = CreateObject("Excel.application")
    Set EXCELWorkBook = EXCELApp.Workbooks.Add
    
    Dim New_Column As Boolean
    
    With Flex
        Rows = .Rows
        Cols = .Cols
        iRow = 0
        iCol = 1
        For hCol = 0 To Cols - 1
            For iRow = 1 To Rows
                EXCELApp.Cells(iRow + 1, iCol + 1).Value = .TextMatrix(iRow - 1, hCol)
            Next iRow
            iCol = iCol + 1
        Next hCol
    End With
    
    EXCELApp.Rows(2).Font.Bold = True
    EXCELApp.Cells.Select
    EXCELApp.Columns.AutoFit
    EXCELApp.Cells(1, 1).Select
    EXCELApp.Application.Visible = True
    
    Set EXCELWorkBook = Nothing
    Set EXCELApp = Nothing
    Flex.SetFocus
    MsgBox "Data has been successfully copied to Excel. ", vbInformation, "Success"
Exit Sub
ErrHandler:
  MsgBox "Unable to open EXCEL. ", vbInformation, "Info "
End Sub
Public Function ReplaceString(ByVal TextString As String, ByVal FromString As String, ByVal ToString As String)
Dim new_TextString As String
Dim Position As Integer

    Do While Len(TextString) > 0
        Position = InStr(TextString, FromString)
        If Position = 0 Then
            new_TextString = new_TextString & TextString
            TextString = ""
        Else
            new_TextString = new_TextString & Left$(TextString, Position - 1) & ToString
            TextString = Mid$(TextString, Position + Len(FromString))
        End If
    Loop

    ReplaceString = new_TextString
End Function
Function LogOnValidate(UserName As String, Password As String) As String

    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer

    LogOnValidate = 0
    
    Set db = OpenDatabase(App.Path & "\Database\UsersDB.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("Users", dbOpenTable)
    
    rec.MoveFirst
    If rec.RecordCount > 0 Then
        For loop1 = 1 To rec.RecordCount
            TDM = DoEvents()
            If rec.Fields("User Name") = UserName And _
                rec.Fields("Password") = Password Then
               
                LogOnValidate = "User ID : " & rec.Fields("UserID") & vbCrLf _
                                & "Date Entered : " & rec.Fields("Date Entered") & vbCrLf _
                                & "User Name : " & rec.Fields("User Name") & vbCrLf _
                                & "Password : " & rec.Fields("Password") & vbCrLf _
                                & "Access Level : " & rec.Fields("AccessLevel") & vbCrLf _
                                & "First Name : " & rec.Fields("First Name") & vbCrLf _
                                & "Middle Name : " & rec.Fields("Middle Name") & vbCrLf _
                                & "Family Name : " & rec.Fields("Family Name") & vbCrLf _
                                & "Birthday : " & rec.Fields("Birthday") & vbCrLf & "Age : " & str(GetAge(rec.Fields("Birthday"))) & vbCrLf _
                                & "Sex : " & rec.Fields("Sex") & vbCrLf _
                                & "Home Address : " & rec.Fields("Home Address") & vbCrLf _
                                & "Contact Number : " & rec.Fields("Contact Number") & vbCrLf _
                                & "Comments : " & rec.Fields("Comments") & vbCrLf
                Exit For
            Else
                LogOnValidate = ""
            End If
            rec.MoveNext  ' Move to the next record
        Next loop1
    End If
    
    db.Close
    
End Function

Public Sub ChangePassword(NewPWD As String, UserName As String)
     
    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer


    
    Set db = OpenDatabase(App.Path & "\Database\UsersDB.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("Users", dbOpenTable)
    
    rec.MoveFirst
    If rec.RecordCount > 0 Then
        For loop1 = 1 To rec.RecordCount
            TDM = DoEvents()
            If rec.Fields("User Name") = UserName Then
                '' Start Change Password
                rec.Edit
                rec.Fields("Password") = NewPWD
                rec.Update
                '' End Change Password
            End If
        
            rec.MoveNext  ' Move to the next record
        Next loop1
     End If
     
      db.Close
End Sub
Function GetAge(myDate As Variant) As Integer
 Dim numyears, numMonths
 myDate = CDate(myDate)
 Dim TotalDays As Long
 TotalDays = DateDiff("y", myDate, Date)
 numyears = Abs(TotalDays / 365.25)
 numMonths = (numyears - Int(numyears)) * 365.25 / 30.4583
 GetAge = Int(numyears)
End Function

Function ReportFileStatus(filespec) As Boolean '' Check if file is present
   Dim fso, msg
   Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FileExists(filespec)) Then
      msg = -1
   Else
      msg = 0
   End If
   ReportFileStatus = msg
End Function
Function GetDayCount(FirstDate As String, SecondDate As String) As Integer
    GetDayCount = Abs(DateSerial(Year(SecondDate), Format(SecondDate, "MM"), Day(SecondDate)) - DateSerial(Year(FirstDate), Format(FirstDate, "MM"), Day(FirstDate)))
End Function
Function Round(RoundMe, RoundTo) As Double
    Round = Int((RoundMe * 10 ^ RoundTo) + 0.5) / 10 ^ RoundTo
   
End Function

Sub LoadUsers(lst As ListBox)  ''load UserNames in a listbox
    On Error GoTo ErrorHandler:
    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    lst.Clear
    
    Set db = OpenDatabase(App.Path & "\Database\UsersDB.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("Users", dbOpenTable)
    
    rec.MoveFirst
    If rec.RecordCount > 0 Then
        For loop1 = 1 To rec.RecordCount
            TDM = DoEvents()
            lst.AddItem rec.Fields("First Name") & " " & Left(rec.Fields("Middle Name"), 1) & ". " & _
            rec.Fields("Family Name") & " (" & rec.Fields("User Name") & ")"
            rec.MoveNext
        Next loop1
     End If
     
  db.Close
  Exit Sub
  
ErrorHandler:
    db.Close

End Sub
Sub getUserInfo(txtUserID As TextBox, txtDateEntered As TextBox, _
                txtUserName As TextBox, txtPassword As TextBox, _
                txtAccessLevel As TextBox, txtFirstName As TextBox, _
                txtMiddleName As TextBox, txtFamilyName As TextBox, _
                txtBirthday As TextBox, txtSex As TextBox, _
                txtHomeAddress As TextBox, txtContactNumber As TextBox, _
                txtComments As TextBox, txtConfirmPassword As TextBox, lst As ListBox)
On Error GoTo Err:
    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    
    Set db = OpenDatabase(App.Path & "\Database\UsersDB.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("Users", dbOpenTable)
    
If rec.BOF = True And rec.EOF = True Then Exit Sub 'chk if rec exists

rec.MoveFirst
    If rec.RecordCount > 0 Then
        For loop1 = 1 To rec.RecordCount
            TDM = DoEvents()
            If lst.Text = rec.Fields("First Name") & " " & Left(rec.Fields("Middle Name"), 1) & ". " & _
                          rec.Fields("Family Name") & " (" & rec.Fields("User Name") & ")" Then
                    txtUserID.Text = rec.Fields("UserID")
                    txtDateEntered.Text = Format(rec.Fields("Date Entered"), "mmm. d, yyyy")
                    txtUserName.Text = rec.Fields("User Name")
                    txtPassword.Text = rec.Fields("Password")
                    txtConfirmPassword.Text = txtPassword.Text
                    txtAccessLevel.Text = rec.Fields("AccessLevel")
                    txtFirstName.Text = rec.Fields("First Name")
                    txtMiddleName.Text = rec.Fields("Middle Name")
                    txtFamilyName.Text = rec.Fields("Family Name")
                    txtBirthday.Text = Format(rec.Fields("Birthday"), "mmm. d, yyyy")
                    txtSex.Text = rec.Fields("Sex")
                    txtHomeAddress.Text = rec.Fields("Home Address")
                    txtContactNumber.Text = rec.Fields("Contact Number")
                    txtComments.Text = rec.Fields("Comments")
                    db.Close
                    Exit Sub
            End If
            rec.MoveNext
        Next loop1
     End If
     
  db.Close
  
Err:
  txtDateEntered.Text = ""
  txtUserID.Text = ""

End Sub

Function UpdateEditedUsersDB(txtUserName As TextBox, txtPassword As TextBox, _
                txtAccessLevel As TextBox, txtFirstName As TextBox, _
                txtMiddleName As TextBox, txtFamilyName As TextBox, _
                txtBirthday As TextBox, txtSex As TextBox, _
                txtHomeAddress As TextBox, txtContactNumber As TextBox, _
                txtComments As TextBox, lst As ListBox) As Boolean
'On Error GoTo ErrorHandler
    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    
    Set db = OpenDatabase(App.Path & "\Database\UsersDB.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("Users", dbOpenTable)
    ''---------------------------------------------
    If rec.BOF = True And rec.EOF = True Then
    rec.AddNew
                 rec.Fields("Date Entered") = Date$
                 rec.Fields("UserID") = 1 '1st record
                 rec.Fields("User Name") = Trim(txtUserName.Text)
                 rec.Fields("Password") = Trim(txtPassword.Text)
                 rec.Fields("AccessLevel") = Trim(txtAccessLevel.Text)
                 rec.Fields("First Name") = Trim(txtFirstName.Text)
                 rec.Fields("Middle Name") = Trim(txtMiddleName.Text)
                 rec.Fields("Family Name") = Trim(txtFamilyName.Text)
                 rec.Fields("Birthday") = Trim(txtBirthday.Text)
                 rec.Fields("Sex") = Trim(txtSex.Text)
                 rec.Fields("Home Address") = Trim(txtHomeAddress.Text)
                 rec.Fields("Contact Number") = Trim(txtContactNumber.Text)
                 rec.Fields("Comments") = Trim(txtComments.Text)
                 rec.Update  ''Update the recordset
               '' End update fields
               db.Close  '' Close DB
    
    UpdateEditedUsersDB = True
    Exit Function
    End If
    ''---------------------------------------------
    rec.MoveFirst
    '' Start Chk for duplicates sql
    '' End -- chk duplicates
    If rec.RecordCount > 0 Then
        For loop1 = 1 To rec.RecordCount
            TDM = DoEvents()
            If lst.Text = rec.Fields("First Name") & " " & Left(rec.Fields("Middle Name"), 1) & ". " & _
                          rec.Fields("Family Name") & " (" & rec.Fields("User Name") & ")" Then
               '' Start update fields
              
                 rec.Edit
                 rec.Fields("User Name") = Trim(txtUserName.Text)
                 rec.Fields("Password") = Trim(txtPassword.Text)
                 rec.Fields("AccessLevel") = Trim(txtAccessLevel.Text)
                 rec.Fields("First Name") = Trim(txtFirstName.Text)
                 rec.Fields("Middle Name") = Trim(txtMiddleName.Text)
                 rec.Fields("Family Name") = Trim(txtFamilyName.Text)
                 rec.Fields("Birthday") = Trim(txtBirthday.Text)
                 rec.Fields("Sex") = Trim(txtSex.Text)
                 rec.Fields("Home Address") = Trim(txtHomeAddress.Text)
                 rec.Fields("Contact Number") = Trim(txtContactNumber.Text)
                 rec.Fields("Comments") = Trim(txtComments.Text)
                 rec.Update  ''Update the recordset
               '' End update fields
               db.Close  '' Close DB
               UpdateEditedUsersDB = True
               Exit Function
            End If
            rec.MoveNext
        Next loop1
     End If
     
  db.Close
  UpdateEditedUsersDB = True
  Exit Function
  
ErrorHandler:
  MsgBox "Your chosen User Name is already in use. " & vbCrLf & vbCrLf & "Change your User Name. ", vbInformation, "Update Error"
UpdateEditedUsersDB = False
End Function

Function AddUserToDB(txtDateEntered As TextBox, txtUserName As TextBox, txtPassword As TextBox, _
                txtAccessLevel As TextBox, txtFirstName As TextBox, _
                txtMiddleName As TextBox, txtFamilyName As TextBox, _
                txtBirthday As TextBox, txtSex As TextBox, _
                txtHomeAddress As TextBox, txtContactNumber As TextBox, _
                txtComments As TextBox) As Boolean
                
    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    Dim AutoNumber As Long
    AutoNumber = 0 ' init to zero
    Dim Valid As Boolean
    Valid = False  ' Init to False
    
    Set db = OpenDatabase(App.Path & "\Database\UsersDB.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("Users", dbOpenTable)
     '' Start -- Check for valid "New ID" to avoid duplicates
     
        If rec.BOF = True And rec.EOF = True Then
                AutoNumber = 1
        Else
        
                Do
                    AutoNumber = AutoNumber + 1
                    rec.MoveFirst
                    Do While (rec.EOF = False)
                        TDM = DoEvents()
                        If Val(rec.Fields("UserID")) = AutoNumber Then
                                Valid = False
                                Exit Do
                        Else
                                Valid = True
                        End If
                        If rec.EOF = False Then rec.MoveNext
                    Loop
                Loop Until Valid = True
        End If
     '' End -- Check for valid "New ID"
     
     '' Start check for username duplicate
                    rec.MoveFirst
                    Do While (rec.EOF = False)
                        TDM = DoEvents()
                        If rec.Fields("User Name") = Trim(txtUserName.Text) Then
                                MsgBox "Someone has already chosen your User Name. " & vbCrLf & vbCrLf & "Change your User Name. ", vbInformation, "Unable to add user"
                                db.Close
                                txtUserName.SetFocus
                                AddUserToDB = False
                                Exit Function
                        End If
                        If rec.EOF = False Then rec.MoveNext
                    Loop
     '' End check for username duplicate
                '' Start update new record fields
                 rec.AddNew
                 rec.Fields("UserID") = AutoNumber
                 rec.Fields("Date Entered") = Format(txtDateEntered.Text, "mm-dd-yyyy")
                 rec.Fields("User Name") = Trim(txtUserName.Text)
                 rec.Fields("Password") = Trim(txtPassword.Text)
                 rec.Fields("AccessLevel") = Trim(txtAccessLevel.Text)
                 rec.Fields("First Name") = Trim(txtFirstName.Text)
                 rec.Fields("Middle Name") = Trim(txtMiddleName.Text)
                 rec.Fields("Family Name") = Trim(txtFamilyName.Text)
                 rec.Fields("Birthday") = Trim(txtBirthday.Text)
                 rec.Fields("Sex") = Trim(txtSex.Text)
                 rec.Fields("Home Address") = Trim(txtHomeAddress.Text)
                 rec.Fields("Contact Number") = Trim(txtContactNumber.Text)
                 rec.Fields("Comments") = Trim(txtComments.Text)
                 rec.Update  ''Update the recordset
               '' End update new record fields
               db.Close  '' Close DB
               MsgBox "Another user has been successfully added. ", vbInformation, "Success"
               AddUserToDB = True
End Function

Sub RemoveUser(UserID As Long)
On Error GoTo Err:
 Dim TDM As Variant
 Dim db As Database
 Dim rec As Recordset
 
 Set db = OpenDatabase(App.Path & "\Database\UsersDB.mdb" _
             , False, False, ";pwd=AdmiN")
 Set rec = db.OpenRecordset("Users", dbOpenTable)
 rec.MoveFirst
    Do While (rec.EOF = False)
       TDM = DoEvents()
       If Val(rec.Fields("UserID")) = UserID Then
           rec.Delete
           Exit Do
       End If
       If rec.EOF = False Then rec.MoveNext
    Loop
 MsgBox "User has been deleted. ", vbInformation, "Removed"
 db.Close

Err:
End Sub

Sub LoadMembers(lst As ListBox) ''load UserNames in a listbox
    On Error GoTo ErrorHandler:
    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    lst.Clear
    
    Set db = OpenDatabase(App.Path & "\Database\MembersDB.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("MembersInfo", dbOpenTable)
    
    rec.MoveFirst
    If rec.RecordCount > 0 Then
        For loop1 = 1 To rec.RecordCount
            TDM = DoEvents()
            'lst.AddItem rec.Fields("First Name") & " " & Left(rec.Fields("Middle Name"), 1) & ". " & _
            rec.Fields("Family Name")
            lst.AddItem rec.Fields("Family Name") & ", " & rec.Fields("First Name") & " " & Left(rec.Fields("Middle Name"), 1) & "."
            rec.MoveNext
        Next loop1
     End If
     
  db.Close
  Exit Sub
  
ErrorHandler:
  db.Close
End Sub

Sub GetMemberInfo(lst As ListBox, txtDateEntered As TextBox, txtIDnumber As TextBox, txtNationality As TextBox, _
                  txtFirstName As TextBox, txtMiddleName As TextBox, txtFamilyName As TextBox, _
                  txtBirthday As TextBox, txtSex As TextBox, txtCivilStatus As TextBox, txtOccupation As TextBox, _
                  txtContactNumber As TextBox, txtHomeAddress As TextBox, txtOfficeSchoolAddress As TextBox, txtComments As TextBox)     'Load member info to textboxes


    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    
    Set db = OpenDatabase(App.Path & "\Database\MembersDB.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("MembersInfo", dbOpenTable)
    
    If rec.BOF = True And rec.EOF = True Then Exit Sub 'chk if rec exists

rec.MoveFirst
    If rec.RecordCount > 0 Then
        For loop1 = 1 To rec.RecordCount
            TDM = DoEvents()
            If lst.Text = rec.Fields("Family Name") & ", " & rec.Fields("First Name") & " " & Left(rec.Fields("Middle Name"), 1) & "." Then
                    txtDateEntered.Text = rec.Fields("Date Entered")
                    txtIDnumber.Text = rec.Fields("ID NUMBER")
                    txtNationality.Text = rec.Fields("Membership Level")
                    txtFirstName.Text = rec.Fields("First Name")
                    txtMiddleName.Text = rec.Fields("Middle Name")
                    txtFamilyName.Text = rec.Fields("Family Name")
                    txtBirthday.Text = Format(rec.Fields("Birthday"), "mmm. dd, yyyy")
                    txtSex.Text = rec.Fields("Sex")
                    txtCivilStatus.Text = rec.Fields("Civil Status")
                    txtOccupation.Text = rec.Fields("Occupation")
                    txtContactNumber.Text = rec.Fields("Contact Number")
                    txtHomeAddress.Text = rec.Fields("Home Address")
                    txtOfficeSchoolAddress.Text = rec.Fields("OfficeOrSchool/Address")
                    txtComments.Text = rec.Fields("Comments")
                    db.Close
                    Exit Sub
            End If
            If rec.EOF = False Then rec.MoveNext
        Next loop1
     End If
     
  db.Close
  Exit Sub


End Sub

Function AddMemberToDB(txtDateEntered As TextBox, txtIDnumber As TextBox, txtNationality As TextBox, _
                  txtFirstName As TextBox, txtMiddleName As TextBox, txtFamilyName As TextBox, _
                  txtBirthday As TextBox, txtSex As TextBox, txtCivilStatus As TextBox, txtOccupation As TextBox, _
                  txtContactNumber As TextBox, txtHomeAddress As TextBox, txtOfficeSchoolAddress As TextBox, txtComments As TextBox) As Boolean
    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    Dim AutoNumber As Long
    AutoNumber = 0 ' init to zero
    Dim Valid As Boolean
    Valid = False  ' Init to False
    
    Set db = OpenDatabase(App.Path & "\Database\MembersDB.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("MembersInfo", dbOpenTable)
     '' Start -- Check for valid "New ID" to avoid duplicates
     
        If rec.BOF = True And rec.EOF = True Then
                AutoNumber = 1
        Else
        
                Do
                    AutoNumber = AutoNumber + 1
                    rec.MoveFirst
                    Do While (rec.EOF = False)
                        TDM = DoEvents()
                        If Val(rec.Fields("ID NUMBER")) = AutoNumber Then
                                Valid = False
                                Exit Do
                        Else
                                Valid = True
                        End If
                        If rec.EOF = False Then rec.MoveNext
                    Loop
                Loop Until Valid = True
        End If
     '' End -- Check for valid "New ID"
     
     
                '' Start update new record fields
                 rec.AddNew
                 rec.Fields("Date Entered") = Format(Trim(txtDateEntered.Text), "mm/dd/yyyy")
                 rec.Fields("ID NUMBER") = AutoNumber
                 rec.Fields("Membership Level") = Trim(txtNationality.Text)
                 rec.Fields("First Name") = Trim(txtFirstName.Text)
                 rec.Fields("Middle Name") = Trim(txtMiddleName.Text)
                 rec.Fields("Family Name") = Trim(txtFamilyName.Text)
                 rec.Fields("Birthday") = Format(Trim(txtBirthday.Text), "mm/dd/yyyy")
                 rec.Fields("Sex") = Trim(txtSex.Text)
                 rec.Fields("Civil Status") = (txtCivilStatus.Text)
                 rec.Fields("Occupation") = Trim(txtOccupation.Text)
                 rec.Fields("Contact Number") = Trim(txtContactNumber.Text)
                 rec.Fields("Home Address") = Trim(txtHomeAddress.Text)
                 rec.Fields("OfficeOrSchool/Address") = Trim(txtOfficeSchoolAddress.Text)
                 rec.Fields("Comments") = Trim(txtComments.Text)
                 rec.Update
               '' End update new record fields
               db.Close  '' Close DB
               MsgBox "Another member has been successfully added. ", vbInformation, "Success"
               AddMemberToDB = True

End Function

Function UpdateEditedMembersDB(txtDateEntered As TextBox, txtIDnumber As TextBox, txtNationality As TextBox, _
                  txtFirstName As TextBox, txtMiddleName As TextBox, txtFamilyName As TextBox, _
                  txtBirthday As TextBox, txtSex As TextBox, txtCivilStatus As TextBox, txtOccupation As TextBox, _
                  txtContactNumber As TextBox, txtHomeAddress As TextBox, txtOfficeSchoolAddress As TextBox, txtComments As TextBox) As Boolean
  'On Error GoTo Err:
    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    
    Set db = OpenDatabase(App.Path & "\Database\MembersDB.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("MembersInfo", dbOpenTable)
        
        rec.MoveFirst
    '' Start Chk for duplicates sql
    '' End -- chk duplicates
    If rec.RecordCount > 0 Then
        For loop1 = 1 To rec.RecordCount
            TDM = DoEvents()
            If Val(txtIDnumber.Text) = rec.Fields("ID NUMBER") Then
               '' Start update fields
              
                 rec.Edit
                 rec.Fields("Date Entered") = Format(Trim(txtDateEntered.Text), "mm/dd/yyyy")
                 ''''rec.Fields("ID NUMBER") = AutoNumber
                 rec.Fields("Membership Level") = Trim(txtNationality.Text) ' now used by mem level
                 rec.Fields("First Name") = Trim(txtFirstName.Text)
                 rec.Fields("Middle Name") = Trim(txtMiddleName.Text)
                 rec.Fields("Family Name") = Trim(txtFamilyName.Text)
                 rec.Fields("Birthday") = Format(Trim(txtBirthday.Text), "mm/dd/yyyy")
                 rec.Fields("Sex") = Trim(txtSex.Text)
                 rec.Fields("Civil Status") = (txtCivilStatus.Text)
                 rec.Fields("Occupation") = Trim(txtOccupation.Text)
                 rec.Fields("Contact Number") = Trim(txtContactNumber.Text)
                 rec.Fields("Home Address") = Trim(txtHomeAddress.Text)
                 rec.Fields("OfficeOrSchool/Address") = Trim(txtOfficeSchoolAddress.Text)
                 rec.Fields("Comments") = Trim(txtComments.Text)
                 rec.Update  ''Update the recordset
               '' End update fields
               db.Close  '' Close DB
               MsgBox "Member Info has been successfully updated. ", vbInformation, "Updated"
               UpdateEditedMembersDB = True
               Exit Function
            End If
            rec.MoveNext
        Next loop1
     End If
     
Err:
 db.Close
 MsgBox "An error occured while updating. ", vbInformation, "Update Error"
 UpdateEditedMembersDB = False
  
End Function

Sub RemoveMember(IDNUMBER As Long)
On Error GoTo Err:
 Dim TDM As Variant
 Dim db As Database
 Dim rec As Recordset
 
 Set db = OpenDatabase(App.Path & "\Database\MembersDB.mdb" _
             , False, False, ";pwd=AdmiN")
 Set rec = db.OpenRecordset("MembersInfo", dbOpenTable)
 rec.MoveFirst
    Do While (rec.EOF = False)
       TDM = DoEvents()
       If Val(rec.Fields("ID NUMBER")) = IDNUMBER Then
           rec.Delete
           Exit Do
       End If
       If rec.EOF = False Then rec.MoveNext
    Loop
 MsgBox "Member has been deleted. ", vbInformation, "Removed"
 db.Close

Err:
End Sub

Sub LoadMoviesList(lst As ListBox)
On Error GoTo ErrorHandler:
    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    lst.Clear
    
    Set db = OpenDatabase(App.Path & "\Database\CD_Tapes.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("CD Tapes Table", dbOpenTable)
    
    rec.MoveFirst
    If rec.RecordCount > 0 Then
        For loop1 = 1 To rec.RecordCount
            TDM = DoEvents()
            'lst.AddItem rec.Fields("First Name") & " " & Left(rec.Fields("Middle Name"), 1) & ". " & _
            rec.Fields("Family Name")
            lst.AddItem rec.Fields("Title") & " [" & rec.Fields("Item Code") & "]"
            rec.MoveNext
        Next loop1
     End If
     
  db.Close
  Exit Sub
  
ErrorHandler:
  db.Close
End Sub

Sub GetCD_TapesInfo(lst As ListBox, txtTitle As TextBox, txtDateEntered As TextBox, txtItemCode As TextBox, _
                    txtActor As TextBox, txtYearReleased As TextBox, txtGenre As TextBox, txtRunTime As TextBox, txtRentalAmount As TextBox, _
                    txtAvailable As TextBox, txtRentalPeriod As TextBox, txtOverdueChargePerDay As TextBox, txtLastDateBorrowed As TextBox, txtLastDateBorrowedAddInfo As TextBox, txtLastDateReturned As TextBox, _
                    txtLastDateReturnedAddInfo As TextBox, txtCondition As TextBox, txtComments As TextBox)

    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    
    Set db = OpenDatabase(App.Path & "\Database\CD_Tapes.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("CD Tapes Table", dbOpenTable)
    
    If rec.BOF = True And rec.EOF = True Then Exit Sub 'chk if rec exists

rec.MoveFirst
    If rec.RecordCount > 0 Then
        For loop1 = 1 To rec.RecordCount
            TDM = DoEvents()
            If lst.Text = rec.Fields("Title") & " [" & rec.Fields("Item Code") & "]" Then
                    txtTitle.Text = rec.Fields("Title")
                    txtDateEntered.Text = rec.Fields("Date Entered")
                    txtItemCode.Text = rec.Fields("Item Code")
                    txtActor.Text = rec.Fields("Actor")
                    txtYearReleased.Text = rec.Fields("Year Released")
                    txtGenre.Text = rec.Fields("Genre")
                    txtRunTime.Text = rec.Fields("Run Time")
                    txtRentalAmount.Text = rec.Fields("Rental Amount")
                    txtAvailable.Text = rec.Fields("Available")
                    If IsNull(rec.Fields("RentalPeriod")) = False Then
                       txtRentalPeriod.Text = rec.Fields("RentalPeriod")
                    Else
                       txtRentalPeriod.Text = ""
                    End If
                    
                    If IsNull(rec.Fields("OverdueChargePerDay")) = False Then
                       txtOverdueChargePerDay.Text = rec.Fields("OverdueChargePerDay")
                    Else
                       txtOverdueChargePerDay.Text = ""
                    End If
                    If IsNull(rec.Fields("LastDateBorrowed")) = False Then
                        txtLastDateBorrowed.Text = rec.Fields("LastDateBorrowed")
                    Else
                       txtLastDateBorrowed.Text = ""
                    End If
                    
                    If IsNull(rec.Fields("LastDateBorrowedAddInfo")) = False Then
                        txtLastDateBorrowedAddInfo.Text = rec.Fields("LastDateBorrowedAddInfo")
                    Else
                       txtLastDateBorrowedAddInfo.Text = ""
                    End If
                    
                    If IsNull(rec.Fields("LastDateReturned")) = False Then
                        txtLastDateReturned.Text = rec.Fields("LastDateReturned")
                    Else
                        txtLastDateReturned.Text = ""
                    End If
                    
                    If IsNull(rec.Fields("LastDateReturnedAddInfo")) = False Then
                        txtLastDateReturnedAddInfo.Text = rec.Fields("LastDateReturnedAddInfo")
                    Else
                       txtLastDateReturnedAddInfo.Text = ""
                    End If
                    txtCondition.Text = rec.Fields("Condition")
                    txtComments.Text = rec.Fields("Comments")
                    db.Close
                    Exit Sub
            End If
            If rec.EOF = False Then rec.MoveNext
        Next loop1
     End If
     
  db.Close
  Exit Sub

End Sub

Function Add_CD_TAPES_MovieToDB(txtTitle As TextBox, txtDateEntered As TextBox, txtItemCode As TextBox, _
                    txtActor As TextBox, txtYearReleased As TextBox, txtGenre As TextBox, txtRunTime As TextBox, txtRentalAmount As TextBox, _
                    txtAvailable As TextBox, txtRentalPeriod As TextBox, txtOverdueChargePerDay As TextBox, txtLastDateBorrowed As TextBox, txtLastDateBorrowedAddInfo As TextBox, txtLastDateReturned As TextBox, _
                    txtLastDateReturnedAddInfo As TextBox, txtCondition As TextBox, txtComments As TextBox) As Boolean
    
    Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    
    Set db = OpenDatabase(App.Path & "\Database\CD_Tapes.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("CD Tapes Table", dbOpenTable)
    
    If rec.BOF = True And rec.EOF = True Then
     ' Do Nothing
    Else
        'chk for item code duplicates
            rec.MoveFirst
            For loop1 = 1 To rec.RecordCount
                TDM = DoEvents()
                If Trim(txtItemCode.Text) = rec.Fields("Item Code") Then
                    MsgBox "Item Code already in use. ", vbInformation, "Adding ERROR!"
                    db.Close
                    Add_CD_TAPES_MovieToDB = False
                    Exit Function
                End If
                If rec.EOF = False Then rec.MoveNext
            Next loop1
        
    End If
 '-------- START ADDING ----------
                    rec.AddNew
                    rec.Fields("Title") = Trim(txtTitle.Text)
                    rec.Fields("Date Entered") = Trim(txtDateEntered.Text)
                    rec.Fields("Item Code") = Trim(txtItemCode.Text)
                    rec.Fields("Actor") = Trim(txtActor.Text)
                    rec.Fields("Year Released") = Trim(txtYearReleased.Text)
                    rec.Fields("Genre") = Trim(txtGenre.Text)
                    rec.Fields("Run Time") = Trim(txtRunTime.Text)
                    rec.Fields("Rental Amount") = Trim(txtRentalAmount.Text)
                    rec.Fields("Available") = Trim(txtAvailable.Text)
                    rec.Fields("RentalPeriod") = Trim(txtRentalPeriod.Text)
                    rec.Fields("OverdueChargePerDay") = Trim(txtOverdueChargePerDay.Text)
                    If IsDate(Trim(txtLastDateBorrowed.Text)) = True Then rec.Fields("LastDateBorrowed") = Trim(txtLastDateBorrowed.Text)
                    rec.Fields("LastDateBorrowedAddInfo") = Trim(txtLastDateBorrowedAddInfo.Text)
                    If IsDate(Trim(txtLastDateReturned.Text)) = True Then rec.Fields("LastDateReturned") = Trim(txtLastDateReturned.Text)
                    rec.Fields("LastDateReturnedAddInfo") = Trim(txtLastDateReturnedAddInfo.Text)
                    rec.Fields("Condition") = Trim(txtCondition.Text)
                    rec.Fields("Comments") = Trim(txtComments.Text)
                    rec.Update
 '-------- END ADDING ------------
     
   db.Close
   MsgBox "Another item has been added. ", vbInformation, "Adding Successful"
   Add_CD_TAPES_MovieToDB = True

End Function

Function UpdateEditedCDtapesInfo(tmpItemCodeB4Edit As String, txtTitle As TextBox, txtDateEntered As TextBox, txtItemCode As TextBox, _
                    txtActor As TextBox, txtYearReleased As TextBox, txtGenre As TextBox, txtRunTime As TextBox, txtRentalAmount As TextBox, _
                    txtAvailable As TextBox, txtRentalPeriod As TextBox, txtOverdueChargePerDay As TextBox, txtLastDateBorrowed As TextBox, txtLastDateBorrowedAddInfo As TextBox, txtLastDateReturned As TextBox, _
                    txtLastDateReturnedAddInfo As TextBox, txtCondition As TextBox, txtComments As TextBox) As Boolean
 Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    
    Set db = OpenDatabase(App.Path & "\Database\CD_Tapes.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("CD Tapes Table", dbOpenTable)
    
    ' Start for ItemCode duplicates
      If tmpItemCodeB4Edit <> Trim(txtItemCode.Text) Then
             
            rec.MoveFirst
            For loop1 = 1 To rec.RecordCount
                TDM = DoEvents()
                If Trim(txtItemCode.Text) = rec.Fields("Item Code") Then
                    MsgBox "Item Code already in use. ", vbInformation, "Update ERROR!"
                    db.Close
                     UpdateEditedCDtapesInfo = False
                    Exit Function
                End If
                If rec.EOF = False Then rec.MoveNext
            Next loop1
      
      End If
    ' End check 4 Item code duplicates
    ' Start - Look 4 record
      rec.MoveFirst
      For loop1 = 1 To rec.RecordCount
                TDM = DoEvents()
                If tmpItemCodeB4Edit = rec.Fields("Item Code") Then
                    Exit For
                End If
                If rec.EOF = False Then rec.MoveNext
      Next loop1

    ' End - Look 4 rec to edit
 '-------- START Update Edit ----------
              
                    rec.Edit
                    rec.Fields("Title") = Trim(txtTitle.Text)
                    rec.Fields("Date Entered") = Trim(txtDateEntered.Text)
                    rec.Fields("Item Code") = Trim(txtItemCode.Text)
                    rec.Fields("Actor") = Trim(txtActor.Text)
                    rec.Fields("Year Released") = Trim(txtYearReleased.Text)
                    rec.Fields("Genre") = Trim(txtGenre.Text)
                    rec.Fields("Run Time") = Trim(txtRunTime.Text)
                    rec.Fields("Rental Amount") = Trim(txtRentalAmount.Text)
                    rec.Fields("Available") = Trim(txtAvailable.Text)
                    rec.Fields("RentalPeriod") = Trim(txtRentalPeriod.Text)
                    rec.Fields("OverdueChargePerDay") = Trim(txtOverdueChargePerDay.Text)
                    If IsDate(Trim(txtLastDateBorrowed.Text)) = False Then
                         rec.Fields("LastDateBorrowed") = Null
                    Else
                         rec.Fields("LastDateBorrowed") = Trim(txtLastDateBorrowed.Text)
                    End If
                    rec.Fields("LastDateBorrowedAddInfo") = Trim(txtLastDateBorrowedAddInfo.Text)
                    If IsDate(Trim(txtLastDateReturned.Text)) = False Then
                       rec.Fields("LastDateReturned") = Null
                    Else
                       rec.Fields("LastDateReturned") = Trim(txtLastDateReturned.Text)
                    End If
                    rec.Fields("LastDateReturnedAddInfo") = Trim(txtLastDateReturnedAddInfo.Text)
                    rec.Fields("Condition") = Trim(txtCondition.Text)
                    rec.Fields("Comments") = Trim(txtComments.Text)
                    rec.Update
 '-------- END Update Edit ------------
     
   db.Close
   MsgBox "Record has been updated. ", vbInformation, "Update Successful"
  UpdateEditedCDtapesInfo = True

End Function

Sub Remove_CD_tape_Items(ItemCode As String)

 Dim db As Database
    Dim rec As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    
    Set db = OpenDatabase(App.Path & "\Database\CD_Tapes.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("CD Tapes Table", dbOpenTable)
   
            rec.MoveFirst
            For loop1 = 1 To rec.RecordCount
                TDM = DoEvents()
                If ItemCode = rec.Fields("Item Code") Then
                    rec.Delete
                    db.Close
                    Exit For
                End If
                If rec.EOF = False Then rec.MoveNext
            Next loop1
      MsgBox "Record has been successfully removed. ", vbInformation, "Record removed"
End Sub

Sub Search_Movies(FlexMovies As MSFlexGrid, SearchString As String, SearchFields As String, SearchMode As Boolean, SortByFields As String, SortMode As Boolean)
    Dim TDM As Variant
    Dim loop1, loop2 As Long
    Dim mySQL As String
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String

    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset

    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase\CD_Tapes.mdb;Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
   

    

With FlexMovies
    .ColWidth(0) = 600
    .ColWidth(1) = 1250
    .ColWidth(2) = 1350
    .ColWidth(3) = 2250
    .ColWidth(4) = 1900
    .ColWidth(5) = 1100
    .ColWidth(6) = 1180
    .ColWidth(7) = 900

    .TextMatrix(0, 0) = "No."
    .TextMatrix(0, 1) = "Item Code"
    .TextMatrix(0, 2) = "Date Entered"
    .TextMatrix(0, 3) = "Title"
    .TextMatrix(0, 4) = "Actor"
    .TextMatrix(0, 5) = "Genre"
    .TextMatrix(0, 6) = "Year Released"
    .TextMatrix(0, 7) = "Available"
    
End With


'' ---------Start SQL
  If SearchString = "[All Movies]" Or SearchString = "*" Or SearchString = "" Then
     mySQL = "SELECT * from [CD Tapes Table] ORDER by [" & SortByFields & "]"
  Else
     If SearchMode = True Then
        mySQL = "SELECT * from [CD Tapes Table] WHERE [" & SearchFields & "] LIKE '%" & SearchString & "%' ORDER by [" & SortByFields & "]"
     Else
        mySQL = "SELECT * from [CD Tapes Table] WHERE [" & SearchFields & "] = " & """" & SearchString & """" & "ORDER by [" & SortByFields & "]"
     End If
  End If
adoRecordset.Open mySQL, adoConnection, adOpenStatic, adLockOptimistic, _
           adCmdText
'' --------End SQL

If adoRecordset.BOF = True And adoRecordset.EOF = True Then
           MsgBox "No records found matching your search. ", vbInformation, "Not found"
           Exit Sub
End If
Select Case SortMode
Case True:
  '---------- Forward Processing ------------------------------------------
        FlexMovies.Rows = 1
        loop1 = 0
        
        adoRecordset.MoveFirst
        Do Until adoRecordset.EOF
             TDM = DoEvents()
             loop1 = loop1 + 1
             FlexMovies.AddItem ""
                For loop2 = 0 To 7
                    Select Case loop2
                        Case 0:
                                FlexMovies.TextMatrix(loop1, loop2) = str(loop1)
                        Case 1:
                                FlexMovies.TextMatrix(loop1, loop2) = adoRecordset![Item Code]
                        Case 2:
                                FlexMovies.TextMatrix(loop1, loop2) = adoRecordset![Date Entered]
                        Case 3:
                                FlexMovies.TextMatrix(loop1, loop2) = adoRecordset!Title
                        Case 4:
                                FlexMovies.TextMatrix(loop1, loop2) = adoRecordset!actor
                        Case 5:
                                FlexMovies.TextMatrix(loop1, loop2) = adoRecordset!Genre
                        Case 6:
                                FlexMovies.TextMatrix(loop1, loop2) = adoRecordset![Year Released]
                        Case 7:
                                FlexMovies.TextMatrix(loop1, loop2) = adoRecordset!Available
                    End Select
                Next loop2
            adoRecordset.MoveNext
         Loop
         'FlexMovies.Rows = FlexMovies.Rows - 1
  '-----------------------------------------------
Case False:
'---START--- Backward Proceesing (Sort)
Dim tmpArray() As Variant
Dim counter, countdown As Long
counter = 0
loop1 = 0
        adoRecordset.MoveFirst
        FlexMovies.Rows = 1
        Do Until adoRecordset.EOF
             TDM = DoEvents()
             loop1 = loop1 + 1
             counter = counter + 8
             ReDim Preserve tmpArray(counter)
             'FlexMovies.AddItem ""
                For loop2 = 0 To 7
                    Select Case loop2
                        Case 0:
                                tmpArray(counter - 7 + loop2) = str(loop1)
                        Case 1:
                                tmpArray(counter - 7 + loop2) = adoRecordset![Item Code]
                        Case 2:
                                tmpArray(counter - 7 + loop2) = adoRecordset![Date Entered]
                        Case 3:
                                tmpArray(counter - 7 + loop2) = adoRecordset!Title
                        Case 4:
                                tmpArray(counter - 7 + loop2) = adoRecordset!actor
                        Case 5:
                                tmpArray(counter - 7 + loop2) = adoRecordset!Genre
                        Case 6:
                                tmpArray(counter - 7 + loop2) = adoRecordset![Year Released]
                        Case 7:
                                tmpArray(counter - 7 + loop2) = adoRecordset!Available
                    End Select
                    
                Next loop2

            adoRecordset.MoveNext
         Loop
       '------ Start Backward display
       For loop1 = 1 To counter
      ' Debug.Print Str(loop1) & ": " & tmpArray(loop1)
       Next
       loop1 = 0
        adoRecordset.MoveFirst
        countdown = counter
        For loop1 = 1 To counter / 8
             TDM = DoEvents()
             FlexMovies.AddItem ""
                                FlexMovies.TextMatrix(loop1, 7) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMovies.TextMatrix(loop1, 6) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMovies.TextMatrix(loop1, 5) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMovies.TextMatrix(loop1, 4) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMovies.TextMatrix(loop1, 3) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMovies.TextMatrix(loop1, 2) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMovies.TextMatrix(loop1, 1) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMovies.TextMatrix(loop1, 0) = str(loop1)
                                countdown = countdown - 1
      Next loop1

'---END--- Backward Proceesing (Sort)
 End Select
 
 
 
adoRecordset.Close
adoConnection.Close
Set adoRecordset = Nothing
Set adoConnection = Nothing

End Sub

Sub Search_Members(FlexMembers As MSFlexGrid, SearchString As String, SearchFields As String, SearchMode As Boolean, SortByFields As String, SortMode As Boolean)
    
    Dim TDM As Variant
    Dim loop1, loop2 As Long
    Dim mySQL As String
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String

    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset

    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase\MembersDB.mdb;Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
   
With FlexMembers
    .ColWidth(0) = 600
    .ColWidth(1) = 1200 ' ID Number
    .ColWidth(2) = 1150 ' Date Entered
    .ColWidth(3) = 1450 ' Membership Level
    .ColWidth(4) = 1500 ' Family Name
    .ColWidth(5) = 1500 ' First Name
    .ColWidth(6) = 1500 ' Middle Name
    .ColWidth(7) = 1190 ' Birthday
    .ColWidth(8) = 400 ' Age
    .ColWidth(9) = 800 ' Sex
    .ColWidth(10) = 1100 ' Civil Status
    .ColWidth(11) = 1450 ' Occupation
    .ColWidth(12) = 3500 ' Home Address
    .ColWidth(13) = 3500 ' Office/School Address
    .ColWidth(14) = 3500 ' Comments

    .TextMatrix(0, 0) = "No."
    .TextMatrix(0, 1) = "ID Number"
    .TextMatrix(0, 2) = "Date Entered"
    .TextMatrix(0, 3) = "Membership Level"
    .TextMatrix(0, 4) = "Family Name"
    .TextMatrix(0, 5) = "First Name"
    .TextMatrix(0, 6) = "Middle Name"
    .TextMatrix(0, 7) = "Birthday"
    .TextMatrix(0, 8) = "Age"
    .TextMatrix(0, 9) = "Sex"
    .TextMatrix(0, 10) = "Civil Status"
    .TextMatrix(0, 11) = "Occupation"
    .TextMatrix(0, 12) = "Home Address"
    .TextMatrix(0, 13) = "Office/School Address"
    .TextMatrix(0, 14) = "Comments"
    
End With
 
'' ---------Start SQL
  If SearchString = "[All Names]" Or SearchString = "*" Or SearchString = "" Then
     mySQL = "SELECT * from [MembersInfo] ORDER by [" & SortByFields & "]"
  Else
     If SearchMode = True Then
        mySQL = "SELECT * from [MembersInfo] WHERE [" & SearchFields & "] LIKE '%" & SearchString & "%' ORDER by [" & SortByFields & "]"
     Else
        mySQL = "SELECT * from [MembersInfo] WHERE [" & SearchFields & "] = " & """" & SearchString & """" & " ORDER by [" & SortByFields & "]"
     End If
  End If
adoRecordset.Open mySQL, adoConnection, adOpenStatic, adLockOptimistic, _
           adCmdText
'' --------End SQL

If adoRecordset.BOF = True And adoRecordset.EOF = True Then
           MsgBox "No records found matching your search. ", vbInformation, "Not found"
           Exit Sub
End If
Select Case SortMode
Case True:
  '---------- Forward Processing ------------------------------------------
        FlexMembers.Rows = 1
        loop1 = 0
        
        adoRecordset.MoveFirst
        Do Until adoRecordset.EOF
             TDM = DoEvents()
             loop1 = loop1 + 1
             FlexMembers.AddItem ""
                For loop2 = 0 To 14
                    Select Case loop2
                        Case 0:
                                FlexMembers.TextMatrix(loop1, loop2) = str(loop1)
                        Case 1:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![ID NUMBER]
                        Case 2:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![Date Entered]
                        Case 3:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![Membership Level]
                        Case 4:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![Family Name]
                        Case 5:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![First Name]
                        Case 6:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![Middle Name]
                        Case 7:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![Birthday]
                        Case 8:
                                FlexMembers.TextMatrix(loop1, loop2) = GetAge(adoRecordset![Birthday])
                        Case 9:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![Sex]
                        Case 10:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![Civil Status]
                        Case 11:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![Occupation]
                        Case 12:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![Home Address]
                        Case 13:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![OfficeOrSchool/Address]
                        Case 14:
                                FlexMembers.TextMatrix(loop1, loop2) = adoRecordset![Comments]
                    End Select
                Next loop2
            adoRecordset.MoveNext
         Loop
         'FlexMembers.Rows = FlexMembers.Rows - 1
  '-----------------------------------------------
Case False:
'---START--- Backward Proceesing (Sort)
Dim tmpArray() As Variant
Dim counter, countdown As Long
counter = 0
loop1 = 0
        adoRecordset.MoveFirst
        FlexMembers.Rows = 1
        Do Until adoRecordset.EOF
             TDM = DoEvents()
             loop1 = loop1 + 1
             counter = counter + 15
             ReDim Preserve tmpArray(counter)
             'FlexMembers.AddItem ""
                For loop2 = 0 To 14
                    Select Case loop2
                        Case 0:
                                tmpArray(counter - 14 + loop2) = str(loop1)
                        Case 1:
                                tmpArray(counter - 14 + loop2) = adoRecordset![ID NUMBER]
                        Case 2:
                                tmpArray(counter - 14 + loop2) = adoRecordset![Date Entered]
                        Case 3:
                                tmpArray(counter - 14 + loop2) = adoRecordset![Membership Level]
                        Case 4:
                                tmpArray(counter - 14 + loop2) = adoRecordset![Family Name]
                        Case 5:
                                tmpArray(counter - 14 + loop2) = adoRecordset![First Name]
                        Case 6:
                                tmpArray(counter - 14 + loop2) = adoRecordset![Middle Name]
                        Case 7:
                                tmpArray(counter - 14 + loop2) = adoRecordset![Birthday]
                        Case 8:
                                tmpArray(counter - 14 + loop2) = GetAge(adoRecordset![Birthday])
                        Case 9:
                                tmpArray(counter - 14 + loop2) = adoRecordset![Sex]
                        Case 10:
                                tmpArray(counter - 14 + loop2) = adoRecordset![Civil Status]
                        Case 11:
                                tmpArray(counter - 14 + loop2) = adoRecordset![Occupation]
                        Case 12:
                                tmpArray(counter - 14 + loop2) = adoRecordset![Home Address]
                        Case 13:
                                tmpArray(counter - 14 + loop2) = adoRecordset![Home Address]
                        Case 14:
                                tmpArray(counter - 14 + loop2) = adoRecordset![OfficeOrSchool/Address]
                        Case 15:
                                tmpArray(counter - 14 + loop2) = adoRecordset![Comments]
                        
                        
                    End Select
               Next loop2

            adoRecordset.MoveNext
         Loop
       '------ Start Backward display
       For loop1 = 1 To counter
      '' Debug.Print Str(loop1) & ": " & tmpArray(loop1)
       Next
       loop1 = 0
        adoRecordset.MoveFirst
        countdown = counter
        For loop1 = 1 To counter / 15
             TDM = DoEvents()
             FlexMembers.AddItem ""
                                FlexMembers.TextMatrix(loop1, 14) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 13) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 12) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 11) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 10) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 9) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 8) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 7) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 6) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 5) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 4) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 3) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 2) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 1) = tmpArray(countdown)
                                countdown = countdown - 1
                                FlexMembers.TextMatrix(loop1, 0) = str(loop1)
                                countdown = countdown - 1
      Next loop1

'---END--- Backward Proceesing (Sort)
 End Select
 
 
 
adoRecordset.Close
adoConnection.Close
Set adoRecordset = Nothing
Set adoConnection = Nothing




End Sub
Function Transaction_GetDateDue(ItemCode As String, RefDate As String) As String
    On Error GoTo ErrHandler
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String
    Dim RentalPeriod As Integer
    Dim DateDue As Long
    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset

    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase\CD_Tapes.mdb;Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
    
    adoRecordset.Open "SELECT RentalPeriod from [CD Tapes Table] WHERE [Item Code] = '" & ItemCode & "'", adoConnection, adOpenStatic, adLockOptimistic, _
           adCmdText
    
    RentalPeriod = adoRecordset![RentalPeriod]
    
    DateDue = Date_CountNumberOfDaysFromJan1Year1ToDec31YearEntered(Format(RefDate, "yyyy") - 1)
    DateDue = DateDue + Date_CountDaysInAYear(Format(RefDate, "mm/dd/yyyy"))
    DateDue = DateDue + RentalPeriod
    Transaction_GetDateDue = Date_GETDATE(DateDue)
    Exit Function
ErrHandler:
    Transaction_GetDateDue = "Unknown"
End Function

Sub Transaction_LoadItemCodes(cbo As ComboBox)
    Dim TDM As Variant
    Dim loop1, loop2 As Long
    Dim mySQL As String
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String

    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset

    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase\CD_Tapes.mdb;Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
    
    adoRecordset.Open "SELECT * from [CD Tapes Table] WHERE [Available] = ""Yes"" ", adoConnection, adOpenStatic, adLockOptimistic, _
           adCmdText
    
    If adoRecordset.BOF = True And adoRecordset.EOF = True Then
           MsgBox "No Item Codes found. ", vbInformation, "Not found"
           Exit Sub
    End If
   
    loop1 = 0
    adoRecordset.MoveFirst
    cbo.Clear
     Do Until adoRecordset.EOF
             TDM = DoEvents()
             loop1 = loop1 + 1
             cbo.AddItem adoRecordset![Item Code]
             adoRecordset.MoveNext
     Loop
     
    adoRecordset.Close
    adoConnection.Close
    Set adoRecordset = Nothing
    Set adoConnection = Nothing

End Sub

Sub Transaction_LoadNameOfMembers(lst As ListBox, ArrayOFNamesAndID(), MembersID())
    Dim TDM As Variant
    Dim loop1, loop2 As Long
    Dim mySQL As String
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String

    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset

    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase\MembersDB.mdb;Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
    
    adoRecordset.Open "SELECT * from [MembersInfo]ORDER BY [Family Name]", adoConnection, adOpenStatic, adLockOptimistic, _
           adCmdText
    
    If adoRecordset.BOF = True And adoRecordset.EOF = True Then
           MsgBox "No members found. ", vbInformation, "Not found"
           Exit Sub
    End If
   
    loop1 = 0
    adoRecordset.MoveFirst
    lst.Clear
     Do Until adoRecordset.EOF
             TDM = DoEvents()
             loop1 = loop1 + 1
             ReDim Preserve ArrayOFNamesAndID(loop1)
             ReDim Preserve MembersID(loop1)
             lst.AddItem adoRecordset![Family Name] & ", " & adoRecordset![First Name] & " " & Mid(adoRecordset![Middle Name], 1, 1) & "."
             ArrayOFNamesAndID(loop1) = adoRecordset![Family Name] & ", " & adoRecordset![First Name] & " " & Mid(adoRecordset![Middle Name], 1, 1) & "." & " ID - " & str(adoRecordset![ID NUMBER])
             MembersID(loop1) = str(adoRecordset![ID NUMBER])
             adoRecordset.MoveNext
     Loop
     
    adoRecordset.Close
    adoConnection.Close
    Set adoRecordset = Nothing
    Set adoConnection = Nothing

End Sub

Sub Transaction_AddItem(Flex As MSFlexGrid, cbo As ComboBox)
    Dim TDM As Variant
    Dim loop1, loop2 As Long
    Dim mySQL As String
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String

    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset

    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase\CD_Tapes.mdb;Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
    
    adoRecordset.Open "SELECT * FROM [CD Tapes Table] WHERE [Item Code] = " & """" & cbo.Text & """", adoConnection, adOpenStatic, adLockOptimistic, _
           adCmdText
    
    If adoRecordset.BOF = True And adoRecordset.EOF = True Then
           MsgBox "Item not found in database.              ", vbInformation, "Not found"
           cbo.SetFocus
           Exit Sub
    End If
    
    If Flex.Rows = 1 Then Flex.AddItem ""
    If Trim(Flex.TextMatrix(1, 0)) = "" And Trim(Flex.TextMatrix(1, 1)) = "" And Trim(Flex.TextMatrix(1, 2)) = "" And Trim(Flex.TextMatrix(1, 3)) = "" Then
            Flex.TextMatrix(1, 0) = "1"
            Flex.TextMatrix(1, 1) = adoRecordset![Item Code]
            Flex.TextMatrix(1, 2) = adoRecordset![Title]
            Flex.TextMatrix(1, 3) = adoRecordset![Rental Amount]
    Else
            Flex.AddItem ""
            Flex.TextMatrix(Flex.Rows - 1, 0) = Flex.Rows - 1
            Flex.TextMatrix(Flex.Rows - 1, 1) = adoRecordset![Item Code]
            Flex.TextMatrix(Flex.Rows - 1, 2) = adoRecordset![Title]
            Flex.TextMatrix(Flex.Rows - 1, 3) = adoRecordset![Rental Amount]
            
    End If

    adoRecordset.Close
    adoConnection.Close
    Set adoRecordset = Nothing
    Set adoConnection = Nothing
End Sub
Sub CheckIfMiscTransSalesDBExistIfNotCreate()
On Error Resume Next
Dim daoWS As dao.Workspace
Dim daoDB As dao.Database
Dim daoTable As New dao.TableDef
Dim daoField As New dao.Field
Call CheckTransactionFolder
Set daoWS = DBEngine.Workspaces(0)

' Create database
 Set daoDB = daoWS.CreateDatabase(App.Path & "\Transaction\MiscTrans.mdb", dbLangGeneral, dbVersion40)

' Create Table
 Set daoTable = daoDB.CreateTableDef("Transaction")

' Create Fields
Set daoField = daoTable.CreateField("Date", dbDate)

' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("Cashier", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("CashierID", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("InvoiceNumber", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("Description", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("Amount", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

' Append table to database
daoDB.TableDefs.Append daoTable

' Clean up objects
Set daoField = Nothing
Set daoTable = Nothing
Set daoDB = Nothing
Set daoWS = Nothing
Call SetDatabasePassword(App.Path & "\Transaction\MiscTrans.mdb", "AdmiN")

End Sub
Sub CheckIfTransactionDBExistIfNotCreate()
On Error Resume Next
Dim daoWS As dao.Workspace
Dim daoDB As dao.Database
Dim daoTable As New dao.TableDef
Dim daoField As New dao.Field
Call CheckTransactionFolder
Set daoWS = DBEngine.Workspaces(0)

' Create database
 Set daoDB = daoWS.CreateDatabase(App.Path & "\Transaction\Transaction.mdb", dbLangGeneral, dbVersion40)

' Create Table
 Set daoTable = daoDB.CreateTableDef("Transaction")

' Create Fields
Set daoField = daoTable.CreateField("Date", dbDate)

' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("Cashier", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("CashierID", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("InvoiceNumber", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("MembersName", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("Members ID", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("Borrowed Items", dbMemo)
' Append field to table
daoTable.Fields.Append daoField

' Append table to database
daoDB.TableDefs.Append daoTable

' Clean up objects
Set daoField = Nothing
Set daoTable = Nothing
Set daoDB = Nothing
Set daoWS = Nothing
Call SetDatabasePassword(App.Path & "\Transaction\Transaction.mdb", "AdmiN")
End Sub
Sub CheckTransactionFolder()
 On Error GoTo C
  MkDir App.Path & "\Transaction"
 Exit Sub
C:
End Sub
Private Function CheckArray(myArray() As String) As Boolean
Dim X As Long
On Error GoTo ErrorHandler
X = UBound(myArray())
CheckArray = True
Exit Function

ErrorHandler:
 CheckArray = False

End Function
Private Function CheckPrevItemCodeArray(myArray()) As Boolean
Dim X As Long
On Error GoTo ErrorHandler
X = UBound(myArray())
CheckPrevItemCodeArray = True
Exit Function

ErrorHandler:
 CheckPrevItemCodeArray = False
 
End Function
Sub SetDatabasePassword(DBPath As String, newPassword As String)
                            
If Dir(DBPath) = "" Then Exit Sub

Dim db As dao.Database

On Error Resume Next
Set db = OpenDatabase(DBPath, True)
If Err.Number <> 0 Then Exit Sub
db.newPassword "", newPassword

db.Close

End Sub
Sub Transaction_SaveNewTransaction(Flex As MSFlexGrid, TransactionDate As String, CashierName As String, ByVal CashierID As String, InvoiceNumber As String, MembersName As String, TotalAmountDue As String, AmountPaid As String, Change As String, MemberFileNameOfHisDB As String, ByVal MembersID As String)
    Dim loop1, loop2 As Long
    Dim BorrowedItems As String
    Dim db As Database
    Dim rec As Recordset
    Dim TDM
    BorrowedItems = ""
    
    ''Increment InvNum
    Open App.Path & "\InvoiceNumber.txt" For Output As #1
        Print #1, InvoiceNumber
    Close #1
    ''End Increment InvNum
    
    
       
    For loop1 = 1 To Flex.Rows - 1
      BorrowedItems = BorrowedItems & "ItemCode::::" & Flex.TextMatrix(loop1, 1) & vbCrLf & "Title::::" & Flex.TextMatrix(loop1, 2) & vbCrLf & "RentalAmount::::" & Flex.TextMatrix(loop1, 3) & vbCrLf
    Next
       BorrowedItems = BorrowedItems & "TotalAmountDue::::" & TotalAmountDue & vbCrLf
       BorrowedItems = BorrowedItems & "AmountPaid::::" & AmountPaid & vbCrLf
       BorrowedItems = BorrowedItems & "Change::::" & Change & vbCrLf
    
    Set db = OpenDatabase(App.Path & "\Transaction\Transaction.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("Transaction", dbOpenTable)
     
    rec.AddNew
    rec.Fields("Date") = TransactionDate
    rec.Fields("Cashier") = CashierName
    rec.Fields("CashierID") = CashierID
    rec.Fields("InvoiceNumber") = Trim(InvoiceNumber)
    rec.Fields("MembersName") = MembersName
    rec.Fields("Members ID") = MembersID
    rec.Fields("Borrowed Items") = BorrowedItems
    rec.Update
    rec.Close
    db.Close
    Set db = Nothing
    Set rec = Nothing
' OK MsgBox "check"

'-------------------------------------------
     '' Start Change Yes to No, AddInfo
    
' OK MsgBox "check"
'    rec.MoveFirst
'    If rec.RecordCount > 0 Then
    Dim mySQL As String
    'Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim adoConnection As ADODB.Connection
    Dim connectString As String
    Set adoConnection = New ADODB.Connection
    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\CD_Tapes.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"
    adoConnection = connectString
    For loop1 = 1 To Flex.Rows - 1
    'Set db = OpenDatabase(App.Path & "\Database\CD_Tapes.mdb" _
             , False, False, ";pwd=AdmiN")
    'Set rec = db.OpenRecordset("CD Tapes Table", dbOpenTable)
    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset
    adoConnection.Open connectString
    mySQL = "Select * FROM [CD Tapes Table] WHERE [Item Code] = " & """" & Flex.TextMatrix(loop1, 1) & """"
    adoRecordset.Open mySQL, adoConnection, adOpenStatic, adLockOptimistic, _
           adCmdText
     'rec.MoveFirst
        'For loop2 = 1 To rec.RecordCount
         '   TDM = DoEvents()
          '  If rec.Fields("Item Code") = Flex.TextMatrix(loop1, 1) Then
                '' Start Change Yes to No, AddInfo
                'MsgBox "check Update"
                'adoRecordset.Edit
                adoRecordset!Available = "No"
                adoRecordset![Borrower's DB] = Trim(MemberFileNameOfHisDB)
                adoRecordset!LastDateBorrowed = TransactionDate
                adoRecordset!LastDateBorrowedAddInfo = "Borrowed by: " & MembersName
                'rec.Edit
                'rec.Fields("Available") = "No"
                'rec.Fields("[Borrower's DB]") = Trim(MemberFileNameOfHisDB)
                'rec.Fields("LastDateBorrowed") = TransactionDate
                'rec.Fields("LastDateBorrowedAddInfo") = "Borrowed by: " & MembersName
                'rec.Update
                adoRecordset.Update
                'MsgBox "check Update End"
                '' End Change Yes to No, AddInfo
            'End If
        'MsgBox "check"
            'If rec.EOF = False Then rec.MoveNext ' Move to the next record
           ' rec.MoveNext
        'Next loop2
               ' MsgBox "check2"
   adoConnection.Close
'   adoRecordset.Close
   Set adoConnection = Nothing
   Set adoRecordset = Nothing

     Next loop1
    ' End If
'      rec.Close
 '     db.Close
    '' End Change Yes to No. AddInfo
Set db = Nothing
Set rec = Nothing
' B4 dis error MsgBox "check"
'------------------------------------------
'------------------------------------------
' Start Append to MembersFileDB
  Call Transaction_CheckForMembersRecordsIfNotExistsCreate(App.Path & "\Transaction\MembersRecords\", MemberFileNameOfHisDB)
  Set db = OpenDatabase(App.Path & "\Transaction\MembersRecords\" & MemberFileNameOfHisDB _
             , False, False, ";pwd=AdmiN")
  Set rec = db.OpenRecordset("Members Borrowed Items History", dbOpenTable)
   For loop1 = 1 To Flex.Rows - 1
        rec.AddNew
        rec.Fields("Date Borrowed") = TransactionDate
        rec.Fields("Invoice Number") = Trim(InvoiceNumber)
        rec.Fields("Item Code") = Flex.TextMatrix(loop1, 1)
        rec.Fields("Title") = Flex.TextMatrix(loop1, 2)
        rec.Fields("Returned") = "No"
        rec.Update
  Next loop1
  
  rec.Close
  db.Close
' End Append to MembersFileDB
'------------------------------------------
    Call Transaction_AppendToBorrowedItemsHistory(Flex, TransactionDate, gVarUserID, Trim(InvoiceNumber), MembersName)
    
    MsgBox "Transaction has been successfully saved. ", vbInformation, "Transaction successful"
Set db = Nothing
Set rec = Nothing
End Sub
Sub Transaction_ChkIfBorrowedItemsHistoryDBExistIfNotCreate()
On Error Resume Next
Dim daoWS As dao.Workspace
Dim daoDB As dao.Database
Dim daoTable As New dao.TableDef
Dim daoField As New dao.Field
Call CheckTransactionFolder
Set daoWS = DBEngine.Workspaces(0)

' Create database
 Set daoDB = daoWS.CreateDatabase(App.Path & "\Transaction\BIH.mdb", dbLangGeneral, dbVersion40)

' Create Table
 Set daoTable = daoDB.CreateTableDef("History")

' Create Fields
Set daoField = daoTable.CreateField("Date", dbDate)

' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("CashierID", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("InvoiceNumber", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField
'''''''''''''''''''''''''''''''''''''''''LLLLL
'Set daoField = daoTable.CreateField("Borrower's ID", dbText, 255)
' Append field to table
'daoTable.Fields.Append daoField
'''''''''''''''''''''''
Set daoField = daoTable.CreateField("Borrower's Name", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("ItemCode", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("Title", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField


Set daoField = daoTable.CreateField("Rental Amount", dbText, 255)
' Append field to table
daoTable.Fields.Append daoField


' Append table to database
daoDB.TableDefs.Append daoTable

' Clean up objects
Set daoField = Nothing
Set daoTable = Nothing
Set daoDB = Nothing
Set daoWS = Nothing
Call SetDatabasePassword(App.Path & "\Transaction\BIH.mdb", "AdmiN")

End Sub
Sub Transaction_AppendToBorrowedItemsHistory(Flex As MSFlexGrid, TransactionDate As String, ByVal CashierID As String, InvoiceNumber As String, BorrowersName As String)
    Dim db As Database
    Dim rec As Recordset
    Dim loop1 As Long
    
    Set db = OpenDatabase(App.Path & "\Transaction\BIH.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("History", dbOpenTable)
     
    
    For loop1 = 1 To Flex.Rows - 1
        rec.AddNew
        rec.Fields("Date") = TransactionDate
        rec.Fields("CashierID") = CashierID
        rec.Fields("InvoiceNumber") = InvoiceNumber
        rec.Fields("Borrower's Name") = BorrowersName
        rec.Fields("ItemCode") = Flex.TextMatrix(loop1, 1)
        rec.Fields("Title") = Flex.TextMatrix(loop1, 2)
        rec.Fields("Rental Amount") = Flex.TextMatrix(loop1, 3)
        rec.Update
    Next loop1
    
    rec.Close
    db.Close
End Sub

Sub Transaction_CheckForMembersRecordsIfNotExistsCreate(FilePath As String, MemberFileDB As String)
     On Error Resume Next
     If ReportFileStatus(FilePath & MemberFileDB) = False Then
        ''Create MembersFileDB
            Dim daoWS As dao.Workspace
            Dim daoDB As dao.Database
            Dim daoTable As New dao.TableDef
            Dim daoField As New dao.Field
            Call CheckTransactionFolder
            Set daoWS = DBEngine.Workspaces(0)

            ' Create database
            Set daoDB = daoWS.CreateDatabase(FilePath & MemberFileDB, dbLangGeneral, dbVersion40)

            ' Create Table
            Set daoTable = daoDB.CreateTableDef("Members Borrowed Items History")

            ' Create Fields
            Set daoField = daoTable.CreateField("Date Borrowed", dbDate)
            ' Append field to table
            daoTable.Fields.Append daoField
 
            ' Create Fields
            Set daoField = daoTable.CreateField("Invoice Number", dbText, 255)
            ' Append field to table
            daoTable.Fields.Append daoField

            Set daoField = daoTable.CreateField("Item Code", dbText, 255)
            ' Append field to table
            daoTable.Fields.Append daoField

            Set daoField = daoTable.CreateField("Title", dbText, 255)
            ' Append field to table
            daoTable.Fields.Append daoField

            Set daoField = daoTable.CreateField("Returned", dbText, 255)
            ' Append field to table
            daoTable.Fields.Append daoField
            
            ' Append table to database
            daoDB.TableDefs.Append daoTable

            ' Clean up objects
            Set daoField = Nothing
            Set daoTable = Nothing
            Set daoDB = Nothing
            Set daoWS = Nothing
            Call SetDatabasePassword(FilePath & MemberFileDB, "AdmiN")

     Else
        '' Do Nothing
     End If
     
End Sub
Sub Transaction_ChkForMembersFIleDBFolderIfNotCreate()
  On Error GoTo ErrHandler
      MkDir (App.Path & "\Transaction\MembersRecords")
   Exit Sub
ErrHandler:
End Sub

Function Transaction_ChkIfItemHasBeenBorrowed(MemberFileDBandPath As String, ItemCode As String) As Boolean
    Dim mySQL As String
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String

    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset

    
    If ReportFileStatus(MemberFileDBandPath) = True Then
        connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MemberFileDBandPath & ";Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"
        adoConnection.Open connectString
        mySQL = "Select * FROM [Members Borrowed Items History] WHERE [Item Code] = " & """" & ItemCode & """"
        adoRecordset.Open mySQL, adoConnection, adOpenStatic, adLockOptimistic, _
           adCmdText
        If adoRecordset.BOF = True And adoRecordset.EOF = True Then
            '' Do nothing
        Else
        adoRecordset.MoveLast
            If MsgBox("The last time you borrowed this item was " & Format(adoRecordset![Date Borrowed], "mmmm dd, yyyy") & "." & vbCrLf & vbCrLf & "Do you want to continue? ", vbYesNo, "Info") = vbNo Then
                Transaction_ChkIfItemHasBeenBorrowed = True
                adoRecordset.Close
                adoConnection.Close
                Set adoRecordset = Nothing
                Set adoConnection = Nothing
               Exit Function
            End If
        End If
        adoRecordset.Close
        adoConnection.Close
        Set adoRecordset = Nothing
        Set adoConnection = Nothing
        Transaction_ChkIfItemHasBeenBorrowed = False
    End If
End Function

Sub Transaction_CheckIfMemberHasUnreturnedItems(MemberFileDBandPath)
    Dim mySQL As String
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String
    Dim loop1 As Integer

    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset
    
    If ReportFileStatus(MemberFileDBandPath) = True Then
        connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MemberFileDBandPath & ";Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"
        adoConnection.Open connectString
        mySQL = "Select * FROM [Members Borrowed Items History] WHERE [Returned] = " & """" & "No" & """"
        adoRecordset.Open mySQL, adoConnection, adOpenStatic, adLockOptimistic, _
           adCmdText
        If adoRecordset.BOF = True And adoRecordset.EOF = True Then
            adoRecordset.Close
            adoConnection.Close
            Set adoRecordset = Nothing
            Set adoConnection = Nothing
            Exit Sub
        Else
           mySQL = ""
           adoRecordset.MoveFirst
           Do While adoRecordset.EOF = False
               loop1 = loop1 + 1
               mySQL = mySQL & str(loop1) & ".) " & adoRecordset![Title] & "  ITEM CODE: " & adoRecordset![Item Code] & "  Date: " & Format(adoRecordset![Date Borrowed], "dddd - mmmm dd, yyyy") & ".      " & vbCrLf
               If adoRecordset.EOF = False Then adoRecordset.MoveNext
           Loop
           MsgBox "Unreturned Item(s) by this member:" & vbCrLf & vbCrLf & mySQL, vbInformation, "Information"
            adoRecordset.Close
            adoConnection.Close
            Set adoRecordset = Nothing
            Set adoConnection = Nothing
        End If
    End If
End Sub

Function Transaction_LoadPrevTransaction(Flex As MSFlexGrid, txtTotalAmountDue As TextBox, txtAmountPaid As TextBox, txtChange As TextBox, txtTransactionDate As TextBox, txtInvoiceNumber As TextBox, txtMembersName As TextBox, MemberID_FindMode) As Boolean
' START -- Chk if transaction file exists
  If ReportFileStatus(App.Path & "\Transaction\Transaction.mdb") = False Then
     txtInvoiceNumber.Text = ""
     MsgBox "Transaction Database is missing.  ", vbInformation, "Database not found."
     Flex.SetFocus
     Exit Function
  End If
' END -- Chk if transaction file exists

Dim Char, StringText, tmpString, FlagField As String
Dim LenStr, counter, loop1 As Long
Dim ItemCode(), Title(), RentalAmount(), _
    TotalAmountDue, AmountPaid, Change As String
    
counter = 0
'StringText = Text1.Text
'LenStr = Len(StringText)

'' Start -- SEARCH TRANSACTION DB
  ''Chk first if DB exist
    Dim mySQL As String
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String
    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset
    
    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Transaction\Transaction.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
    'NOTE: Spc b4 invoicenumber in BD                                      d2 ang spc
    adoRecordset.Open "SELECT * from [Transaction] WHERE [InvoiceNumber]= '" & txtInvoiceNumber.Text & "'", adoConnection, adOpenStatic, adLockOptimistic, adCmdText
   If adoRecordset.BOF = True And adoRecordset.EOF = True Then
      MsgBox "Transaction Not Found. ", vbInformation, "Info"
      Flex.SetFocus
      adoRecordset.Close
      adoConnection.Close
      txtInvoiceNumber.Text = ""
      Set adoRecordset = Nothing
      Set adoConnection = Nothing
      Transaction_LoadPrevTransaction = False
      Exit Function
   Else
      adoRecordset.MoveFirst
      StringText = adoRecordset![Borrowed Items]
      txtMembersName.Text = adoRecordset![MembersName]
      MemberID_FindMode = adoRecordset![Members ID]
      txtTransactionDate.Text = Format(adoRecordset![Date], "mmm. dd, yyyy")
      
      LenStr = Len(StringText)
   End If
   
   adoRecordset.Close
   adoConnection.Close
   Set adoRecordset = Nothing
   Set adoConnection = Nothing
'' END -- SEARCH TRANSACTION DB

For loop1 = 1 To LenStr
    Char = Mid(StringText, loop1, 1)
 ' ------ Start "::::" Detect -----------------------
    If Char = ":" Then
         If loop1 + 3 <= LenStr Then
            If Mid(StringText, loop1, 4) = "::::" Then
               'Text2.Text = Text2.Text & Trim(tmpString)
               FlagField = Trim(tmpString)
               If Trim(tmpString) = "ItemCode" Then
                    counter = counter + 1
                    ReDim Preserve ItemCode(counter)
                    ReDim Preserve Title(counter)
                    ReDim Preserve RentalAmount(counter)
               End If
               tmpString = ""
               loop1 = loop1 + 3
            End If
         End If
    Else
       If Char <> Chr(13) Or Char <> Chr(10) Then tmpString = tmpString & Char
    End If
 ' ------ End "::::" Detect --------------------------
 ' ------ Start vbLFCR detect ------------------------
    If Char = Chr(13) Then
       If loop1 + 1 <= LenStr Then
          If Mid(StringText, loop1, 2) = vbCrLf Then
               If counter = 0 Then counter = 1
                    
                    Select Case FlagField
                      Case "ItemCode": ItemCode(counter) = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                      Case "Title": Title(counter) = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                      Case "RentalAmount": RentalAmount(counter) = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                      Case "TotalAmountDue": TotalAmountDue = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                      Case "AmountPaid": AmountPaid = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                      Case "Change": Change = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                    End Select
                    
                    tmpString = ""
                    loop1 = loop1 + 1
                End If
          End If
   
    End If
  ' ------ END vbLFCR detect ------------------------
Next loop1
 '========== Display OutPut =========================
 ' Flex.ColAlignment(1) = 1
 ' Flex.ColAlignment(2) = 1
 ' Flex.ColAlignment(3) = 6
 ' Flex.Rows = 2
 ' Flex.ColWidth(0) = 500
 ' Flex.ColWidth(1) = 1300
 ' Flex.ColWidth(2) = 2500
 ' Flex.ColWidth(3) = 1300
 ' Flex.TextMatrix(0, 0) = "No."
 ' Flex.TextMatrix(0, 1) = "Item Code"
 ' Flex.TextMatrix(0, 2) = "Title"
 ' Flex.TextMatrix(0, 3) = "Rental Amount"

  For loop1 = 0 To counter - 1
      Flex.TextMatrix(loop1 + 1, 0) = str(loop1 + 1)
      Flex.TextMatrix(loop1 + 1, 1) = ItemCode(loop1 + 1)
      Flex.TextMatrix(loop1 + 1, 2) = Title(loop1 + 1)
      Flex.TextMatrix(loop1 + 1, 3) = RentalAmount(loop1 + 1)
      If loop1 < counter - 1 Then Flex.AddItem ""
  Next
     txtTotalAmountDue.Text = TotalAmountDue
     txtAmountPaid.Text = AmountPaid
     txtChange.Text = Change
     Transaction_LoadPrevTransaction = True

End Function

Sub Transaction_UpdatePrevTrasaction(Flex As MSFlexGrid, TransactionDate As String, CashierName As String, ByVal CashierID As String, InvoiceNumber As String, MembersName As String, TotalAmountDue As String, AmountPaid As String, Change As String, MemberFileNameOfHisDB As String, ByVal MembersID As String, ArrayItems() As String, DeletedItemCodes() As String)
 '' START -- Update Transaction Record Variable Declaration
    Dim loop1, loop2, loop3 As Long
    Dim BorrowedItems As String
    Dim db As Database
    Dim dbcloseFLAG As Boolean
    Dim rec As Recordset
    Dim TDM
    Dim PrevItemCode(), PrevReturnedStatus() As String
    BorrowedItems = ""
 '' END -- Update Transaction Record Variable Declaration
    
 
 '' Start -- Update BIH : (delete)
    '' Start -- Check if File exists
    '' End -- Check if File exists
    ' START ---- DELETE BIH ----
    Call Transaction_ChkIfBorrowedItemsHistoryDBExistIfNotCreate
    Set db = OpenDatabase(App.Path & "\Transaction\BIH.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("History", dbOpenTable)
    
    '''Set rs = db.OpenRecordset("SELECT * FROM Tb WHERE tdID=10")
        If rec.EOF = False And rec.BOF = False Then
            rec.MoveFirst
            While rec.EOF = False
                If rec.Fields("InvoiceNumber") = InvoiceNumber Then rec.Delete
                If rec.EOF = False Then rec.MoveNext
            Wend
        End If
    db.Close
 '' rec.Close
' END ---- DELETE BIH ----
    
 '' END -- Update BIH (DELETE)
    
 '' Start -- Update MBIH : (delete)
    '' Start -- Check if File exists
    '' End -- Check if File exists
    '------------------------------------------
' Start Delete MembersFileDB
  Call Transaction_CheckForMembersRecordsIfNotExistsCreate(App.Path & "\Transaction\MembersRecords\", MemberFileNameOfHisDB)
  Set db = OpenDatabase(App.Path & "\Transaction\MembersRecords\" & MemberFileNameOfHisDB _
             , False, False, ";pwd=AdmiN")
  Set rec = db.OpenRecordset("Members Borrowed Items History", dbOpenTable)
    ' Start -- GetPrev Returned Status
      If rec.BOF = False And rec.EOF = False Then
         rec.MoveFirst
         ReDim PrevItemCode(rec.RecordCount)
         ReDim PrevReturnedStatus(rec.RecordCount)
         For loop1 = 1 To rec.RecordCount
             PrevItemCode(loop1) = rec.Fields("Item Code")
             PrevReturnedStatus(loop1) = rec.Fields("Returned")
             If rec.EOF = False Then rec.MoveNext
         Next
      End If
    ' END -- Get Prev Returned Status
    If rec.BOF = False Then rec.MoveFirst
    While rec.BOF = False And rec.EOF = False
      If Trim(rec.Fields("Invoice Number")) = Trim(InvoiceNumber) Then
          rec.Delete
      End If
      If rec.EOF = False Then rec.MoveNext
    Wend
    rec.Close
    db.Close
  ' End Delete MembersFileDB
'------------------------------------------
    
 '' END -- Update MBIH
 
 '' Start -- Update Unreturned item(s) : (No to Yes)
    '' Start -- Check if File exists
    '' End -- Check if File exists
    Set db = OpenDatabase(App.Path & "\Database\CD_Tapes.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("CD Tapes Table", dbOpenTable)
    loop2 = rec.RecordCount
  
    If CheckArray(DeletedItemCodes()) = True Then
        rec.MoveFirst
        For loop1 = 1 To loop2
            For loop3 = 1 To UBound(DeletedItemCodes())
                If Trim(rec.Fields("Item Code")) = Trim(DeletedItemCodes(loop3)) Then
                    rec.Edit
                    rec.Fields("Available") = "Yes"
                    rec.Fields("[Borrower's DB]") = " "
                    rec.Fields("LastDateReturned") = Format(Now, "mm/dd/yyyy")
                    rec.Fields("LastDateReturnedAddInfo") = "Altered Transaction by " & UCase(gVarFirstName) & " " & UCase(Left(gVarMiddleName, 1)) & ". " & UCase(gVarFamilyName)
                    rec.Update
                End If
            Next
        If rec.EOF = False Then rec.MoveNext
        Next
    End If
 
    rec.Close
    db.Close
 '' END -- Update Unreturned item(s) : (No to Yes)

 
 '' START -- Update Transaction Record
    For loop1 = 1 To Flex.Rows - 1
      BorrowedItems = BorrowedItems & "ItemCode::::" & Flex.TextMatrix(loop1, 1) & vbCrLf & "Title::::" & Flex.TextMatrix(loop1, 2) & vbCrLf & "RentalAmount::::" & Flex.TextMatrix(loop1, 3) & vbCrLf
    Next
       BorrowedItems = BorrowedItems & "TotalAmountDue::::" & TotalAmountDue & vbCrLf
       BorrowedItems = BorrowedItems & "AmountPaid::::" & AmountPaid & vbCrLf
       BorrowedItems = BorrowedItems & "Change::::" & Change & vbCrLf
    
    Set db = OpenDatabase(App.Path & "\Transaction\Transaction.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("Transaction", dbOpenTable)
    
   rec.MoveFirst
    If rec.RecordCount > 0 Then
        For loop1 = 1 To rec.RecordCount
            TDM = DoEvents()
            If rec.Fields("InvoiceNumber") = InvoiceNumber Then
                '' Start Update
                rec.Edit
                rec.Fields("Date") = TransactionDate
                rec.Fields("Cashier") = CashierName
                rec.Fields("CashierID") = CashierID
                rec.Fields("InvoiceNumber") = Trim(InvoiceNumber)
                rec.Fields("MembersName") = MembersName
                rec.Fields("Members ID") = MembersID
                rec.Fields("Borrowed Items") = BorrowedItems
                rec.Update
                '' End Update
            End If
        
            rec.MoveNext  ' Move to the next record
        Next loop1
     End If
     db.Close
'-------------------------------------------
     '' Start Change Yes to No, AddInfo
    Set db = OpenDatabase(App.Path & "\Database\CD_Tapes.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("CD Tapes Table", dbOpenTable)
    
    rec.MoveFirst
    If rec.RecordCount > 0 Then
    For loop1 = 1 To Flex.Rows - 1
     rec.MoveFirst
        For loop2 = 1 To rec.RecordCount
            TDM = DoEvents()
            If rec.Fields("Item Code") = Flex.TextMatrix(loop1, 1) Then
                '' Start Change Yes to No, AddInfo
                rec.Edit
        ''''''''''''''''''''''''''''''''''''''''''
                rec.Fields("Available") = "No"
                rec.Fields("[Borrower's DB]") = MemberFileNameOfHisDB
            If CheckPrevItemCodeArray(PrevItemCode()) = True Then
                For loop3 = 1 To UBound(PrevItemCode())
                    If Trim(Flex.TextMatrix(loop1, 1)) = PrevItemCode(loop3) Then
                       rec.Fields("Available") = PrevReturnedStatus(loop3)
                       Exit For
                    End If
                Next
            End If
         ''''''''''''''''''''''''''''''''''''''''''
                rec.Fields("LastDateBorrowed") = TransactionDate
                rec.Fields("LastDateBorrowedAddInfo") = "Borrowed by: " & MembersName
                rec.Update
                '' End Change Yes to No, AddInfo
            End If
        
            If rec.EOF = False Then rec.MoveNext ' Move to the next record
        Next loop2
     Next loop1
     End If
      rec.Close
      db.Close
    '' End Change Yes to No. AddInfo
'------------------------------------------
'------------------------------------------
' Start Append to MembersFileDB
  Call Transaction_CheckForMembersRecordsIfNotExistsCreate(App.Path & "\Transaction\MembersRecords\", MemberFileNameOfHisDB)
  Set db = OpenDatabase(App.Path & "\Transaction\MembersRecords\" & MemberFileNameOfHisDB _
             , False, False, ";pwd=AdmiN")
  Set rec = db.OpenRecordset("Members Borrowed Items History", dbOpenTable)
   For loop1 = 1 To Flex.Rows - 1
        rec.AddNew
        rec.Fields("Date Borrowed") = TransactionDate
        rec.Fields("Invoice Number") = Trim(InvoiceNumber)
        rec.Fields("Item Code") = Flex.TextMatrix(loop1, 1)
        rec.Fields("Title") = Flex.TextMatrix(loop1, 2)
       ''''''''''''''''''''''''''''''''''''''''''
         rec.Fields("Returned") = "No"
      If CheckPrevItemCodeArray(PrevItemCode()) = True Then
         For loop3 = 1 To UBound(PrevItemCode())
             If Trim(Flex.TextMatrix(loop1, 1)) = PrevItemCode(loop3) Then
                 rec.Fields("Returned") = PrevReturnedStatus(loop3)
                 Exit For
             End If
         Next
       End If
      ''''''''''''''''''''''''''''''''''''''''''
        rec.Update
  Next loop1
  
  rec.Close
  db.Close
' End Append to MembersFileDB
'------------------------------------------
    Call Transaction_AppendToBorrowedItemsHistory(Flex, TransactionDate, gVarUserID, InvoiceNumber, MembersName)
    
    MsgBox "Edited transaction has been successfully saved. ", vbInformation, "Transaction successful"
 '' END -- Update Transaction Record
    
  
End Sub

Public Sub Transaction_DeletePrevTransaction(InvoiceNumber As String, MemberFileNameOfHisDB As String)
   '' Locate InvoiceNumber
   '' Get List of items
   '' Delete items in MBIH and Modify CD_Tapes Table
   '' Delete Transaction Record
   '=================================================
   Dim loop111, loop1, rcount, LenStr, counter As Long
   Dim db, db2 As Database
   Dim rec, rec2 As Recordset
   Dim StringText, ItemCode(), Char, tmpString, FlagField As String
   ' START LOCATE INVOICENUMBER
   ' Chk if file exists if not create
   Set db = OpenDatabase(App.Path & "\Transaction\Transaction.mdb" _
             , False, False, ";pwd=AdmiN")
   Set rec = db.OpenRecordset("Transaction", dbOpenTable)
       
   If rec.BOF = False And rec.EOF = False Then
        rcount = rec.RecordCount
        rec.MoveFirst
        For loop111 = 1 To rcount
            If rec.Fields("InvoiceNumber") = Trim(InvoiceNumber) Then  '+++++++++++++++++++++
               StringText = rec.Fields("Borrowed Items")
               LenStr = Len(StringText)
               For loop1 = 1 To LenStr
                    Char = Mid(StringText, loop1, 1)
                 ' ------ Start "::::" Detect -----------------------
                    If Char = ":" Then
                        If loop1 + 3 <= LenStr Then
                            If Mid(StringText, loop1, 4) = "::::" Then
                                'Text2.Text = Text2.Text & Trim(tmpString)
                                FlagField = Trim(tmpString)
                                    If Trim(tmpString) = "ItemCode" Then
                                        counter = counter + 1
                                        ReDim Preserve ItemCode(counter)
                                       ' ReDim Preserve Title(Counter)
                                       ' ReDim Preserve RentalAmount(Counter)
                                    End If
                              tmpString = ""
                              loop1 = loop1 + 3
                            End If
                    End If
            Else
                If Char <> Chr(13) Or Char <> Chr(10) Then tmpString = tmpString & Char
            End If
                ' ------ End "::::" Detect --------------------------
                ' ------ Start vbLFCR detect ------------------------
                If Char = Chr(13) Then
                    If loop1 + 1 <= LenStr Then
                        If Mid(StringText, loop1, 2) = vbCrLf Then
                            If counter = 0 Then counter = 1
                    
                                Select Case FlagField
                                    Case "ItemCode": ItemCode(counter) = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                                    'Case "Title": Title(Counter) = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                                    'Case "RentalAmount": RentalAmount(Counter) = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                                    'Case "TotalAmountDue": TotalAmountDue = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                                    'Case "AmountPaid": AmountPaid = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                                    'Case "Change": Change = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                                    '''MsgBox ItemCode(Counter)
                                End Select
                    
                                    tmpString = ""
                                    loop1 = loop1 + 1
                            End If
                        End If
   
                End If
                ' ------ END vbLFCR detect ------------------------
                Next loop1
            End If '+++++++++++++++++++++
            If rec.EOF = False Then rec.MoveNext
        Next
        
   Else
      'do nothing
   End If
     rec.Close
     db.Close
   ' END LOCATE INVOICENUMBER
   '--------------------------------------------------------------------
   
   ' Start Delete MembersFileDB
  
  Call Transaction_CheckForMembersRecordsIfNotExistsCreate(App.Path & "\Transaction\MembersRecords\", MemberFileNameOfHisDB)
  Set db = OpenDatabase(App.Path & "\Transaction\MembersRecords\" & MemberFileNameOfHisDB _
             , False, False, ";pwd=AdmiN")
  Set rec = db.OpenRecordset("Members Borrowed Items History", dbOpenTable)
   ' Open CD_TApes TABLE ///////////////////////////////////////
    Set db2 = OpenDatabase(App.Path & "\Database\CD_Tapes.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec2 = db2.OpenRecordset("CD Tapes Table", dbOpenTable)
   ' End Open CD_Tapes TABLE////////////////////////////////
      If rec.BOF = False Then rec.MoveFirst
    While rec.BOF = False And rec.EOF = False
      If Trim(rec.Fields("Invoice Number")) = Trim(InvoiceNumber) Then
          If rec.Fields("Returned") = "No" Then
             If rec2.BOF = False And rec2.EOF = False Then
                rec2.MoveFirst
                Do While rec2.EOF = False 'Edit CD_Tapes Table: No to Yes
                  If rec2.Fields("Item Code") = rec.Fields("Item Code") Then
                     rec2.Edit
                     rec2.Fields("Available") = "Yes"
                     rec2.Fields("[Borrower's DB]") = " "
                     rec2.Update
                     Exit Do
                  End If
                  If rec2.BOF = False And rec2.EOF = False Then rec2.MoveNext
                Loop
             End If
          End If
          rec.Delete
      End If
      If rec.EOF = False Then rec.MoveNext
    Wend
    rec.Close
    db.Close
    rec2.Close
    db2.Close
  ' End Delete MembersFileDB
'------------------------------------------

' START ---- DELETE BIH ----
    Call Transaction_ChkIfBorrowedItemsHistoryDBExistIfNotCreate
    Set db = OpenDatabase(App.Path & "\Transaction\BIH.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("History", dbOpenTable)
    
    '''Set rs = db.OpenRecordset("SELECT * FROM Tb WHERE tdID=10")
        If rec.EOF = False And rec.BOF = False Then
            rec.MoveFirst
            While rec.EOF = False
                If rec.Fields("InvoiceNumber") = InvoiceNumber Then rec.Delete
                If rec.EOF = False Then rec.MoveNext
            Wend
        End If
    db.Close
 '' rec.Close
' END ---- DELETE BIH ----


'-----------------------------------------------------
' START -- Delete Transaction Record
 Set db = OpenDatabase(App.Path & "\Transaction\Transaction.mdb" _
             , False, False, ";pwd=AdmiN")
   Set rec = db.OpenRecordset("Transaction", dbOpenTable)
       
   If rec.BOF = False And rec.EOF = False Then
        rcount = rec.RecordCount
        rec.MoveFirst
        For loop111 = 1 To rcount
            If rec.Fields("InvoiceNumber") = Trim(InvoiceNumber) Then
               rec.Delete
               Exit For
            End If
            If rec.EOF = False Then rec.MoveNext
        Next
    End If
' END -- Delete Transaction Record
'-----------------------------------------------------
MsgBox "Transaction has been successfully deleted.  ", vbInformation, "Deleted"
   
   
End Sub

Sub Report_LoadUnreturnedItems(Flex As MSFlexGrid, mySQL As String)
    On Error GoTo ErrHandler
    Dim TDM As Variant
    Dim loop1, loop2 As Long
    Dim OverDue As Double
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String

    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset

    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase\CD_Tapes.mdb;Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
   

    

'' ---------Start SQL
''mySQL = "SELECT * FROM [CD TAPES TABLE] WHERE Available = 'No' ORDER BY LastDateBorrowed "
adoRecordset.Open mySQL, adoConnection, adOpenStatic, adLockOptimistic, _
           adCmdText
'' --------End SQL

If adoRecordset.BOF = True And adoRecordset.EOF = True Then
           MsgBox "No records found matching your search. ", vbInformation, "Not found"
           Flex.Rows = 1
           Exit Sub
End If
  '---------- Forward Processing ------------------------------------------
        Flex.Rows = 1
        loop1 = 0
        
        adoRecordset.MoveFirst
        Do Until adoRecordset.EOF
             TDM = DoEvents()
             loop1 = loop1 + 1
             Flex.AddItem ""
                For loop2 = 0 To 8
                    Select Case loop2
                        Case 0:
                                Flex.TextMatrix(loop1, loop2) = str(loop1)
                        Case 1:
                                Flex.TextMatrix(loop1, loop2) = adoRecordset![Item Code]
                        Case 2:
                                Flex.TextMatrix(loop1, loop2) = adoRecordset![Title]
                        Case 3:
                                Flex.TextMatrix(loop1, loop2) = adoRecordset!LastDateBorrowed
                                
                        Case 4:
                                Flex.TextMatrix(loop1, loop2) = Transaction_GetDateDue(adoRecordset![Item Code], Format(adoRecordset!LastDateBorrowed, "mm/dd/yy"))
                        Case 5:
                                ' Start - Calulate No. of days Overdue
                                OverDue = GetDayCount(adoRecordset!LastDateBorrowed, Now) - Val(adoRecordset!RentalPeriod)
                                If OverDue < 0 Then
                                  Flex.TextMatrix(loop1, loop2) = "0"
                                Else
                                  Flex.TextMatrix(loop1, loop2) = Trim(str(OverDue))
                                End If
                                ' End - Calulate No. of days Overdue
                                ' Start - Calculate Overdue Charges
                                  If OverDue > 0 Then
                                     OverDue = OverDue * Val(adoRecordset!OverdueChargePerDay)
                                  Else
                                     OverDue = 0
                                  End If
                                ' End - Calculate Overdue Charges
                        Case 6:
                                'Display OverDue Charge
                                Flex.TextMatrix(loop1, loop2) = Format(Trim(str(Round(OverDue, 2))), "0.00")
                        Case 7:
                                Flex.TextMatrix(loop1, loop2) = "  " & adoRecordset!LastDateBorrowedAddInfo
                        
                    End Select
                Next loop2
            adoRecordset.MoveNext
         Loop
         'Flex.Rows = Flex.Rows - 1
  '-----------------------------------------------
  Exit Sub
ErrHandler:
End Sub
Sub Report_LoadItemsToBeReturnedToday(Flex As MSFlexGrid)
    'On Error GoTo ErrHandler
    Dim TDM As Variant
    Dim mySQL As String
    Dim loop1, loop2 As Long
    Dim OverDue As Double
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String

    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset

    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase\CD_Tapes.mdb;Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
   

    

'' ---------Start SQL
mySQL = "SELECT * FROM [CD TAPES TABLE] WHERE Available = 'No' ORDER BY LastDateBorrowed "
adoRecordset.Open mySQL, adoConnection, adOpenStatic, adLockOptimistic, _
           adCmdText
'' --------End SQL

If adoRecordset.BOF = True And adoRecordset.EOF = True Then
           MsgBox "No records found matching your search. ", vbInformation, "Not found"
           Flex.Rows = 1
           Exit Sub
End If

' Start -- Populate Flex
Dim counter As Long
'Dim OverDue As Double
counter = 0
Flex.Rows = 1
Do While adoRecordset.EOF = False
    OverDue = GetDayCount(adoRecordset!LastDateBorrowed, Now) - Val(adoRecordset!RentalPeriod)
    If OverDue < 0 Then
       If adoRecordset.EOF = False Then adoRecordset.MoveNext
    Else
        OverDue = OverDue * Val(adoRecordset!OverdueChargePerDay)
        Flex.AddItem ""
        counter = counter + 1
        Flex.TextMatrix(counter, 0) = str(counter) 'No. Col
        Flex.TextMatrix(counter, 1) = adoRecordset![Item Code]
        Flex.TextMatrix(counter, 2) = adoRecordset!Title
        Flex.TextMatrix(counter, 3) = adoRecordset!LastDateBorrowed
        Flex.TextMatrix(counter, 4) = Transaction_GetDateDue(adoRecordset![Item Code], Format(adoRecordset!LastDateBorrowed, "mm/dd/yy"))
        Flex.TextMatrix(counter, 5) = Format(OverDue, "0.00") & "          "
        Flex.TextMatrix(counter, 6) = adoRecordset!LastDateBorrowedAddInfo
    End If
    If adoRecordset.EOF = False Then adoRecordset.MoveNext
Loop
End Sub
Function Return_Items(lst As ListBox) As Boolean
 Return_Items = False
 Dim listlines, loop1, counter As Long
 Dim ItemCodesToBeReturned(), tmpString, Char As String
 
'Start -- // Get ItemCodes and assign to ItemCodesToBeReturned()
 counter = 0 ' initialize
 Char = "A" ' initialize
 listlines = lst.ListCount
 For loop1 = 0 To listlines - 1
    Do While Char <> " "
        counter = counter + 1
        Char = Mid(Trim(lst.List(loop1)), counter, 1)
        tmpString = tmpString & Char
     Loop
     Char = "A" ' reset
     counter = 0 ' reset
     ReDim Preserve ItemCodesToBeReturned(loop1 + 1)
     ItemCodesToBeReturned(loop1 + 1) = Trim(tmpString) ' assign to array
     tmpString = "" 'reset
 Next
 'End -- // Get ItemCodes and assign to ItemCodesToBeReturned()
 
 '' Start - Update CD_Tapes and MBIH
    Dim db1 As Database
    Dim rec1 As Recordset
    Dim db2 As Database
    Dim rec2 As Recordset
    Dim TDM As Variant
    Dim tmpMemberDB As String
    Set db1 = OpenDatabase(App.Path & "\Database\CD_Tapes.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec1 = db1.OpenRecordset("CD Tapes Table", dbOpenTable)
    
    For loop1 = 1 To listlines
     rec1.MoveFirst
     Do While rec1.EOF = False
        TDM = DoEvents()
        If Trim(rec1.Fields("Available")) = "No" And Trim(rec1.Fields("Item Code")) = ItemCodesToBeReturned(loop1) Then
              tmpMemberDB = Trim(rec1.Fields("Borrower's DB"))
           ' Start - Update CD_TAPES_TABLE
              rec1.Edit
              rec1.Fields("Available") = "Yes"
              rec1.Fields("LastDateReturned") = Format(Now, "mm/dd/yyyy")
              rec1.Fields("LastDateReturnedAddInfo") = "Updated by: " & gVarFirstName & " " & Mid(gVarMiddleName, 1, 1) & ". " & gVarFamilyName
              rec1.Fields("Borrower's DB") = ""
              rec1.Update
           ' End - Update CD_TAPES_TABLE
           
           ' Start - Update MBIH
             Call CheckTransactionFolder
             Call Transaction_CheckForMembersRecordsIfNotExistsCreate(App.Path & "\Transaction\MembersRecords\", tmpMemberDB)
             Set db2 = OpenDatabase(App.Path & "\Transaction\MembersRecords\" & tmpMemberDB _
                        , False, False, ";pwd=AdmiN")
             Set rec2 = db2.OpenRecordset("Members Borrowed Items History", dbOpenTable)
               If rec2.EOF = False And rec2.BOF = False Then
                  rec2.MoveFirst
                  Do While rec2.EOF = False
                     If rec2.Fields("Item Code") = ItemCodesToBeReturned(loop1) And rec2.Fields("Returned") = "No" Then
                        rec2.Edit
                        rec2.Fields("Returned") = "Yes"
                        rec2.Update
                        rec2.Close
                        db2.Close
                        Exit Do
                     End If
                     rec2.MoveNext
                  Loop
                  
               End If
             
             
           ' End - Update MBIH
        End If
        rec1.MoveNext
     Loop
    Next
 '' End - Update CD_Tapes and MBIH
 
   rec1.Close
   db1.Close
   
   Set rec1 = Nothing
   Set db1 = Nothing
   Set rec2 = Nothing
   Set db2 = Nothing
   lst.Clear
   
   Return_Items = True
End Function
Sub Return_SaveReturnTransaction(Flex As MSFlexGrid, InvNum As String)
    Dim loop1 As Long
    Dim db As Database
    Dim rec As Recordset
    ''Increment InvNum
    Open App.Path & "\InvoiceNumber.txt" For Output As #1
        Print #1, InvNum
    Close #1
    ''End Increment InvNum
    Set db = OpenDatabase(App.Path & "\Transaction\MiscTrans.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec = db.OpenRecordset("Transaction", dbOpenTable)
    For loop1 = 1 To Flex.Rows - 1
        rec.AddNew
        rec.Fields("Date") = Format(Now, "mm/dd/yyyy")
        rec.Fields("Cashier") = UCase(gVarFirstName) & " " & UCase(Mid(gVarMiddleName, 1, 1)) & ". " & UCase(gVarFamilyName)
        rec.Fields("CashierID") = gVarUserID
        rec.Fields("InvoiceNumber") = Trim(InvNum)
        rec.Fields("Description") = "Overdue : " & Trim(Flex.TextMatrix(loop1, 0))
        rec.Fields("Amount") = Trim(Flex.TextMatrix(loop1, 1))
        rec.Update
    Next loop1
    rec.Close
    db.Close
    Set rec = Nothing
    Set db = Nothing
End Sub
Sub REPORT_MOVIESTAT_FILLCBO(cboGenre As ComboBox, cboItemCode As ComboBox, cboActor As ComboBox)
    Dim loop1 As Long
    Dim db1 As Database
    Dim rec1 As Recordset
    Dim TDM
    Set db1 = OpenDatabase(App.Path & "\Database\CD_Tapes.mdb" _
             , False, False, ";pwd=AdmiN")
    Set rec1 = db1.OpenRecordset("CD Tapes Table", dbOpenTable)
    
    Do While rec1.EOF = False
       If cboGenre.ListCount > 0 Then
          For loop1 = 0 To cboGenre.ListCount - 1
              TDM = DoEvents()
              If Trim(UCase(rec1.Fields("Genre"))) = Trim(UCase(cboGenre.List(loop1))) Then
                 GoTo JUMP1
              End If
          Next loop1
       End If
       cboGenre.AddItem rec1.Fields("Genre")
JUMP1:
       cboItemCode.AddItem rec1.Fields("Item Code")
       rec1.MoveNext
    Loop
    
    If rec1.BOF = False Then rec1.MoveFirst
    Do While rec1.EOF = False
       If cboActor.ListCount > 0 Then
          For loop1 = 0 To cboActor.ListCount - 1
              TDM = DoEvents()
              If Trim(UCase(rec1.Fields("Actor"))) = Trim(UCase(cboActor.List(loop1))) Then
                 GoTo JUMP2
              End If
          Next loop1
       End If
       cboActor.AddItem rec1.Fields("Actor")
JUMP2:
       rec1.MoveNext
    Loop

End Sub


Sub REPORT_GETMOVIESTAT(Flex As MSFlexGrid, DateStart As DTPicker, DateEnd As DTPicker, optDescending As OptionButton, mySQL As String)
'Flex As MSFlexGrid, StartDate As String, EndDate As String, mySQL As String
    Dim loop1, loop2 As Long
    Dim sTitle(), sItemCode() As String
    Dim TDM
  
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String

    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset

    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase\CD_Tapes.mdb;Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
   
'' ---------Start SQL
''mySQL = "SELECT * FROM [CD TAPES TABLE] WHERE Available = 'No' ORDER BY LastDateBorrowed "
''Dim mySQL As String
''mySQL = "SELECT DISTINCT Title FROM [CD TAPES TABLE] "
adoRecordset.Open mySQL, adoConnection, adOpenStatic, adLockOptimistic, _
           adCmdText
'' --------End SQL
    
'' START STORE IN Array
   Dim counter As Long
   '''Dim tmpstr As String
   counter = 0
  If adoRecordset.EOF = False And adoRecordset.EOF = False Then
       adoRecordset.MoveFirst
  Else
       MsgBox "No item(s) found.    ", vbInformation, " "
       Exit Sub
  End If
  
   Do While adoRecordset.EOF = False
        counter = counter + 1
        ReDim Preserve sTitle(counter)
        ReDim Preserve sItemCode(counter)
        sTitle(counter) = adoRecordset!Title
        'sItemCode(Counter) = adoRecordset![Item Code]
        '''tmpstr = tmpstr & vbCrLf & sItemcode(counter)
        adoRecordset.MoveNext
   Loop
   adoConnection.Close
   '''MsgBox tmpstr
'' END STORE IN Array

'Start OPEN HISTORY TABLE
Call Transaction_ChkIfBorrowedItemsHistoryDBExistIfNotCreate
connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Transaction\BIH.mdb;Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"
adoConnection.Open connectString
' END OPEN HISTORY TABLE

'HISTORY TABLE SQL
Dim mySQL2, Frequency() As String
For loop1 = 1 To counter
    mySQL2 = "SELECT * FROM History WHERE Title = """ & sTitle(loop1) & """" & " AND (Date >= #" & DateStart.Value & "# AND Date <= #" & DateEnd.Value & "#)"
    ' Detect if item code is to be search
    If Trim(Mid(mySQL, 42, 11)) = "[Item Code]" Then
       mySQL2 = "SELECT * FROM History WHERE [ItemCode] = """ & Trim(Mid(mySQL, 57, Len(mySQL) - 57)) & """" & " AND (Date >= #" & DateStart.Value & "# AND Date <= #" & DateEnd.Value & "#)"
    End If
     ''MsgBox Trim(Mid(mySQL, 57, Len(mySQL) - 57))
    adoRecordset.Open mySQL2, adoConnection, adOpenStatic, adLockOptimistic, _
                  adCmdText
    'MsgBox "found " & Str(adoRecordset.RecordCount)
    ReDim Preserve Frequency(loop1)
    Frequency(loop1) = str(adoRecordset.RecordCount)
    adoRecordset.Close
Next loop1

Select Case optDescending.Value

Case 1:
    '' START -- SORT DESC
        Dim tmpStr As String
            For loop1 = 1 To counter
                For loop2 = loop1 + 1 To counter
                    If Val(Frequency(loop1)) < Val(Frequency(loop2)) Then
                        'SORT FREQUENCY
                        tmpStr = Frequency(loop1)
                        Frequency(loop1) = Frequency(loop2)
                        Frequency(loop2) = tmpStr
                        'SORT sTitle
                        tmpStr = sTitle(loop1)
                        sTitle(loop1) = sTitle(loop2)
                        sTitle(loop2) = tmpStr
                        'SORT sItemCode
                    End If
                Next loop2
            Next loop1
    '' END -- SORT DESC
Case 0:
        '' START -- SORT ASC
                For loop1 = 1 To counter
                    For loop2 = loop1 + 1 To counter
                        If Val(Frequency(loop1)) > Val(Frequency(loop2)) Then
                            'SORT FREQUENCY
                            tmpStr = Frequency(loop1)
                            Frequency(loop1) = Frequency(loop2)
                            Frequency(loop2) = tmpStr
                            'SORT sTitle
                            tmpStr = sTitle(loop1)
                            sTitle(loop1) = sTitle(loop2)
                            sTitle(loop2) = tmpStr
                            'SORT sItemCode
                        End If
                    Next loop2
                Next loop1
        '' END -- SORT ASC

End Select

Flex.Rows = 1
For loop1 = 1 To counter
    Flex.AddItem ""
    Flex.TextMatrix(loop1, 0) = str(loop1)
    Flex.TextMatrix(loop1, 1) = sTitle(loop1)
    Flex.TextMatrix(loop1, 2) = Frequency(loop1)
Next loop1
End Sub
Sub REPORT_GETMOVIESTAT_FillCBOcashier(cbo As ComboBox)
' START -- Chk if transaction file exists
  If ReportFileStatus(App.Path & "\Transaction\Transaction.mdb") = False Then
     'Exit w/o msg
     Exit Sub
  End If
' END -- Chk if transaction file exists


'' Start -- SEARCH TRANSACTION DB
  ''Chk first if DB exist
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString, mySQL As String
    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset
    
    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Transaction\Transaction.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
    'NOTE: Spc b4 invoicenumber in BD                                      d2 ang spc
    mySQL = "SELECT DISTINCT Cashier FROM [Transaction]"
    adoRecordset.Open mySQL, adoConnection, adOpenStatic, adLockOptimistic, adCmdText '"SELECT * from [Transaction] WHERE [InvoiceNumber]= '" & txtInvoiceNumber.Text & "'", adoConnection, adOpenStatic, adLockOptimistic, adCmdText
   If adoRecordset.BOF = True And adoRecordset.EOF = True Then
      ' Exit w/o msg
      adoRecordset.Close
      adoConnection.Close
      Set adoRecordset = Nothing
      Set adoConnection = Nothing
      Exit Sub
   Else
      adoRecordset.MoveFirst
      cbo.Clear
      Do While adoRecordset.EOF = False
        cbo.AddItem adoRecordset!Cashier
        adoRecordset.MoveNext
      Loop
   End If
End Sub
Sub Report_GetSalesDetailed(Flex As MSFlexGrid, mySQL As String)
Dim Char, StringText, tmpString, FlagField As String
Dim LenStr, counter, loop1 As Long
Dim ItemCode(), Title(), RentalAmount(), _
    TotalAmountDue, AmountPaid, Change As String

' START -- Chk if transaction file exists
  If ReportFileStatus(App.Path & "\Transaction\Transaction.mdb") = False Then
  
     MsgBox "Transaction Database is missing.  ", vbInformation, "Database not found."
     Flex.SetFocus
     Exit Sub
  End If
' END -- Chk if transaction file exists


'' Start -- SEARCH TRANSACTION DB
  ''Chk first if DB exist
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String
    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset
    
    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Transaction\Transaction.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
    'NOTE: Spc b4 invoicenumber in BD                                      d2 ang spc
    adoRecordset.Open mySQL, adoConnection, adOpenStatic, adLockOptimistic, adCmdText '"SELECT * from [Transaction] WHERE [InvoiceNumber]= '" & txtInvoiceNumber.Text & "'", adoConnection, adOpenStatic, adLockOptimistic, adCmdText
   If adoRecordset.BOF = True And adoRecordset.EOF = True Then
      MsgBox "Transaction Not Found. ", vbInformation, "Info"
      Flex.Rows = 1
      Flex.AddItem ""
      Flex.SetFocus
      adoRecordset.Close
      adoConnection.Close

      Set adoRecordset = Nothing
      Set adoConnection = Nothing
      Exit Sub
   Else
      adoRecordset.MoveFirst
      Flex.Rows = 1
      Dim FlexRowCount As Long
      FlexRowCount = 0
      Do While adoRecordset.EOF = False
      
            StringText = adoRecordset![Borrowed Items]
            'MemberID_FindMode = adoRecordset![Members ID]
            LenStr = Len(StringText)
'=================================================
For loop1 = 1 To LenStr
    Char = Mid(StringText, loop1, 1)
 ' ------ Start "::::" Detect -----------------------
    If Char = ":" Then
         If loop1 + 3 <= LenStr Then
            If Mid(StringText, loop1, 4) = "::::" Then
               'Text2.Text = Text2.Text & Trim(tmpString)
               FlagField = Trim(tmpString)
               If Trim(tmpString) = "ItemCode" Then
                    counter = counter + 1
                    ReDim Preserve ItemCode(counter)
                    ReDim Preserve Title(counter)
                    ReDim Preserve RentalAmount(counter)
               End If
               tmpString = ""
               loop1 = loop1 + 3
            End If
         End If
    Else
       If Char <> Chr(13) Or Char <> Chr(10) Then tmpString = tmpString & Char
    End If
 ' ------ End "::::" Detect --------------------------
 ' ------ Start vbLFCR detect ------------------------
    If Char = Chr(13) Then
       If loop1 + 1 <= LenStr Then
          If Mid(StringText, loop1, 2) = vbCrLf Then
               If counter = 0 Then counter = 1
                    
                    Select Case FlagField
                      Case "ItemCode": ItemCode(counter) = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                      Case "Title": Title(counter) = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                      Case "RentalAmount": RentalAmount(counter) = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                      Case "TotalAmountDue": TotalAmountDue = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                      Case "AmountPaid": AmountPaid = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                      Case "Change": Change = Mid(Trim(tmpString), 1, Len(Trim(tmpString)) - 1)
                    End Select
                    
                    tmpString = ""
                    loop1 = loop1 + 1
                End If
          End If
   
    End If
  ' ------ END vbLFCR detect ------------------------
Next loop1
 
  For loop1 = 0 To counter - 1
      Flex.AddItem ""
      FlexRowCount = FlexRowCount + 1
      Flex.TextMatrix(FlexRowCount, 1) = adoRecordset!InvoiceNumber
      Flex.TextMatrix(FlexRowCount, 2) = Format(adoRecordset![Date], "mmm. dd, yyyy")
      Flex.TextMatrix(FlexRowCount, 3) = adoRecordset!Cashier
      Flex.TextMatrix(FlexRowCount, 4) = adoRecordset![MembersName]

      Flex.TextMatrix(FlexRowCount, 0) = FlexRowCount 'Str(loop1 + 1)
      Flex.TextMatrix(FlexRowCount, 5) = ItemCode(loop1 + 1)
      Flex.TextMatrix(FlexRowCount, 6) = Title(loop1 + 1)
      Flex.TextMatrix(FlexRowCount, 7) = RentalAmount(loop1 + 1) & " "
      If loop1 < counter - 1 Then Flex.AddItem ""
  Next
'====================================================
            counter = 0 ' reset
            adoRecordset.MoveNext
      Loop
   End If
   Flex.Rows = FlexRowCount + 1
   adoRecordset.Close
   adoConnection.Close
   Set adoRecordset = Nothing
   Set adoConnection = Nothing
'' END -- SEARCH TRANSACTION DB

End Sub

Sub Report_GetMiscTrans(Flex As MSFlexGrid, mySQL As String)
' START -- Chk if MiscTransaction file exists
  If ReportFileStatus(App.Path & "\Transaction\MiscTrans.mdb") = False Then
     MsgBox "Overdue/Misc Transaction Database is missing.  ", vbInformation, "Database not found."
     Flex.SetFocus
     Exit Sub
  End If
' END -- Chk if MiscTransaction file exists

'Start DB Search
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim connectString As String
    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset
    
    connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Transaction\MiscTrans.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=AdmiN"

    adoConnection.Open connectString
    'NOTE: Spc b4 invoicenumber in BD                                      d2 ang spc
    adoRecordset.Open mySQL, adoConnection, adOpenStatic, adLockOptimistic, adCmdText '"SELECT * from [Transaction] WHERE [InvoiceNumber]= '" & txtInvoiceNumber.Text & "'", adoConnection, adOpenStatic, adLockOptimistic, adCmdText
   If adoRecordset.BOF = True And adoRecordset.EOF = True Then
      MsgBox "Transaction Not Found. ", vbInformation, "Info"
      Flex.Rows = 1
      Flex.AddItem ""
      Flex.SetFocus
      adoRecordset.Close
      adoConnection.Close

      Set adoRecordset = Nothing
      Set adoConnection = Nothing
      Exit Sub
   Else
      adoRecordset.MoveFirst
      Flex.Rows = 1
      Dim FlexRowCount As Long
      FlexRowCount = 0
      Do While adoRecordset.EOF = False
         Flex.AddItem ""
         FlexRowCount = FlexRowCount + 1
         Flex.TextMatrix(FlexRowCount, 0) = Trim(str(FlexRowCount))
         Flex.TextMatrix(FlexRowCount, 1) = adoRecordset!InvoiceNumber
         Flex.TextMatrix(FlexRowCount, 2) = Format(adoRecordset!Date, "mmm. dd, yyyy")
         Flex.TextMatrix(FlexRowCount, 3) = adoRecordset!Cashier
         Flex.TextMatrix(FlexRowCount, 4) = adoRecordset!Description
         Flex.TextMatrix(FlexRowCount, 5) = adoRecordset!Amount & "    "
         If adoRecordset.EOF = False Then adoRecordset.MoveNext
      Loop
      Flex.SetFocus
   End If
'End DB Search



End Sub


