Attribute VB_Name = "VRENTAL_MODULE1"
Option Explicit

Global LogOnResult As String ''User Full Info
Global gVarUserID, gVarDateEntered, gVarUserName, gVarPassword, gVarAccessLevel, _
       gVarFirstName, gVarMiddleName, gVarFamilyName, gVarBirthday, gVarAge, gVarSex, gVarHomeAddress, gVarContactNumber, _
       gVarComments, gVarLogInDate, gVarLogInTime As String

Sub AssignUserInfoToGlobalVar(UserInfo As String)
    Dim TDM
    Dim loop1, Counter As Integer
    Dim tmpString, Char As String
    Counter = 1
    gVarLogInTime = Format$(Now, "hh:mm:ss") ' Get LogIn Time
    gVarLogInDate = Format$(Now, "mm/dd/yyyy") ' Get LogIn Date
    For loop1 = 1 To Len(UserInfo)
        TDM = DoEvents()
        Char = Mid(UserInfo, loop1, 1)
        If Char = Chr(10) Then
           tmpString = Trim(Mid(tmpString, InStr(1, tmpString, ":", 1) + 2, Len(tmpString) - InStr(1, tmpString, ":", 1)))
              If Counter = 1 Then gVarUserID = tmpString
              If Counter = 2 Then gVarDateEntered = tmpString
              If Counter = 3 Then gVarUserName = tmpString
              If Counter = 4 Then gVarPassword = tmpString
              If Counter = 5 Then gVarAccessLevel = tmpString
              If Counter = 6 Then gVarFirstName = tmpString
              If Counter = 7 Then gVarMiddleName = tmpString
              If Counter = 8 Then gVarFamilyName = tmpString
              If Counter = 9 Then gVarBirthday = tmpString
              If Counter = 10 Then gVarAge = tmpString
              If Counter = 11 Then gVarSex = tmpString
              If Counter = 12 Then gVarHomeAddress = tmpString
              If Counter = 13 Then gVarContactNumber = tmpString
              If Counter = 14 Then gVarComments = tmpString
              If Counter = 15 Then gVarLogInDate = tmpString
              If Counter = 16 Then gVarLogInTime = tmpString
           tmpString = ""
           Counter = Counter + 1
        End If
        If Char <> Chr(13) And Char <> Chr(10) Then tmpString = tmpString & Char
    Next loop1
    
 
 ' MsgBox gVarUserID & vbCrLf & gVarDateEntered & vbCrLf & gVarUserName _
          & vbCrLf & gVarPassword & vbCrLf & gVarAccessLevel _
          & vbCrLf & gVarFirstName & vbCrLf & gVarMiddleName _
          & vbCrLf & gVarFamilyName & vbCrLf & gVarBirthday _
          & vbCrLf & gVarAge & vbCrLf & gVarSex & vbCrLf & gVarHomeAddress _
          & vbCrLf & gVarContactNumber & vbCrLf & gVarComments _
          & vbCrLf & gVarLogInDate & vbCrLf & gVarLogInTime
    

End Sub



