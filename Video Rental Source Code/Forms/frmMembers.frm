VERSION 5.00
Begin VB.Form frmMembers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Members"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   Icon            =   "frmMembers.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   10575
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   840
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddMember 
      Caption         =   "&Add Member"
      Height          =   375
      Left            =   5420
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6630
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7840
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9060
      TabIndex        =   33
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txtNationality 
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Nationality"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtComments 
      Height          =   1245
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   17
      ToolTipText     =   "Comments"
      Top             =   4440
      Width           =   7095
   End
   Begin VB.TextBox txtOfficeSchoolAddress 
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   16
      ToolTipText     =   "Office or School and address"
      Top             =   3840
      Width           =   7095
   End
   Begin VB.TextBox txtHomeAddress 
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   15
      ToolTipText     =   "Home Address"
      Top             =   3240
      Width           =   7095
   End
   Begin VB.TextBox txtContactNumber 
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "Contact Number"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtOccupation 
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "Occupation"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtCivilStatus 
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   12
      ToolTipText     =   "Civil Status"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtSex 
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   11
      ToolTipText     =   "Sex"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtAge 
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Age"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtBirthday 
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "Date of birth (e.g. Dec. 25, 1976"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtFamilyName 
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Family Name"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtMiddleName 
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Middle Name"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "First Name"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtIDnumber 
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "ID Number"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtDateEntered 
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Date Entered"
      Top             =   840
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   " Detailed Information "
      Height          =   5775
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   7575
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Membership Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   5040
         TabIndex        =   32
         Tag             =   "&Password:"
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   14
         Left            =   240
         TabIndex        =   31
         Tag             =   "&Password:"
         Top             =   4080
         Width           =   750
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   13
         Left            =   5040
         TabIndex        =   30
         Tag             =   "&Password:"
         Top             =   2280
         Width           =   1155
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Office/School Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   12
         Left            =   240
         TabIndex        =   29
         Tag             =   "&Password:"
         Top             =   3480
         Width           =   1665
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   11
         Left            =   240
         TabIndex        =   28
         Tag             =   "&Password:"
         Top             =   2880
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   10
         Left            =   2640
         TabIndex        =   27
         Tag             =   "&Password:"
         Top             =   2280
         Width           =   825
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   240
         TabIndex        =   26
         Tag             =   "&Password:"
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   8
         Left            =   5040
         TabIndex        =   25
         Tag             =   "&Password:"
         Top             =   1680
         Width           =   285
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   7
         Left            =   2640
         TabIndex        =   24
         Tag             =   "&Password:"
         Top             =   1680
         Width           =   300
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birthday"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   240
         TabIndex        =   23
         Tag             =   "&Password:"
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Family Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   5040
         TabIndex        =   22
         Tag             =   "&Password:"
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   2640
         TabIndex        =   21
         Tag             =   "&Password:"
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Tag             =   "&Password:"
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   2640
         TabIndex        =   19
         Tag             =   "&Password:"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Entered"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Tag             =   "&Password:"
         Top             =   480
         Width           =   1170
      End
   End
   Begin VB.ListBox lstMembers 
      Height          =   5130
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "List of members"
      Top             =   480
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   " Members "
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim lstTEXT As String

Private Sub cmdAddMember_Click()

    If cmdAddMember.Caption = "&Add Member" Then
        Me.Caption = "Members (Add Mode)"
        lstMembers.Enabled = False
        cmdEdit.Enabled = False
        cmdAddMember.Caption = "C&ancel Add..."
        txtDateEntered.Text = Format(Now, "mmm. dd, yyyy")
        txtIDnumber.Text = "(AUTO-NUMBER)"
        cmdUpdate.Enabled = True
        cmdClear.Enabled = True
        Call cmdClear_Click
        txtAge.Text = "(SKIP THIS FEILD)"
        Call UnlockTextboxes
        txtNationality.SetFocus
        Exit Sub
    End If
    
    If cmdAddMember.Caption = "C&ancel Add..." Then
        Me.Caption = "Members"
        Call cmdClear_Click
        txtDateEntered.Text = ""
        txtIDnumber.Text = ""
        txtAge.Text = ""
        Call LockTextboxes
        
        lstMembers.Enabled = True
        If Trim(lstTEXT) <> "" Then
                lstMembers.Text = lstTEXT
        Else
                lstMembers.SetFocus
        End If
        cmdAddMember.Caption = "&Add Member"
        cmdUpdate.Enabled = False
        cmdClear.Enabled = False
        
        Exit Sub
    End If

End Sub

Private Sub cmdClear_Click()
     
     txtNationality.Text = ""
     txtFirstName.Text = ""
     txtMiddleName.Text = ""
     txtFamilyName.Text = ""
     txtBirthday.Text = ""
     txtSex.Text = ""
     txtCivilStatus.Text = ""
     txtOccupation.Text = ""
     txtContactNumber.Text = ""
     txtHomeAddress.Text = ""
     txtOfficeSchoolAddress.Text = ""
     txtComments.Text = ""
     txtNationality.SetFocus
End Sub

Private Sub cmdEdit_Click()
  
'------------------------------------------
  If cmdEdit.Caption = "&Edit" Then
     Me.Caption = "Members (Edit Mode)"
     cmdEdit.Caption = "Ca&ncel Edit"
     lstMembers.Enabled = False
     cmdAddMember.Enabled = False
     cmdClear.Enabled = True
     cmdUpdate.Enabled = True
     Call UnlockTextboxes
     txtNationality.SetFocus
     Exit Sub
  End If
'------------------------------------------
  If cmdEdit.Caption = "Ca&ncel Edit" Then
     Me.Caption = "Members"
     cmdEdit.Caption = "&Edit"
     lstMembers.Enabled = True
     cmdAddMember.Enabled = True
     cmdClear.Enabled = False
     cmdUpdate.Enabled = False
     If Trim(lstTEXT) <> "" Then
                lstMembers.Text = lstTEXT
                lstMembers.SetFocus
        Else
                lstMembers.SetFocus
        End If
     Call LockTextboxes
     Exit Sub
  End If
 '-----------------------------------------
  
End Sub

Private Sub cmdRemove_Click()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE

If IsNumeric(txtIDnumber.Text) = True And lstMembers.Enabled = True Then
   If MsgBox("Are you sure you want to delete this user? ", vbYesNo) = vbNo Then Exit Sub
  Call vr_engine.RemoveMember(Int(txtIDnumber.Text))
  Call vr_engine.LoadMembers(lstMembers) ' Refresh Members list
    Call cmdClear_Click
    txtDateEntered.Text = ""
    txtIDnumber.Text = ""
    txtAge.Text = ""
    cmdEdit.Enabled = False
    lstMembers.SetFocus
End If

End Sub

Private Sub cmdUpdate_Click()
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    Dim X As Boolean
'-----------------------------------------------------
    If cmdAddMember.Caption = "C&ancel Add..." Then
    
    If ValidateMembersInfo = True Then
        X = vr_engine.AddMemberToDB(txtDateEntered, txtIDnumber, txtNationality, _
                      txtFirstName, txtMiddleName, txtFamilyName, _
                      txtBirthday, txtSex, txtCivilStatus, txtOccupation, _
                      txtContactNumber, txtHomeAddress, txtOfficeSchoolAddress, txtComments)
        Call vr_engine.LoadMembers(lstMembers) ' Refresh Members list
        Me.Caption = "Members"
    Else
        Exit Sub
    End If
    
    ' Start - Enable Controls
    lstMembers.Enabled = True
    cmdAddMember.Caption = "&Add Member"
    cmdAddMember.Enabled = True
    lstMembers.Enabled = True
    lstMembers.Text = Trim(txtFamilyName.Text) & ", " & Trim(txtFirstName.Text) & " " & Left(Trim(txtMiddleName.Text), 1) & "."
    ' End - Enable Controls
    ' Start - Disable Controls
    cmdUpdate.Enabled = False
    cmdClear.Enabled = False
    ' End - Disable Controls
    Call LockTextboxes
    lstMembers.SetFocus
    Exit Sub
    End If
'-----------------------------------------------------
'-----------------------------------------------------
  If cmdEdit.Caption = "Ca&ncel Edit" Then
  
    If ValidateMembersInfo = True Then
        If vr_engine.UpdateEditedMembersDB(txtDateEntered, txtIDnumber, txtNationality, _
                      txtFirstName, txtMiddleName, txtFamilyName, _
                      txtBirthday, txtSex, txtCivilStatus, txtOccupation, _
                      txtContactNumber, txtHomeAddress, txtOfficeSchoolAddress, txtComments) = True Then
                 '' Do Nothing
        Else
                Exit Sub
        End If
        Call vr_engine.LoadMembers(lstMembers) ' Refresh Members list
        Me.Caption = "Members"
        cmdEdit.Caption = "&Edit"
        lstMembers.Enabled = True
        cmdAddMember.Enabled = True
        cmdClear.Enabled = False
        cmdUpdate.Enabled = False
        Call LockTextboxes
        lstMembers.Text = Trim(txtFamilyName.Text) & ", " & Trim(txtFirstName.Text) & " " & Left(Trim(txtMiddleName.Text), 1) & "."
        lstMembers.SetFocus
    Else
        Exit Sub
    End If
  End If
'-----------------------------------------------------
 
End Sub

Private Sub Form_Load()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE

Call vr_engine.LoadMembers(lstMembers)



End Sub
Sub LockTextboxes()

     txtNationality.Locked = True
     txtFirstName.Locked = True
     txtMiddleName.Locked = True
     txtFamilyName.Locked = True
     txtBirthday.Locked = True
     txtSex.Locked = True
     txtCivilStatus.Locked = True
     txtOccupation.Locked = True
     txtContactNumber.Locked = True
     txtHomeAddress.Locked = True
     txtOfficeSchoolAddress.Locked = True
     txtComments.Locked = True
     
End Sub
Sub UnlockTextboxes()

     txtNationality.Locked = False
     txtFirstName.Locked = False
     txtMiddleName.Locked = False
     txtFamilyName.Locked = False
     txtBirthday.Locked = False
     txtSex.Locked = False
     txtCivilStatus.Locked = False
     txtOccupation.Locked = False
     txtContactNumber.Locked = False
     txtHomeAddress.Locked = False
     txtOfficeSchoolAddress.Locked = False
     txtComments.Locked = False
     
End Sub
Function ValidateMembersInfo() As Boolean
  ValidateMembersInfo = False
    
    If Trim(txtNationality.Text) = "" Then
        MsgBox "Please Enter your Membership Level. ", vbInformation, "Incomplete Information"
        txtNationality.SetFocus
        Exit Function
    End If
    
    If Trim(txtFirstName.Text) = "" Then
        MsgBox "Please Enter your Name. ", vbInformation, "Incomplete Information"
        txtFirstName.SetFocus
        Exit Function
    End If
    
    If Trim(txtMiddleName.Text) = "" Then
        MsgBox "Please Enter your Middle Name. ", vbInformation, "Incomplete Information"
        txtMiddleName.SetFocus
        Exit Function
    End If
    
    If Trim(txtFamilyName.Text) = "" Then
        MsgBox "Please Enter your Family Name. ", vbInformation, "Incomplete Information"
        txtFamilyName.SetFocus
        Exit Function
    End If
    
    If IsDate(Trim(txtBirthday.Text)) Then
       txtBirthday.Text = Format(txtBirthday.Text, "mmm. d, yyyy")
    Else
        MsgBox "Invalid birth date. ", vbInformation, "Incomplete Information"
        txtBirthday.SetFocus
        Exit Function
    End If
    
    If UCase(Trim(txtSex.Text)) <> "MALE" And UCase(Trim(txtSex.Text)) <> "FEMALE" _
       And UCase(Trim(txtSex.Text)) <> "M" And UCase(Trim(txtSex.Text)) <> "F" Then
        'MsgBox txtSex.Text
        MsgBox "Undetermined Sex!!! ", vbCritical, "Error"
        '''ValidateUserFields = 0
        txtSex.SetFocus
        Exit Function
    Else
           If UCase(Left(txtSex.Text, 1)) = "M" Then txtSex.Text = "Male"
           If UCase(Left(txtSex.Text, 1)) = "F" Then txtSex.Text = "Female"
    End If
    
    If Trim(txtCivilStatus.Text) = "" Then
        MsgBox "Please Enter your Civil Status. ", vbInformation, "Incomplete Information"
        txtCivilStatus.SetFocus
        Exit Function
    End If
    
    If txtHomeAddress.Text = "" Then
        MsgBox "Please Enter your Home Address. ", vbInformation, "Incomplete Information"
        txtHomeAddress.SetFocus
        Exit Function
    End If
    
  ValidateMembersInfo = True
End Function

Private Sub lstMembers_Click()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
lstTEXT = lstMembers.Text  '
Call vr_engine.GetMemberInfo(lstMembers, txtDateEntered, txtIDnumber, txtNationality, _
                  txtFirstName, txtMiddleName, txtFamilyName, _
                  txtBirthday, txtSex, txtCivilStatus, txtOccupation, _
                  txtContactNumber, txtHomeAddress, txtOfficeSchoolAddress, txtComments)
                  
txtAge.Text = vr_engine.GetAge(Format(txtBirthday.Text, "mmm. dd, yyyy"))
cmdEdit.Enabled = True
End Sub


