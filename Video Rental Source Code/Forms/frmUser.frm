VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   10710
   Begin VB.CommandButton cmdSetPermission 
      Caption         =   "&Permissions"
      Height          =   375
      Left            =   9120
      TabIndex        =   36
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtSex 
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "Sex (Male or Female)"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtAccessLevel 
      Height          =   285
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Access Level - (Integer value (1 - 6)"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddUser 
      Caption         =   "&Add User"
      Height          =   375
      Left            =   3840
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1145
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtHomeAddress 
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   12
      ToolTipText     =   "Home Address"
      Top             =   3240
      Width           =   6735
   End
   Begin VB.TextBox txtComments 
      Height          =   1605
      Left            =   3720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "Comments"
      Top             =   3840
      Width           =   6735
   End
   Begin VB.TextBox txtContactNumber 
      Height          =   285
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   11
      ToolTipText     =   "Contact Number"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtBirthday 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "d. MMMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   10
      ToolTipText     =   "Birth Date Example:  December 25, 1976"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtFamilyName 
      Height          =   285
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Family Name"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtMiddleName 
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Middle Name"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "First Name"
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6000
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   4
      ToolTipText     =   "Re-Enter password for confirmation"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3720
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Password"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "User Name"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtDateEntered 
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Date"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "User Identification"
      Top             =   840
      Width           =   2175
   End
   Begin VB.ListBox lstUsers 
      Height          =   4740
      ItemData        =   "frmUser.frx":030A
      Left            =   360
      List            =   "frmUser.frx":0311
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "List of Users"
      Top             =   650
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   " Detailed Information "
      Height          =   5415
      Left            =   3480
      TabIndex        =   15
      Top             =   240
      Width           =   7150
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access Level"
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
         Left            =   4800
         TabIndex        =   29
         Tag             =   "&Password:"
         Top             =   960
         Width           =   1005
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
         Index           =   12
         Left            =   240
         TabIndex        =   28
         Tag             =   "&Password:"
         Top             =   3360
         Width           =   750
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
         TabIndex        =   27
         Tag             =   "&Password:"
         Top             =   2760
         Width           =   1080
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
         Index           =   10
         Left            =   4800
         TabIndex        =   26
         Tag             =   "&Password:"
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birthday (Ex. Dec. 25, 1976)"
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
         Left            =   2520
         TabIndex        =   25
         Tag             =   "&Password:"
         Top             =   2160
         Width           =   2040
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
         Left            =   240
         TabIndex        =   24
         Tag             =   "&Password:"
         Top             =   2160
         Width           =   285
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
         Index           =   7
         Left            =   4800
         TabIndex        =   23
         Tag             =   "&Password:"
         Top             =   1560
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
         Index           =   6
         Left            =   2520
         TabIndex        =   22
         Tag             =   "&Password:"
         Top             =   1560
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
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Tag             =   "&Password:"
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
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
         Left            =   2520
         TabIndex        =   20
         Tag             =   "&Password:"
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   240
         TabIndex        =   19
         Tag             =   "&Password:"
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Left            =   4800
         TabIndex        =   18
         Tag             =   "&Password:"
         Top             =   360
         Width           =   795
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
         TabIndex        =   17
         Tag             =   "&Password:"
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
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
         Left            =   2520
         TabIndex        =   16
         Tag             =   "&Password:"
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Users "
      Height          =   5415
      Left            =   120
      TabIndex        =   34
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim frmGVAR_lstUsersText As String

Private Sub cmdAddUser_Click()

''-------------------------------------------
If cmdAddUser.Caption = "&Add User" Then
   frmUser.Caption = "Users (Add User Mode)"
   Call cmdClear_Click
   txtDateEntered.Text = Format(Now, "mmm. d, yyyy")
   txtUserID.Text = "(AUTO-ASSIGN)"
'' Disable other controls
   cmdEdit.Enabled = False
   lstUsers.Enabled = False
'' Disable controls
''Enable update
  cmdUpdate.Enabled = True
  cmdClear.Enabled = True
''
   cmdAddUser.Caption = "&Cancel Add..."
   Call UnlockTextboxes
   txtUserName.SetFocus
   Exit Sub
End If
''-------------------------------------------
''-------------------------------------------
If cmdAddUser.Caption = "&Cancel Add..." Then
   frmUser.Caption = "Users"
'' Enable other controls
   cmdEdit.Enabled = True
   lstUsers.Enabled = True
   cmdAddUser.Caption = "&Add User"
   lstUsers.SetFocus
   lstUsers.Text = frmGVAR_lstUsersText
'' Enable controls
'' Disable on error
   If Trim(lstUsers.Text) = "" Then
      cmdEdit.Enabled = False
      cmdUpdate.Enabled = False
      cmdClear.Enabled = False
    Else
      cmdUpdate.Enabled = False
      cmdClear.Enabled = False
      Call LockTextboxes
   End If
   Exit Sub
End If
''-------------------------------------------



End Sub

Private Sub cmdClear_Click()
'' Start Clear Txt/Cbo boxes


txtUserName.Text = ""
txtPassword.Text = ""
txtConfirmPassword.Text = ""
txtAccessLevel.Text = ""
txtFirstName.Text = ""
txtMiddleName.Text = ""
txtFamilyName.Text = ""
txtSex.Text = ""
txtBirthday.Text = ""
txtContactNumber.Text = ""
txtHomeAddress.Text = ""
txtComments.Text = ""

txtUserName.SetFocus
End Sub

Private Sub cmdEdit_Click()
 
 If cmdEdit.Caption = "&Edit" Then
    cmdEdit.Caption = "&Cancel Edit"
    lstUsers.Enabled = False
    Call UnlockTextboxes
    cmdAddUser.Enabled = False
    cmdUpdate.Enabled = True
    cmdClear.Enabled = True
    txtUserName.SetFocus
    frmUser.Caption = "Users (Edit Mode)"
    Exit Sub
 End If
 
 If cmdEdit.Caption = "&Cancel Edit" Then
    Call cmdClear_Click
    cmdEdit.Caption = "&Edit"
    Call LockTextboxes
    cmdAddUser.Enabled = True
    lstUsers.Enabled = True
    cmdUpdate.Enabled = False
    cmdClear.Enabled = False
    lstUsers.SetFocus
    frmUser.Caption = "Users"
 lstUsers.Text = frmGVAR_lstUsersText

    Exit Sub
 End If
 
End Sub

Private Sub cmdRemove_Click()
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    
    If IsNumeric(txtUserID.Text) = True And cmdEdit.Caption = "&Edit" _
       And cmdAddUser.Caption = "&Add User" Then
       If MsgBox("Are you sure you want to delete this user? ", vbYesNo) = vbNo Then Exit Sub
       Call vr_engine.RemoveUser(Int(txtUserID.Text))
       Call vr_engine.LoadUsers(lstUsers) ''Refresh listbox
       txtDateEntered.Text = ""
       txtUserID.Text = ""
       Call cmdClear_Click
       cmdEdit.Enabled = False
       cmdUpdate.Enabled = False
       cmdClear.Enabled = False
       If lstUsers.Enabled = True Then lstUsers.SetFocus
    Else
      '' Put Message Here
       If lstUsers.Enabled = True Then
          lstUsers.SetFocus
       Else
         txtUserName.SetFocus
       End If
    End If

End Sub

Private Sub cmdSetPermission_Click()
If gVarAccessLevel < 6 Then
    
    If lstUsers.Enabled = True Then
        lstUsers.SetFocus
    Else
        txtDateEntered.SetFocus
    End If
    Exit Sub
End If
frmPermission.Show 1
If lstUsers.Enabled = True Then
lstUsers.SetFocus
Else
txtDateEntered.SetFocus
End If

End Sub

Private Sub cmdUpdate_Click()

Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE

If ValidateUserFields = 0 Then Exit Sub  '' Chk fields

''-------------------------------------------
If cmdAddUser.Caption = "&Cancel Add..." Then
    txtDateEntered.Text = Format(Now, "mmm. d, yyyy")
    txtUserID.Text = "(Auto-assign)"
    If ValidateUserFields = 0 Then Exit Sub  '' Chk fields
    If vr_engine.AddUserToDB(txtDateEntered, txtUserName, txtPassword, _
                    txtAccessLevel, txtFirstName, _
                    txtMiddleName, txtFamilyName, _
                    txtBirthday, txtSex, _
                    txtHomeAddress, txtContactNumber, _
                    txtComments) Then
    '' Do Nothing
    Else
            Exit Sub
    End If

    '' Enable other controls
        cmdEdit.Enabled = True
        lstUsers.Enabled = True
        Call vr_engine.LoadUsers(lstUsers) ''Refresh listbox
        frmGVAR_lstUsersText = Trim(txtFirstName.Text) & " " & Trim(Left(txtMiddleName.Text, 1)) & ". " & Trim(txtFamilyName.Text) & " (" & Trim(txtUserName.Text) & ")"
        lstUsers.Text = frmGVAR_lstUsersText
        lstUsers.SetFocus
        cmdAddUser.Caption = "&Add User"
        
    '' Enable controls
    '' Disable cmdUpdate/cmdClear
       cmdUpdate.Enabled = False
       cmdClear.Enabled = False
       Exit Sub
Else
'' ========= Start Edit Update ============
'' ========= Start Edit Update ============
If vr_engine.UpdateEditedUsersDB(txtUserName, txtPassword, _
                txtAccessLevel, txtFirstName, _
                txtMiddleName, txtFamilyName, _
                txtBirthday, txtSex, _
                txtHomeAddress, txtContactNumber, _
                txtComments, lstUsers) = True Then
'' Do Nothing
Else
  Exit Sub
End If
                
 Call vr_engine.LoadUsers(lstUsers) ''Refresh listbox
 frmGVAR_lstUsersText = Trim(txtFirstName.Text) & " " & Trim(Left(txtMiddleName.Text, 1)) & ". " & Trim(txtFamilyName.Text) & " (" & Trim(txtUserName.Text) & ")"
    '''MsgBox frmGVAR_lstUsersText
    lstUsers.Enabled = True
    cmdUpdate.Enabled = False
    lstUsers.SetFocus
    MsgBox "Record has been successfully updated. ", vbOKOnly, "Success"
    Call LockTextboxes
    cmdEdit.Caption = "&Edit"
    frmUser.Caption = "Users"
    cmdUpdate.Enabled = False
    cmdClear.Enabled = False
    cmdAddUser.Enabled = True
    txtDateEntered.Locked = True
    lstUsers.Text = frmGVAR_lstUsersText
'' ========= End Edit Update ============
'' ========= End Edit Update ============



End If
''-------------------------------------------
End Sub

Private Sub Form_Load()
  lstUsers.Clear '' Clears listbox
  Dim vr_engine As VRENTAL_ENGINE
  Set vr_engine = New VRENTAL_ENGINE
  Call vr_engine.LoadUsers(lstUsers)
      
End Sub



Private Sub lstUsers_Click()
Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    
Call vr_engine.getUserInfo(txtUserID, txtDateEntered, _
                txtUserName, txtPassword, _
                txtAccessLevel, txtFirstName, _
                txtMiddleName, txtFamilyName, _
                txtBirthday, txtSex, _
                txtHomeAddress, txtContactNumber, _
                txtComments, txtConfirmPassword, lstUsers)

cmdEdit.Enabled = True

frmGVAR_lstUsersText = lstUsers.Text


End Sub

Sub UnlockTextboxes()
    'txtDateEntered.Locked = False
    txtUserName.Locked = False
    txtPassword.Locked = False
    txtConfirmPassword.Locked = False
    txtAccessLevel.Locked = False
    txtFirstName.Locked = False
    txtMiddleName.Locked = False
    txtFamilyName.Locked = False
    txtSex.Locked = False
    txtBirthday.Locked = False
    txtContactNumber.Locked = False
    txtHomeAddress.Locked = False
    txtComments.Locked = False

End Sub

Sub LockTextboxes()

    txtUserName.Locked = True
    txtPassword.Locked = True
    txtConfirmPassword.Locked = True
    txtAccessLevel.Locked = True
    txtFirstName.Locked = True
    txtMiddleName.Locked = True
    txtFamilyName.Locked = True
    txtSex.Locked = True
    txtBirthday.Locked = True
    txtContactNumber.Locked = True
    txtHomeAddress.Locked = True
    txtComments.Locked = True

End Sub

Function ValidateUserFields()
    If IsDate(txtDateEntered.Text) Then
        txtDateEntered.Text = Format(txtDateEntered.Text, "mmm. d, yyyy")
    Else
        MsgBox "Invalid Date. ", vbCritical, "Error"
        ValidateUserFields = 0
        Exit Function
    End If
    
    If Trim(txtUserName.Text = "") Then
        MsgBox "Invalid User Name. ", vbCritical, "Error"
        ValidateUserFields = 0
        Exit Function
    End If
    
    If Trim(txtPassword.Text) <> Trim(txtConfirmPassword.Text) Then
        MsgBox "Mismatch between Password and Confirm Password. ", vbCritical, "Error"
        ValidateUserFields = 0
        Exit Function
    End If
    
    If txtPassword.Text = "" Then
        MsgBox "You must have a password. ", vbCritical, "Error"
        ValidateUserFields = 0
        Exit Function
    End If
    
    If IsNumeric(Trim(txtAccessLevel.Text)) Then
        txtAccessLevel.Text = Int(Trim(txtAccessLevel.Text))
    Else
        MsgBox "Invalid Access Level. ", vbCritical, "Error"
        ValidateUserFields = 0
        Exit Function
    End If

    
    If Trim(txtFirstName.Text) = "" Then
        MsgBox "First Name field is empty.", vbCritical, "Error"
        ValidateUserFields = 0
        Exit Function
    End If
    
    If Trim(txtMiddleName.Text) = "" Then
        MsgBox "Middle Name field is empty. ", vbCritical, "Error"
        ValidateUserFields = 0
        Exit Function
    End If

    If Trim(txtFamilyName.Text) = "" Then
        MsgBox "Family Name field is empty. ", vbCritical, "Error"
        ValidateUserFields = 0
        Exit Function
    End If
    
    If Trim(txtMiddleName.Text) = "" Then
        MsgBox "Middle Name field is empty. ", vbCritical, "Error"
        ValidateUserFields = 0
        Exit Function
    End If
    
    If IsDate(txtBirthday.Text) Then
        txtBirthday.Text = Format(txtBirthday.Text, "mmm. d, yyyy")
    Else
        MsgBox "Invalid Birth Date. ", vbCritical, "Error"
        ValidateUserFields = 0
        Exit Function
    End If


    If UCase(Trim(txtSex.Text)) <> "MALE" And UCase(Trim(txtSex.Text)) <> "FEMALE" _
       And UCase(Trim(txtSex.Text)) <> "M" And UCase(Trim(txtSex.Text)) <> "F" Then
        'MsgBox txtSex.Text
        MsgBox "Undetermined Sex!!! ", vbCritical, "Error"
        ValidateUserFields = 0
        Exit Function
    Else
           If UCase(Left(txtSex.Text, 1)) = "M" Then txtSex.Text = "Male"
           If UCase(Left(txtSex.Text, 1)) = "F" Then txtSex.Text = "Female"
    End If
    
    If Trim(txtHomeAddress.Text = "") Then
        MsgBox "Enter your address. ", vbCritical, "Error"
        ValidateUserFields = 0
        Exit Function
    End If
    
     
        ValidateUserFields = 1
        
End Function


Private Sub txtAccessLevel_LostFocus()
If Val(txtAccessLevel.Text) > 6 Then
    MsgBox "'Access Level' should not be greater than 6. ", vbInformation, "Invalid Entry"
    txtAccessLevel.Text = ""
    txtAccessLevel.SetFocus
    Exit Sub
End If
If Val(txtAccessLevel.Text) > gVarAccessLevel Then
   MsgBox "'Access Level' should not be greater than your own access level. ", vbInformation, "Invalid Entry"
    txtAccessLevel.Text = ""
    txtAccessLevel.SetFocus
    Exit Sub
End If
End Sub
