VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3870
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton Button3 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdCPOK 
         Caption         =   "OK"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtOldPWD 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   240
         Width           =   1845
      End
      Begin VB.TextBox txtNewPWD 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   600
         Width           =   1845
      End
      Begin VB.TextBox txtConfirmNewPWD 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
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
         Left            =   120
         TabIndex        =   13
         Tag             =   "&Password:"
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
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
         Left            =   120
         TabIndex        =   12
         Tag             =   "&Password:"
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re type Password"
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
         Left            =   120
         TabIndex        =   11
         Tag             =   "&Password:"
         Top             =   960
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton Button2 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "OK"
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   1080
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3240
         Picture         =   "frmLogin.frx":1272
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   5
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1080
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   240
         Width           =   1845
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1200
         MouseIcon       =   "frmLogin.frx":13B8
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Tag             =   "&User Name:"
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Tag             =   "&Password:"
         Top             =   600
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub Button2_Click()
End
End Sub

Private Sub Button3_Click()
frmLogin.Caption = "Login"
Frame1.Visible = -1
Frame2.Visible = 0#
txtUserName.Text = ""
txtPassword.Text = ""
txtUserName.SetFocus
End Sub

Private Sub cmdCPOK_Click()
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    
 If txtOldPWD.Text = txtPassword.Text And txtNewPWD.Text = txtConfirmNewPWD.Text Then
        '' ----- Start Update User Password Here
         Call vr_engine.ChangePassword(txtNewPWD.Text, txtUserName.Text)
        '' ----- End Update User Password Here
        MsgBox "Your password has been changed. ", vbOKOnly, "Success"
        Button3_Click
 Else
        MsgBox "Letters in passwords must be typed in correct case. ", vbCritical, "Change Password Error"
        txtNewPWD.Text = ""
        txtConfirmNewPWD.Text = ""
        txtOldPWD.Text = ""
        txtOldPWD.SetFocus
 End If
 
End Sub

Private Sub cmdLogin_Click()
    'Dim LogOnResult As String
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    LogOnResult = vr_engine.LogOnValidate(Trim(txtUserName.Text), Trim(txtPassword.Text))
    If Trim(LogOnResult) <> "" Then
      LogOnResult = LogOnResult & "LogInDateDate : " & Format(Now, "dddd, mmm d yyyy") & vbCrLf & "LogInTime : " & Format(Now, "hh:mm:ss AMPM") & vbCrLf
      AssignUserInfoToGlobalVar (LogOnResult)
     '' Append user to logfile
       If vr_engine.ReportFileStatus(App.Path & "\UserLog.txt") = -1 Then
          Open App.Path & "\UserLog.txt" For Append As #1
          Print #1, gVarFirstName & " " & Left(gVarMiddleName, 1) & ". " & gVarFamilyName & " : " & Format(Now, "dddd, mmm d yyyy") & " : " & Format(Now, "hh:mm:ss AMPM") & vbCrLf     '' Append to logfile
          Close #1
       Else
          Open App.Path & "\UserLog.txt" For Output As #1
          Print #1, txtUserName.Text & " : " & Format(Now, "dddd, mmm d yyyy") & " : " & Format(Now, "hh:mm:ss AMPM") & vbCrLf     '' Append to logfile
          Close #1
       End If
     '' End Appen to logfile format$
       MDIForm1.Show
       Unload Me
    Else
       MsgBox "Invalid User Name or Password.   ", vbInformation, "Access denied"
       txtUserName.Text = ""
       txtPassword.Text = ""
       txtUserName.SetFocus
    End If

End Sub

Private Sub Label1_Click()
    Dim LogOnResult As String
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    LogOnResult = vr_engine.LogOnValidate(Trim(txtUserName.Text), Trim(txtPassword.Text))
    If Trim(LogOnResult) <> "" Then
        '''4Erase MsgBox LogOnResult, vbOKOnly, "test"
    Else
       MsgBox "Invalid User Name or Password.   ", vbInformation, "Access denied"
       txtUserName.Text = ""
       txtPassword.Text = ""
       txtUserName.SetFocus
       Exit Sub
    End If
    
Frame2.Visible = -1
Frame1.Visible = 0
frmLogin.Caption = "Change Password"
txtOldPWD.Text = ""
txtNewPWD.Text = ""
txtConfirmNewPWD.Text = ""
End Sub

Private Sub txtConfirmNewPWD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdCPOK_Click
End Sub

Private Sub txtNewPWD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdCPOK_Click
End Sub

Private Sub txtOldPWD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdCPOK_Click
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdLogin_Click
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdLogin_Click
End Sub
