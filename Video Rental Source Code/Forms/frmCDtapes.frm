VERSION 5.00
Begin VB.Form frmCDtapes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Available Movies"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "frmCDtapes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   10620
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1320
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddMovie 
      Caption         =   "&Add Entry"
      Height          =   375
      Left            =   5430
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6645
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7860
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9075
      TabIndex        =   35
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtComments 
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   18
      ToolTipText     =   "Additional Remarks/Comments"
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox txtLastDateReturnedAddInfo 
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   16
      ToolTipText     =   "Last Date Returned Additional Information (Comments)"
      Top             =   4920
      Width           =   2775
   End
   Begin VB.TextBox txtLastDateReturned 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   15
      ToolTipText     =   "Last Date Returned  (Example: Dec. 25, 2001)"
      Top             =   4920
      Width           =   2775
   End
   Begin VB.TextBox txtCondition 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   17
      ToolTipText     =   "Condition (G - good, D - damaged, DF - defective,  L - lost)"
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox txtLastDateBorrowedAddInfo 
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "Last Date Borrowed Additional Information (Comments)"
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox txtLastDateBorrowed 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "Last Date Borrowed (Example: Dec. 25, 2001)"
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox txtAvailable 
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   10
      ToolTipText     =   "Available"
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox txtRentalAmount 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "Rental Amount"
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox txtGenre 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Genre"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox txtRunTime 
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Run Time"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox txtYearReleased 
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Year Released"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtActor 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Name of Actor"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Title"
      Top             =   720
      Width           =   5895
   End
   Begin VB.TextBox txtItemCode 
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "ID Number"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtDateEntered 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Date Entered"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   " Detailed Information "
      Height          =   5895
      Left            =   4080
      TabIndex        =   19
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtRentalPeriod 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox txtOverdueChargePerDay 
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3600
         Width           =   2775
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rental Period (in days)"
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
         Index           =   16
         Left            =   240
         TabIndex        =   41
         Tag             =   "&Password:"
         Top             =   3360
         Width           =   1635
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Overdue Charge Per Day"
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
         Index           =   15
         Left            =   3360
         TabIndex        =   40
         Tag             =   "&Password:"
         Top             =   3360
         Width           =   1815
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
         Left            =   3360
         TabIndex        =   34
         Tag             =   "&Password:"
         Top             =   5160
         Width           =   750
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condition  (Good, Defective, Lost)"
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
         Left            =   240
         TabIndex        =   33
         Tag             =   "&Password:"
         Top             =   5160
         Width           =   2445
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Date Return Add. Info."
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
         Left            =   3360
         TabIndex        =   32
         Tag             =   "&Password:"
         Top             =   4560
         Width           =   1965
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Date Returned (mmm. dd, yyyy)"
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
         TabIndex        =   31
         Tag             =   "&Password:"
         Top             =   4560
         Width           =   2640
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Date Borrowed Add. Info."
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
         Left            =   3360
         TabIndex        =   30
         Tag             =   "&Password:"
         Top             =   3960
         Width           =   2220
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Date Borrowed (mmm. dd, yyyy)"
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
         TabIndex        =   29
         Tag             =   "&Password:"
         Top             =   3960
         Width           =   2715
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available"
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
         Left            =   3360
         TabIndex        =   28
         Tag             =   "&Password:"
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rental Amount"
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
         Left            =   240
         TabIndex        =   27
         Tag             =   "&Password:"
         Top             =   2760
         Width           =   1050
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Run Time"
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
         Left            =   3360
         TabIndex        =   26
         Tag             =   "&Password:"
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Genre"
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
         TabIndex        =   25
         Tag             =   "&Password:"
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year Released"
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
         Left            =   3360
         TabIndex        =   24
         Tag             =   "&Password:"
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actor"
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
         TabIndex        =   23
         Tag             =   "&Password:"
         Top             =   1560
         Width           =   405
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
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
         Left            =   3360
         TabIndex        =   22
         Tag             =   "&Password:"
         Top             =   960
         Width           =   705
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
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Tag             =   "&Password:"
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
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
         TabIndex        =   20
         Tag             =   "&Password:"
         Top             =   360
         Width           =   285
      End
   End
   Begin VB.ListBox lstTitles 
      Height          =   5325
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "List of Titles"
      Top             =   480
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Titles "
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmCDtapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim lstTEXT, tmpItemCodeB4Edit As String

Private Sub cmdAddMovie_Click()
      If cmdAddMovie.Caption = "&Add Entry" Then
         cmdAddMovie.Caption = "&Cancel Add..."
         Call cmdClear_Click
         txtDateEntered = Format(Now, "mmm. dd, yyyy")
         txtItemCode.Text = ""
         Call UnlockTextboxes
         lstTitles.Enabled = False
         cmdUpdate.Enabled = True
         cmdClear.Enabled = True
         cmdEdit.Enabled = False
         txtTitle.SetFocus
         Exit Sub
      End If
      
      If cmdAddMovie.Caption = "&Cancel Add..." Then
         cmdAddMovie.Caption = "&Add Entry"
         Call cmdClear_Click
         txtItemCode.Text = ""
         txtDateEntered.Text = ""
         txtItemCode.Text = ""
         Call LockTextboxes
         lstTitles.Enabled = True
         cmdUpdate.Enabled = False
         cmdClear.Enabled = False
         lstTitles.Text = lstTEXT
         If Trim(lstTitles.Text) = "" Then cmdEdit.Enabled = False
         lstTitles.SetFocus
         Exit Sub
      End If
End Sub

Private Sub cmdClear_Click()
    txtTitle.Text = ""
    txtActor.Text = ""
    txtYearReleased.Text = ""
    txtGenre.Text = ""
    txtRunTime.Text = ""
    txtRentalAmount.Text = ""
    txtAvailable.Text = ""
    txtOverdueChargePerDay.Text = ""
    txtRentalPeriod.Text = ""
    txtLastDateBorrowed.Text = ""
    txtLastDateBorrowedAddInfo.Text = ""
    txtLastDateReturned.Text = ""
    txtLastDateReturnedAddInfo.Text = ""
    txtCondition.Text = ""
    If cmdAddMovie.Caption = "&Cancel Add..." Then txtItemCode.Text = ""
    If cmdEdit.Caption = "Ca&ncel Edit" Then txtItemCode.Text = ""
    txtComments.Text = ""
    txtTitle.SetFocus
    
End Sub

Private Sub cmdEdit_Click()
  If cmdEdit.Caption = "&Edit" Then
      cmdEdit.Caption = "Ca&ncel Edit"
      tmpItemCodeB4Edit = Trim(txtItemCode.Text)
      lstTitles.Enabled = False
      cmdAddMovie.Enabled = False
      cmdUpdate.Enabled = True
      cmdClear.Enabled = True
      Call UnlockTextboxes
      txtTitle.SetFocus
      Exit Sub
  End If
  
  If cmdEdit.Caption = "Ca&ncel Edit" Then
      cmdEdit.Caption = "&Edit"
      lstTitles.Enabled = True
      cmdAddMovie.Enabled = True
      cmdUpdate.Enabled = False
      cmdClear.Enabled = False
      Call LockTextboxes
      lstTitles.Text = lstTEXT
      lstTitles.SetFocus
  End If
      
End Sub

Private Sub cmdRemove_Click()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
If lstTitles.Enabled = True And cmdUpdate.Enabled = False And cmdClear.Enabled = False And Trim(lstTitles.Text) <> "" Then
    If MsgBox("Are you sure you want to delete this user? ", vbYesNo) = vbNo Then Exit Sub
    Call vr_engine.Remove_CD_tape_Items(txtItemCode.Text)
    Call cmdClear_Click
    txtDateEntered.Text = ""
    txtItemCode.Text = ""
    cmdEdit.Enabled = False
    Call vr_engine.LoadMoviesList(lstTitles) 'Refresh items list
    lstTitles.SetFocus
Else
    If lstTitles.Enabled = False Then
       txtTitle.SetFocus
    Else
       lstTitles.SetFocus
    End If
End If

End Sub

Private Sub cmdUpdate_Click()
  Dim vr_engine As VRENTAL_ENGINE
  Set vr_engine = New VRENTAL_ENGINE
 
       If cmdAddMovie.Caption = "&Cancel Add..." Then
         'Start - Validate input
         If ValidateTextBoxesEntries = True Then
            '' Do nothing
         Else
            Exit Sub
         End If
         'End - Validate input
         'Update DB
         If vr_engine.Add_CD_TAPES_MovieToDB(txtTitle, txtDateEntered, txtItemCode, _
                    txtActor, txtYearReleased, txtGenre, txtRunTime, txtRentalAmount, _
                    txtAvailable, txtRentalPeriod, txtOverdueChargePerDay, txtLastDateBorrowed, txtLastDateBorrowedAddInfo, txtLastDateReturned, _
                    txtLastDateReturnedAddInfo, txtCondition, txtComments) = True Then
                    Call vr_engine.LoadMoviesList(lstTitles)
                    lstTitles.Text = Trim(txtTitle.Text) & " [" & Trim(txtItemCode.Text) & "]"
                    lstTitles.Enabled = True
                    lstTitles.SetFocus
                    cmdUpdate.Enabled = False
                    cmdClear.Enabled = False
                    cmdAddMovie.Caption = "&Add Entry"
                    Call LockTextboxes
         Else
             '' Do Nothing
         End If
         'End update
          Exit Sub
       End If
'-----------------------------------------------------------------------------------------------------------------
       If cmdEdit.Caption = "Ca&ncel Edit" Then
         
         'Start - Validate input
         If ValidateTextBoxesEntries = True Then
            '' Do nothing
         Else
            Exit Sub
         End If
         'End - Validate input
         
            If vr_engine.UpdateEditedCDtapesInfo(tmpItemCodeB4Edit, txtTitle, txtDateEntered, txtItemCode, _
                        txtActor, txtYearReleased, txtGenre, txtRunTime, txtRentalAmount, _
                        txtAvailable, txtRentalPeriod, txtOverdueChargePerDay, txtLastDateBorrowed, txtLastDateBorrowedAddInfo, txtLastDateReturned, _
                        txtLastDateReturnedAddInfo, txtCondition, txtComments) = True Then
                    
                        Call vr_engine.LoadMoviesList(lstTitles)
                        lstTitles.Text = Trim(txtTitle.Text) & " [" & Trim(txtItemCode.Text) & "]"
                        lstTitles.Enabled = True
                         cmdAddMovie.Enabled = True
                        lstTitles.SetFocus
                        cmdUpdate.Enabled = False
                        cmdClear.Enabled = False
                        cmdEdit.Caption = "&Edit"
                        Call LockTextboxes
            Else
            End If
       End If
       
       
'-----------------------------------------------------------------------------------------------------------------
End Sub

Private Sub Form_Load()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE

Call vr_engine.LoadMoviesList(lstTitles)
End Sub

Private Sub lstTitles_Click()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE

Call vr_engine.GetCD_TapesInfo(lstTitles, txtTitle, txtDateEntered, txtItemCode, _
                    txtActor, txtYearReleased, txtGenre, txtRunTime, txtRentalAmount, _
                    txtAvailable, txtRentalPeriod, txtOverdueChargePerDay, txtLastDateBorrowed, txtLastDateBorrowedAddInfo, txtLastDateReturned, _
                    txtLastDateReturnedAddInfo, txtCondition, txtComments)
                    
    lstTEXT = lstTitles.Text
    cmdEdit.Enabled = True
    
End Sub

Sub UnlockTextboxes()
            txtTitle.Locked = False
            ''txtDateEntered.Locked = False
            txtItemCode.Locked = False
            txtActor.Locked = False
            txtYearReleased.Locked = False
            txtGenre.Locked = False
            txtRunTime.Locked = False
            txtRentalAmount.Locked = False
            txtAvailable.Locked = False
            txtOverdueChargePerDay.Locked = False
            txtRentalPeriod.Locked = False
            txtLastDateBorrowed.Locked = False
            txtLastDateBorrowedAddInfo.Locked = False
            txtLastDateReturned.Locked = False
            txtLastDateReturnedAddInfo.Locked = False
            txtCondition.Locked = False
            txtComments.Locked = False
End Sub
Sub LockTextboxes()
            txtTitle.Locked = True
            txtDateEntered.Locked = True
            txtItemCode.Locked = True
            txtActor.Locked = True
            txtYearReleased.Locked = True
            txtGenre.Locked = True
            txtRunTime.Locked = True
            txtRentalAmount.Locked = True
            txtAvailable.Locked = True
            txtOverdueChargePerDay.Locked = True
            txtRentalPeriod.Locked = True
            txtLastDateBorrowed.Locked = True
            txtLastDateBorrowedAddInfo.Locked = True
            txtLastDateReturned.Locked = True
            txtLastDateReturnedAddInfo.Locked = True
            txtCondition.Locked = True
            txtComments.Locked = True
End Sub

Function ValidateTextBoxesEntries() As Boolean
      '-----------------------------------------------------------------
          If Trim(txtTitle.Text) = "" Then
             MsgBox "You must enter a title. ", vbInformation, "Cannot update"
             ValidateTextBoxesEntries = False
             txtTitle.SetFocus
             Exit Function
          End If
      '-----------------------------------------------------------------
          If Trim(txtItemCode.Text) = "" Then
             MsgBox "You must enter your Item Code. ", vbInformation, "Cannot update"
             ValidateTextBoxesEntries = False
             txtItemCode.SetFocus
             Exit Function
          End If
           'START - CHECKt Itemcode Format
      Dim str2 As String
      str2 = Trim(txtItemCode.Text)
      If Len(str2) <> 8 Then
         MsgBox "Item Code must be 8 characters long. ", vbInformation, "Invalid Item Code"
         txtItemCode.SetFocus
         Exit Function
      Else
         Dim t As String
         
         If UCase(Mid(str2, 1, 4)) <> UCase("VHS-") _
            And UCase(Mid(str2, 1, 4)) <> UCase("VCD-") _
            And UCase(Mid(str2, 1, 4)) <> UCase("DVD-") Then
               MsgBox "Fisrt four characters of item code " & vbCrLf & "must be either 'VHS-', 'VCD-', or 'DVD-'. ", vbInformation, "Invalid Item Code"
               txtItemCode.SetFocus
            Exit Function
         Else
                If IsNumeric(Mid(str2, 5, 1)) = False Then
                    MsgBox "Last four characters of item code " & vbCrLf & "must be a numeric character. ", vbInformation, "Invalid Item Code"
                    txtItemCode.SetFocus
                    Exit Function
                End If
                If IsNumeric(Mid(str2, 6, 1)) = False Then
                    MsgBox "Last four characters of item code " & vbCrLf & "must be a numeric character. ", vbInformation, "Invalid Item Code"
                    txtItemCode.SetFocus
                    Exit Function
                End If
                If IsNumeric(Mid(str2, 7, 1)) = False Then
                    MsgBox "Last four characters of item code " & vbCrLf & "must be a numeric character. ", vbInformation, "Invalid Item Code"
                    txtItemCode.SetFocus
                    Exit Function
                End If
                If IsNumeric(Mid(str2, 8, 1)) = False Then
                    MsgBox "Last four characters of item code " & vbCrLf & "must be a numeric character. ", vbInformation, "Invalid Item Code"
                    txtItemCode.SetFocus
                Exit Function
            End If
         End If
      End If
     txtItemCode.Text = UCase(txtItemCode.Text)
    'END - CHECK Itemcode format
       '-----------------------------------------------------------------
          If Trim(txtActor.Text) = "" Then
             MsgBox "You must enter name of the actor. ", vbInformation, "Cannot update"
             ValidateTextBoxesEntries = False
             txtActor.SetFocus
             Exit Function
          End If
       '-----------------------------------------------------------------
          If Trim(txtYearReleased.Text) = "" Then
             If MsgBox("""Year Released"" field is blank.  " & vbCrLf & vbCrLf & "Do you want to continue?", vbYesNo, "Update Info") = vbNo Then
                ValidateTextBoxesEntries = False
                txtYearReleased.SetFocus
                Exit Function
             End If
          End If
       '-----------------------------------------------------------------
          If Trim(txtGenre.Text) = "" Then
             If MsgBox("""Genre"" field is blank.  " & vbCrLf & vbCrLf & "Do you want to continue?", vbYesNo, "Update Info") = vbNo Then
                ValidateTextBoxesEntries = False
                txtGenre.SetFocus
                Exit Function
             End If
          End If
      '-----------------------------------------------------------------
          If Trim(txtRunTime.Text) = "" Then
             If MsgBox("""Run Time"" field is blank.  " & vbCrLf & vbCrLf & "Do you want to continue?", vbYesNo, "Update Info") = vbNo Then
                ValidateTextBoxesEntries = False
                txtRunTime.SetFocus
                Exit Function
             End If
          End If
      '-----------------------------------------------------------------
          If Trim(txtRentalAmount) = "" Then
             MsgBox """Rental Amount"" is blank. ", vbInformation, "Cannot update"
             txtRentalAmount.SetFocus
             ValidateTextBoxesEntries = False
             Exit Function
          Else
             If IsNumeric(Trim(txtRentalAmount.Text)) Then
             '' do nothing
             Else
             MsgBox """Rental Amount"" is invalid. ", vbInformation, "Cannot update"
             txtRentalAmount.SetFocus
             ValidateTextBoxesEntries = False
             Exit Function
             End If
          End If
      '-----------------------------------------------------------------
          If Trim(txtAvailable.Text) = "" Then
             MsgBox """Available"" field is blank.     " & vbCrLf & vbCrLf & "Please enter 'Yes' or 'No'    ", vbInformation, "Cannot update"
             txtAvailable.SetFocus
             ValidateTextBoxesEntries = False
             Exit Function
          Else
             If UCase(Trim(txtAvailable.Text)) = "YES" Or UCase(Trim(txtAvailable.Text)) = "NO" Then
                    If Left(UCase(Trim(txtAvailable.Text)), 1) = "Y" Then txtAvailable.Text = "Yes"
                    If Left(UCase(Trim(txtAvailable.Text)), 1) = "N" Then txtAvailable.Text = "No"
             Else
                        MsgBox """Available"" field entry is invalid.     " & vbCrLf & vbCrLf & "Please enter 'Yes' or 'No'    ", vbInformation, "Cannot update"
                        txtAvailable.SetFocus
                        ValidateTextBoxesEntries = False
                        Exit Function
             End If
             
          End If
          
      ' ----------------------------------------------------------------
        If IsNumeric(Trim(txtRentalPeriod.Text)) = True Then
            If Int(Val(Trim(txtRentalPeriod.Text))) <= 0 Then
                MsgBox "Rental Period field is blank or invalid.  ", vbInformation, "Input Error: "
                ValidateTextBoxesEntries = False
                txtRentalPeriod.SetFocus
                Exit Function
            End If
            txtRentalPeriod.Text = Trim(str(Int(Trim(txtRentalPeriod.Text))))
        Else
                MsgBox "Rental Period field is blank or invalid.  ", vbInformation, "Input Error: "
                ValidateTextBoxesEntries = False
                txtRentalPeriod.SetFocus
                Exit Function
             
        End If
      ' ----------------------------------------------------------------
        'txtOverdueChargePerDay
        If IsNumeric(Trim(txtOverdueChargePerDay.Text)) = True Then
            If Int(Val(Trim(txtOverdueChargePerDay.Text))) < 0 Then
                MsgBox "Overdue Charge Per Day field is blank or invalid.  ", vbInformation, "Input Error: "
                ValidateTextBoxesEntries = False
                txtOverdueChargePerDay.SetFocus
                Exit Function
            End If
          txtOverdueChargePerDay.Text = Trim(txtOverdueChargePerDay.Text)
        Else
                MsgBox "Overdue Charge Per Day field is blank or invalid.  ", vbInformation, "Input Error: "
                ValidateTextBoxesEntries = False
                txtOverdueChargePerDay.SetFocus
                Exit Function
        End If
      '-----------------------------------------------------------------
          If IsDate(Trim(txtLastDateBorrowed.Text)) Then
             txtLastDateBorrowed.Text = Format(Trim(txtLastDateBorrowed.Text), "mmm. dd, yyyy")
          Else
             If Trim(txtLastDateBorrowed.Text) <> "" Then
                MsgBox """Last Date Borrowed"" entry is invalid. ", vbInformation, "Cannot update"
                txtLastDateBorrowed.SetFocus
                ValidateTextBoxesEntries = False
                Exit Function
             End If
          End If
      '-----------------------------------------------------------------
            If IsDate(Trim(txtLastDateReturned.Text)) Then
                 txtLastDateReturned.Text = Format(Trim(txtLastDateReturned.Text), "mmm. dd, yyyy")
            Else
                 If Trim(txtLastDateReturned.Text) <> "" Then
                    MsgBox """Last Date Returned"" entry is invalid. ", vbInformation, "Cannot update"
                    txtLastDateReturned.SetFocus
                    ValidateTextBoxesEntries = False
                    Exit Function
                End If
            End If
      
      '-----------------------------------------------------------------
            If Trim(txtCondition.Text) = "" Then
               MsgBox "Please enter the condition of the item. " & vbCrLf & vbCrLf & "Type: Good, Defective, Damaged, or Lost ", vbInformation, "Cannot update"
               txtCondition.SetFocus
               ValidateTextBoxesEntries = False
               Exit Function
            Else
               If Trim(UCase(txtCondition.Text)) = "GOOD" Or Trim(UCase(txtCondition.Text)) = "DEFECTIVE" Or Trim(UCase(txtCondition.Text)) = "DAMAGED" Or _
                  Trim(UCase(txtCondition.Text)) = "LOST" Or Trim(UCase(txtCondition.Text)) = "D" Or Trim(UCase(txtCondition.Text)) = "DF" Or Trim(UCase(txtCondition.Text)) = "G" Or _
                  Trim(UCase(txtCondition.Text)) = "L" Then
                   If Trim(UCase(txtCondition.Text)) = "GOOD" Then txtCondition.Text = "Good"
                   If Trim(UCase(txtCondition.Text)) = "G" Then txtCondition.Text = "Good"
                   If Trim(UCase(txtCondition.Text)) = "DEFECTIVE" Then txtCondition.Text = "Defective"
                   If Trim(UCase(txtCondition.Text)) = "DF" Then txtCondition.Text = "Defective"
                   If Trim(UCase(txtCondition.Text)) = "DAMAGED" Then txtCondition.Text = "Damaged"
                   If Trim(UCase(txtCondition.Text)) = "D" Then txtCondition.Text = "Damaged"
                   If Trim(UCase(txtCondition.Text)) = "LOST" Then txtCondition.Text = "Lost"
                   If Trim(UCase(txtCondition.Text)) = "L" Then txtCondition.Text = "Lost"
                  
               Else
                    MsgBox "Invalid condition for the item. " & vbCrLf & vbCrLf & "Type: Good, Defective, Damaged, or Lost ", vbInformation, "Cannot update"
                    txtCondition.SetFocus
                    ValidateTextBoxesEntries = False
                    Exit Function
               End If
            End If
      '-----------------------------------------------------------------
      
      
          ValidateTextBoxesEntries = True
End Function


Private Sub txtAvailable_LostFocus()
If UCase(txtAvailable.Text) = "Y" Then txtAvailable.Text = "Yes"
If UCase(txtAvailable.Text) = "N" Then txtAvailable.Text = "No"
End Sub

Private Sub txtItemCode_LostFocus()
 Dim str As String
 Dim loop1 As Integer
 If Trim(txtItemCode.Text) <> "" Then
    For loop1 = 1 To Len(txtItemCode.Text)
       If Mid(Trim(txtItemCode.Text), loop1, 1) = " " Then
          MsgBox "Item Code should not contain a 'space' character. ", vbInformation, "Invalid Entry"
          txtItemCode.Text = ""
          txtItemCode.SetFocus
       End If
    Next
   
 End If
End Sub

Private Sub txtRentalAmount_LostFocus()
  If IsNumeric(txtRentalAmount) = True Then txtRentalAmount.Text = Format(txtRentalAmount.Text, "##,##0.00")
End Sub
