VERSION 5.00
Begin VB.Form frmPermission 
   Caption         =   "Assign User(s)  Permission(s)"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   Icon            =   "frmPermission.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8940
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSaveSettings 
      Caption         =   "&Save Settings"
      Height          =   375
      Left            =   7560
      TabIndex        =   57
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " Check to enable "
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   41
         Left            =   7920
         TabIndex        =   50
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   40
         Left            =   6960
         TabIndex        =   49
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   39
         Left            =   5880
         TabIndex        =   48
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   38
         Left            =   5040
         TabIndex        =   47
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   37
         Left            =   4080
         TabIndex        =   46
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   36
         Left            =   3240
         TabIndex        =   45
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   35
         Left            =   2040
         TabIndex        =   44
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   34
         Left            =   7920
         TabIndex        =   43
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   33
         Left            =   6960
         TabIndex        =   42
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   32
         Left            =   5880
         TabIndex        =   41
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   31
         Left            =   5040
         TabIndex        =   40
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   30
         Left            =   4080
         TabIndex        =   39
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   29
         Left            =   3240
         TabIndex        =   38
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   28
         Left            =   2040
         TabIndex        =   37
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   27
         Left            =   7920
         TabIndex        =   36
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   26
         Left            =   6960
         TabIndex        =   35
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   25
         Left            =   5880
         TabIndex        =   34
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   24
         Left            =   5040
         TabIndex        =   33
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   23
         Left            =   4080
         TabIndex        =   32
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   22
         Left            =   3240
         TabIndex        =   31
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   21
         Left            =   2040
         TabIndex        =   30
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   20
         Left            =   7920
         TabIndex        =   29
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   19
         Left            =   6960
         TabIndex        =   28
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   18
         Left            =   5880
         TabIndex        =   27
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   17
         Left            =   5040
         TabIndex        =   26
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   16
         Left            =   4080
         TabIndex        =   25
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   15
         Left            =   3240
         TabIndex        =   24
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   23
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   13
         Left            =   7920
         TabIndex        =   22
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   12
         Left            =   6960
         TabIndex        =   21
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   11
         Left            =   5880
         TabIndex        =   20
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   19
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   9
         Left            =   4080
         TabIndex        =   18
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   8
         Left            =   3240
         TabIndex        =   17
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   16
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   6
         Left            =   7920
         TabIndex        =   15
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   5
         Left            =   6960
         TabIndex        =   14
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   4
         Left            =   5880
         TabIndex        =   13
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   3
         Left            =   5040
         TabIndex        =   12
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   10
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   56
         Top             =   3270
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   55
         Top             =   2800
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   54
         Top             =   2310
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   53
         Top             =   1830
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   52
         Top             =   1350
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   51
         Top             =   860
         Width           =   120
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Access Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Return Items"
         Height          =   195
         Left            =   7560
         TabIndex        =   7
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Users"
         Height          =   195
         Left            =   6840
         TabIndex        =   6
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Members"
         Height          =   195
         Left            =   5640
         TabIndex        =   5
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Items"
         Height          =   195
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Search"
         Height          =   195
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Report"
         Height          =   195
         Left            =   3120
         TabIndex        =   2
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transaction (Rent)"
         Height          =   195
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub CreateDB()
On Error Resume Next
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Dim daoWS As dao.Workspace
Dim daoDB As dao.Database
Dim daoTable As New dao.TableDef
Dim daoField As New dao.Field
Set daoWS = DBEngine.Workspaces(0)

' Create database
 Set daoDB = daoWS.CreateDatabase(App.Path & "\Permission.mdb", dbLangGeneral, dbVersion40)

' Create Table
 Set daoTable = daoDB.CreateTableDef("PermissionTable")

' Create Fields
Set daoField = daoTable.CreateField("AccessLevel", dbInteger)

' Append field to table
daoTable.Fields.Append daoField

Set daoField = daoTable.CreateField("Permissions", dbText, 255)

' Append field to table
daoTable.Fields.Append daoField

' Append table to database
daoDB.TableDefs.Append daoTable

' Clean up objects
Set daoField = Nothing
Set daoTable = Nothing
Set daoDB = Nothing
Set daoWS = Nothing
Call vr_engine.SetDatabasePassword(App.Path & "\Permission.mdb", "AdmiN")
End Sub
Sub initDB()
Dim loop1 As Integer
Dim db As Database
Dim rec As Recordset
Set db = OpenDatabase(App.Path & "\Permission.mdb" _
             , False, False, ";pwd=AdmiN")
Set rec = db.OpenRecordset("PermissionTable", dbOpenTable)

For loop1 = 1 To 6
    rec.AddNew
    rec.Fields("AccessLevel") = loop1
    rec.Fields("Permissions") = "0000000"
    rec.Update
Next loop1

Set db = Nothing
Set rec = Nothing
End Sub
Sub UpdateDB()

Dim loop1, loop3, counter As Integer
Dim strn As String
Dim db As Database
Dim rec As Recordset
Set db = OpenDatabase(App.Path & "\Permission.mdb" _
             , False, False, ";pwd=AdmiN")
Set rec = db.OpenRecordset("PermissionTable", dbOpenTable)
    counter = 0
    rec.MoveFirst
    strn = ""
For loop1 = 1 To 6
      rec.Edit
      For loop3 = 1 To 7
        strn = strn & Trim(str(Check1(counter + loop3 - 1).Value))
      Next loop3
      rec.Fields("Permissions") = strn
      strn = ""
      rec.Update
      counter = counter + 7
    If rec.EOF = False Then rec.MoveNext
Next loop1

Set db = Nothing
Set rec = Nothing
End Sub
Sub loadvaluesToChkbox()
Dim loop1, loop2, counter As Integer
Dim str As String
Dim db As Database
Dim rec As Recordset
Set db = OpenDatabase(App.Path & "\Permission.mdb" _
             , False, False, ";pwd=AdmiN")
Set rec = db.OpenRecordset("PermissionTable", dbOpenTable)
counter = 0
    rec.MoveFirst
For loop1 = 1 To 6
    str = Trim(rec.Fields("Permissions"))
    For loop2 = 1 To 7
        Check1(counter).Value = Int(Val(Mid(str, loop2, 1)))
        counter = counter + 1
    Next loop2
    If rec.EOF = False Then rec.MoveNext
Next loop1

Set db = Nothing
Set rec = Nothing
End Sub



Private Sub Check1_Click(Index As Integer)
Check1(40).Value = 1
End Sub

Private Sub cmdSaveSettings_Click()
Call UpdateDB
MsgBox "Settings have been successfully changed.  ", vbInformation, "Settings changed"
Check1(0).SetFocus
End Sub
Private Sub Form_Activate()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Dim loop1 As Integer
' Check if file exist
If vr_engine.ReportFileStatus(App.Path & "\Permission.mdb") = False Then
    Call CreateDB
    Call initDB
    Call loadvaluesToChkbox
End If
'Save to DB
Call loadvaluesToChkbox
Check1(0).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me

End Sub
