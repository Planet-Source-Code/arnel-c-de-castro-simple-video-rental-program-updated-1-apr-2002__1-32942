VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTransaction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   11850
   Begin VB.Frame FrameTranSactionMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   11605
      Begin VB.Frame FrameBorrowedItemsList 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   240
         TabIndex        =   24
         Top             =   2040
         Width           =   11175
         Begin VB.ComboBox cboDeleteRow 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6120
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   3360
            Width           =   975
         End
         Begin VB.CommandButton cmdDeleteRow 
            Caption         =   "Delete Row"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   385
            Left            =   4800
            TabIndex        =   7
            Top             =   3360
            Width           =   1255
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2175
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   3836
            _Version        =   393216
            Cols            =   4
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdAddItem 
            Caption         =   "Add Item"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   385
            Left            =   3120
            TabIndex        =   6
            Top             =   3360
            Width           =   1215
         End
         Begin VB.CommandButton cmdRefreshList 
            Caption         =   "Refresh List"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9240
            TabIndex        =   4
            Top             =   930
            Width           =   1695
         End
         Begin VB.ListBox lstMembers 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1740
            ItemData        =   "frmTransaction.frx":0442
            Left            =   7920
            List            =   "frmTransaction.frx":0444
            Sorted          =   -1  'True
            TabIndex        =   3
            ToolTipText     =   "Highlight and double-click or press enter to select. "
            Top             =   1290
            Width           =   3015
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cboItemCode 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   600
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   3360
            Width           =   2295
         End
         Begin VB.TextBox txtChange 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9720
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   3960
            Width           =   1215
         End
         Begin VB.TextBox txtAmountPaid 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9720
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   3600
            Width           =   1215
         End
         Begin VB.TextBox txtTotalAmountDue 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9720
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Frame frameOptionsControls 
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   0
            TabIndex        =   27
            Top             =   3720
            Width           =   7080
            Begin VB.CommandButton cmdCancel 
               Caption         =   "&Cancel"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4140
               TabIndex        =   17
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton cmdPrint 
               Caption         =   "&Print"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   6090
               TabIndex        =   19
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "&Delete"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5120
               TabIndex        =   18
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "&Save"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3160
               TabIndex        =   16
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton cmdEdit 
               Caption         =   "&Edit"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2190
               TabIndex        =   15
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton cmdFind 
               Caption         =   "&Find"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1210
               TabIndex        =   14
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton cmdNew 
               Caption         =   "&New"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   13
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "F1                   F2                F3                  F4                 F5                 F6                  F7"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   600
               TabIndex        =   34
               Top             =   100
               Width           =   6375
            End
         End
         Begin VB.TextBox txtDate 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9240
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtInvoiceNumber 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Select Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7920
            TabIndex        =   33
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Member's Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   260
            TabIndex        =   32
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Enter/Select Item Code:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   31
            Top             =   3120
            Width           =   2295
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Amount Due:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   7800
            TabIndex        =   30
            Top             =   3300
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount Paid:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   8280
            TabIndex        =   29
            Top             =   3680
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Change:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   8400
            TabIndex        =   28
            Top             =   4020
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   8698
            TabIndex        =   26
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Invoice Number:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4480
            TabIndex        =   25
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame FrameDisplay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   11175
         Begin VB.Label lblDisplay1 
            BackColor       =   &H00000000&
            Caption         =   "Total Amount Due"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   495
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   10695
         End
         Begin VB.Label lblDisplay2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   815
            Left            =   240
            TabIndex        =   22
            Top             =   840
            Width           =   10695
         End
      End
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ArrayOFNamesAndID(), MembersID(), MemberID_FindMode As String
Dim PrevTransItems() As String  'List of PrevTransaction items
Dim DeletedItems() As String
Dim PrevTransMode As Boolean

Private Sub cboDeleteRow_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdDeleteRow_Click
End Sub

Private Sub cboItemCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then Call cmdAddItem_Click
End Sub

Private Sub cmdAddItem_Click()
    Dim strMembersFile As String
    Dim loop1 As Long
    Dim TotalAmountDue As Double
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    If Trim(cboItemCode.Text) = "" Then
        cboItemCode.SetFocus
        Exit Sub
    End If
    
  'Start - Chk if item has been borrowed already.
  If PrevTransMode = False Then
     strMembersFile = txtName.Text & " ID - " & MembersID(lstMembers.ListIndex + 1)
  Else
     strMembersFile = txtName.Text & " ID - " & MemberID_FindMode
  End If
  
   If vr_engine.Transaction_ChkIfItemHasBeenBorrowed(App.Path & "\Transaction\MembersRecords\" & strMembersFile & ".mdb", cboItemCode.Text) = True Then
     cboItemCode.Text = ""
     cboItemCode.SetFocus
     Exit Sub
   End If
  'End - Chk if item has been borrowed already.
  
  'Start - Chk if item code is already selected
    For loop1 = 1 To MSFlexGrid1.Rows
     If loop1 <= MSFlexGrid1.Rows - 1 Then
        If cboItemCode.Text = MSFlexGrid1.TextMatrix(loop1, 1) Then
            MsgBox "Item code is already in the list. ", vbInformation, "Item Code Has Been Loaded."
            cboItemCode.SetFocus
            Exit Sub
        End If
     End If
    Next
  'End - Chk if item code is already selected

    Call vr_engine.Transaction_AddItem(MSFlexGrid1, cboItemCode)
    '' caculate Total Amnt.
    For loop1 = 1 To MSFlexGrid1.Rows - 1
        TotalAmountDue = TotalAmountDue + Val(MSFlexGrid1.TextMatrix(loop1, 3))
    Next
     If MSFlexGrid1.Rows = 1 Then MSFlexGrid1.AddItem ""
     If Trim(MSFlexGrid1.TextMatrix(1, 0)) = "" And Trim(MSFlexGrid1.TextMatrix(1, 1)) = "" And Trim(MSFlexGrid1.TextMatrix(1, 2)) = "" And Trim(MSFlexGrid1.TextMatrix(1, 3)) = "" Then
        cmdDeleteRow.Enabled = False
        cboDeleteRow.Enabled = False
        cboDeleteRow.Clear
     Else
        cmdDeleteRow.Enabled = True
        cboDeleteRow.Enabled = True
        cboDeleteRow.Clear
        For loop1 = 1 To MSFlexGrid1.Rows - 1
            cboDeleteRow.AddItem str(loop1)
        Next
     End If
     
        txtTotalAmountDue.Text = str(TotalAmountDue)
        cboItemCode.Text = ""
        cboItemCode.SetFocus
    
End Sub

Private Sub cmdCancel_Click()
    cmdCancel.Enabled = False
    txtName.Text = ""
    txtInvoiceNumber.Text = ""
    txtDate.Text = ""
    cmdRefreshList.Enabled = False
    lstMembers.Clear
    lstMembers.Enabled = False
    cboItemCode.Clear
    cboItemCode.Enabled = False
    cmdAddItem.Enabled = False
    cmdDeleteRow.Enabled = False
    cboDeleteRow.Clear
    cboDeleteRow.Enabled = False
    txtTotalAmountDue.Text = ""
    txtAmountPaid.Text = ""
    txtChange.Text = ""
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.TextMatrix(1, 0) = ""
    MSFlexGrid1.TextMatrix(1, 1) = ""
    MSFlexGrid1.TextMatrix(1, 2) = ""
    MSFlexGrid1.TextMatrix(1, 3) = ""
    lblDisplay2.Caption = "0.00"

    
    cmdNew.Enabled = True
    cmdFind.Enabled = True
    If cmdEdit.Enabled = True Then cmdEdit.Enabled = False
    If cmdDelete.Enabled = True Then cmdDelete.Enabled = False
    If cmdPrint.Enabled = True Then cmdPrint.Enabled = False
    
End Sub

Private Sub cmdDelete_Click()
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    Dim X
    X = MsgBox("Are you sure you want to delete this transaction?  ", vbYesNo, "Confirm Delete")
    If X = vbNo Then
       MSFlexGrid1.SetFocus
       Exit Sub
    End If
    Call vr_engine.Transaction_DeletePrevTransaction(Trim(txtInvoiceNumber.Text), Trim(txtName.Text) & " ID - " & MemberID_FindMode & ".mdb")
    Call cmdCancel_Click
    
End Sub

Private Sub cmdDeleteRow_Click()
Dim loop1 As Integer
Dim FlagFORone As Boolean
Dim TotalAmountDue As Double
FlagFORone = False
   If IsNumeric(cboDeleteRow.Text) = True Then
       If Int(cboDeleteRow.Text) + 1 = MSFlexGrid1.Rows Then
            MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1
            cboDeleteRow.RemoveItem (cboDeleteRow.ListCount - 1)
            For loop1 = 1 To MSFlexGrid1.Rows - 1
                TotalAmountDue = TotalAmountDue + Val(MSFlexGrid1.TextMatrix(loop1, 3))
            Next
                txtTotalAmountDue.Text = str(TotalAmountDue)
                cboItemCode.Text = ""
                cboItemCode.SetFocus
   Else
                If Int(cboDeleteRow.Text) = 1 Then FlagFORone = True
            For loop1 = Int(cboDeleteRow.Text) + 1 To cboDeleteRow.ListCount
                If FlagFORone = True Then loop1 = 2
                FlagFORone = False
                MSFlexGrid1.TextMatrix(loop1 - 1, 1) = MSFlexGrid1.TextMatrix(loop1, 1)
                MSFlexGrid1.TextMatrix(loop1 - 1, 2) = MSFlexGrid1.TextMatrix(loop1, 2)
                MSFlexGrid1.TextMatrix(loop1 - 1, 3) = MSFlexGrid1.TextMatrix(loop1, 3)
                
            Next
                MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1
                For loop1 = 1 To MSFlexGrid1.Rows - 1
                    TotalAmountDue = TotalAmountDue + Val(MSFlexGrid1.TextMatrix(loop1, 3))
                Next
                cboDeleteRow.RemoveItem (cboDeleteRow.ListCount - 1)
                txtTotalAmountDue.Text = str(TotalAmountDue)
                cboItemCode.Text = ""
                cboItemCode.SetFocus
       End If
   End If
   
   
   
   cboItemCode.SetFocus
End Sub

Private Sub cmdEdit_Click()
  Dim vr_engine As VRENTAL_ENGINE
  Set vr_engine = New VRENTAL_ENGINE
  Dim loop1 As Long
  cmdEdit.Enabled = False
  cmdSave.Enabled = True
  cmdPrint.Enabled = True
  Call vr_engine.Transaction_LoadItemCodes(cboItemCode)
  cboItemCode.Enabled = True
  cmdAddItem.Enabled = True
  '' Load cboRow Number
     If MSFlexGrid1.Rows = 1 Then MSFlexGrid1.AddItem ""
     If Trim(MSFlexGrid1.TextMatrix(1, 0)) = "" And Trim(MSFlexGrid1.TextMatrix(1, 1)) = "" And Trim(MSFlexGrid1.TextMatrix(1, 2)) = "" And Trim(MSFlexGrid1.TextMatrix(1, 3)) = "" Then
        cmdDeleteRow.Enabled = False
        cboDeleteRow.Enabled = False
        cboDeleteRow.Clear
     Else
        cmdDeleteRow.Enabled = True
        cboDeleteRow.Enabled = True
        cboDeleteRow.Clear
        For loop1 = 1 To MSFlexGrid1.Rows - 1
            cboDeleteRow.AddItem str(loop1)
        Next
     End If
  '' End cboLoad Row Number
  
  '' START -- Store Prev Borrowed items to Memory
     ReDim PrevTransItems(MSFlexGrid1.Rows - 1)
     For loop1 = 1 To MSFlexGrid1.Rows - 1
         'Stores Item Code
         PrevTransItems(loop1) = MSFlexGrid1.TextMatrix(loop1, 1)
         ''Debug.Print PrevTransItems(loop1)
     Next
  '' END -- Store Prev Borrowed items to Memory
  
  
End Sub
Private Sub cmdFind_Click()
  Dim vr_engine As VRENTAL_ENGINE
  Set vr_engine = New VRENTAL_ENGINE
  Dim InvoiceNumber As String
  InvoiceNumber = InputBox("Enter Invoice Number:", "Find Previous Transaction")
  If Trim(InvoiceNumber) <> "" Then
     txtInvoiceNumber.Text = InvoiceNumber
  Else
     MSFlexGrid1.SetFocus
     Exit Sub
  End If
  
  If vr_engine.Transaction_LoadPrevTransaction(MSFlexGrid1, txtTotalAmountDue, txtAmountPaid, txtChange, txtDate, txtInvoiceNumber, txtName, MemberID_FindMode) = True Then
    txtAmountPaid.Locked = True
    cmdFind.Enabled = False
    cmdNew.Enabled = False
    cmdEdit.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = True
    PrevTransMode = True
  End If
  
End Sub

Private Sub cmdNew_Click()
  cmdNew.Enabled = False
  Dim vr_engine As VRENTAL_ENGINE
  Set vr_engine = New VRENTAL_ENGINE
  
  
 '' SetFlag
    PrevTransMode = False
  '' Load ItemCodes
      Call vr_engine.Transaction_LoadItemCodes(cboItemCode)
  '' Load Members ' Once
      Call vr_engine.Transaction_LoadNameOfMembers(lstMembers, ArrayOFNamesAndID(), MembersID())
      cmdRefreshList.Enabled = True
      cmdFind.Enabled = False
      cmdCancel.Enabled = True
      lstMembers.Enabled = True
      lblDisplay1.Caption = "NEW TRANSACTION"
      lblDisplay2.Caption = "SELECT NAME FROM LIST "
      If lstMembers.Enabled = True Then lstMembers.SetFocus
  
End Sub

Private Sub cmdPrint_Click()
'--------------------------------------------
If MsgBox("Please insert 8 1/2"" by 13""  paper. ", vbOKCancel, "Insert Paper ") = vbCancel Then
       MSFlexGrid1.SetFocus
       Exit Sub
End If
'--------------------------------------------
MousePointer = vbHourglass
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Printer.Font = "Lucida Console"
Printer.PaperSize = vbPRPSLegal ' 8.5 by 14 inc
Printer.FontSize = 9
Printer.Orientation = 1 'Portrait
    
Dim LeftMargin, PageCount, BlankLines As Integer
Dim DateDue() As String
Dim loop1, loop2, Lines, Flag As Long
LeftMargin = 10
Lines = MSFlexGrid1.Rows - 1
' Start Calculate DateDue
For loop1 = 1 To Lines
    ReDim Preserve DateDue(loop1)
    DateDue(loop1) = vr_engine.Transaction_GetDateDue(MSFlexGrid1.TextMatrix(loop1, 1), Format(Now, "mm/dd/yyyy"))
Next loop1

' End Calculate DateDue
 'Start - Rcpt Header
  Printer.Print ""
  Printer.Print "" ' You can put your company name here.
  Printer.Print Tab(LeftMargin); "______________________________________________________________________________________________"
  Printer.Print Tab(LeftMargin); "INVOICE No.   : " & UCase(Trim(txtInvoiceNumber.Text))
  Printer.Print Tab(LeftMargin); "NAME          : " & UCase(txtName.Text) & "  __________"
  Printer.Print Tab(LeftMargin); "DATE          : " & UCase(txtDate.Text)
  Printer.Print Tab(LeftMargin); "CASHIER       : " & UCase(Mid(gVarFirstName, 1, 1)) & ". " & UCase(gVarFamilyName) & "  __________"
  Printer.Print Tab(LeftMargin); "=============================================================================================="
  Printer.Print Tab(LeftMargin); "Date Due        Item Code       Film Title                                     Amount      "
 'End - Rcpt Header
 'Detailed Section
  For loop2 = 1 To Lines
    ''Printer.Print Tab(LeftMargin); "FEB. 25, 2002"; Tab(LeftMargin + 16); "VHS-0001"; Tab(LeftMargin + 32); "FErdies' Wave Fage"; Tab(LeftMargin + 85 - Len("50.45")); "50.45"
    Printer.Print Tab(LeftMargin); UCase(Format(DateDue(loop2), "mmm. dd, yyyy")); Tab(LeftMargin + 16); MSFlexGrid1.TextMatrix(loop2, 1); Tab(LeftMargin + 32); MSFlexGrid1.TextMatrix(loop2, 2); Tab(LeftMargin + 85 - Len(MSFlexGrid1.TextMatrix(loop2, 3))); MSFlexGrid1.TextMatrix(loop2, 3)
    If loop2 = Lines Then
    Printer.Print Tab(LeftMargin); "----------------------------------------------------------------------------------------------"
    Printer.Print Tab(LeftMargin); "TOTAL         : "; Tab(LeftMargin + 34); Trim(str(Lines)) & " - Item(s)"; Tab(LeftMargin + 85 - Len("P  " & Trim(txtTotalAmountDue.Text))); "P  " & Trim(txtTotalAmountDue.Text)
    Printer.Print Tab(LeftMargin); "______________________________________________________________________________________________"
    End If
  Next loop2
 'End Detiled Section
 Printer.EndDoc
 MSFlexGrid1.SetFocus
 MousePointer = vbDefault
End Sub

Private Sub cmdRefreshList_Click()
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    Call vr_engine.Transaction_LoadNameOfMembers(lstMembers, ArrayOFNamesAndID(), MembersID())
    lblDisplay2.Caption = "SELECT NAME FROM LIST "
    If lstMembers.Enabled = True Then lstMembers.SetFocus
  
End Sub

Private Sub cmdSave_Click()
'--------------------------------------------
Dim MsgResponse
MsgResponse = MsgBox("Do you want to print out the receipt? ", vbYesNoCancel, App.Title)
If MsgResponse = vbCancel Then
       MSFlexGrid1.SetFocus
       Exit Sub
End If
If MsgResponse = vbYes Then
    Call cmdPrint_Click
End If
'--------------------------------------------
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
Select Case PrevTransMode
Case False
    Call vr_engine.CheckIfTransactionDBExistIfNotCreate
    Call vr_engine.Transaction_ChkForMembersFIleDBFolderIfNotCreate
    Call vr_engine.Transaction_ChkIfBorrowedItemsHistoryDBExistIfNotCreate
' '  Call vr_engine.Transaction_CheckForMembersRecordsIfNotExistsCreate(App.Path & "\Transaction\MembersRecords\", "Decastro.mdb")
    Call vr_engine.Transaction_SaveNewTransaction(MSFlexGrid1, txtDate, gVarFirstName & " " & Mid(gVarMiddleName, 1, 1) & ". " & gVarFamilyName, gVarUserID, txtInvoiceNumber, txtName, txtTotalAmountDue, txtAmountPaid, txtChange, ArrayOFNamesAndID(lstMembers.ListIndex + 1) & ".mdb", MembersID(lstMembers.ListIndex + 1))
    Call cmdCancel_Click
Case True
    ' Start -- Chk 4 deleted prv itemcodes
      Dim delcount As Integer
      Dim loop1, loop2 As Integer
      Dim DelFlag As Boolean
      delcount = 0
      For loop1 = 1 To UBound(PrevTransItems())
        DelFlag = True
        For loop2 = 1 To MSFlexGrid1.Rows - 1
          If Trim(MSFlexGrid1.TextMatrix(loop2, 1)) = Trim(PrevTransItems(loop1)) Then
             DelFlag = False
          End If
        Next loop2
        If DelFlag = True Then
             delcount = delcount + 1
             ReDim Preserve DeletedItems(delcount)
             DeletedItems(delcount) = PrevTransItems(loop1)
             '' Stores deleted items in array -- not used
             ''Debug.Print DeletedItems(delcount)
        End If
      Next loop1
    ' End -- Chk 4 deleted prv itemcodes
    'Save Edited Previous Transaction
    Call vr_engine.Transaction_UpdatePrevTrasaction(MSFlexGrid1, txtDate, gVarFirstName & " " & Mid(gVarMiddleName, 1, 1) & ". " & gVarFamilyName, gVarUserID, txtInvoiceNumber, txtName, txtTotalAmountDue, txtAmountPaid, txtChange, Trim(txtName.Text) & " ID - " & MemberID_FindMode & ".mdb", MemberID_FindMode, PrevTransItems(), DeletedItems())
    Call cmdCancel_Click
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF9 Then
  If txtAmountPaid.Enabled = True Then txtAmountPaid.SetFocus
End If
If KeyCode = vbKeyF1 Then
   If cmdNew.Enabled = True Then Call cmdNew_Click
End If

If KeyCode = vbKeyF2 Then
   If cmdFind.Enabled = True Then Call cmdFind_Click
End If

If KeyCode = vbKeyF3 Then
   If cmdEdit.Enabled = True Then Call cmdEdit_Click
End If

If KeyCode = vbKeyF4 Then
   If cmdSave.Enabled = True Then Call cmdSave_Click
End If

If KeyCode = vbKeyF5 Then
   If cmdCancel.Enabled = True Then Call cmdCancel_Click
End If

If KeyCode = vbKeyF6 Then
   If cmdDelete.Enabled = True Then Call cmdDelete_Click
End If

If KeyCode = vbKeyF7 Then
   If cmdPrint.Enabled = True Then Call cmdPrint_Click
End If

End Sub
Private Sub Form_Load()
  MSFlexGrid1.ColAlignment(0) = 5
  MSFlexGrid1.ColAlignment(1) = 5
  MSFlexGrid1.ColAlignment(2) = 5
  MSFlexGrid1.ColAlignment(3) = 5
  MSFlexGrid1.Rows = 2
  MSFlexGrid1.ColWidth(0) = 800
  MSFlexGrid1.ColWidth(1) = 1550
  MSFlexGrid1.ColWidth(2) = 2900
  MSFlexGrid1.ColWidth(3) = 1750
  MSFlexGrid1.TextMatrix(0, 0) = "No."
  MSFlexGrid1.TextMatrix(0, 1) = "Item Code"
  MSFlexGrid1.TextMatrix(0, 2) = "Title"
  MSFlexGrid1.TextMatrix(0, 3) = "Rental Amount"
End Sub

Private Sub lblDisplay2_Change()
If IsNumeric(lblDisplay2.Caption) = True Then lblDisplay2.Caption = Format(lblDisplay2.Caption, "##,##0.00")
End Sub

Private Sub lstMembers_Click()
  If Trim(lstMembers.Text) <> "" Then lblDisplay2.Caption = lstMembers.Text
  ''MsgBox ArrayOFNamesAndID(lstMembers.ListIndex + 1)
End Sub

Private Sub lstMembers_DblClick()
Dim vr_rental As VRENTAL_ENGINE
Set vr_rental = New VRENTAL_ENGINE
Dim InvNum As String
Dim intInvNum As Long

'' Start -- Chk if Members has Unreturned Items
   Call vr_rental.Transaction_CheckIfMemberHasUnreturnedItems(App.Path & "\Transaction\MembersRecords\" & ArrayOFNamesAndID(lstMembers.ListIndex + 1) & ".mdb")
'' End -- Chk if Members has Unreturned Items

           txtDate.Text = Format(Now, "mmm. dd, yyyy")
            '' Load Invoice Number
        If vr_rental.ReportFileStatus(App.Path & "\InvoiceNumber.txt") = True Then
               Open App.Path & "\InvoiceNumber.txt" For Input As #1
               Line Input #1, InvNum
               Close #1
                 If IsNumeric(InvNum) = True Then
                    intInvNum = Int(Val(InvNum)) + 1
                 Else
                    Open App.Path & "\InvoiceNumber.txt" For Output As #1
                    Print #1, "1"
                    Close #1
                    intInvNum = 1
                 End If
               
        Else
                Open App.Path & "\InvoiceNumber.txt" For Output As #1
                Print #1, "1"
                Close #1
                intInvNum = 1
        End If
            
            
            
            txtInvoiceNumber.Text = str(intInvNum)
            '' End Load Invoice Number
            
            lblDisplay1.Caption = "Total Amount Due"
            lblDisplay2.Caption = "0.00"
            txtName.Text = lstMembers.Text
            lstMembers.Enabled = False
            cmdRefreshList.Enabled = False
            cboItemCode.Enabled = True
            cmdAddItem.Enabled = True
            cboItemCode.SetFocus
         
End Sub

Private Sub lstMembers_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn And Trim(lstMembers.Text) <> "" Then
Dim vr_rental As VRENTAL_ENGINE
Set vr_rental = New VRENTAL_ENGINE
Dim InvNum As String
Dim intInvNum As Long

'' Start -- Chk if Members has Unreturned Items
   Call vr_rental.Transaction_CheckIfMemberHasUnreturnedItems(App.Path & "\Transaction\MembersRecords\" & ArrayOFNamesAndID(lstMembers.ListIndex + 1) & ".mdb")
'' End -- Chk if Members has Unreturned Items

           txtDate.Text = Format(Now, "mmm. dd, yyyy")
            '' Load Invoice Number
        If vr_rental.ReportFileStatus(App.Path & "\InvoiceNumber.txt") = True Then
               Open App.Path & "\InvoiceNumber.txt" For Input As #1
               Line Input #1, InvNum
               Close #1
                 If IsNumeric(InvNum) = True Then
                    intInvNum = Int(Val(InvNum)) + 1
                 Else
                    Open App.Path & "\InvoiceNumber.txt" For Output As #1
                    Print #1, "1"
                    Close #1
                    intInvNum = 1
                 End If
               
        Else
                Open App.Path & "\InvoiceNumber.txt" For Output As #1
                Print #1, "1"
                Close #1
                intInvNum = 1
        End If
            
            
            
            txtInvoiceNumber.Text = str(intInvNum)
            '' End Load Invoice Number
            
            lblDisplay1.Caption = "Total Amount Due"
            lblDisplay2.Caption = "0.00"
            txtName.Text = lstMembers.Text
            lstMembers.Enabled = False
            cmdRefreshList.Enabled = False
            cboItemCode.Enabled = True
            cmdAddItem.Enabled = True
            cboItemCode.SetFocus
End If
End Sub


Private Sub txtAmountPaid_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then Call txtAmountPaid_LostFocus
End Sub

Private Sub txtAmountPaid_LostFocus()
    If Trim(txtAmountPaid.Text) = "" Then
        txtChange.Text = ""
    Else
       If IsNumeric(Trim(txtAmountPaid.Text)) Then
          txtAmountPaid.Text = Format(txtAmountPaid.Text, "0.00")
          If IsNumeric(txtTotalAmountDue.Text) Then
             txtChange.Text = str(Val(Val(txtAmountPaid.Text) - Val(txtTotalAmountDue.Text)))
             txtChange.Text = Format(txtChange.Text, "0.00")
          End If
       Else
          MsgBox "Invalid input.  ", vbInformation, "Amount paid is invalid. "
          txtAmountPaid.Text = ""
          txtChange.Text = ""
          txtAmountPaid.SetFocus
       End If
    End If
End Sub

Private Sub txtChange_Change()
    txtChange.Text = Format(txtChange.Text, "0.00")
    If Val(txtChange.Text) < 0 Or Trim(txtChange.Text) = "" Then
        cmdSave.Enabled = False
        cmdPrint.Enabled = False
        lblDisplay1.Caption = "Total Amount Due"
    Else
       If Val(txtTotalAmountDue.Text) > 0 Then
          If cboItemCode.Enabled = True Then
             cmdSave.Enabled = True
             cmdPrint.Enabled = True
          End If
          lblDisplay1.Caption = "Change"
          lblDisplay2.Caption = txtChange.Text
       Else
          cmdSave.Enabled = False
          cmdPrint.Enabled = False
          lblDisplay1.Caption = "Total Amount Due"
       End If
    End If
End Sub

Private Sub txtTotalAmountDue_Change()

lblDisplay2.Caption = txtTotalAmountDue.Text
 txtTotalAmountDue.Text = Format(txtTotalAmountDue.Text, "0.00")
If IsNumeric(txtTotalAmountDue.Text) Then
    txtAmountPaid.Locked = False
Else
    txtAmountPaid.Locked = True
End If
If IsNumeric(txtTotalAmountDue.Text) = True And IsNumeric(txtAmountPaid.Text) = True Then
     txtChange.Text = str(Val(Val(txtAmountPaid.Text) - Val(txtTotalAmountDue.Text)))
     txtTotalAmountDue.Text = Format(txtTotalAmountDue.Text, "0.00")
End If
End Sub
