VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReturnItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return Items"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   Icon            =   "frmReturnItems.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   11580
   Begin VB.Frame FrameII 
      Height          =   3135
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   11295
      Begin VB.TextBox txtChange 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtAmountPaid 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9240
         TabIndex        =   13
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox chkEnablePrintReceipt 
         Caption         =   "Enable Print Receipt"
         Height          =   255
         Left            =   9240
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   2040
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid FlexOverdue 
         Height          =   1410
         Left            =   7560
         TabIndex        =   11
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2487
         _Version        =   393216
         FixedCols       =   0
      End
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Execute Return"
         Height          =   975
         Left            =   7560
         TabIndex        =   15
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdTransferToLeft 
         Caption         =   "<=="
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdTransferToRight 
         Caption         =   "==>"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.ListBox lst2 
         Height          =   2400
         ItemData        =   "frmReturnItems.frx":030A
         Left            =   4320
         List            =   "frmReturnItems.frx":030C
         TabIndex        =   8
         Top             =   480
         Width           =   3015
      End
      Begin VB.ListBox lst1 
         Height          =   2400
         ItemData        =   "frmReturnItems.frx":030E
         Left            =   240
         List            =   "frmReturnItems.frx":0310
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Change:"
         Height          =   255
         Index           =   4
         Left            =   8560
         TabIndex        =   24
         Top             =   2775
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Amnt Pd:"
         Height          =   255
         Index           =   3
         Left            =   8520
         TabIndex        =   23
         Top             =   2430
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Total:"
         Height          =   255
         Index           =   2
         Left            =   8760
         TabIndex        =   22
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Overdue Charge(s)"
         Height          =   255
         Index           =   1
         Left            =   7560
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Item Code(s) to be returned"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Unreturned Item Code(s)"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   11295
      Begin VB.CommandButton cmdFlexToExcel 
         Caption         =   "&Copy FlexGrid Contents to Excel"
         Height          =   375
         Left            =   8515
         TabIndex        =   25
         Top             =   210
         Width           =   2535
      End
      Begin VB.ComboBox cboUnreturnedItemsSortBy 
         Height          =   315
         ItemData        =   "frmReturnItems.frx":0312
         Left            =   960
         List            =   "frmReturnItems.frx":031F
         TabIndex        =   2
         Text            =   "LastDateBorrowed"
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton OptUnreturnedItemsAsc 
         Caption         =   "Ascending"
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptUnreturnedItemsDesc 
         Caption         =   "Descending"
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   5990
         TabIndex        =   5
         Top             =   210
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Sort by: "
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frmReturnItems1 
      Caption         =   " Unreturned Item(s) "
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1935
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   8
         AllowUserResizing=   1
      End
   End
End
Attribute VB_Name = "frmReturnItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FlagPrevActivate As Boolean
Private Sub cmdExecute_Click()
    MousePointer = vbHourglass
    Dim ReturnSuccess As Boolean
    Dim TDM
    Dim strInvNum As String
    Dim intInvNum, loop1 As Long
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    
    If lst2.ListCount = 0 Then
        MousePointer = vbDefault
        MSFlexGrid1.SetFocus
        Exit Sub
    End If
    
  ' SAVE TRANS TO DB
  If Trim(FlexOverdue.TextMatrix(1, 0)) <> "" Then ' SAVE TRANS
  '' Load Invoice Number
            If vr_engine.ReportFileStatus(App.Path & "\InvoiceNumber.txt") = True Then
                Open App.Path & "\InvoiceNumber.txt" For Input As #1
                Line Input #1, strInvNum
                Close #1
                    If IsNumeric(strInvNum) = True Then
                        intInvNum = Int(Val(strInvNum)) + 1
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
            strInvNum = str(intInvNum)
  '' End Load Invoice Number
            Call vr_engine.CheckIfMiscTransSalesDBExistIfNotCreate
            Call vr_engine.Return_SaveReturnTransaction(FlexOverdue, str(strInvNum))
  End If ' SAVE TRANS
  ' END SAVE TRANS TO DB
   If chkEnablePrintReceipt.Value = False Then
       ReturnSuccess = vr_engine.Return_Items(lst2)
      If ReturnSuccess = True Then
       'Start - Clear Flex
        TDM = DoEvents
        FlexOverdue.Rows = 2
        FlexOverdue.TextMatrix(1, 0) = ""
        FlexOverdue.TextMatrix(1, 1) = ""
        TDM = DoEvents
       'End - Clear Flex
        txtTotal.Text = "0.00"
        txtAmountPaid.Text = ""
        txtChange.Text = ""
       
        MsgBox "Item(s) has been sucessfully returned.  ", vbInformation, "Database Updated. "
        
        Call cmdRefresh_Click
        MousePointer = vbDefault
        MSFlexGrid1.SetFocus
      End If
   Else '-------------------------------------------------------
        ReturnSuccess = vr_engine.Return_Items(lst2)
        If ReturnSuccess = True Then
            
      'Start - Print Receipt
      If Trim(FlexOverdue.TextMatrix(1, 0)) <> "" Then 'Print Rcpt
      '--------------------------------------------
            If MsgBox("Please insert 8 1/2"" by 13""  paper. ", vbOKCancel, "Insert Paper ") = vbCancel Then
                        'Start - Clear Flex
                            TDM = DoEvents
                            FlexOverdue.Rows = 2
                            FlexOverdue.TextMatrix(1, 0) = ""
                            FlexOverdue.TextMatrix(1, 1) = ""
                            TDM = DoEvents
                        'End - Clear Flex
                            txtTotal.Text = "0.00"
                            txtAmountPaid.Text = ""
                            txtChange.Text = ""
                            MsgBox "Item(s) has been sucessfully returned.  ", vbInformation, "Database Updated. "
                            Call cmdRefresh_Click
                            MousePointer = vbDefault
                            MSFlexGrid1.SetFocus
                            Exit Sub
            End If
      'End If
     '--------------------------------------------
            Printer.Font = "Lucida Console"
            Printer.PaperSize = vbPRPSLegal ' 8.5 by 14 inc
            Printer.FontSize = 9
            Printer.Orientation = 1 'Portrait
      'Start - Rcpt Header
            Dim LeftMargin As Integer
            LeftMargin = 10
            Printer.Print ""
            Printer.Print "" ' You can put your company name here.
            Printer.Print Tab(LeftMargin); "______________________________________________________________________________________________"
            Printer.Print Tab(LeftMargin); "INVOICE No.   : " & UCase(Trim(strInvNum))
            Printer.Print Tab(LeftMargin); "NAME          : " & UCase("***** Member *****") & "  __________"
            Printer.Print Tab(LeftMargin); "DATE          : " & UCase(Format(Now, "mmm. dd, yyyy"))
            Printer.Print Tab(LeftMargin); "CASHIER       : " & UCase(Mid(gVarFirstName, 1, 1)) & ". " & UCase(gVarFamilyName) & "  __________"
            Printer.Print Tab(LeftMargin); "=============================================================================================="
            Printer.Print Tab(LeftMargin); "Date Due        Item Code       Film Title                                     Amount      "
            'End - Rcpt Header
            'Detailed Section
            For loop1 = 1 To FlexOverdue.Rows - 1
                ''Printer.Print Tab(LeftMargin); "FEB. 25, 2002"; Tab(LeftMargin + 16); "VHS-0001"; Tab(LeftMargin + 32); "FErdies' Wave Fage"; Tab(LeftMargin + 85 - Len("50.45")); "50.45"
                Printer.Print Tab(LeftMargin); "*********"; Tab(LeftMargin + 16); "*********"; Tab(LeftMargin + 32); "Overdue : " & FlexOverdue.TextMatrix(loop1, 0); Tab(LeftMargin + 85 - Len(Trim(FlexOverdue.TextMatrix(loop1, 1)))); Trim(FlexOverdue.TextMatrix(loop1, 1)); ""
                If loop1 = FlexOverdue.Rows - 1 Then
                    Printer.Print Tab(LeftMargin); "----------------------------------------------------------------------------------------------"
                    Printer.Print Tab(LeftMargin); "TOTAL         : "; Tab(LeftMargin + 34); Trim(str(FlexOverdue.Rows - 1)) & " - Item(s)"; Tab(LeftMargin + 85 - Len("P  " & Trim(txtTotal))); "P  " & Trim(txtTotal)
                    Printer.Print Tab(LeftMargin); "______________________________________________________________________________________________"
                End If
            Next loop1
            'End Detiled Section
            Printer.EndDoc
            'End - Print Receipt
        End If 'Print Rcpt
     '--------------------------------------------
      End If 'ReturnSuccess
      'Start - Clear Flex
            TDM = DoEvents
            FlexOverdue.Rows = 2
            FlexOverdue.TextMatrix(1, 0) = ""
            FlexOverdue.TextMatrix(1, 1) = ""
            TDM = DoEvents
            'End - Clear Flex
            txtTotal.Text = "0.00"
            txtAmountPaid.Text = ""
            txtChange.Text = ""
      MsgBox "Item(s) has been sucessfully returned.  ", vbInformation, "Database Updated. "
      Call cmdRefresh_Click
      MousePointer = vbDefault
      MSFlexGrid1.SetFocus
      
   End If
    
End Sub

Private Sub cmdFlexToExcel_Click()
MousePointer = vbHourglass
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Call vr_engine.CopyFlexDataToExcel(MSFlexGrid1)
MousePointer = vbDefault
MSFlexGrid1.SetFocus
End Sub

Private Sub cmdRefresh_Click()
    MousePointer = vbHourglass
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    Call vr_engine.Report_LoadUnreturnedItems(MSFlexGrid1, UnreturnedItems_SQL())
    Call LoadUnreturnedItemcodesToListBox
    lst2.Clear
    MSFlexGrid1.SetFocus
    MousePointer = vbDefault
End Sub

Sub UnreturnedItems_initializeFLEXGRID(Flex As MSFlexGrid)
MousePointer = vbHourglass
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Dim mySQL As String
      
With Flex
     .ColWidth(0) = 600 'No.
     .ColWidth(1) = 1600 'Item Code
     .ColWidth(2) = 3000 'Title
     .ColWidth(3) = 1600 'Date Borrowed
     
     .ColWidth(4) = 1600 'Date Due
     
     .ColWidth(5) = 1600 'Overdue (Days)
     .ColWidth(6) = 1600 'Overdue Charges
     .ColWidth(7) = 3000 'Additional Info
     .ColAlignment(0) = 5 'No.
     .ColAlignment(1) = 5 'Item Code
     .ColAlignment(2) = 5 'Title
     .ColAlignment(3) = 5 'Date Borrowed
     
     .ColAlignment(4) = 5
     
     .ColAlignment(5) = 5 'Overdue (Days)
     .ColAlignment(6) = 5 'Overdue Charges
     '.ColAlignment(4) = 5 'Additional Info
     
     .TextMatrix(0, 0) = " No. "
     .TextMatrix(0, 1) = "Item Code"
     .TextMatrix(0, 2) = "Title"
     .TextMatrix(0, 3) = "Date Borrowed"
     
     .TextMatrix(0, 4) = "Date Due"
     
     .TextMatrix(0, 5) = "Overdue (Days)"
     .TextMatrix(0, 6) = "Overdue Charges"
     .TextMatrix(0, 7) = "                      Additional Info"
End With
     'mySQL = "SELECT * FROM [CD TAPES TABLE] WHERE Available = 'No' ORDER BY LastDateBorrowed "
     Call vr_engine.Report_LoadUnreturnedItems(MSFlexGrid1, UnreturnedItems_SQL())
     ' Start - Count Unreturned Items
     If MSFlexGrid1.Rows > 1 Then
        If Trim(MSFlexGrid1.TextMatrix(1, 0)) <> "" Then
            frmReturnItems1.Caption = "Unreturned Item(s) " & str(MSFlexGrid1.Rows - 1) & " "
        Else
            frmReturnItems1.Caption = "Unreturned Item(s) 0 "
        End If
    End If
     'End - Count Unreturned Items
     Call LoadUnreturnedItemcodesToListBox
MousePointer = vbDefault
End Sub
Function UnreturnedItems_SQL() As String
  If OptUnreturnedItemsAsc = True Then
    UnreturnedItems_SQL = "SELECT * FROM [CD TAPES TABLE] WHERE Available = 'No' ORDER BY [" & Trim(cboUnreturnedItemsSortBy.Text) & "]"
  Else
    UnreturnedItems_SQL = "SELECT * FROM [CD TAPES TABLE] WHERE Available = 'No' ORDER BY [" & Trim(cboUnreturnedItemsSortBy.Text) & "] Desc"
  End If
End Function

Private Sub cmdTransferToLeft_Click()
   On Error GoTo ErrorHandler
   Dim Char, tmpString As String
   Dim tmpLst2text As String
   Dim counter As Integer
   Dim loop1 As Long
   Dim Str1, str2 As String
   counter = 0
   Char = "A"
   tmpString = ""
   If Trim(lst2.Text <> "") Then
     tmpLst2text = lst2.Text
     
     Do While Char <> " "
        counter = counter + 1
        Char = Mid(Trim(lst2.Text), counter, 1)
        tmpString = tmpString & Char
     Loop
     lst1.AddItem Trim(tmpString)
     lst1.Text = Trim(tmpString)
     lst2.RemoveItem lst2.ListIndex
   End If
   ' Start Remove items from FlexOverdue
     
        For loop1 = 1 To FlexOverdue.Rows - 1
            Str1 = Trim(FlexOverdue.TextMatrix(loop1, 0))
            str2 = Trim(Mid(tmpLst2text, Len(tmpString), Len(tmpLst2text) + 1 - Len(tmpString)))
            If Len(Trim(Str1)) > 0 Then
                If Str1 = str2 Then
                    If FlexOverdue.Rows = 2 Then
                        FlexOverdue.TextMatrix(1, 0) = ""
                        FlexOverdue.TextMatrix(1, 1) = ""
                    Else
                        FlexOverdue.RemoveItem (loop1)
                    End If
                    Exit For
                End If
            End If
        Next loop1
    ' End Remove items from FlexOverdue
ErrorHandler:
   Call TotalOverDue
   lst1.SetFocus
End Sub

Private Sub cmdTransferToRight_Click()
  Dim FlexRow, loop1 As Long
  If Trim(lst1.Text <> "") Then
     For loop1 = 1 To MSFlexGrid1.Rows - 1
         If Trim(MSFlexGrid1.TextMatrix(loop1, 1)) = Trim(lst1.Text) Then
            lst2.AddItem lst1.Text & "  " & Trim(MSFlexGrid1.TextMatrix(loop1, 2))
            lst2.Text = lst1.Text & "  " & Trim(MSFlexGrid1.TextMatrix(loop1, 2))
            Exit For
         End If
     Next
     
   'Start - Add Overdue items
    If Trim(MSFlexGrid1.TextMatrix(loop1, 6)) <> "0.00" Then
       If Trim(FlexOverdue.TextMatrix(1, 0)) = "" And FlexOverdue.Rows = 2 Then
          ' Do Nothing
       Else
          FlexOverdue.AddItem ""
       End If
       
       'Title Col
       FlexOverdue.TextMatrix(FlexOverdue.Rows - 1, 0) = MSFlexGrid1.TextMatrix(loop1, 2)
       'Overdue Amount Col
       FlexOverdue.TextMatrix(FlexOverdue.Rows - 1, 1) = MSFlexGrid1.TextMatrix(loop1, 6) & "   "
    End If
   'End - Add Overdue items
     lst1.RemoveItem lst1.ListIndex
  End If
  Call TotalOverDue
  lst1.SetFocus
End Sub

Private Sub FlexOverdue_Click()
'On Error Resume Next
 Dim loop1, counter As Long
 Dim Char As String
 
  For loop1 = 1 To lst2.ListCount
   If FlexOverdue.Rows > 1 Then
      counter = 0
      Char = "a" ' initialize to some values not equal to SPACE
      Do While Char <> " "
        counter = counter + 1
        Char = Mid(Trim(lst2.List(loop1 - 1)), counter, 1)
      Loop
        If Trim(Mid(lst2.List(loop1 - 1), counter, Len(lst2.List(loop1 - 1)) + 1 - counter)) = Trim(FlexOverdue.TextMatrix(FlexOverdue.RowSel, 0)) Then
           lst2.Text = lst2.List(loop1 - 1)
        End If
   End If
  Next
End Sub

Private Sub FlexOverdue_SelChange()
   Call FlexOverdue_Click
End Sub

Private Sub Form_Activate()
If FlagPrevActivate = True Then
Else
  FlagPrevActivate = True
  Call UnreturnedItems_initializeFLEXGRID(MSFlexGrid1)
  ' Start - FlexOverdue initialize
    FlexOverdue.ColWidth(0) = 2270
    FlexOverdue.ColWidth(1) = 950
    FlexOverdue.TextMatrix(0, 0) = "Title"
    FlexOverdue.TextMatrix(0, 1) = "Amount    "
    FlexOverdue.ColAlignment(0) = 5
    FlexOverdue.ColAlignment(1) = 6
  ' End - FlexOverdue initialize
End If
End Sub

Private Sub LoadUnreturnedItemcodesToListBox()
   Dim FlexRow, loop1 As Long
    lst1.Clear  ' Erase list
    FlexRow = MSFlexGrid1.Rows
    If FlexRow > 1 Then
        For loop1 = 1 To FlexRow - 1
            If Trim(MSFlexGrid1.TextMatrix(loop1, 1)) <> "" Then
               lst1.AddItem Trim(MSFlexGrid1.TextMatrix(loop1, 1))
            End If
        Next
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
FlagPrevActivate = False
End Sub

Private Sub lst1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   Call cmdTransferToRight_Click
End If
End Sub
Private Sub lst2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   Call cmdTransferToLeft_Click
   lst2.SetFocus
End If
End Sub

Private Sub MSFlexGrid1_Click()
   On Error Resume Next
   If MSFlexGrid1.Rows > 1 Then lst1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 1)
End Sub


Private Sub MSFlexGrid1_EnterCell()
MSFlexGrid1.CellBackColor = vbCyan
End Sub

Private Sub MSFlexGrid1_GotFocus()

If MSFlexGrid1.Rows > 1 Then
   If Trim(MSFlexGrid1.TextMatrix(1, 0)) <> "" Then
        frmReturnItems1.Caption = "Unreturned Item(s) " & str(MSFlexGrid1.Rows - 1) & " "
     Else
        frmReturnItems1.Caption = "Unreturned Item(s) 0 "
   End If
End If

End Sub

Private Sub MSFlexGrid1_LeaveCell()
MSFlexGrid1.CellBackColor = vbWhite
End Sub

Private Sub MSFlexGrid1_LostFocus()
MSFlexGrid1.CellBackColor = vbWhite
End Sub

Private Sub MSFlexGrid1_SelChange()
   On Error Resume Next
   If MSFlexGrid1.Rows > 1 Then
      lst1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 1)
      MSFlexGrid1.CellBackColor = vbCyan
   End If
End Sub

Sub TotalOverDue()
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    Dim Total As Double
    Total = 0
    For loop1 = 1 To FlexOverdue.Rows - 1
       Total = Total + Val(FlexOverdue.TextMatrix(loop1, 1))
    Next
    Call vr_engine.Round(Total, 2)
    txtTotal.Text = Format(str(Total), "0.00")
    
End Sub

Private Sub OptUnreturnedItemsAsc_Click()
Call cmdRefresh_Click

End Sub

Private Sub OptUnreturnedItemsDesc_Click()
Call cmdRefresh_Click
End Sub

Private Sub txtAmountPaid_Change()
If Val(txtTotal.Text) > 0 And Val(txtAmountPaid.Text) < Val(txtTotal.Text) Then
     cmdExecute.Enabled = False
  Else
     cmdExecute.Enabled = True
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
          If IsNumeric(txtTotal.Text) Then
             txtChange.Text = str(Val(Val(txtAmountPaid.Text) - Val(txtTotal.Text)))
             txtChange.Text = Format(txtChange.Text, "0.00")
          End If
       Else
          MsgBox "Amount paid is invalid. ", vbInformation, "Invalid input.  "
          txtAmountPaid.Text = ""
          txtChange.Text = ""
          txtAmountPaid.SetFocus
       End If
    End If
End Sub

Private Sub txtChange_Change()
txtChange.Text = Format(txtChange.Text, "0.00")
    If Val(txtChange.Text) < 0 Or Trim(txtChange.Text) = "" Then
        '
    Else
       If Val(txtTotal.Text) > 0 Then
          '
       Else
          '
       End If
    End If
End Sub

Private Sub txtTotal_Change()
  If Val(txtTotal.Text) > 0 And Val(txtAmountPaid.Text) < Val(txtTotal.Text) Then
     cmdExecute.Enabled = False
  Else
     cmdExecute.Enabled = True
  End If
  Call txtAmountPaid_LostFocus
End Sub
