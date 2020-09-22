VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   11715
   Visible         =   0   'False
   Begin VB.Frame frmMovieStatistics 
      Height          =   6375
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   11385
      Begin VB.Frame frmControlHolder 
         Height          =   5895
         Left            =   5520
         TabIndex        =   25
         Top             =   240
         Width           =   5625
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   1015
            Left            =   3120
            TabIndex        =   30
            Top             =   4650
            Width           =   2295
         End
         Begin VB.Frame frmOptions 
            Caption         =   "Options "
            Height          =   3015
            Left            =   240
            TabIndex        =   29
            Top             =   2520
            Width           =   2775
            Begin VB.ComboBox cboActor 
               Enabled         =   0   'False
               Height          =   315
               Left            =   300
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   2400
               Width           =   2280
            End
            Begin VB.CheckBox chkActor 
               Caption         =   "Filter by Actor"
               Height          =   255
               Left            =   300
               TabIndex        =   21
               Top             =   2160
               Width           =   1935
            End
            Begin VB.ComboBox cboItemCode 
               Enabled         =   0   'False
               Height          =   315
               Left            =   300
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   1540
               Width           =   2280
            End
            Begin VB.CheckBox chkItemCode 
               Caption         =   "Filter by Item Code"
               Height          =   255
               Left            =   300
               TabIndex        =   19
               Top             =   1290
               Width           =   1935
            End
            Begin VB.ComboBox cboGenre 
               Enabled         =   0   'False
               Height          =   315
               Left            =   300
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   690
               Width           =   2280
            End
            Begin VB.CheckBox chkGenre 
               Caption         =   "Filter by Genre"
               Height          =   255
               Left            =   300
               TabIndex        =   17
               Top             =   435
               Width           =   1335
            End
         End
         Begin VB.OptionButton optDescending 
            Caption         =   "Descending"
            Height          =   255
            Left            =   3360
            TabIndex        =   16
            Top             =   2040
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptAscending 
            Caption         =   "Ascending"
            Height          =   255
            Left            =   960
            TabIndex        =   15
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton cmdTransferToExcel 
            Caption         =   "&Copy Flex Contents to Excel"
            Height          =   1015
            Left            =   3120
            TabIndex        =   24
            Top             =   3650
            Width           =   2295
         End
         Begin VB.CommandButton cmdGenerateStat 
            Caption         =   "&Generate Statistics"
            Height          =   1015
            Left            =   3120
            TabIndex        =   23
            Top             =   2640
            Width           =   2295
         End
         Begin VB.Frame frmPeriod 
            Caption         =   " Period "
            Height          =   1455
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   5160
            Begin MSComCtl2.DTPicker DTPickerEnd 
               Height          =   375
               Left            =   3120
               TabIndex        =   14
               Top             =   600
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               _Version        =   393216
               Format          =   22609921
               CurrentDate     =   37299
            End
            Begin MSComCtl2.DTPicker DTPickerStart 
               Height          =   375
               Left            =   720
               TabIndex        =   13
               Top             =   600
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               _Version        =   393216
               Format          =   22609921
               CurrentDate     =   36526
            End
            Begin VB.Label lblTo 
               Caption         =   "To"
               Height          =   255
               Left            =   2760
               TabIndex        =   28
               Top             =   600
               Width           =   375
            End
            Begin VB.Label lblFrom 
               Caption         =   "From"
               Height          =   375
               Left            =   240
               TabIndex        =   27
               Top             =   600
               Width           =   495
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FlexMovie 
         Height          =   5775
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   10186
         _Version        =   393216
         Cols            =   3
      End
   End
   Begin VB.Frame frameMiscTrans 
      Caption         =   " Ovedue/Misc Transactions "
      Height          =   6375
      Left            =   120
      TabIndex        =   41
      Top             =   360
      Visible         =   0   'False
      Width           =   11385
      Begin VB.Frame Frame1 
         Height          =   1455
         Index           =   1
         Left            =   240
         TabIndex        =   43
         Top             =   4800
         Width           =   10935
         Begin VB.CommandButton cmdMiscTransReport 
            Caption         =   "&Generate Report"
            Height          =   375
            Left            =   4680
            TabIndex        =   54
            Top             =   960
            Width           =   2045
         End
         Begin VB.CommandButton cmdMiscTransFlexToExcel 
            Caption         =   "&Flexgrid To Excel"
            Height          =   375
            Left            =   6720
            TabIndex        =   53
            Top             =   960
            Width           =   2040
         End
         Begin VB.Frame Frame2 
            Caption         =   " Report Period "
            Height          =   1215
            Left            =   120
            TabIndex        =   48
            Top             =   120
            Width           =   4455
            Begin MSComCtl2.DTPicker MiscTransDTPicker2 
               Height          =   375
               Left            =   2760
               TabIndex        =   49
               Top             =   480
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   661
               _Version        =   393216
               Format          =   22609921
               CurrentDate     =   37310
            End
            Begin MSComCtl2.DTPicker MiscTransDTPicker1 
               Height          =   375
               Left            =   840
               TabIndex        =   50
               Top             =   480
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   661
               _Version        =   393216
               Format          =   22609921
               CurrentDate     =   37310
            End
            Begin VB.Label Label4 
               Caption         =   "From"
               Height          =   255
               Left            =   240
               TabIndex        =   52
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label3 
               Caption         =   "To"
               Height          =   255
               Left            =   2400
               TabIndex        =   51
               Top             =   480
               Width           =   255
            End
         End
         Begin VB.CommandButton cmdMiscTransPrint 
            Caption         =   "&Print"
            Height          =   375
            Left            =   8760
            TabIndex        =   47
            Top             =   960
            Width           =   2040
         End
         Begin VB.TextBox txtMiscTransTotal 
            Alignment       =   1  'Right Justify
            Height          =   305
            Left            =   8880
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   "0.00 "
            Top             =   360
            Width           =   1935
         End
         Begin VB.ComboBox cboMiscCashier 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5880
            TabIndex        =   45
            Text            =   "cboMiscCashier"
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox chkMiscTransByCashier 
            Caption         =   "by Cashier:"
            Height          =   255
            Left            =   4800
            TabIndex        =   44
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Total Sales:"
            Height          =   255
            Left            =   7945
            TabIndex        =   55
            Top             =   360
            Width           =   855
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FlexMiscTrans 
         Height          =   4215
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7435
         _Version        =   393216
         Cols            =   6
      End
      Begin VB.Label lblMiscFlex 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   240
         TabIndex        =   56
         Top             =   360
         Width           =   45
      End
   End
   Begin VB.Frame frmIncome 
      Caption         =   " Video Rental Sales Report "
      Height          =   6375
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   11385
      Begin VB.Frame Frame1 
         Height          =   1455
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   4800
         Width           =   10935
         Begin VB.CheckBox chkFilterByCashier 
            Caption         =   "by Cashier:"
            Height          =   255
            Left            =   4800
            TabIndex        =   4
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cboCashier 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5880
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox txtTotalSales 
            Alignment       =   1  'Right Justify
            Height          =   305
            Left            =   8880
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "0.00 "
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmdPrintSales 
            Caption         =   "&Print"
            Height          =   375
            Left            =   8760
            TabIndex        =   9
            Top             =   960
            Width           =   2040
         End
         Begin VB.Frame frameSalesPeriod 
            Caption         =   " Report Period "
            Height          =   1215
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   4455
            Begin MSComCtl2.DTPicker DTSalesEnd 
               Height          =   375
               Left            =   2760
               TabIndex        =   3
               Top             =   480
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   661
               _Version        =   393216
               Format          =   22609921
               CurrentDate     =   37310
            End
            Begin MSComCtl2.DTPicker DTSalesStart 
               Height          =   375
               Left            =   840
               TabIndex        =   2
               Top             =   480
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   661
               _Version        =   393216
               Format          =   22609921
               CurrentDate     =   37310
            End
            Begin VB.Label Label2 
               Caption         =   "To"
               Height          =   255
               Left            =   2400
               TabIndex        =   35
               Top             =   480
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "From"
               Height          =   255
               Left            =   240
               TabIndex        =   34
               Top             =   480
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdFlexSalesToExcel 
            Caption         =   "&Flexgrid To Excel"
            Height          =   375
            Left            =   6720
            TabIndex        =   8
            Top             =   960
            Width           =   2040
         End
         Begin VB.CommandButton cmdGenerateReport 
            Caption         =   "&Generate Report"
            Height          =   375
            Left            =   4680
            TabIndex        =   7
            Top             =   960
            Width           =   2045
         End
         Begin VB.Label lblTotalSales 
            Caption         =   "Total Sales:"
            Height          =   255
            Left            =   7945
            TabIndex        =   36
            Top             =   360
            Width           =   855
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FlexSales 
         Height          =   4215
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7435
         _Version        =   393216
         Cols            =   8
         AllowUserResizing=   3
      End
      Begin VB.Label lblFlexSalesCaption 
         AutoSize        =   -1  'True
         Caption         =   "From :"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   340
         Width           =   435
      End
   End
   Begin VB.Frame frameListofItemsToBeReturnedToday 
      Height          =   6375
      Left            =   125
      TabIndex        =   37
      Top             =   360
      Visible         =   0   'False
      Width           =   11385
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   495
         Left            =   8370
         TabIndex        =   40
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdFlexToExcel 
         Caption         =   "&Flex To Excel"
         Height          =   495
         Left            =   9720
         TabIndex        =   39
         Top             =   5520
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid FlexListOfItemsToBEReturnedToday 
         Height          =   5055
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8916
         _Version        =   393216
         Cols            =   7
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Rental Sales"
            Key             =   "Income"
            Object.ToolTipText     =   "Income"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Overdue/Misc Trans"
            Key             =   "MiscTrans"
            Object.ToolTipText     =   "Overdue/Miscellaneous Transactions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Movie Statistics"
            Key             =   "Movie Statistics"
            Object.ToolTipText     =   "Movie Statistics"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&List of Items to be Returned Today"
            Key             =   "LIRT"
            Object.ToolTipText     =   "List of Items to be Returned Today"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkActor_Click()
If chkActor.Value = 1 Then
    cboActor.Enabled = True
    chkGenre.Value = False
    chkItemCode.Value = False
Else
    cboActor.Enabled = False
End If
End Sub

Private Sub chkFilterByCashier_Click()
If chkFilterByCashier.Value = 1 Then
   cboCashier.Enabled = True
Else
   cboCashier.Enabled = False
End If
End Sub

Private Sub chkGenre_Click()

If chkGenre.Value = 1 Then
    cboGenre.Enabled = True
    chkItemCode.Value = False
    chkActor.Value = False
Else
    cboGenre.Enabled = False
End If
End Sub

Private Sub chkItemCode_Click()
If chkItemCode.Value = 1 Then
   cboItemCode.Enabled = True
   chkGenre.Value = False
   chkActor.Value = False
Else
   cboItemCode.Enabled = False
End If
End Sub

Private Sub chkMiscTransByCashier_Click()
If chkMiscTransByCashier.Value = 1 Then
   cboMiscCashier.Enabled = True
Else
   cboMiscCashier.Enabled = False
End If
End Sub

Private Sub cmdFlexSalesToExcel_Click()
If Trim(FlexSales.TextMatrix(1, 0)) = "" Then
   FlexSales.SetFocus
   Exit Sub
End If
MousePointer = vbHourglass
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Call vr_engine.CopyFlexDataToExcel(FlexSales)
MousePointer = vbDefault
End Sub

Private Sub cmdFlexToExcel_Click()

MousePointer = vbHourglass
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Call vr_engine.CopyFlexDataToExcel(FlexListOfItemsToBEReturnedToday)
MousePointer = vbDefault
FlexListOfItemsToBEReturnedToday.SetFocus

End Sub

Private Sub cmdGenerateReport_Click()
MousePointer = vbHourglass
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE

'START SQL
Dim mySQL As String
mySQL = "SELECT * FROM [Transaction] WHERE (Date >= #" & DTSalesStart.Value & "# AND Date <= #" & DTSalesEnd.Value & "#) ORDER BY Date"
If chkFilterByCashier.Value = 1 And Trim(cboCashier.Text) <> "" Then
mySQL = "SELECT * FROM [Transaction] WHERE Cashier = '" & cboCashier.Text & "' AND (Date >= #" & DTSalesStart.Value & "# AND Date <= #" & DTSalesEnd.Value & "#)ORDER BY Date"
End If
'End SQL


Call vr_engine.Report_GetSalesDetailed(FlexSales, mySQL)

Call FlexSales_TotalSales
FlexSales.SetFocus
MousePointer = vbDefault
End Sub

Private Sub cmdGenerateStat_Click()
MousePointer = vbHourglass
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE

  Dim mySQL As String
  
  ' START SQL
  mySQL = "SELECT DISTINCT Title FROM [CD TAPES TABLE] "
  If chkGenre.Value = 1 Then
     mySQL = "SELECT DISTINCT Title FROM [CD TAPES TABLE] WHERE Genre = '" & cboGenre.Text & "'"
  End If
  If chkItemCode.Value = 1 Then
     mySQL = "SELECT Title FROM [CD TAPES TABLE] WHERE [Item Code] = '" & cboItemCode.Text & "'"
     ''MsgBox Trim(Mid(mySQL, 42, 11)) 'Detect [Item Code]
     ''MsgBox Trim(Mid(mySQL, 57, Len(mySQL) - 57)) 'Detect [Item Code] value
  End If
  If chkActor.Value = 1 Then
     mySQL = "SELECT DISTINCT Title FROM [CD TAPES TABLE] WHERE Actor = '" & vr_engine.ReplaceString(cboActor.Text, "'", "''") & "'"
  End If
  
  ' END SQL
  Call vr_engine.REPORT_GETMOVIESTAT(FlexMovie, DTPickerStart, DTPickerEnd, optDescending, mySQL)
  FlexMovie.SetFocus
  MousePointer = vbDefault
End Sub

Private Sub cmdMiscTransFlexToExcel_Click()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
MousePointer = vbHourglass
If FlexMiscTrans.Rows > 1 Then
    If Trim(FlexMiscTrans.TextMatrix(1, 0)) = "" Then
        FlexMiscTrans.SetFocus
        MousePointer = vbDefault
        Exit Sub
    End If
    Call vr_engine.CopyFlexDataToExcel(FlexMiscTrans)
End If
FlexMiscTrans.SetFocus
MousePointer = vbDefault
End Sub

Private Sub cmdMiscTransPrint_Click()
'Not Yet Done

If Trim(FlexMiscTrans.TextMatrix(1, 0)) = "" Then
   FlexMiscTrans.SetFocus
   Exit Sub
End If
If MsgBox("Please insert 8 1/2"" by 13""  paper. ", vbOKCancel, "Insert Paper ") = vbCancel Then
       FlexMiscTrans.SetFocus
       Exit Sub
End If
MousePointer = vbHourglass
  Dim c1, BlankLines, LoopBlankLines, AmntPos, PageCount, PrintedPageCount As Integer
  Dim loop1, loopPrintPage, Flag As Long
    c1 = 8
    Printer.Font = "Sans Serif"
    Printer.PaperSize = vbPRPSLegal ' 8.5 by 14 inc
    Printer.FontSize = 9
    Printer.Orientation = 2 'LandScape
    
    'Count No. of Pages
    If (FlexMiscTrans.Rows - 1) Mod 35 = 0 Then
        PageCount = (FlexMiscTrans.Rows - 1) / 35
    Else
        PageCount = Int((FlexMiscTrans.Rows - 1) / 35) + 1
    End If
    PrintedPageCount = 0
 '''''''''''''''' START PRINTING
 For loopPrintPage = 1 To PageCount
    'START Report Header
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(c1); "Overdue/Misc Sales Report - " & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    Printer.Print Tab(c1); "Period : " & Format(MiscTransDTPicker1.Value, "mm/dd/yyyy") & " to " & Format(MiscTransDTPicker2.Value, "mm/dd/yyyy") & "  " & "TOTAL SALES   P " & Trim(txtMiscTransTotal.Text)
    Printer.Print ""
    Printer.Print Tab(c1); "No."; Tab(c1 + 16); "Invoice No. "; Tab(c1 + 36); "Date" _
                  ; Tab(c1 + 55); "Cashier"; Tab(c1 + 106); "Description" _
                  ; Tab(c1 + 179); "Amount" '''; Tab(c1 + 167); "Title"; Tab(c1 + 216); "Amount"
    Printer.Print Tab(c1); "======="; Tab(c1 + 16); "========="; Tab(c1 + 36); "========" _
                  ; Tab(c1 + 55); "=========================="; Tab(c1 + 106); "==================================" _
                  ; Tab(c1 + 173); "============" '''; Tab(c1 + 167); "========================"; Tab(c1 + 216); "======"
    ' END REPORT HEADER
    
    ' START DETAILED SECTION
    
    For loop1 = 1 To 35
        Flag = Flag + 1
        AmntPos = (c1 + 171) - Len(FlexMiscTrans.TextMatrix(Flag, 5))
        Printer.Print Tab(c1); FlexMiscTrans.TextMatrix(Flag, 0); Tab(c1 + 16); FlexMiscTrans.TextMatrix(Flag, 1); Tab(c1 + 36); Format(FlexMiscTrans.TextMatrix(Flag, 2), "mm/dd/yyyy") _
                  ; Tab(c1 + 55); FlexMiscTrans.TextMatrix(Flag, 3); Tab(c1 + 106); FlexMiscTrans.TextMatrix(Flag, 4) _
                  ; Tab(c1 + 190 - Len(FlexMiscTrans.TextMatrix(Flag, 5))); FlexMiscTrans.TextMatrix(Flag, 5) '; Tab(c1 + 167); FlexMiscTrans.TextMatrix(loop1, 6); Tab(AmntPos); FlexMiscTrans.TextMatrix(Flag, 7)
        If Flag = FlexMiscTrans.Rows - 1 Then
            Printer.Print Tab(c1 + 172); "-------------------------"
            AmntPos = (c1 + 177) - Len("TOTAL SALES   " & Trim(txtMiscTransTotal.Text))
            Printer.Print Tab(AmntPos); "TOTAL SALES   " & Trim(txtMiscTransTotal.Text)
            If Flag Mod 35 <> 0 Then
               BlankLines = 35 - Flag Mod 35
               For LoopBlankLines = 1 To BlankLines
                   Printer.Print "" 'Print Blank line
               Next
            End If
            Exit For
        End If
    Next loop1
    ' END DETAILED SECTION
    
    ' START FOOTER
    PrintedPageCount = PrintedPageCount + 1
    If Flag < FlexMiscTrans.Rows - 1 Then Printer.Print ""
    If Flag < FlexMiscTrans.Rows - 1 Then Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(c1 + 211); "Page " & str(PrintedPageCount) & " of " & str(PageCount)
    ' END FOOTER
    Printer.NewPage
Next loopPrintPage
    Printer.EndDoc
'''''''''''''''' END PRINTING

MousePointer = vbDefault

FlexMiscTrans.SetFocus
End Sub

Private Sub cmdMiscTransReport_Click()
MousePointer = vbHourglass
Dim mySQL As String
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
'START SQL

mySQL = "SELECT * FROM [Transaction] WHERE (Date >= #" & MiscTransDTPicker1.Value & "# AND Date <= #" & MiscTransDTPicker2.Value & "#) ORDER BY Date"
If chkMiscTransByCashier.Value = 1 And Trim(cboMiscCashier.Text) <> "" Then
mySQL = "SELECT * FROM [Transaction] WHERE Cashier = '" & cboMiscCashier.Text & "' AND (Date >= #" & MiscTransDTPicker1.Value & "# AND Date <= #" & MiscTransDTPicker2.Value & "#)ORDER BY Date"
End If
'End SQL
Call vr_engine.Report_GetMiscTrans(FlexMiscTrans, mySQL)
' START -- Calulate Total
Dim Total As Double
Dim loop1 As Long
Total = 0
For loop1 = 1 To FlexMiscTrans.Rows - 1
    Total = Total + Val(FlexMiscTrans.TextMatrix(loop1, 5))
Next loop1
txtMiscTransTotal.Text = Format(Total, "0.00")
' END -- Calculate Total
MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
'--------------------------------------------
If Trim(FlexMovie.TextMatrix(1, 0)) = "" Then
   FlexMovie.SetFocus
   Exit Sub
End If
'--------------------------------------------
'--------------------------------------------
If MsgBox("Please insert 8 1/2"" by 13""  paper. ", vbOKCancel, "Insert Paper ") = vbCancel Then
       FlexMovie.SetFocus
       Exit Sub
End If
'--------------------------------------------
MousePointer = vbHourglass
Printer.Font = "Sans Serif"
    Printer.PaperSize = vbPRPSLegal ' 8.5 by 14 inc
    Printer.FontSize = 12
    Printer.Orientation = 1 'Portrait
    
Dim LeftMargin, PageCount, BlankLines As Integer
Dim loop1, loop2, Lines, Flag As Long
LeftMargin = 10

Lines = FlexMovie.Rows - 1
If Lines Mod 55 = 0 Then
   PageCount = Lines / 55
Else
   PageCount = Int(Lines / 55) + 1
End If
'START - HEADING
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print Tab(LeftMargin); "MOVIE STATISTICS FROM " & UCase(Format(DTPickerStart.Value, "mmm. dd, yyyy")) & " TO " & UCase(Format(DTPickerEnd, "mmm. dd, yyyy"))
Printer.Print ""
Printer.Print Tab(LeftMargin); "No."; Tab(LeftMargin + 9); "TITLE"; Tab(LeftMargin + 45); "FREQUENCY"
Printer.Print Tab(LeftMargin); "======"; Tab(LeftMargin + 9); "=========================="; Tab(LeftMargin + 45); "=========="

'END - HEADING


For loop1 = 1 To PageCount
  For loop2 = 1 To 55
    Flag = Flag + 1
    Printer.Print Tab(LeftMargin); FlexMovie.TextMatrix(Flag, 0); Tab(LeftMargin + 9); FlexMovie.TextMatrix(Flag, 1); Tab(LeftMargin + 45); FlexMovie.TextMatrix(Flag, 2)
    If Flag = Lines Then
       If Lines Mod 55 > 0 Then
            For BlankLines = (Lines Mod 55) To 55
               Printer.Print ""
            Next BlankLines
       End If
       Exit For
    End If
  Next loop2

' START - FOOTER
Printer.Print ""
Printer.Print Tab(106 - Len("Page " & str(loop1) & " of " & str(PageCount))); "Page " & str(loop1) & " of " & str(PageCount)
' END - FOOTER
Next loop1
Printer.EndDoc
FlexMovie.SetFocus
MousePointer = vbDefault
End Sub

Private Sub cmdPrintSales_Click()
If Trim(FlexSales.TextMatrix(1, 0)) = "" Then
   FlexSales.SetFocus
   Exit Sub
End If
If MsgBox("Please insert 8 1/2"" by 13""  paper. ", vbOKCancel, "Insert Paper ") = vbCancel Then
       FlexSales.SetFocus
       Exit Sub
End If
MousePointer = vbHourglass
  Dim c1, BlankLines, LoopBlankLines, AmntPos, PageCount, PrintedPageCount As Integer
  Dim loop1, loopPrintPage, Flag As Long
    c1 = 8
    Printer.Font = "Sans Serif"
    Printer.PaperSize = vbPRPSLegal ' 8.5 by 14 inc
    Printer.FontSize = 9
    Printer.Orientation = 2 'LandScape
    
    'Count No. of Pages
    If (FlexSales.Rows - 1) Mod 35 = 0 Then
        PageCount = (FlexSales.Rows - 1) / 35
    Else
        PageCount = Int((FlexSales.Rows - 1) / 35) + 1
    End If
    PrintedPageCount = 0
 '''''''''''''''' START PRINTING
 For loopPrintPage = 1 To PageCount
    'START Report Header
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(c1); "Rental Sales Report - " & Format(Now, "mm/dd/yyyy hh:mm:ss AM/PM")
    Printer.Print Tab(c1); "Rental Period : " & Format(DTSalesStart.Value, "mm/dd/yyyy") & " to " & Format(DTSalesEnd.Value, "mm/dd/yyyy") & "  " & "TOTAL SALES   P " & Trim(txtTotalSales.Text)
    Printer.Print ""
    Printer.Print Tab(c1); "Item No."; Tab(c1 + 16); "Invoice No. "; Tab(c1 + 36); "Date" _
                  ; Tab(c1 + 55); "Cashier"; Tab(c1 + 96); "Borrower's Name" _
                  ; Tab(c1 + 141); "Item Code"; Tab(c1 + 167); "Title"; Tab(c1 + 216); "Amount"
    Printer.Print Tab(c1); "======="; Tab(c1 + 16); "========="; Tab(c1 + 36); "========" _
                  ; Tab(c1 + 55); "===================="; Tab(c1 + 96); "======================" _
                  ; Tab(c1 + 141); "============"; Tab(c1 + 167); "========================"; Tab(c1 + 216); "======"
    ' END REPORT HEADER
    
    ' START DETAILED SECTION
    
    For loop1 = 1 To 35
        Flag = Flag + 1
        AmntPos = (c1 + 224) - Len(FlexSales.TextMatrix(Flag, 7))
        Printer.Print Tab(c1); FlexSales.TextMatrix(Flag, 0); Tab(c1 + 16); FlexSales.TextMatrix(Flag, 1); Tab(c1 + 36); Format(FlexSales.TextMatrix(Flag, 2), "mm/dd/yyyy") _
                  ; Tab(c1 + 55); FlexSales.TextMatrix(Flag, 3); Tab(c1 + 96); FlexSales.TextMatrix(Flag, 4) _
                  ; Tab(c1 + 141); FlexSales.TextMatrix(Flag, 5); Tab(c1 + 167); FlexSales.TextMatrix(loop1, 6); Tab(AmntPos); FlexSales.TextMatrix(Flag, 7)
        If Flag = FlexSales.Rows - 1 Then
            Printer.Print Tab(c1 + 211); "--------------------"
            AmntPos = (c1 + 213) - Len("TOTAL SALES   " & Trim(txtTotalSales.Text))
            Printer.Print Tab(AmntPos); "TOTAL SALES   " & Trim(txtTotalSales.Text)
            If Flag Mod 35 <> 0 Then
               BlankLines = 35 - Flag Mod 35
               For LoopBlankLines = 1 To BlankLines
                   Printer.Print "" 'Print Blank line
               Next
            End If
            Exit For
        End If
    Next loop1
    ' END DETAILED SECTION
    
    ' START FOOTER
    PrintedPageCount = PrintedPageCount + 1
    If Flag < FlexSales.Rows - 1 Then Printer.Print ""
    If Flag < FlexSales.Rows - 1 Then Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(c1 + 211); "Page " & str(PrintedPageCount) & " of " & str(PageCount)
    ' END FOOTER
    Printer.NewPage
Next loopPrintPage
    Printer.EndDoc
'''''''''''''''' END PRINTING
FlexSales.SetFocus
MousePointer = vbDefault
End Sub

Private Sub cmdRefresh_Click()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Call InitializeFlexItemsToBeReturnedToday
MousePointer = vbHourglass
Call vr_engine.Report_LoadItemsToBeReturnedToday(FlexListOfItemsToBEReturnedToday)
MousePointer = vbDefault
End Sub

Private Sub cmdTransferToExcel_Click()
If Trim(FlexMovie.TextMatrix(1, 0) = "") Then
   FlexMovie.SetFocus
   Exit Sub
End If
MousePointer = vbHourglass
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Call vr_engine.CopyFlexDataToExcel(FlexMovie)
MousePointer = vbDefault
FlexMovie.SetFocus
End Sub

Private Sub FlexSales_TotalSales()
Dim Total As Double
Dim loop1 As Long
Total = 0
For loop1 = 1 To FlexSales.Rows - 1
    Total = Total + Val(FlexSales.TextMatrix(loop1, 7))
Next loop1
txtTotalSales.Text = Total
End Sub


Private Sub DTSalesEnd_Change()
lblFlexSalesCaption.Caption = "From: " & Format(DTSalesStart.Value, "mmmm dd, yyyy") & " to " & Format(DTSalesEnd.Value, "mmmm dd, yyyy")
End Sub

Private Sub DTSalesStart_Change()
lblFlexSalesCaption.Caption = "From: " & Format(DTSalesStart.Value, "mmmm dd, yyyy") & " to " & Format(DTSalesEnd.Value, "mmmm dd, yyyy")
End Sub

Private Sub Form_Activate()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
If FlexSales.Visible = True Then
   MousePointer = vbHourglass
   Call InitializeFlexSalesDetiled
   Call vr_engine.REPORT_GETMOVIESTAT_FillCBOcashier(cboCashier)
   MousePointer = vbDefault
End If
End Sub

Private Sub MiscTransDTPicker1_Change()
lblMiscFlex.Caption = "From " & Format(MiscTransDTPicker1.Value, "mmmm dd, yyyy") & " To " & Format(MiscTransDTPicker2.Value, "mmmm dd, yyyy")
End Sub


Private Sub MiscTransDTPicker2_Change()
lblMiscFlex.Caption = "From " & Format(MiscTransDTPicker1.Value, "mmmm dd, yyyy") & " To " & Format(MiscTransDTPicker2.Value, "mmmm dd, yyyy")
End Sub

Private Sub OptAscending_Click()
If FlexMovie.Rows > 1 Then
  If Trim(FlexMovie.TextMatrix(1, 0)) <> "" Then Call cmdGenerateStat_Click
End If
End Sub

Private Sub optDescending_Click()
If FlexMovie.Rows > 1 Then
  If Trim(FlexMovie.TextMatrix(1, 0)) <> "" Then Call cmdGenerateStat_Click
End If
End Sub

Private Sub TabStrip1_Click()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
'-------------------------------------------------
If TabStrip1.SelectedItem.Caption = "&Rental Sales" Then
    frmMovieStatistics.Visible = False
    frameListofItemsToBeReturnedToday.Visible = False
    frameMiscTrans.Visible = False
    frmIncome.Visible = True
    FlexSales.SetFocus
End If
'-------------------------------------------------
If TabStrip1.SelectedItem.Caption = "&Movie Statistics" Then
    frmIncome.Visible = False
    frameMiscTrans.Visible = False
    frameListofItemsToBeReturnedToday.Visible = False
    frmMovieStatistics.Visible = True
    Call vr_engine.REPORT_MOVIESTAT_FILLCBO(cboGenre, cboItemCode, cboActor)
    FlexMovie.SetFocus
End If
'-------------------------------------------------
If TabStrip1.SelectedItem.Caption = "&List of Items to be Returned Today" Then
    frmIncome.Visible = False
    frameMiscTrans.Visible = False
    frmMovieStatistics.Visible = False
    frameListofItemsToBeReturnedToday.Visible = True
    Call InitializeFlexItemsToBeReturnedToday
    MousePointer = vbHourglass
    Call vr_engine.Report_LoadItemsToBeReturnedToday(FlexListOfItemsToBEReturnedToday)
    MousePointer = vbDefault
End If
'-------------------------------------------------
If TabStrip1.SelectedItem.Caption = "Overdue/Misc Trans" Then
    frmIncome.Visible = False
    frmMovieStatistics.Visible = False
    frameListofItemsToBeReturnedToday.Visible = False
    frameMiscTrans.Visible = True
    Call InittializeFlexMiscTrans
    Call vr_engine.REPORT_GETMOVIESTAT_FillCBOcashier(cboMiscCashier)
    MiscTransDTPicker1.Value = Format(Now, "mm/dd/yyyy")
    MiscTransDTPicker2.Value = Format(Now, "mm/dd/yyyy")
    lblMiscFlex.Caption = "From " & Format(MiscTransDTPicker1.Value, "mmmm dd, yyyy") & " To " & Format(MiscTransDTPicker1.Value, "mmmm dd, yyyy")
    FlexMiscTrans.SetFocus
End If
'-------------------------------------------------
Dim StartDate As String
StartDate = "1/1/" & Format(Now, "yyyy")
DTPickerStart.Value = StartDate
DTPickerEnd.Value = Format(Now, "mm/dd/yyyy")

Call InitializeFlexMovie
Call InitializeFlexSalesDetiled
End Sub
Private Sub InitializeFlexItemsToBeReturnedToday()
    MousePointer = vbHourglass
    If Trim(FlexListOfItemsToBEReturnedToday.TextMatrix(0, 0)) = "" Then
       FlexListOfItemsToBEReturnedToday.ColWidth(0) = 600  ' No.
       FlexListOfItemsToBEReturnedToday.ColWidth(1) = 1200 ' Item Code
       FlexListOfItemsToBEReturnedToday.ColWidth(2) = 2900 ' Title
       FlexListOfItemsToBEReturnedToday.ColWidth(3) = 1300 ' Date Borrowed
       FlexListOfItemsToBEReturnedToday.ColWidth(4) = 1300 ' Date Due
       FlexListOfItemsToBEReturnedToday.ColWidth(5) = 1500 ' Overdue Charges
       FlexListOfItemsToBEReturnedToday.ColWidth(6) = 2900 ' Borrowed By
       
       FlexListOfItemsToBEReturnedToday.ColAlignment(3) = 5
       FlexListOfItemsToBEReturnedToday.ColAlignment(4) = 5
       
       FlexListOfItemsToBEReturnedToday.TextMatrix(0, 0) = "No."
       FlexListOfItemsToBEReturnedToday.TextMatrix(0, 1) = "Item Code"
       FlexListOfItemsToBEReturnedToday.TextMatrix(0, 2) = "Title"
       FlexListOfItemsToBEReturnedToday.TextMatrix(0, 3) = "Date Borrowed"
       FlexListOfItemsToBEReturnedToday.TextMatrix(0, 4) = "Date Due"
       FlexListOfItemsToBEReturnedToday.TextMatrix(0, 5) = "Overdue Charges"
       FlexListOfItemsToBEReturnedToday.TextMatrix(0, 6) = "Remarks"
    End If
    
    FlexListOfItemsToBEReturnedToday.SetFocus
    MousePointer = vbDefault
End Sub
Private Sub InitializeFlexMovie()
   If FlexMovie.TextMatrix(0, 0) = "" Then
        FlexMovie.ColWidth(0) = 550
        FlexMovie.ColWidth(1) = 2700
        FlexMovie.ColWidth(2) = 1200
        FlexMovie.ColAlignment(0) = 5
     '' FlexMovie.ColAlignment(1) = 5
        FlexMovie.ColAlignment(2) = 5
        FlexMovie.TextMatrix(0, 0) = "No."
        FlexMovie.TextMatrix(0, 1) = "Title"
        FlexMovie.TextMatrix(0, 2) = "Frequency"
    'Start - Clear 2nd row
        FlexMovie.TextMatrix(1, 0) = ""
        FlexMovie.TextMatrix(1, 1) = ""
        FlexMovie.TextMatrix(1, 2) = ""
    'End - Clear 2nd row
   End If
End Sub
Sub InittializeFlexMiscTrans()
    FlexMiscTrans.ColWidth(0) = 700 'No
    FlexMiscTrans.ColWidth(1) = 1600 ' Invoice Number
    FlexMiscTrans.ColWidth(2) = 1400 ' Date
    FlexMiscTrans.ColWidth(3) = 2400 ' Cashier
    FlexMiscTrans.ColWidth(4) = 2800 ' Description
    FlexMiscTrans.ColWidth(5) = 1100 ' Amount
    
    FlexMiscTrans.ColAlignment(0) = 1
    FlexMiscTrans.ColAlignment(1) = 1
    

    FlexMiscTrans.TextMatrix(0, 0) = "No."
    FlexMiscTrans.TextMatrix(0, 1) = "Invoice Number"
    FlexMiscTrans.TextMatrix(0, 2) = "Date"
    FlexMiscTrans.TextMatrix(0, 3) = "Cashier"
    FlexMiscTrans.TextMatrix(0, 4) = "Description"
    FlexMiscTrans.TextMatrix(0, 5) = "      Amount"
    
End Sub
Sub InitializeFlexSalesDetiled()
If FlexSales.TextMatrix(0, 0) = "" Then
        FlexSales.ColWidth(0) = 500
        FlexSales.ColWidth(1) = 1400
        FlexSales.ColWidth(2) = 1300
        FlexSales.ColWidth(3) = 2150
        FlexSales.ColWidth(4) = 2150
        FlexSales.ColWidth(5) = 1500
        FlexSales.ColWidth(6) = 2200
        FlexSales.ColWidth(7) = 800
        FlexSales.ColAlignment(0) = 1
        FlexSales.ColAlignment(1) = 1
        FlexSales.ColAlignment(7) = 6
        FlexSales.TextMatrix(0, 0) = "No."
        FlexSales.TextMatrix(0, 1) = "Invoice No."
        FlexSales.TextMatrix(0, 2) = "Date"
        FlexSales.TextMatrix(0, 3) = "Cashier"
        FlexSales.TextMatrix(0, 4) = "Borrower' Name"
        FlexSales.TextMatrix(0, 5) = "Item Code"
        FlexSales.TextMatrix(0, 6) = "Title"
        FlexSales.TextMatrix(0, 7) = " Amount "
    'Start - Clear 2nd row
        FlexSales.TextMatrix(1, 0) = ""
        FlexSales.TextMatrix(1, 1) = ""
        FlexSales.TextMatrix(1, 2) = ""
    'End - Clear 2nd row
    DTSalesStart.Value = Format(Now, "mm/dd/yyyy")
    DTSalesEnd.Value = Format(Now, "mm/dd/yyyy")
    lblFlexSalesCaption.Caption = "From: " & Format(DTSalesStart.Value, "mmmm dd, yyyy") & " to " & Format(DTSalesEnd.Value, "mmmm dd, yyyy")
    FlexSales.SetFocus
   End If

End Sub

Private Sub txtTotalSales_Change()
txtTotalSales.Text = Format(txtTotalSales.Text, "0.00")
End Sub
