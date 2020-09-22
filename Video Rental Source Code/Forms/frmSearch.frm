VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   11610
   Begin VB.Frame frameUnreturnedItems 
      Caption         =   " Unreturned Items "
      Height          =   6135
      Left            =   120
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   11295
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   240
         TabIndex        =   33
         Top             =   5280
         Width           =   10815
         Begin VB.CommandButton cmdFlexToExcel 
            Caption         =   "&Copy FlexGrid Contents to Excel"
            Height          =   375
            Left            =   8160
            TabIndex        =   39
            Top             =   210
            Width           =   2535
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
            Height          =   375
            Left            =   5640
            TabIndex        =   38
            Top             =   210
            Width           =   2535
         End
         Begin VB.OptionButton OptUnreturnedItemsDesc 
            Caption         =   "Descending"
            Height          =   255
            Left            =   4200
            TabIndex        =   37
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OptUnreturnedItemsAsc 
            Caption         =   "Ascending"
            Height          =   255
            Left            =   3000
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.ComboBox cboUnreturnedItemsSortBy 
            Height          =   315
            ItemData        =   "frmSearch.frx":0742
            Left            =   960
            List            =   "frmSearch.frx":074F
            TabIndex        =   35
            Text            =   "LastDateBorrowed"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "Sort by: "
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Width           =   615
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4695
         Left            =   240
         TabIndex        =   32
         Top             =   480
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   8
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame frameMovies 
      Caption         =   "Movies"
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   11295
      Begin VB.CommandButton cmdFlexMovieToExcel 
         Caption         =   "&Flex to Excel"
         Height          =   615
         Left            =   9720
         TabIndex        =   40
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Frame frameMoviesFilter 
         Caption         =   " Filter "
         Height          =   1335
         Left            =   240
         TabIndex        =   12
         Top             =   4680
         Width           =   5535
         Begin VB.OptionButton optMoviesUsePatternMatching 
            Caption         =   "Use pattern matching"
            Height          =   255
            Left            =   3480
            TabIndex        =   6
            Top             =   840
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton optMoviesWholeWord 
            Caption         =   "Find whole word only"
            Height          =   255
            Left            =   960
            TabIndex        =   5
            Top             =   840
            Width           =   1815
         End
         Begin VB.ComboBox cboMovieSearch 
            Height          =   315
            ItemData        =   "frmSearch.frx":0777
            Left            =   3600
            List            =   "frmSearch.frx":078D
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtMovieSearch 
            Height          =   285
            Left            =   960
            TabIndex        =   3
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "from:"
            Height          =   255
            Left            =   3120
            TabIndex        =   15
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Search:"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame frameMovieSort 
         Caption         =   " Sort "
         Height          =   1335
         Left            =   5880
         TabIndex        =   11
         Top             =   4680
         Width           =   3735
         Begin VB.OptionButton OptMoviesSortDecending 
            Caption         =   "Descending"
            Height          =   255
            Left            =   2160
            TabIndex        =   9
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optMovieSortAscending 
            Caption         =   "Ascending"
            Height          =   375
            Left            =   600
            TabIndex        =   8
            Top             =   790
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.ComboBox cboMoviesSort 
            Height          =   315
            ItemData        =   "frmSearch.frx":07CE
            Left            =   1200
            List            =   "frmSearch.frx":07E1
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Sort by"
            Height          =   255
            Left            =   600
            TabIndex        =   13
            Top             =   270
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdMovies_Refresh 
         Caption         =   "&Refresh"
         Height          =   615
         Left            =   9720
         TabIndex        =   10
         Top             =   4800
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid FlexMovies 
         Height          =   4095
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   8
         FillStyle       =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame frameMembers 
      Caption         =   "Members"
      Height          =   6135
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   11295
      Begin VB.CommandButton cmdFleMembersToExcel 
         Caption         =   "&Flex to Excel"
         Height          =   615
         Left            =   9720
         TabIndex        =   41
         Top             =   5400
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid FlexMembers 
         Height          =   4095
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   15
         WordWrap        =   -1  'True
         AllowUserResizing=   1
      End
      Begin VB.Frame Frame1 
         Caption         =   " Sort "
         Height          =   1335
         Left            =   5880
         TabIndex        =   25
         Top             =   4680
         Width           =   3735
         Begin VB.OptionButton optMembersSortDecending 
            Caption         =   "Descending"
            Height          =   255
            Left            =   2280
            TabIndex        =   29
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optMembersSortAscending 
            Caption         =   "Ascending"
            Height          =   255
            Left            =   720
            TabIndex        =   28
            Top             =   840
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.ComboBox cboMembersSortField 
            Height          =   315
            ItemData        =   "frmSearch.frx":0817
            Left            =   1200
            List            =   "frmSearch.frx":0845
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "Sort by"
            Height          =   255
            Left            =   480
            TabIndex        =   26
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame frameMembersSearchFilter 
         Caption         =   " Filter "
         Height          =   1335
         Left            =   240
         TabIndex        =   18
         Top             =   4680
         Width           =   5535
         Begin VB.OptionButton optUsePatternMatching 
            Caption         =   "Use pattern matching"
            Height          =   375
            Left            =   3600
            TabIndex        =   24
            Top             =   840
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optMembersFindWholeWordOnly 
            Caption         =   "Find whole word only"
            Height          =   375
            Left            =   840
            TabIndex        =   23
            Top             =   840
            Width           =   1815
         End
         Begin VB.ComboBox cboMembersSearchField 
            Height          =   315
            ItemData        =   "frmSearch.frx":08F8
            Left            =   3600
            List            =   "frmSearch.frx":0923
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtMembersSearchString 
            Height          =   285
            Left            =   840
            TabIndex        =   19
            Text            =   "[All Names]"
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label5 
            Caption         =   "from"
            Height          =   255
            Left            =   3240
            TabIndex        =   21
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Search:"
            Height          =   255
            Left            =   210
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdRefreshMembers 
         Caption         =   "&Refresh"
         Height          =   615
         Left            =   9720
         TabIndex        =   17
         Top             =   4800
         Width           =   1335
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11668
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Movies"
            Key             =   "Movie"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Mem&bers"
            Key             =   "Members"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Unreturned &Items"
            Key             =   "UnreturnedItems"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cboMembersSearchField_Click()
If cboMembersSearchField.Text = "Date Entered" Or cboMembersSearchField.Text = "Birthday" Then
    optMembersFindWholeWordOnly.Enabled = False
    optUsePatternMatching.Value = True
Else
    optMembersFindWholeWordOnly.Enabled = True
End If
End Sub

Private Sub cboMovieSearch_click()
If cboMovieSearch.Text = "Date Entered" Then
   optMoviesWholeWord.Enabled = False
   optMoviesUsePatternMatching.Value = True
Else
   optMoviesWholeWord.Enabled = True
End If
End Sub

Private Sub cmdFleMembersToExcel_Click()
MousePointer = vbHourglass
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Call vr_engine.CopyFlexDataToExcel(FlexMembers)
MousePointer = vbDefault
FlexMembers.SetFocus
End Sub

Private Sub cmdFlexMovieToExcel_Click()
MousePointer = vbHourglass
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Call vr_engine.CopyFlexDataToExcel(FlexMovies)
MousePointer = vbDefault
FlexMovies.SetFocus
End Sub

Private Sub cmdFlexToExcel_Click()
MousePointer = vbHourglass
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Call vr_engine.CopyFlexDataToExcel(MSFlexGrid1)
MousePointer = vbDefault
MSFlexGrid1.SetFocus
End Sub

Private Sub cmdMovies_Refresh_Click()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Call vr_engine.Search_Movies(FlexMovies, Trim(txtMovieSearch.Text), cboMovieSearch.Text, optMoviesUsePatternMatching.Value, cboMoviesSort.Text, optMovieSortAscending.Value)
FlexMovies.SetFocus
End Sub


Private Sub cmdRefresh_Click()
    MousePointer = vbHourglass
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE
    Call vr_engine.Report_LoadUnreturnedItems(MSFlexGrid1, UnreturnedItems_SQL())
    MousePointer = vbDefault
    MSFlexGrid1.SetFocus
End Sub

Private Sub cmdRefreshMembers_Click()
    Dim vr_engine As VRENTAL_ENGINE
    Set vr_engine = New VRENTAL_ENGINE

    Call vr_engine.Search_Members(FlexMembers, Trim(txtMembersSearchString.Text), Trim(cboMembersSearchField.Text), optUsePatternMatching.Value, Trim(cboMembersSortField.Text), optMembersSortAscending.Value)
    'Call vr_engine.Search_Movies(FlexMembers, Trim(txtMovieSearch.Text), cboMovieSearch.Text, optMoviesUsePatternMatching.Value, cboMoviesSort.Text, optMovieSortAscending.Value)
    If FlexMembers.Visible = True Then FlexMembers.SetFocus

End Sub

Private Sub Command1_Click()

MSFlexGrid1.SetFocus
End Sub

Private Sub Form_Activate()
If FlexMovies.Visible = True Then FlexMovies.SetFocus
End Sub

Private Sub Form_Load()
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
Call vr_engine.Search_Movies(FlexMovies, "", "Title", True, "Title", True)

' initialize controls
  
  txtMovieSearch.Text = "[All Movies]"
  cboMovieSearch.Text = "Title"
  cboMoviesSort.Text = "Title"
  
  cboMembersSearchField.Text = "Family Name"
  cboMembersSortField.Text = "Family Name"
  
' end initialize controls
End Sub




Private Sub OptUnreturnedItemsAsc_Click()
 Call cmdRefresh_Click
End Sub

Private Sub OptUnreturnedItemsDesc_Click()
Call cmdRefresh_Click
End Sub

Private Sub TabStrip1_Click()
'--------------------------------------------------
If TabStrip1.SelectedItem.Caption = "&Movies" Then
    frameMovies.Visible = True
    frameMembers.Visible = False
    frameUnreturnedItems.Visible = False
    FlexMovies.SetFocus
End If
'--------------------------------------------------
'--------------------------------------------------
If TabStrip1.SelectedItem.Caption = "Mem&bers" Then
    frameMovies.Visible = False
    frameUnreturnedItems.Visible = False
    frameMembers.Visible = True
    Call cmdRefreshMembers_Click


End If
'-------------------------------------------------
'-------------------------------------------------
If TabStrip1.SelectedItem.Caption = "Unreturned &Items" Then
    frameMovies.Visible = False
    frameUnreturnedItems.Visible = True
    frameMembers.Visible = False
    MSFlexGrid1.SetFocus
    Call UnreturnedItems_initializeFLEXGRID(MSFlexGrid1)
    
End If
'-------------------------------------------------
End Sub
Sub UnreturnedItems_initializeFLEXGRID(Flex As MSFlexGrid)
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
     
     
End Sub

Function UnreturnedItems_SQL() As String
  If OptUnreturnedItemsAsc = True Then
    UnreturnedItems_SQL = "SELECT * FROM [CD TAPES TABLE] WHERE Available = 'No' ORDER BY [" & Trim(cboUnreturnedItemsSortBy.Text) & "]"
  Else
    UnreturnedItems_SQL = "SELECT * FROM [CD TAPES TABLE] WHERE Available = 'No' ORDER BY [" & Trim(cboUnreturnedItemsSortBy.Text) & "] Desc"
  End If
End Function
