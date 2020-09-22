VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Video Rental Information Console"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   Icon            =   "MDIfrmVideoInfo.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1650
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   2910
      ButtonWidth     =   2514
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Transaction (Rent)"
            Key             =   "Transaction"
            Object.ToolTipText     =   "Transaction"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            Key             =   "Report"
            Object.ToolTipText     =   "Report"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Info"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Items"
            Key             =   "AddCD_Tapes"
            Object.ToolTipText     =   "Items"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Members"
            Key             =   "Members"
            Object.ToolTipText     =   "Members"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Users"
            Key             =   "Users"
            Object.ToolTipText     =   "Users"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Return Items"
            Key             =   "ReturnItems"
            Object.ToolTipText     =   "Return Items"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmVideoInfo.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmVideoInfo.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmVideoInfo.frx":0BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmVideoInfo.frx":1000
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmVideoInfo.frx":1452
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmVideoInfo.frx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmVideoInfo.frx":1CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIfrmVideoInfo.frx":2148
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
MDIForm1.WindowState = 2
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim vr_engine As VRENTAL_ENGINE
Set vr_engine = New VRENTAL_ENGINE
On Error Resume Next

Select Case Button.Key

Case "Users"
  If vr_engine.CheckPermission(Val(gVarAccessLevel), 6) = False Then Exit Sub
  If frmUser.WindowState = vbMinimized Then frmUser.WindowState = vbNormal: Exit Sub
  frmUser.Show
  frmUser.SetFocus
  frmUser.Left = 0
  frmUser.Top = 0

Case "Members"
  If vr_engine.CheckPermission(Val(gVarAccessLevel), 5) = False Then Exit Sub
 If frmMembers.WindowState = vbMinimized Then frmMembers.WindowState = vbNormal: Exit Sub
  frmMembers.Show
  frmMembers.SetFocus
  frmMembers.Left = 0
  frmMembers.Top = 0
  

Case "AddCD_Tapes"
  If vr_engine.CheckPermission(Val(gVarAccessLevel), 4) = False Then Exit Sub
If frmCDtapes.WindowState = vbMinimized Then frmCDtapes.WindowState = vbNormal: Exit Sub
  frmCDtapes.Show
  frmCDtapes.SetFocus
  frmCDtapes.Left = 0
  frmCDtapes.Top = 0

Case "Info"
  If vr_engine.CheckPermission(Val(gVarAccessLevel), 3) = False Then Exit Sub
If frmSearch.WindowState = vbMinimized Then frmSearch.WindowState = vbNormal: Exit Sub
  frmSearch.Show
  frmSearch.SetFocus
  frmSearch.Left = 0
  frmSearch.Top = 0
  
Case "Transaction"
  If vr_engine.CheckPermission(Val(gVarAccessLevel), 1) = False Then Exit Sub
If frmTransaction.WindowState = vbMinimized Then frmTransaction.WindowState = vbNormal: Exit Sub
 frmTransaction.Show
 frmTransaction.SetFocus
 frmTransaction.Left = 0
 frmTransaction.Top = 0
 
Case "ReturnItems"
If vr_engine.CheckPermission(Val(gVarAccessLevel), 7) = False Then Exit Sub
If frmReturnItems.WindowState = vbMinimized Then frmReturnItems.WindowState = vbNormal: Exit Sub
 frmReturnItems.Show
 frmReturnItems.SetFocus
 frmReturnItems.Left = 0
 frmReturnItems.Top = 0
  
Case "Report"
  If vr_engine.CheckPermission(Val(gVarAccessLevel), 2) = False Then Exit Sub
If frmReport.WindowState = vbMinimized Then frmReport.WindowState = vbNormal: Exit Sub
 frmReport.Show
 frmReport.SetFocus
 frmReport.Left = 0
 frmReport.Top = 0

Case "Purchase"
If frmPurchase.WindowState = vbMinimized Then frmPurchase.WindowState = vbNormal: Exit Sub
 frmPurchase.Show
 frmPurchase.SetFocus
 frmPurchase.Left = 0
 frmPurchase.Top = 0

  
End Select

End Sub
