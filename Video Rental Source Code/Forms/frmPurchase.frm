VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase/Misc"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   Icon            =   "frmPurchase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   11610
   Begin VB.Frame FrameAddItems 
      Caption         =   "Add/Remove Items"
      Height          =   6495
      Left            =   5640
      TabIndex        =   16
      Top             =   120
      Width           =   5865
      Begin VB.Frame FrameAddRemove 
         Height          =   1335
         Left            =   240
         TabIndex        =   24
         Top             =   4920
         Width           =   5415
         Begin VB.CommandButton Command1 
            Caption         =   "Add"
            Height          =   375
            Left            =   4080
            TabIndex        =   15
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txt_Amount 
            Height          =   285
            Left            =   4080
            TabIndex        =   14
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txt_Description 
            Height          =   285
            Left            =   1570
            TabIndex        =   13
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox txt_ItemCode 
            Height          =   285
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
            Height          =   195
            Left            =   4080
            TabIndex        =   27
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            Height          =   195
            Left            =   1560
            TabIndex        =   26
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Item Code"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   720
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FlexAddItem 
         Height          =   4455
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   3
      End
   End
   Begin VB.Frame FrameTransaction 
      Caption         =   "Transaction"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5385
      Begin VB.Frame Frame1 
         Caption         =   " Others "
         Height          =   2175
         Left            =   2760
         TabIndex        =   21
         Top             =   4200
         Width           =   2415
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   1560
            Width           =   1935
         End
         Begin VB.TextBox txtMiscCharge 
            Height          =   285
            Left            =   1200
            TabIndex        =   5
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtDescription 
            Height          =   285
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount"
            Height          =   255
            Left            =   480
            TabIndex        =   23
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Description"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.ListBox lstItemCode 
         Height          =   2790
         Left            =   360
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   3490
         Width           =   2055
      End
      Begin VB.TextBox txtItemCode 
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtChange 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtAmountPaid 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   8
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid FlexTransaction 
         Height          =   2295
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   3
      End
      Begin VB.Label lblItemCode 
         Caption         =   "Item Code:"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Change"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount Paid"
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   2760
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub InitFlex(Flex As MSFlexGrid)
    With Flex
        .ColAlignment(0) = 5
        .ColAlignment(1) = 5
        .ColAlignment(2) = 6
        .ColWidth(0) = 940
        .ColWidth(1) = 2950
        .ColWidth(2) = 900
        .TextMatrix(0, 0) = "Item Code"
        .TextMatrix(0, 1) = "Description"
        .TextMatrix(0, 2) = "Amount "
    End With
End Sub

Private Sub Form_Load()
Call InitFlex(FlexTransaction)
Call InitFlex(FlexAddItem)
End Sub

Private Sub txtAmountPaid_LostFocus()
If Trim(txtAmountPaid.Text) <> "" And Trim(txtTotal.Text) <> "" Then
    If IsNumeric(Trim(txtAmountPaid.Text)) = False Then
        MsgBox "Invalid input.  ", vbInformation, "Error"
        txtAmountPaid.Text = ""
        txtAmountPaid.SetFocus
    Else
        txtAmountPaid.Text = Format(txtAmountPaid.Text, "0.00")
        txtChange.Text = Format(Val(txtAmountPaid.Text) - Val(txtTotal.Text), "0.00")
        If Val(txtTotal.Text) <= Val(txtAmountPaid.Text) Then
            cmdSave.Enabled = True
        Else
            cmdSave.Enabled = False
        End If
    End If

End If
End Sub

Private Sub txtChange_Change()
    txtChange.Text = Format(txtChange.Text, "0.00")
End Sub

Private Sub txtTotal_Change()
txtTotal.Text = Format(txtTotal.Text, "0.00")
End Sub
