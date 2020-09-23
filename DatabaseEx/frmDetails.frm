VERSION 5.00
Begin VB.Form frmDetails 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customer Details"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Customer ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.Line Line1 
         Index           =   11
         X1              =   1440
         X2              =   1440
         Y1              =   2760
         Y2              =   360
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   4560
         X2              =   120
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   4560
         X2              =   120
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   4560
         X2              =   120
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   4560
         X2              =   120
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   4560
         X2              =   120
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   4560
         X2              =   120
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   4560
         X2              =   120
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   4560
         X2              =   120
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4560
         X2              =   4560
         Y1              =   2760
         Y2              =   360
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4560
         X2              =   120
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   120
         Y1              =   360
         Y2              =   2760
      End
      Begin VB.Label lblLastName 
         Caption         =   "Last Name"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblFirstName 
         Caption         =   "First Name"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblIdentityNo 
         Caption         =   "Identity No."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblHomeNo 
         Caption         =   "Home No."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblMobileNo 
         Caption         =   "Mobile No."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblOfficeNo 
         Caption         =   "Office No."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblFirstName 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblEmail 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   18
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label lblOfficeNo 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   17
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label lblMobileNo 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   16
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label lblHomeNo 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   15
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label lblAddress 
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   14
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label lblAddress 
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   13
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label lblAddress 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   12
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label lblIdentityNo 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   11
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label lblLastName 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   10
         Top             =   600
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function Prompt(ByVal custID As String)
On Error GoTo Err_Handler
    
    Frame1.Caption = custID
    
    Call DisplayData(custID)
    
    Me.Show
    
    Exit Function
    
Err_Handler:
    LogError "frmDetails", "Prompt", Err.Number, Err.Description, Erl()
    Resume Next
End Function

Private Function DisplayData(ByVal cID As String)
On Error GoTo Err_Handler

    Dim strSQL As String
    
    Set connDB = New Connection
    Set recSet = New Recordset
    
    strSQL = "SELECT * FROM Customer WHERE c_id = '" & cID & "'"
    connDB.Open VIDEO_DB
    recSet.Open strSQL, connDB, adOpenKeyset, adLockReadOnly
    If Not recSet.EOF Then
        lblFirstName(1).Caption = IIf(IsNull(recSet.Fields("c_firstName").Value), "", recSet.Fields("c_firstName").Value)
        lblLastName(1).Caption = IIf(IsNull(recSet.Fields("c_lastName").Value), "", recSet.Fields("c_lastName").Value)
        lblIdentityNo(1).Caption = IIf(IsNull(recSet.Fields("c_identityNo").Value), "", recSet.Fields("c_identityNo").Value)
        lblAddress(1).Caption = IIf(IsNull(recSet.Fields("c_address1").Value), "", recSet.Fields("c_address1").Value)
        lblAddress(2).Caption = IIf(IsNull(recSet.Fields("c_address2").Value), "", recSet.Fields("c_address2").Value)
        lblAddress(3).Caption = IIf(IsNull(recSet.Fields("c_address3").Value), "", recSet.Fields("c_address3").Value)
        lblHomeNo(1).Caption = IIf(IsNull(recSet.Fields("c_homeNo").Value), "", recSet.Fields("c_homeNo").Value)
        lblMobileNo(1).Caption = IIf(IsNull(recSet.Fields("c_mobileNo").Value), "", recSet.Fields("c_mobileNo").Value)
        lblOfficeNo(1).Caption = IIf(IsNull(recSet.Fields("c_officeNo").Value), "", recSet.Fields("c_officeNo").Value)
        lblEmail(1).Caption = IIf(IsNull(recSet.Fields("c_email").Value), "", recSet.Fields("c_email").Value)
    End If
    
    recSet.Close
    connDB.Close
    
    Set recSet = Nothing
    Set connDB = Nothing
    
    Exit Function

Err_Handler:
    LogError "frmDetails", "DisplayData", Err.Number, Err.Description, Erl()
    End
End Function
