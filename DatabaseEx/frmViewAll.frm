VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmViewAll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdDetails 
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         Height          =   375
         Left            =   6240
         TabIndex        =   8
         Top             =   5880
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Top             =   5880
         Width           =   615
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<"
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   5880
         Width           =   615
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   5880
         Width           =   615
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4695
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8281
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblRecordNo 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   5880
         Width           =   3735
      End
      Begin VB.Label lblCustName 
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblCustName 
         Caption         =   "Name"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblCustNo 
         Caption         =   "Cust No."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmViewAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim topID, bottomID, currentID As Long

Private Sub cmdDetails_Click()
On Error GoTo Err_Handler
    
    frmDetails.Prompt cmdDetails.Caption
    
    Exit Sub
    
Err_Handler:
    LogError "frmViewAll", "cmdDetails_Click", Err.Number, Err.Description, Erl()
    Resume Next
End Sub

Private Sub cmdFirst_Click()
On Error GoTo Err_Handler
    
    If currentID <> topID Then
        currentID = topID
        Call DisplayData(currentID)
    End If
    
    Exit Sub
    
Err_Handler:
    LogError "frmViewAll", "cmdFirst_Click", Err.Number, Err.Description, Erl()
    Resume Next
End Sub

Private Sub cmdLast_Click()
On Error GoTo Err_Handler

    If currentID <> bottomID Then
        currentID = bottomID
        Call DisplayData(currentID)
    End If
    
    Exit Sub
    
Err_Handler:
    LogError "frmViewAll", "cmdLast_Click", Err.Number, Err.Description, Erl()
    Resume Next
End Sub

Private Sub cmdNext_Click()
On Error GoTo Err_Handler

    If currentID < bottomID Then
        currentID = currentID + 1
        Call DisplayData(currentID)
    End If
    
    Exit Sub
    
Err_Handler:
    LogError "frmViewAll", "cmdNext_Click", Err.Number, Err.Description, Erl()
    Resume Next
End Sub

Private Sub cmdPrev_Click()
On Error GoTo Err_Handler

    If currentID > topID Then
        currentID = currentID - 1
        Call DisplayData(currentID)
    End If
    
    Exit Sub
    
Err_Handler:
    LogError "frmViewAll", "cmdPrev_Click", Err.Number, Err.Description, Erl()
    Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
    
    currentID = 1
    
    Call InitializeAll
    
    Call GetDataBorder
    Call DisplayData(currentID)
    
    Exit Sub

Err_Handler:
    LogError "frmViewAll", "Form_Load", Err.Number, Err.Description, Erl()
    End
End Sub

Private Function GetDataBorder()
On Error GoTo Err_Handler

    Dim strSQL As String
    
    Set connDB = New Connection
    Set recSet = New Recordset
    
    strSQL = "SELECT COUNT(*) AS totalRegistered FROM Customer"
    connDB.Open VIDEO_DB
    recSet.Open strSQL, connDB, adOpenKeyset, adLockReadOnly
    If Not recSet.EOF Then
        topID = 1
        bottomID = recSet.Fields("totalRegistered")
    Else
        topID = 0
        bottomID = 0
    End If
    recSet.Close
    connDB.Close
    
    Set recSet = Nothing
    Set connDB = Nothing
    
    Exit Function

Err_Handler:
    LogError "frmViewAll", "GetDataBorder", Err.Number, Err.Description, Erl()
    End
End Function

Private Function DisplayData(ByVal curID As Long)
On Error GoTo Err_Handler

    Dim count As Long
    Dim strSQL As String
    
    count = 0

    If topID > 0 Then
        Set connDB = New Connection
        Set recSet = New Recordset
        
        strSQL = "SELECT * FROM Customer ORDER BY c_id"
        connDB.Open VIDEO_DB
        recSet.Open strSQL, connDB, adOpenKeyset, adLockReadOnly
        Do While Not recSet.EOF
            count = count + 1
            If count = curID Then
                cmdDetails.Caption = recSet.Fields("c_id")
                lblCustName(1).Caption = recSet.Fields("c_firstName") & " " & recSet.Fields("c_lastName")
                GoTo exit_loop
            End If
            recSet.MoveNext
        Loop
        
exit_loop:
        recSet.Close
        connDB.Close
        
        Set recSet = Nothing
        Set connDB = Nothing
        
        Call DisplayTable(Trim(cmdDetails.Caption))
    Else
        MsgBox "No data to be displayed!"
    End If
    
    Call RefreshUI
    
    Exit Function

Err_Handler:
    LogError "frmViewAll", "DisplayData", Err.Number, Err.Description, Erl()
    End
End Function

Private Function DisplayTable(ByVal cID As String)
On Error GoTo Err_Handler

    Dim Index As Long

    Dim strSQL As String
    
    Set connDB = New Connection
    Set recSet = New Recordset
    
    strSQL = "SELECT m_name, r_payment, r_dateBorrowed, r_dateReturned, r_billNo FROM Rent_Info " & _
             "INNER JOIN Movie " & _
             "ON Rent_Info.m_id = Movie.m_id " & _
             "WHERE c_id = '" & cID & "'"
    connDB.Open VIDEO_DB
    recSet.Open strSQL, connDB, adOpenKeyset, adLockReadOnly
    
    MSHFlexGrid1.FixedRows = 1
    MSHFlexGrid1.FixedCols = 1
    
    MSHFlexGrid1.AllowUserResizing = flexResizeColumns
    
    If recSet.RecordCount = 0 Then
        MSHFlexGrid1.Rows = recSet.RecordCount + 2
        MSHFlexGrid1.Cols = recSet.Fields.count + 1
        
        MSHFlexGrid1.TextMatrix(1, 0) = ""
        MSHFlexGrid1.TextMatrix(1, 1) = ""
        MSHFlexGrid1.TextMatrix(1, 2) = ""
        MSHFlexGrid1.TextMatrix(1, 3) = ""
        MSHFlexGrid1.TextMatrix(1, 4) = ""
        MSHFlexGrid1.TextMatrix(1, 5) = ""
    Else
        MSHFlexGrid1.Rows = recSet.RecordCount + 1
        MSHFlexGrid1.Cols = recSet.Fields.count + 1
    End If
    
    MSHFlexGrid1.ColWidth(0) = 400
    MSHFlexGrid1.ColWidth(1) = 2000
    MSHFlexGrid1.ColWidth(2) = 1000
    MSHFlexGrid1.ColWidth(3) = 1000
    MSHFlexGrid1.ColWidth(4) = 1000
    MSHFlexGrid1.ColWidth(5) = 1000
    
    MSHFlexGrid1.TextMatrix(0, 0) = "No."
    MSHFlexGrid1.TextMatrix(0, 1) = "Movie Name"
    MSHFlexGrid1.TextMatrix(0, 2) = "Borrowed"
    MSHFlexGrid1.TextMatrix(0, 3) = "Returned"
    MSHFlexGrid1.TextMatrix(0, 4) = "Payment"
    MSHFlexGrid1.TextMatrix(0, 5) = "Bill No."
    
    Index = 1
    Do While Not recSet.EOF
        MSHFlexGrid1.TextMatrix(Index, 0) = Index
        MSHFlexGrid1.TextMatrix(Index, 1) = IIf(IsNull(recSet.Fields("m_name").Value), "", recSet.Fields("m_name").Value)
        MSHFlexGrid1.TextMatrix(Index, 2) = IIf(IsNull(recSet.Fields("r_dateBorrowed").Value), "", recSet.Fields("r_dateBorrowed").Value)
        MSHFlexGrid1.TextMatrix(Index, 3) = IIf(IsNull(recSet.Fields("r_dateReturned").Value), "", recSet.Fields("r_dateReturned").Value)
        MSHFlexGrid1.TextMatrix(Index, 4) = IIf(IsNull(recSet.Fields("r_payment").Value), "", recSet.Fields("r_payment").Value)
        MSHFlexGrid1.TextMatrix(Index, 5) = IIf(IsNull(recSet.Fields("r_billNo").Value), "", recSet.Fields("r_billNo").Value)
        
        Index = Index + 1
        recSet.MoveNext
    Loop
        
    recSet.Close
    connDB.Close
    
    Exit Function

Err_Handler:
    LogError "frmViewAll", "DisplayTable", Err.Number, Err.Description, Erl()
    End
End Function

Private Function RefreshUI()
On Error GoTo Err_Handler

    If currentID = bottomID Then
        cmdNext.Enabled = False
        cmdPrev.Enabled = True
    ElseIf currentID = topID Then
        cmdNext.Enabled = True
        cmdPrev.Enabled = False
    ElseIf topID = 0 Then
        cmdNext.Enabled = False
        cmdPrev.Enabled = False
        cmdFirst.Enabled = False
        cmdLast.Enabled = False
    Else
        cmdNext.Enabled = True
        cmdPrev.Enabled = True
    End If
        
    lblRecordNo.Caption = currentID & " / " & bottomID
    
    Exit Function

Err_Handler:
    LogError "frmViewAll", "RefreshUI", Err.Number, Err.Description, Erl()
    Resume Next
End Function
