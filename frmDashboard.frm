VERSION 5.00
Begin VB.Form frmDashboard 
   Caption         =   "Task Dashboard"
   ClientHeight    =   6135
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraProgress 
      Caption         =   "Progress Overview"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.PictureBox picProgress 
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   3555
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         Caption         =   "0 of 0 tasks completed (0%)"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame fraPriority 
      Caption         =   "Tasks by Priority"
      Height          =   2775
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   4095
      Begin VB.Label lblLowCount 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblMediumCount 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblHighCount 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblLow 
         Caption         =   "Low Priority"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label lblMedium 
         Caption         =   "Medium Priority"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label lblHigh 
         Caption         =   "High Priority"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame fraCategories 
      Caption         =   "Tasks by Category"
      Height          =   2775
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   4095
      Begin VB.ListBox lstCategories 
         Height          =   2400
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame fraPinned 
      Caption         =   "Pinned Tasks"
      Height          =   2775
      Left            =   4320
      TabIndex        =   12
      Top             =   3000
      Width           =   4095
      Begin VB.ListBox lstPinned 
         Height          =   2400
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mUserID As Long
Private mTeamID As Long

Public Property Let UserID(ByVal Value As Long)
    mUserID = Value
    If Me.Visible Then RefreshDashboard
End Property

Public Property Let TeamID(ByVal Value As Long)
    mTeamID = Value
    If Me.Visible Then RefreshDashboard
End Property

Private Sub Form_Load()
    RefreshDashboard
End Sub

Private Sub RefreshDashboard()
    On Error GoTo RefreshError
    
    Dim stats As Dictionary
    Set stats = GetTaskStats(mUserID, mTeamID)
    
    If stats Is Nothing Then Exit Sub
    
    ' Update progress
    Dim total As Long, completed As Long
    total = stats("Total")
    completed = stats("Completed")
    
    If total > 0 Then
        Dim percent As Integer
        percent = (completed * 100) \ total
        lblProgress.Caption = completed & " of " & total & _
                            " tasks completed (" & percent & "%)"
        
        ' Draw progress bar
        picProgress.Cls
        picProgress.Line (0, 0)-(picProgress.ScaleWidth * percent / 100, _
                                picProgress.ScaleHeight), _
                        RGB(0, 192, 0), BF
    Else
        lblProgress.Caption = "No tasks"
    End If
    
    ' Update priority counts
    lblHighCount.Caption = stats("HighPriority")
    lblMediumCount.Caption = stats("MediumPriority")
    lblLowCount.Caption = stats("LowPriority")
    
    ' Update categories
    Dim catStats As Dictionary
    Set catStats = stats("Categories")
    
    lstCategories.Clear
    Dim key As Variant
    For Each key In catStats.Keys
        lstCategories.AddItem key & ": " & catStats(key)
    Next key
    
    ' Load pinned tasks
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT TaskText, Priority, DueDate FROM Tasks " & _
           "WHERE IsPinned = True AND IsDeleted = False AND " & _
           IIf(mTeamID > 0, "TeamID = " & mTeamID, _
               "UserID = " & mUserID & " AND TeamID IS NULL"), _
           cn, adOpenStatic, adLockReadOnly
           
    lstPinned.Clear
    Do While Not rs.EOF
        lstPinned.AddItem "[" & rs("Priority") & "] " & rs("TaskText") & _
                         IIf(IsNull(rs("DueDate")), "", _
                             " (Due: " & Format$(rs("DueDate"), "dd/mm/yyyy") & ")")
        rs.MoveNext
    Loop
    rs.Close
    
    Exit Sub
    
RefreshError:
    MsgBox "Error refreshing dashboard: " & Err.Description, vbCritical
End Sub
