VERSION 5.00
Begin VB.Form frmTeam 
   Caption         =   "Team Management"
   ClientHeight    =   6135
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTeams 
      Caption         =   "Teams"
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top            =   120
      Width           =   4095
      Begin VB.ListBox lstTeams 
         Height          =   2400
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame fraMembers 
      Caption         =   "Team Members"
      Height          =   2775
      Left            =   4320
      TabIndex        =   5
      Top            =   120
      Width           =   4095
      Begin VB.ListBox lstMembers 
         Height          =   2400
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame fraNewTeam 
      Caption         =   "Create New Team"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top            =   3000
      Width           =   8295
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create Team"
         Height          =   375
         Left            =   6840
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtTeamName 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label lblTeamName 
         Caption         =   "Team Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame fraAddMember 
      Caption         =   "Add Team Member"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top            =   4320
      Width           =   8295
      Begin VB.ComboBox cboRole 
         Height          =   315
         ItemData        =   "frmTeam.frx":0000
         Left            =   4920
         List            =   "frmTeam.frx":000D
         Style          =   2  'Dropdown List
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Member"
         Height          =   375
         Left            =   6840
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblRole 
         Caption         =   "Role:"
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TeamInfo
    ID As Long
    Name As String
End Type

Private Teams() As TeamInfo
Private TeamCount As Integer
Private mUserID As Long

Public Property Let UserID(ByVal Value As Long)
    mUserID = Value
    If Me.Visible Then LoadTeams
End Property

Private Sub Form_Load()
    ReDim Teams(0)
    TeamCount = 0
    
    ' Initialize role combo
    If cboRole.ListCount = 0 Then
        cboRole.AddItem "Member"
        cboRole.AddItem "Admin"
        cboRole.AddItem "Viewer"
    End If
    cboRole.ListIndex = 0
    
    LoadTeams
End Sub

Private Sub LoadTeams()
    On Error GoTo LoadError
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Get teams user is member of
    rs.Open "SELECT DISTINCT T.TeamID, T.TeamName " & _
           "FROM Teams T INNER JOIN TeamMembers TM ON T.TeamID = TM.TeamID " & _
           "WHERE TM.UserID = " & mUserID & " OR T.CreatedBy = " & mUserID & " " & _
           "ORDER BY T.TeamName", cn, adOpenStatic, adLockReadOnly
           
    ReDim Teams(0)
    TeamCount = 0
    lstTeams.Clear
    
    Do While Not rs.EOF
        TeamCount = TeamCount + 1
        ReDim Preserve Teams(TeamCount - 1)
        
        With Teams(TeamCount - 1)
            .ID = rs("TeamID")
            .Name = rs("TeamName")
            lstTeams.AddItem .Name
        End With
        
        rs.MoveNext
    Loop
    
    rs.Close
    Exit Sub
    
LoadError:
    MsgBox "Error loading teams: " & Err.Description, vbCritical
End Sub

Private Sub lstTeams_Click()
    If lstTeams.ListIndex < 0 Then Exit Sub
    LoadTeamMembers Teams(lstTeams.ListIndex).ID
End Sub

Private Sub LoadTeamMembers(ByVal TeamID As Long)
    On Error GoTo LoadError
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT U.Email, U.DisplayName, TM.Role, TM.JoinDate " & _
           "FROM Users U INNER JOIN TeamMembers TM ON U.UserID = TM.UserID " & _
           "WHERE TM.TeamID = " & TeamID & " " & _
           "ORDER BY TM.Role, U.DisplayName", cn, adOpenStatic, adLockReadOnly
           
    lstMembers.Clear
    
    Do While Not rs.EOF
        lstMembers.AddItem rs("DisplayName") & " (" & rs("Email") & ") - " & rs("Role")
        rs.MoveNext
    Loop
    
    rs.Close
    Exit Sub
    
LoadError:
    MsgBox "Error loading team members: " & Err.Description, vbCritical
End Sub

Private Sub cmdCreate_Click()
    If Trim(txtTeamName.Text) = "" Then
        MsgBox "Please enter a team name!", vbExclamation
        txtTeamName.SetFocus
        Exit Sub
    End If
    
    Dim TeamID As Long
    TeamID = CreateTeam(txtTeamName.Text, mUserID)
    
    If TeamID > 0 Then
        ' Add creator as admin
        If AddTeamMember(TeamID, mUserID, "Admin") Then
            MsgBox "Team created successfully!", vbInformation
            txtTeamName.Text = ""
            LoadTeams
        Else
            MsgBox "Error adding team member!", vbCritical
        End If
    Else
        MsgBox "Error creating team!", vbCritical
    End If
End Sub

Private Sub cmdAdd_Click()
    If lstTeams.ListIndex < 0 Then
        MsgBox "Please select a team first!", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtEmail.Text) = "" Then
        MsgBox "Please enter member's email!", vbExclamation
        txtEmail.SetFocus
        Exit Sub
    End If
    
    ' Get UserID for email
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT UserID FROM Users WHERE Email = '" & _
           Replace(txtEmail.Text, "'", "''") & "'", cn, adOpenStatic, adLockReadOnly
           
    If rs.EOF Then
        MsgBox "User not found with this email!", vbExclamation
        rs.Close
        Exit Sub
    End If
    
    Dim NewUserID As Long
    NewUserID = rs("UserID")
    rs.Close
    
    ' Check if already a member
    rs.Open "SELECT UserID FROM TeamMembers WHERE TeamID = " & _
           Teams(lstTeams.ListIndex).ID & " AND UserID = " & NewUserID, _
           cn, adOpenStatic, adLockReadOnly
           
    If Not rs.EOF Then
        MsgBox "This user is already a team member!", vbExclamation
        rs.Close
        Exit Sub
    End If
    rs.Close
    
    ' Add member
    If AddTeamMember(Teams(lstTeams.ListIndex).ID, NewUserID, cboRole.Text) Then
        MsgBox "Team member added successfully!", vbInformation
        txtEmail.Text = ""
        LoadTeamMembers Teams(lstTeams.ListIndex).ID
    Else
        MsgBox "Error adding team member!", vbCritical
    End If
End Sub
