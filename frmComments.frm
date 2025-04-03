VERSION 5.00
Begin VB.Form frmComments 
   Caption         =   "Task Comments"
   ClientHeight    =   6135
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight    =   6135
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtComments 
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   720
      Width           =   8295
   End
   Begin VB.Frame fraNewComment 
      Caption         =   "Add Comment"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   8295
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Comment"
         Height          =   375
         Left            =   6840
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtMentions 
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtNewComment 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblMentions 
         Caption         =   "@Mentions (comma-separated emails):"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblComment 
         Caption         =   "Comment:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Label lblTask 
      Caption         =   "Task: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Weight          =   700
         Charset         =   0
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "frmComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTaskID As Long
Private mUserID As Long
Private mTaskText As String

Public Property Let TaskID(ByVal Value As Long)
    mTaskID = Value
    If Me.Visible Then LoadComments
End Property

Public Property Let UserID(ByVal Value As Long)
    mUserID = Value
End Property

Public Property Let TaskText(ByVal Value As String)
    mTaskText = Value
    lblTask.Caption = "Task: " & mTaskText
End Property

Private Sub Form_Load()
    LoadComments
End Sub

Private Sub LoadComments()
    On Error GoTo LoadError
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT C.CommentText, C.CommentDate, U.DisplayName, " & _
           "(SELECT GROUP_CONCAT(MU.DisplayName) " & _
           "FROM Mentions M INNER JOIN Users MU ON M.MentionedUserID = MU.UserID " & _
           "WHERE M.CommentID = C.CommentID) AS Mentions " & _
           "FROM Comments C INNER JOIN Users U ON C.UserID = U.UserID " & _
           "WHERE C.TaskID = " & mTaskID & " " & _
           "ORDER BY C.CommentDate DESC", cn, adOpenStatic, adLockReadOnly
           
    txtComments.Text = ""
    
    Do While Not rs.EOF
        txtComments.Text = txtComments.Text & _
                          Format$(rs("CommentDate"), "dd/mm/yyyy hh:nn") & " - " & _
                          rs("DisplayName") & ":" & vbCrLf & _
                          rs("CommentText") & vbCrLf
                          
        If Not IsNull(rs("Mentions")) Then
            txtComments.Text = txtComments.Text & _
                              "@Mentions: " & rs("Mentions") & vbCrLf
        End If
        
        txtComments.Text = txtComments.Text & String$(50, "-") & vbCrLf & vbCrLf
        
        rs.MoveNext
    Loop
    
    rs.Close
    Exit Sub
    
LoadError:
    MsgBox "Error loading comments: " & Err.Description, vbCritical
End Sub

Private Sub cmdAdd_Click()
    If Trim(txtNewComment.Text) = "" Then
        MsgBox "Please enter a comment!", vbExclamation
        txtNewComment.SetFocus
        Exit Sub
    End If
    
    If AddComment(mTaskID, mUserID, txtNewComment.Text, txtMentions.Text) Then
        txtNewComment.Text = ""
        txtMentions.Text = ""
        LoadComments
    Else
        MsgBox "Error adding comment!", vbCritical
    End If
End Sub
