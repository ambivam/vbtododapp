VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTaskDetails 
   Caption         =   "Task Details"
   ClientHeight    =   7935
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSubtasks 
      Caption         =   "Subtasks"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top            =   120
      Width           =   9375
      Begin VB.ListBox lstSubtasks 
         Height          =   1425
         Left            =   120
         TabIndex        =   3
         Top            =   600
         Width           =   9135
      End
      Begin VB.TextBox txtNewSubtask 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top            =   240
         Width           =   8055
      End
      Begin VB.CommandButton cmdAddSubtask 
         Caption         =   "Add"
         Height          =   285
         Left            =   8280
         TabIndex        =   1
         Top            =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraNotes 
      Caption         =   "Notes"
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top            =   2400
      Width           =   9375
      Begin RichTextLib.RichTextBox rtbNotes 
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top            =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4260
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmTaskDetails.frx":0000
      End
   End
   Begin VB.Frame fraRecurrence 
      Caption         =   "Recurrence"
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top            =   5280
      Width           =   9375
      Begin VB.ComboBox cboRecurType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top            =   360
         Width           =   1935
      End
      Begin VB.TextBox txtInterval 
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Text            =   "1"
         Top            =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdSetRecurrence 
         Caption         =   "Set"
         Height          =   375
         Left            =   8280
         TabIndex        =   7
         Top            =   360
         Width           =   975
      End
      Begin VB.Label lblRecurType 
         Caption         =   "Repeat:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top            =   360
         Width           =   735
      End
      Begin VB.Label lblInterval 
         Caption         =   "Every"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top            =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraAttachments 
      Caption         =   "Attachments"
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top            =   6600
      Width           =   9375
      Begin VB.ListBox lstAttachments 
         Height          =   840
         Left            =   120
         TabIndex        =   14
         Top            =   240
         Width           =   8055
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "Attach File"
         Height          =   375
         Left            =   8280
         TabIndex        =   13
         Top            =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmTaskDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTaskID As Long
Private mUserID As Long
Private WithEvents mTimer As Timer

Public Property Let TaskID(ByVal Value As Long)
    mTaskID = Value
    If Me.Visible Then RefreshDetails
End Property

Public Property Let UserID(ByVal Value As Long)
    mUserID = Value
End Property

Private Sub Form_Load()
    ' Initialize recurrence types
    With cboRecurType
        .Clear
        .AddItem "Daily"
        .AddItem "Weekly"
        .AddItem "Monthly"
        .AddItem "Yearly"
        .ListIndex = 0
    End With
    
    ' Create auto-save timer
    Set mTimer = Me.Controls.Add("VB.Timer", "tmrAutoSave")
    mTimer.Interval = 30000  ' 30 seconds
    mTimer.Enabled = True
    
    RefreshDetails
End Sub

Private Sub RefreshDetails()
    ' Load subtasks
    LoadSubtasks
    
    ' Load notes
    LoadNotes
    
    ' Load attachments
    LoadAttachments
End Sub

Private Sub LoadSubtasks()
    Dim rs As ADODB.Recordset
    Set rs = GetSubtasks(mTaskID)
    
    lstSubtasks.Clear
    Do While Not rs.EOF
        lstSubtasks.AddItem IIf(rs("IsCompleted"), "[✓] ", "[  ] ") & _
                           rs("TaskText")
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub LoadNotes()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT Notes FROM Tasks WHERE TaskID = " & mTaskID, _
            cn, adOpenStatic, adLockReadOnly
            
    If Not rs.EOF Then
        rtbNotes.TextRTF = IIf(IsNull(rs("Notes")), "", rs("Notes"))
    End If
    rs.Close
End Sub

Private Sub LoadAttachments()
    Dim rs As ADODB.Recordset
    Set rs = GetAttachments(mTaskID)
    
    lstAttachments.Clear
    Do While Not rs.EOF
        lstAttachments.AddItem rs("FileName") & " (" & _
                              Format$(rs("FileSize") / 1024, "#,##0.0") & " KB) - " & _
                              "Added by " & rs("UploaderName") & " on " & _
                              Format$(rs("UploadDate"), "dd/mm/yyyy")
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub cmdAddSubtask_Click()
    If Trim$(txtNewSubtask.Text) = "" Then
        MsgBox "Please enter a subtask description!", vbExclamation
        Exit Sub
    End If
    
    If AddSubtask(mTaskID, txtNewSubtask.Text) Then
        txtNewSubtask.Text = ""
        LoadSubtasks
    Else
        MsgBox "Error adding subtask!", vbCritical
    End If
End Sub

Private Sub cmdSetRecurrence_Click()
    If Not IsNumeric(txtInterval.Text) Then
        MsgBox "Please enter a valid interval!", vbExclamation
        Exit Sub
    End If
    
    Dim interval As Integer
    interval = CInt(txtInterval.Text)
    
    If interval <= 0 Then
        MsgBox "Interval must be greater than 0!", vbExclamation
        Exit Sub
    End If
    
    If SetTaskRecurrence(mTaskID, cboRecurType.Text, interval) Then
        MsgBox "Recurrence set successfully!", vbInformation
    Else
        MsgBox "Error setting recurrence!", vbCritical
    End If
End Sub

Private Sub cmdAttach_Click()
    Dim dlg As Object
    Set dlg = Application.FileDialog(3)  ' msoFileDialogFilePicker
    
    With dlg
        .AllowMultiSelect = False
        .Title = "Select File to Attach"
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        
        If .Show = -1 Then
            Dim fso As New FileSystemObject
            Dim fil As File
            Set fil = fso.GetFile(.SelectedItems(1))
            
            If AddAttachment(mTaskID, mUserID, fil.Name, .SelectedItems(1), fil.Size) Then
                LoadAttachments
            Else
                MsgBox "Error attaching file!", vbCritical
            End If
        End If
    End With
End Sub

Private Sub mTimer_Timer()
    ' Auto-save notes
    UpdateTaskNotes mTaskID, rtbNotes.TextRTF
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Save notes one last time
    UpdateTaskNotes mTaskID, rtbNotes.TextRTF
End Sub
