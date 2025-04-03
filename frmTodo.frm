VERSION 5.00
Begin VB.Form frmTodo 
   Caption         =   "Todo Application"
   ClientHeight    =   6195
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTask 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox cboCategory 
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cboPriority 
         Height          =   315
         Left            =   2280
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtDueDate 
         Height          =   285
         Left            =   3600
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Task"
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtNewTask 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label lblCategory 
         Caption         =   "Category:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblPriority 
         Caption         =   "Priority:"
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblDueDate 
         Caption         =   "Due Date (dd/mm/yyyy):"
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame fraSearch 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   6735
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   3255
      End
      Begin VB.ComboBox cboFilterCategory 
         Height          =   315
         Left            =   3600
         TabIndex        =   15
         Text            =   "All Categories"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdClearFilters 
         Caption         =   "Clear Filters"
         Height          =   315
         Left            =   5640
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ListBox lstTodos 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   6735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdMarkComplete 
      Caption         =   "Mark Complete"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Tasks"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Tasks"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Timer tmrReminder 
      Interval        =   60000
      Left            =   6480
      Top             =   120
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdTeams 
      Caption         =   "Teams"
      Height          =   375
      Left            =   7200
      TabIndex        =   19
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdShare 
      Caption         =   "Share"
      Height          =   375
      Left            =   6120
      TabIndex        =   20
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdComments 
      Caption         =   "Comments"
      Height          =   375
      Left            =   5040
      TabIndex        =   21
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   3960
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdDashboard 
      Caption         =   "Dashboard"
      Height          =   375
      Left            =   2880
      TabIndex        =   23
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdPin 
      Caption         =   "Pin"
      Height          =   375
      Left            =   1800
      TabIndex        =   24
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cboSort 
      Height          =   315
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   5580
      Width           =   1935
   End
   Begin VB.ComboBox cboGroup 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   5580
      Width           =   1935
   End
   Begin VB.CommandButton cmdTheme 
      Caption         =   "Theme"
      Height          =   375
      Left            =   720
      TabIndex        =   27
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblSort 
      Caption         =   "Sort by:"
      Height          =   255
      Left            =   5640
      TabIndex        =   28
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lblGroup 
      Caption         =   "Group by:"
      Height          =   255
      Left            =   3480
      TabIndex        =   29
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "Details"
      Height          =   375
      Left            =   3960
      TabIndex        =   30
      Top             =   120
      Width           =   975
   End
   Begin VB.HScrollBar scrWeight 
      Height          =   255
      Left            =   5040
      Max             =   5
      Min             =   1
      TabIndex        =   31
      Top             =   180
      Value           =   1
      Width           =   1215
   End
   Begin VB.Label lblWeight 
      Caption         =   "Weight: 1"
      Height          =   255
      Left            =   6360
      TabIndex        =   32
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmTodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TodoItem
    ID As Long
    Text As String
    IsCompleted As Boolean
    Priority As String
    CategoryID As Long
    Category As String
    DueDate As String
    CreatedDate As Date
    IsPinned As Boolean
    Weight As Integer
    IsRecurring As Boolean
End Type

Private TodoItems() As TodoItem
Private ItemCount As Integer
Private Categories() As String
Private CategoryCount As Integer

Private mUserID As Long
Private mUserEmail As String
Private mTheme As String

Public Property Let UserID(ByVal Value As Long)
    mUserID = Value
    LoadTasksFromDB
End Property

Public Property Let UserEmail(ByVal Value As String)
    mUserEmail = Value
End Property

Private Sub Form_Load()
    ' Initialize database
    If Not InitializeDatabase Then
        MsgBox "Failed to initialize database. Application will close.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    ReDim TodoItems(0)
    ReDim Categories(0)
    ItemCount = 0
    CategoryCount = 0
    
    ' Initialize Priority ComboBox
    With cboPriority
        .AddItem "High"
        .AddItem "Medium"
        .AddItem "Low"
        .Text = "Medium"
    End With
    
    ' Initialize sorting options
    With cboSort
        .Clear
        .AddItem "Created Date (Newest)"
        .AddItem "Created Date (Oldest)"
        .AddItem "Due Date (Soonest)"
        .AddItem "Due Date (Latest)"
        .AddItem "Priority (High to Low)"
        .AddItem "Priority (Low to High)"
        .ListIndex = 0
    End With
    
    ' Initialize grouping options
    With cboGroup
        .Clear
        .AddItem "(None)"
        .AddItem "Priority"
        .AddItem "Category"
        .AddItem "Due Date"
        .ListIndex = 0
    End With
    
    ' Load theme
    mTheme = GetUserTheme(mUserID)
    ApplyTheme
    
    LoadCategories
    LoadTasksFromDB
    
    ' Initialize filter combo
    RefreshCategoryFilter
    
    ' Start checking for reminders
    tmrReminder.Enabled = True
    
    ' Process any recurring tasks
    ProcessRecurringTasks
End Sub

Private Sub LoadCategoriesFromDB()
    Dim rs As ADODB.Recordset
    Set rs = LoadCategories
    
    cboCategory.Clear
    cboFilterCategory.Clear
    cboFilterCategory.AddItem "All Categories"
    
    Do While Not rs.EOF
        cboCategory.AddItem rs("CategoryName")
        cboFilterCategory.AddItem rs("CategoryName")
        rs.MoveNext
    Loop
    
    If Not rs.BOF Then
        rs.MoveFirst
        cboCategory.Text = rs("CategoryName")
    End If
    
    cboFilterCategory.Text = "All Categories"
    rs.Close
End Sub

Private Sub LoadTasksFromDB()
    Dim rs As ADODB.Recordset
    Dim sortClause As String
    
    ' Determine sort order
    Select Case cboSort.ListIndex
        Case 0: sortClause = "CreatedDate DESC"
        Case 1: sortClause = "CreatedDate ASC"
        Case 2: sortClause = "DueDate ASC"
        Case 3: sortClause = "DueDate DESC"
        Case 4: sortClause = "CASE Priority WHEN 'High' THEN 1 WHEN 'Medium' THEN 2 ELSE 3 END"
        Case 5: sortClause = "CASE Priority WHEN 'Low' THEN 1 WHEN 'Medium' THEN 2 ELSE 3 END"
    End Select
    
    ' Determine grouping
    Dim groupClause As String
    Select Case cboGroup.ListIndex
        Case 1: groupClause = "Priority"
        Case 2: groupClause = "CategoryID"
        Case 3: groupClause = "FORMAT(DueDate, 'yyyy-mm-dd')"
        Case Else: groupClause = ""
    End Select
    
    Set rs = LoadTasks(mUserID, False, 0, sortClause, groupClause)
    
    lstTodos.Clear
    ReDim TodoItems(0)
    ItemCount = 0
    
    Dim lastGroup As String
    lastGroup = ""
    
    Do While Not rs.EOF
        ' Add group header if needed
        If groupClause <> "" Then
            Dim currentGroup As String
            Select Case cboGroup.ListIndex
                Case 1: currentGroup = rs("Priority") & ""
                Case 2: currentGroup = Nz(rs("CategoryName"), "(No Category)")
                Case 3: currentGroup = Format$(rs("DueDate"), "mmmm dd, yyyy")
            End Select
            
            If currentGroup <> lastGroup Then
                lstTodos.AddItem "--- " & currentGroup & " ---"
                lastGroup = currentGroup
            End If
        End If
        
        ' Add task
        ItemCount = ItemCount + 1
        ReDim Preserve TodoItems(ItemCount - 1)
        
        With TodoItems(ItemCount - 1)
            .ID = rs("TaskID")
            .Text = rs("TaskText")
            .IsCompleted = rs("IsCompleted")
            .Priority = rs("Priority") & ""
            .CategoryID = Nz(rs("CategoryID"), 0)
            .DueDate = Format$(rs("DueDate"), "dd/mm/yyyy")
            .IsPinned = rs("IsPinned")
            .Weight = rs("Weight")
            .IsRecurring = rs("IsRecurring")
            
            ' Format display text
            Dim displayText As String
            displayText = IIf(.IsCompleted, "[✓] ", "[  ] ")
            displayText = displayText & IIf(.IsPinned, "📌 ", "")
            displayText = displayText & "[" & GetPrioritySymbol(.Priority) & "] "
            If .CategoryID > 0 Then
                displayText = displayText & "[" & rs("CategoryName") & "] "
            End If
            displayText = displayText & .Text
            If .DueDate <> "" Then
                displayText = displayText & " (Due: " & .DueDate & ")"
            End If
            If .IsRecurring Then
                displayText = displayText & " 🔄"
            End If
            If Not IsNull(rs("ParentTaskID")) Then
                displayText = "    " & displayText  ' Indent subtasks
            End If
            
            lstTodos.AddItem displayText
        End With
        
        rs.MoveNext
    Loop
    
    rs.Close
End Sub

Private Sub cmdAdd_Click()
    If Trim(txtNewTask.Text) = "" Then
        MsgBox "Please enter a task!", vbExclamation
        Exit Sub
    End If
    
    ' Validate due date if entered
    If Trim(txtDueDate.Text) <> "" Then
        If Not IsDate(txtDueDate.Text) Then
            MsgBox "Please enter a valid date in dd/mm/yyyy format!", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Get CategoryID
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "SELECT CategoryID FROM Categories WHERE CategoryName = '" & _
            Replace(cboCategory.Text, "'", "''") & "'", cn, adOpenStatic, adLockReadOnly
            
    If rs.EOF Then
        MsgBox "Invalid category selected!", vbExclamation
        rs.Close
        Exit Sub
    End If
    
    Dim categoryID As Long
    categoryID = rs("CategoryID")
    rs.Close
    
    ' Save task to database
    If SaveTask(txtNewTask.Text, False, cboPriority.Text, categoryID, txtDueDate.Text) Then
        txtNewTask.Text = ""
        txtDueDate.Text = ""
        LoadTasksFromDB  ' Refresh the list from database
    End If
End Sub

Private Sub cmdDelete_Click()
    If lstTodos.ListIndex < 0 Then
        MsgBox "Please select a task to delete!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete this task?", _
              vbQuestion + vbYesNo) = vbYes Then
        
        Dim taskID As Long
        taskID = TodoItems(lstTodos.ListIndex).ID
        
        cn.Execute "DELETE FROM Tasks WHERE TaskID = " & taskID
        
        LoadTasksFromDB  ' Refresh the list from database
    End If
End Sub

Private Sub cmdMarkComplete_Click()
    If lstTodos.ListIndex < 0 Then
        MsgBox "Please select a task to mark as complete!", vbExclamation
        Exit Sub
    End If
    
    Dim taskID As Long
    Dim newStatus As Boolean
    
    taskID = TodoItems(lstTodos.ListIndex).ID
    newStatus = Not TodoItems(lstTodos.ListIndex).IsCompleted
    
    cn.Execute "UPDATE Tasks SET IsCompleted = " & _
              IIf(newStatus, "True", "False") & _
              " WHERE TaskID = " & taskID
              
    LoadTasksFromDB  ' Refresh the list from database
End Sub

Private Sub tmrReminder_Timer()
    Dim i As Integer
    Dim msg As String
    Dim dueSoonCount As Integer
    
    For i = 0 To ItemCount - 1
        With TodoItems(i)
            If Not .IsCompleted And Trim$(.DueDate) <> "" Then
                If IsDueSoon(.DueDate) Then
                    msg = msg & "- " & .Text & " (Due: " & .DueDate & ")" & vbCrLf
                    dueSoonCount = dueSoonCount + 1
                End If
            End If
        End With
    Next i
    
    If dueSoonCount > 0 Then
        msg = "You have " & dueSoonCount & " task(s) due soon:" & vbCrLf & vbCrLf & msg
        If MsgBox(msg, vbExclamation + vbOKCancel, "Task Reminders") = vbCancel Then
            ' User clicked Cancel, disable reminders
            tmrReminder.Enabled = False
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    If lstTodos.ListIndex < 0 Then
        MsgBox "Please select a task to edit!", vbExclamation
        Exit Sub
    End If
    
    Dim frmEdit As New frmEditTask
    frmEdit.TaskID = TodoItems(lstTodos.ListIndex).ID
    frmEdit.Show vbModal
    
    If Not frmEdit.Cancelled Then
        LoadTasksFromDB  ' Refresh the list
    End If
End Sub

Private Sub cmdTeams_Click()
    Dim frmT As New frmTeam
    frmT.UserID = mUserID
    frmT.Show vbModal
    LoadTasksFromDB  ' Refresh in case team tasks were modified
End Sub

Private Sub cmdShare_Click()
    If lstTodos.ListIndex < 0 Then
        MsgBox "Please select a task to share!", vbExclamation
        Exit Sub
    End If
    
    Dim ShareLink As String
    ShareLink = ShareTask(TodoItems(lstTodos.ListIndex).ID, mUserID)
    
    If ShareLink <> "" Then
        MsgBox "Task shared successfully!" & vbCrLf & vbCrLf & _
               "Share Link: " & ShareLink, vbInformation
    Else
        MsgBox "Error sharing task!", vbCritical
    End If
End Sub

Private Sub cmdComments_Click()
    If lstTodos.ListIndex < 0 Then
        MsgBox "Please select a task to view comments!", vbExclamation
        Exit Sub
    End If
    
    Dim frmC As New frmComments
    frmC.TaskID = TodoItems(lstTodos.ListIndex).ID
    frmC.UserID = mUserID
    frmC.TaskText = TodoItems(lstTodos.ListIndex).Text
    frmC.Show vbModal
End Sub

Private Sub cmdExport_Click()
    Dim sfd As Object
    Set sfd = Application.GetSaveAsFilename( _
        FileFilter:="CSV Files (*.csv)|*.csv|All Files (*.*)|*.*", _
        FilterIndex:=1, _
        Title:="Export Tasks")
        
    If sfd <> False Then
        If ExportTasksToCSV(mUserID, sfd) Then
            MsgBox "Tasks exported successfully!", vbInformation
        Else
            MsgBox "Error exporting tasks!", vbCritical
        End If
    End If
End Sub

Private Sub cmdDashboard_Click()
    Dim frmD As New frmDashboard
    frmD.UserID = mUserID
    frmD.Show vbModal
End Sub

Private Sub cmdPin_Click()
    If lstTodos.ListIndex < 0 Then
        MsgBox "Please select a task to pin!", vbExclamation
        Exit Sub
    End If
    
    If ToggleTaskPin(TodoItems(lstTodos.ListIndex).ID) Then
        LoadTasksFromDB  ' Refresh to show new pin status
    Else
        MsgBox "Error toggling pin status!", vbCritical
    End If
End Sub

Private Sub cmdTheme_Click()
    If mTheme = "Light" Then
        mTheme = "Dark"
    Else
        mTheme = "Light"
    End If
    
    SetUserTheme mUserID, mTheme
    ApplyTheme
End Sub

Private Sub cmdDetails_Click()
    If lstTodos.ListIndex < 0 Then
        MsgBox "Please select a task to view details!", vbExclamation
        Exit Sub
    End If
    
    Dim frmD As New frmTaskDetails
    frmD.TaskID = TodoItems(lstTodos.ListIndex).ID
    frmD.UserID = mUserID
    frmD.Show vbModal
    
    LoadTasksFromDB  ' Refresh in case subtasks were added
End Sub

Private Sub scrWeight_Change()
    lblWeight.Caption = "Weight: " & scrWeight.Value
    
    If lstTodos.ListIndex >= 0 Then
        UpdateTaskWeight TodoItems(lstTodos.ListIndex).ID, scrWeight.Value
    End If
End Sub

Private Sub lstTodos_Click()
    If lstTodos.ListIndex >= 0 Then
        scrWeight.Value = TodoItems(lstTodos.ListIndex).Weight
    End If
End Sub

Private Sub ApplyTheme()
    Dim backColor As Long
    Dim foreColor As Long
    Dim ctrlBackColor As Long
    
    If mTheme = "Dark" Then
        backColor = RGB(32, 32, 32)
        foreColor = RGB(255, 255, 255)
        ctrlBackColor = RGB(64, 64, 64)
    Else
        backColor = RGB(240, 240, 240)
        foreColor = RGB(0, 0, 0)
        ctrlBackColor = RGB(255, 255, 255)
    End If
    
    Me.BackColor = backColor
    lstTodos.BackColor = ctrlBackColor
    lstTodos.ForeColor = foreColor
    txtNewTask.BackColor = ctrlBackColor
    txtNewTask.ForeColor = foreColor
    
    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is Label Then
            ctl.ForeColor = foreColor
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.BackColor = ctrlBackColor
            ctl.ForeColor = foreColor
        End If
    Next ctl
End Sub

Private Sub cboSort_Click()
    LoadTasksFromDB
End Sub

Private Sub cboGroup_Click()
    LoadTasksFromDB
End Sub

Private Sub RefreshList()
    Dim i As Integer
    Dim displayText As String
    Dim shouldShow As Boolean
    
    lstTodos.Clear
    
    For i = 0 To ItemCount - 1
        shouldShow = True
        
        ' Apply category filter
        If cboFilterCategory.Text <> "All Categories" Then
            If TodoItems(i).Category <> cboFilterCategory.Text Then
                shouldShow = False
            End If
        End If
        
        ' Apply search filter
        If Trim(txtSearch.Text) <> "" Then
            If InStr(1, LCase(TodoItems(i).Text), LCase(txtSearch.Text)) = 0 Then
                shouldShow = False
            End If
        End If
        
        If shouldShow Then
            displayText = ""
            
            ' Status
            If TodoItems(i).IsCompleted Then
                displayText = "[✓] "
            Else
                displayText = "[ ] "
            End If
            
            ' Priority indicator
            Select Case TodoItems(i).Priority
                Case "High"
                    displayText = displayText & "[!] "
                Case "Medium"
                    displayText = displayText & "[-] "
                Case "Low"
                    displayText = displayText & "[.] "
            End Select
            
            ' Category
            displayText = displayText & "[" & TodoItems(i).Category & "] "
            
            ' Task text
            displayText = displayText & TodoItems(i).Text
            
            ' Due date if present
            If Trim(TodoItems(i).DueDate) <> "" Then
                displayText = displayText & " (Due: " & TodoItems(i).DueDate & ")"
            End If
            
            lstTodos.AddItem displayText
        End If
    Next i
End Sub

Private Sub txtSearch_Change()
    RefreshList
End Sub

Private Sub cboFilterCategory_Click()
    RefreshList
End Sub

Private Sub cmdClearFilters_Click()
    txtSearch.Text = ""
    cboFilterCategory.Text = "All Categories"
    RefreshList
End Sub

Private Sub RefreshCategoryFilter()
    Dim i As Integer
    
    cboFilterCategory.Clear
    cboFilterCategory.AddItem "All Categories"
    
    For i = 0 To CategoryCount - 1
        cboFilterCategory.AddItem Categories(i)
    Next i
    
    cboFilterCategory.Text = "All Categories"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseDatabase
End Sub

Private Function GetPrioritySymbol(Priority As String) As String
    Select Case Priority
        Case "High"
            GetPrioritySymbol = "⚡"  ' Lightning bolt for high priority
        Case "Medium"
            GetPrioritySymbol = "●"   ' Bullet for medium priority
        Case "Low"
            GetPrioritySymbol = "○"   ' Circle for low priority
        Case Else
            GetPrioritySymbol = "●"
    End Select
End Function
