Attribute VB_Name = "modDatabase"
Option Explicit

Private Const DB_PATH As String = "Todo.mdb"
Private Const CONNECTION_STRING As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="

Public cn As ADODB.Connection

Public Function InitializeDatabase() As Boolean
    On Error GoTo InitError
    
    Set cn = New ADODB.Connection
    cn.ConnectionString = CONNECTION_STRING & App.Path & "\" & DB_PATH
    cn.Open
    
    ' Create tables if they don't exist
    CreateTables
    
    InitializeDatabase = True
    Exit Function
    
InitError:
    MsgBox "Database initialization error: " & Err.Description, vbCritical
    InitializeDatabase = False
End Function

Private Sub CreateTables()
    On Error Resume Next
    
    ' Create Users table
    cn.Execute "CREATE TABLE Users (" & _
               "UserID COUNTER PRIMARY KEY, " & _
               "Email TEXT(100) NOT NULL, " & _
               "Password TEXT(50) NOT NULL, " & _
               "Phone TEXT(20), " & _
               "DisplayName TEXT(50), " & _
               "Theme TEXT(10) DEFAULT 'Light', " & _
               "LastLogin DATE)"
               
    ' Create Teams table
    cn.Execute "CREATE TABLE Teams (" & _
               "TeamID COUNTER PRIMARY KEY, " & _
               "TeamName TEXT(50) NOT NULL, " & _
               "CreatedBy LONG NOT NULL, " & _
               "CreatedDate DATE, " & _
               "CONSTRAINT FK_TeamCreator FOREIGN KEY (CreatedBy) " & _
               "REFERENCES Users(UserID))"
               
    ' Create TeamMembers table
    cn.Execute "CREATE TABLE TeamMembers (" & _
               "TeamID LONG NOT NULL, " & _
               "UserID LONG NOT NULL, " & _
               "Role TEXT(20), " & _
               "JoinDate DATE, " & _
               "CONSTRAINT PK_TeamMembers PRIMARY KEY (TeamID, UserID), " & _
               "CONSTRAINT FK_Team FOREIGN KEY (TeamID) " & _
               "REFERENCES Teams(TeamID), " & _
               "CONSTRAINT FK_TeamUser FOREIGN KEY (UserID) " & _
               "REFERENCES Users(UserID))"
               
    ' Create Categories table
    cn.Execute "CREATE TABLE Categories (" & _
               "CategoryID COUNTER PRIMARY KEY, " & _
               "CategoryName TEXT(50) NOT NULL)"
               
    ' Create Tasks table with team support
    cn.Execute "CREATE TABLE Tasks (" & _
               "TaskID COUNTER PRIMARY KEY, " & _
               "UserID LONG NOT NULL, " & _
               "TeamID LONG, " & _
               "ParentTaskID LONG, " & _
               "TaskText TEXT(255) NOT NULL, " & _
               "IsCompleted BIT DEFAULT 0, " & _
               "Priority TEXT(10), " & _
               "CategoryID LONG, " & _
               "DueDate DATE, " & _
               "CreatedDate DATE, " & _
               "IsDeleted BIT DEFAULT 0, " & _
               "DeletedDate DATE, " & _
               "IsShared BIT DEFAULT 0, " & _
               "ShareLink TEXT(50), " & _
               "IsPinned BIT DEFAULT 0, " & _
               "PinnedDate DATE, " & _
               "CompletedDate DATE, " & _
               "Weight INTEGER DEFAULT 1, " & _
               "IsRecurring BIT DEFAULT 0, " & _
               "RecurringType TEXT(10), " & _
               "RecurringInterval INTEGER, " & _
               "LastRecurrence DATE, " & _
               "NextRecurrence DATE, " & _
               "Notes MEMO, " & _
               "CONSTRAINT FK_Category FOREIGN KEY (CategoryID) " & _
               "REFERENCES Categories(CategoryID), " & _
               "CONSTRAINT FK_User FOREIGN KEY (UserID) " & _
               "REFERENCES Users(UserID), " & _
               "CONSTRAINT FK_TaskTeam FOREIGN KEY (TeamID) " & _
               "REFERENCES Teams(TeamID), " & _
               "CONSTRAINT FK_ParentTask FOREIGN KEY (ParentTaskID) " & _
               "REFERENCES Tasks(TaskID))"
               
    ' Create Comments table
    cn.Execute "CREATE TABLE Comments (" & _
               "CommentID COUNTER PRIMARY KEY, " & _
               "TaskID LONG NOT NULL, " & _
               "UserID LONG NOT NULL, " & _
               "CommentText TEXT(500), " & _
               "CommentDate DATE, " & _
               "CONSTRAINT FK_CommentTask FOREIGN KEY (TaskID) " & _
               "REFERENCES Tasks(TaskID), " & _
               "CONSTRAINT FK_CommentUser FOREIGN KEY (UserID) " & _
               "REFERENCES Users(UserID))"
               
    ' Create Mentions table
    cn.Execute "CREATE TABLE Mentions (" & _
               "MentionID COUNTER PRIMARY KEY, " & _
               "CommentID LONG, " & _
               "TaskID LONG, " & _
               "UserID LONG NOT NULL, " & _
               "MentionedUserID LONG NOT NULL, " & _
               "MentionDate DATE, " & _
               "IsRead BIT DEFAULT 0, " & _
               "CONSTRAINT FK_MentionComment FOREIGN KEY (CommentID) " & _
               "REFERENCES Comments(CommentID), " & _
               "CONSTRAINT FK_MentionTask FOREIGN KEY (TaskID) " & _
               "REFERENCES Tasks(TaskID), " & _
               "CONSTRAINT FK_MentionUser FOREIGN KEY (UserID) " & _
               "REFERENCES Users(UserID), " & _
               "CONSTRAINT FK_MentionedUser FOREIGN KEY (MentionedUserID) " & _
               "REFERENCES Users(UserID))"
               
    ' Create SharedTasks table for external sharing
    cn.Execute "CREATE TABLE SharedTasks (" & _
               "ShareID COUNTER PRIMARY KEY, " & _
               "TaskID LONG NOT NULL, " & _
               "ShareLink TEXT(50) NOT NULL, " & _
               "SharedBy LONG NOT NULL, " & _
               "SharedDate DATE, " & _
               "ExpiryDate DATE, " & _
               "AccessCount LONG DEFAULT 0, " & _
               "CONSTRAINT FK_SharedTask FOREIGN KEY (TaskID) " & _
               "REFERENCES Tasks(TaskID), " & _
               "CONSTRAINT FK_TaskSharer FOREIGN KEY (SharedBy) " & _
               "REFERENCES Users(UserID))"
               
    ' Create DeletedTasks table for undo functionality
    cn.Execute "CREATE TABLE DeletedTasks (" & _
               "DeleteID COUNTER PRIMARY KEY, " & _
               "TaskID LONG NOT NULL, " & _
               "UserID LONG NOT NULL, " & _
               "TaskText TEXT(255) NOT NULL, " & _
               "IsCompleted BIT, " & _
               "Priority TEXT(10), " & _
               "CategoryID LONG, " & _
               "DueDate DATE, " & _
               "CreatedDate DATE, " & _
               "DeletedDate DATE)"
               
    ' Create Attachments table
    cn.Execute "CREATE TABLE Attachments (" & _
               "AttachmentID COUNTER PRIMARY KEY, " & _
               "TaskID LONG NOT NULL, " & _
               "FileName TEXT(255) NOT NULL, " & _
               "FilePath TEXT(255) NOT NULL, " & _
               "FileSize LONG, " & _
               "UploadDate DATE, " & _
               "UploadedBy LONG, " & _
               "CONSTRAINT FK_AttachmentTask FOREIGN KEY (TaskID) " & _
               "REFERENCES Tasks(TaskID), " & _
               "CONSTRAINT FK_AttachmentUser FOREIGN KEY (UploadedBy) " & _
               "REFERENCES Users(UserID))"
End Sub

Public Function SaveCategory(ByVal CategoryName As String) As Long
    On Error GoTo SaveCatError
    
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    ' Check if category exists
    sql = "SELECT CategoryID FROM Categories WHERE CategoryName = ?"
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    rs.AddNew
    rs.Fields("CategoryName") = CategoryName
    rs.Update
    SaveCategory = rs.Fields("CategoryID")
    rs.Close
    Exit Function
    
SaveCatError:
    MsgBox "Error saving category: " & Err.Description, vbCritical
    SaveCategory = -1
End Function

Public Function SaveTask(TaskText As String, IsCompleted As Boolean, _
                        Priority As String, CategoryID As Long, _
                        DueDate As String, UserID As Long, Optional ByVal TeamID As Long = 0) As Boolean
    On Error GoTo SaveTaskError
    
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = "INSERT INTO Tasks (UserID, TeamID, TaskText, IsCompleted, Priority, CategoryID, " & _
          "DueDate, CreatedDate, IsDeleted, DeletedDate, IsShared, ShareLink, IsPinned, PinnedDate, CompletedDate) VALUES (?, ?, ?, ?, ?, ?, ?, Now(), False, NULL, False, NULL, False, NULL, NULL)"
          
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    rs.AddNew
    rs.Fields("UserID") = UserID
    rs.Fields("TeamID") = TeamID
    rs.Fields("TaskText") = TaskText
    rs.Fields("IsCompleted") = IsCompleted
    rs.Fields("Priority") = Priority
    rs.Fields("CategoryID") = CategoryID
    If IsDate(DueDate) Then
        rs.Fields("DueDate") = CDate(DueDate)
    End If
    rs.Update
    rs.Close
    
    SaveTask = True
    Exit Function
    
SaveTaskError:
    MsgBox "Error saving task: " & Err.Description, vbCritical
    SaveTask = False
End Function

Public Function LoadCategories() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM Categories ORDER BY CategoryName", cn, adOpenStatic, adLockReadOnly
    Set LoadCategories = rs
End Function

Public Function LoadTasks(ByVal UserID As Long, Optional ByVal ShowDeleted As Boolean = False, _
                        Optional ByVal TeamID As Long = 0, Optional ByVal SortBy As String = "CreatedDate DESC", _
                        Optional ByVal GroupBy As String = "") As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim sql As String
    sql = "SELECT T.*, C.CategoryName FROM Tasks T " & _
          "LEFT JOIN Categories C ON T.CategoryID = C.CategoryID " & _
          "WHERE T.IsDeleted = " & IIf(ShowDeleted, "True", "False") & " AND "
          
    If TeamID > 0 Then
        sql = sql & "T.TeamID = " & TeamID
    Else
        sql = sql & "T.UserID = " & UserID
    End If
    
    ' Add grouping if specified
    If GroupBy <> "" Then
        sql = sql & " GROUP BY " & GroupBy
    End If
    
    ' Add sorting with pinned tasks always first
    sql = sql & " ORDER BY T.IsPinned DESC, " & SortBy
    
    rs.Open sql, cn, adOpenStatic, adLockReadOnly
    Set LoadTasks = rs
End Function

Public Function MoveTaskToTrash(ByVal TaskID As Long) As Boolean
    On Error GoTo TrashError
    
    ' First, copy the task to DeletedTasks
    cn.Execute "INSERT INTO DeletedTasks " & _
              "(TaskID, UserID, TaskText, IsCompleted, Priority, CategoryID, DueDate, CreatedDate, DeletedDate) " & _
              "SELECT TaskID, UserID, TaskText, IsCompleted, Priority, CategoryID, DueDate, CreatedDate, Now() " & _
              "FROM Tasks WHERE TaskID = " & TaskID
              
    ' Then mark the task as deleted
    cn.Execute "UPDATE Tasks SET IsDeleted = True, DeletedDate = Now() " & _
              "WHERE TaskID = " & TaskID
              
    MoveTaskToTrash = True
    Exit Function
    
TrashError:
    MoveTaskToTrash = False
End Function

Public Function RestoreTaskFromTrash(ByVal TaskID As Long) As Boolean
    On Error GoTo RestoreError
    
    cn.Execute "UPDATE Tasks SET IsDeleted = False, DeletedDate = NULL " & _
              "WHERE TaskID = " & TaskID
              
    ' Remove from DeletedTasks
    cn.Execute "DELETE FROM DeletedTasks WHERE TaskID = " & TaskID
    
    RestoreTaskFromTrash = True
    Exit Function
    
RestoreError:
    RestoreTaskFromTrash = False
End Function

Public Function PermanentlyDeleteTask(ByVal TaskID As Long) As Boolean
    On Error GoTo DeleteError
    
    cn.Execute "DELETE FROM Tasks WHERE TaskID = " & TaskID
    cn.Execute "DELETE FROM DeletedTasks WHERE TaskID = " & TaskID
    
    PermanentlyDeleteTask = True
    Exit Function
    
DeleteError:
    PermanentlyDeleteTask = False
End Function

Public Function AddComment(ByVal TaskID As Long, ByVal UserID As Long, _
                         ByVal CommentText As String, Optional ByVal MentionedUsers As String) As Boolean
    On Error GoTo AddCommentError
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Add the comment
    rs.Open "SELECT * FROM Comments", cn, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs("TaskID") = TaskID
    rs("UserID") = UserID
    rs("CommentText") = CommentText
    rs("CommentDate") = Now
    rs.Update
    
    Dim CommentID As Long
    CommentID = rs("CommentID")
    rs.Close
    
    ' Process mentions if any
    If Len(Trim(MentionedUsers)) > 0 Then
        Dim Users() As String
        Users = Split(MentionedUsers, ",")
        
        Dim i As Integer
        For i = 0 To UBound(Users)
            ' Get UserID for mentioned user
            Set rs = New ADODB.Recordset
            rs.Open "SELECT UserID FROM Users WHERE Email = '" & _
                   Replace(Trim(Users(i)), "'", "''") & "'", cn, adOpenStatic, adLockReadOnly
            
            If Not rs.EOF Then
                ' Add mention
                cn.Execute "INSERT INTO Mentions (CommentID, TaskID, UserID, MentionedUserID, MentionDate) " & _
                          "VALUES (" & CommentID & ", " & TaskID & ", " & UserID & ", " & _
                          rs("UserID") & ", Now())"
                          
                ' Send notification
                Dim MentionedUserID As Long
                MentionedUserID = rs("UserID")
                rs.Close
                
                ' Get task details
                Set rs = New ADODB.Recordset
                rs.Open "SELECT TaskText FROM Tasks WHERE TaskID = " & TaskID, cn, adOpenStatic, adLockReadOnly
                
                If Not rs.EOF Then
                    NotifyMention UserID, MentionedUserID, rs("TaskText"), CommentText
                End If
                rs.Close
            End If
        Next i
    End If
    
    AddComment = True
    Exit Function
    
AddCommentError:
    AddComment = False
End Function

Public Function CreateTeam(ByVal TeamName As String, ByVal CreatedBy As Long) As Long
    On Error GoTo CreateTeamError
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT * FROM Teams", cn, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs("TeamName") = TeamName
    rs("CreatedBy") = CreatedBy
    rs("CreatedDate") = Now
    rs.Update
    
    CreateTeam = rs("TeamID")
    rs.Close
    Exit Function
    
CreateTeamError:
    CreateTeam = 0
End Function

Public Function AddTeamMember(ByVal TeamID As Long, ByVal UserID As Long, _
                            Optional ByVal Role As String = "Member") As Boolean
    On Error GoTo AddMemberError
    
    cn.Execute "INSERT INTO TeamMembers (TeamID, UserID, Role, JoinDate) " & _
              "VALUES (" & TeamID & ", " & UserID & ", '" & _
              Replace(Role, "'", "''") & "', Now())"
              
    AddTeamMember = True
    Exit Function
    
AddMemberError:
    AddTeamMember = False
End Function

Public Function ShareTask(ByVal TaskID As Long, ByVal UserID As Long, _
                        Optional ByVal ExpiryDays As Integer = 7) As String
    On Error GoTo ShareError
    
    ' Generate unique share link
    Dim ShareLink As String
    ShareLink = "TODO-" & Format$(Now, "YYYYMMDD") & "-" & _
               Format$(Timer * 100, "0000") & "-" & _
               Format$(Rnd * 1000, "000")
               
    ' Add share record
    cn.Execute "INSERT INTO SharedTasks (TaskID, ShareLink, SharedBy, SharedDate, ExpiryDate) " & _
              "VALUES (" & TaskID & ", '" & ShareLink & "', " & UserID & ", Now(), " & _
              "DateAdd('d', " & ExpiryDays & ", Now()))"
              
    ' Update task
    cn.Execute "UPDATE Tasks SET IsShared = True, ShareLink = '" & ShareLink & "' " & _
              "WHERE TaskID = " & TaskID
              
    ShareTask = ShareLink
    Exit Function
    
ShareError:
    ShareTask = ""
End Function

Public Function ExportTasksToCSV(ByVal UserID As Long, ByVal FilePath As String, _
                               Optional ByVal TeamID As Long = 0) As Boolean
    On Error GoTo ExportError
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim sql As String
    sql = "SELECT T.TaskText, T.IsCompleted, T.Priority, C.CategoryName, " & _
          "T.DueDate, U.DisplayName AS CreatedBy, Teams.TeamName " & _
          "FROM Tasks T " & _
          "LEFT JOIN Categories C ON T.CategoryID = C.CategoryID " & _
          "LEFT JOIN Users U ON T.UserID = U.UserID " & _
          "LEFT JOIN Teams ON T.TeamID = Teams.TeamID " & _
          "WHERE T.IsDeleted = False AND "
          
    If TeamID > 0 Then
        sql = sql & "T.TeamID = " & TeamID
    Else
        sql = sql & "T.UserID = " & UserID
    End If
    
    rs.Open sql, cn, adOpenStatic, adLockReadOnly
    
    ' Create CSV file
    Dim FileNum As Integer
    FileNum = FreeFile
    
    Open FilePath For Output As #FileNum
    
    ' Write header
    Print #FileNum, "Task,Status,Priority,Category,Due Date,Created By,Team"
    
    ' Write data
    Do While Not rs.EOF
        Print #FileNum, _
              """" & Replace(rs("TaskText"), """", """""") & ""","; _
              IIf(rs("IsCompleted"), "Complete", "Incomplete") & ","; _
              rs("Priority") & ","; _
              """" & Replace(Nz(rs("CategoryName"), ""), """", """""") & ""","; _
              Format$(rs("DueDate"), "yyyy-mm-dd") & ","; _
              """" & Replace(rs("CreatedBy"), """", """""") & ""","; _
              """" & Replace(Nz(rs("TeamName"), ""), """", """""") & """"
        rs.MoveNext
    Loop
    
    Close #FileNum
    rs.Close
    
    ExportTasksToCSV = True
    Exit Function
    
ExportError:
    ExportTasksToCSV = False
End Function

Public Function GetUserTheme(ByVal UserID As Long) As String
    On Error GoTo ThemeError
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT Theme FROM Users WHERE UserID = " & UserID, _
            cn, adOpenStatic, adLockReadOnly
            
    If Not rs.EOF Then
        GetUserTheme = rs("Theme")
    Else
        GetUserTheme = "Light"
    End If
    
    rs.Close
    Exit Function
    
ThemeError:
    GetUserTheme = "Light"
End Function

Public Function SetUserTheme(ByVal UserID As Long, ByVal Theme As String) As Boolean
    On Error GoTo SetThemeError
    
    cn.Execute "UPDATE Users SET Theme = '" & Theme & "' " & _
              "WHERE UserID = " & UserID
              
    SetUserTheme = True
    Exit Function
    
SetThemeError:
    SetUserTheme = False
End Function

Public Function ToggleTaskPin(ByVal TaskID As Long) As Boolean
    On Error GoTo PinError
    
    cn.Execute "UPDATE Tasks SET " & _
              "IsPinned = NOT IsPinned, " & _
              "PinnedDate = " & _
              "IIF(NOT IsPinned, Now(), NULL) " & _
              "WHERE TaskID = " & TaskID
              
    ToggleTaskPin = True
    Exit Function
    
PinError:
    ToggleTaskPin = False
End Function

Public Function GetTaskStats(ByVal UserID As Long, Optional ByVal TeamID As Long = 0) As Dictionary
    On Error GoTo StatsError
    
    Dim stats As New Dictionary
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Base SQL for either personal or team tasks
    Dim whereClause As String
    If TeamID > 0 Then
        whereClause = "TeamID = " & TeamID
    Else
        whereClause = "UserID = " & UserID & " AND TeamID IS NULL"
    End If
    
    ' Total tasks
    rs.Open "SELECT COUNT(*) AS Total FROM Tasks " & _
            "WHERE " & whereClause & " AND IsDeleted = False", _
            cn, adOpenStatic, adLockReadOnly
    stats.Add "Total", rs(0)
    rs.Close
    
    ' Completed tasks
    rs.Open "SELECT COUNT(*) AS Completed FROM Tasks " & _
            "WHERE " & whereClause & " AND IsCompleted = True AND IsDeleted = False", _
            cn, adOpenStatic, adLockReadOnly
    stats.Add "Completed", rs(0)
    rs.Close
    
    ' Tasks by priority
    rs.Open "SELECT Priority, COUNT(*) AS Count FROM Tasks " & _
            "WHERE " & whereClause & " AND IsDeleted = False " & _
            "GROUP BY Priority", cn, adOpenStatic, adLockReadOnly
    
    stats.Add "HighPriority", 0
    stats.Add "MediumPriority", 0
    stats.Add "LowPriority", 0
    
    Do While Not rs.EOF
        Select Case rs("Priority")
            Case "High"
                stats("HighPriority") = rs("Count")
            Case "Medium"
                stats("MediumPriority") = rs("Count")
            Case "Low"
                stats("LowPriority") = rs("Count")
        End Select
        rs.MoveNext
    Loop
    rs.Close
    
    ' Tasks by category
    rs.Open "SELECT C.CategoryName, COUNT(*) AS Count " & _
            "FROM Tasks T LEFT JOIN Categories C ON T.CategoryID = C.CategoryID " & _
            "WHERE " & whereClause & " AND T.IsDeleted = False " & _
            "GROUP BY C.CategoryName", cn, adOpenStatic, adLockReadOnly
            
    Dim catStats As New Dictionary
    Do While Not rs.EOF
        catStats.Add rs("CategoryName"), rs("Count")
        rs.MoveNext
    Loop
    rs.Close
    
    stats.Add "Categories", catStats
    
    Set GetTaskStats = stats
    Exit Function
    
StatsError:
    Set GetTaskStats = Nothing
End Function

Private Function Nz(ByVal Value As Variant, Optional ByVal DefaultValue As Variant = "") As Variant
    If IsNull(Value) Then
        Nz = DefaultValue
    Else
        Nz = Value
    End If
End Function

Public Sub CloseDatabase()
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Sub

Public Function AddSubtask(ByVal ParentTaskID As Long, ByVal TaskText As String, _
                         Optional ByVal Priority As String = "Medium") As Boolean
    On Error GoTo AddSubtaskError
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Get parent task info
    rs.Open "SELECT UserID, TeamID FROM Tasks WHERE TaskID = " & ParentTaskID, _
            cn, adOpenStatic, adLockReadOnly
            
    If rs.EOF Then
        AddSubtask = False
        Exit Function
    End If
    
    Dim userID As Long, teamID As Variant
    userID = rs("UserID")
    teamID = rs("TeamID")
    rs.Close
    
    ' Add subtask
    cn.Execute "INSERT INTO Tasks (UserID, TeamID, ParentTaskID, TaskText, " & _
              "Priority, CreatedDate) VALUES (" & _
              userID & ", " & _
              IIf(IsNull(teamID), "NULL", teamID) & ", " & _
              ParentTaskID & ", '" & _
              Replace(TaskText, "'", "''") & "', '" & _
              Priority & "', Now())"
              
    AddSubtask = True
    Exit Function
    
AddSubtaskError:
    AddSubtask = False
End Function

Public Function GetSubtasks(ByVal ParentTaskID As Long) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT * FROM Tasks WHERE ParentTaskID = " & ParentTaskID & _
            " AND IsDeleted = False ORDER BY CreatedDate", _
            cn, adOpenStatic, adLockReadOnly
            
    Set GetSubtasks = rs
End Function

Public Function AddAttachment(ByVal TaskID As Long, ByVal UserID As Long, _
                            ByVal FileName As String, ByVal FilePath As String, _
                            ByVal FileSize As Long) As Boolean
    On Error GoTo AddAttachmentError
    
    cn.Execute "INSERT INTO Attachments (TaskID, FileName, FilePath, " & _
              "FileSize, UploadDate, UploadedBy) VALUES (" & _
              TaskID & ", '" & _
              Replace(FileName, "'", "''") & "', '" & _
              Replace(FilePath, "'", "''") & "', " & _
              FileSize & ", Now(), " & UserID & ")"
              
    AddAttachment = True
    Exit Function
    
AddAttachmentError:
    AddAttachment = False
End Function

Public Function GetAttachments(ByVal TaskID As Long) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT A.*, U.DisplayName AS UploaderName " & _
            "FROM Attachments A INNER JOIN Users U " & _
            "ON A.UploadedBy = U.UserID " & _
            "WHERE A.TaskID = " & TaskID & _
            " ORDER BY A.UploadDate DESC", _
            cn, adOpenStatic, adLockReadOnly
            
    Set GetAttachments = rs
End Function

Public Function UpdateTaskNotes(ByVal TaskID As Long, ByVal Notes As String) As Boolean
    On Error GoTo UpdateNotesError
    
    cn.Execute "UPDATE Tasks SET Notes = '" & _
              Replace(Notes, "'", "''") & "' " & _
              "WHERE TaskID = " & TaskID
              
    UpdateTaskNotes = True
    Exit Function
    
UpdateNotesError:
    UpdateTaskNotes = False
End Function

Public Function SetTaskRecurrence(ByVal TaskID As Long, ByVal RecurringType As String, _
                                ByVal Interval As Integer) As Boolean
    On Error GoTo SetRecurrenceError
    
    Dim nextDate As Date
    nextDate = CalculateNextRecurrence(Now, RecurringType, Interval)
    
    cn.Execute "UPDATE Tasks SET " & _
              "IsRecurring = True, " & _
              "RecurringType = '" & RecurringType & "', " & _
              "RecurringInterval = " & Interval & ", " & _
              "LastRecurrence = Now(), " & _
              "NextRecurrence = #" & Format$(nextDate, "yyyy/mm/dd") & "# " & _
              "WHERE TaskID = " & TaskID
              
    SetTaskRecurrence = True
    Exit Function
    
SetRecurrenceError:
    SetTaskRecurrence = False
End Function

Public Function UpdateTaskWeight(ByVal TaskID As Long, ByVal Weight As Integer) As Boolean
    On Error GoTo UpdateWeightError
    
    cn.Execute "UPDATE Tasks SET Weight = " & Weight & _
              " WHERE TaskID = " & TaskID
              
    UpdateTaskWeight = True
    Exit Function
    
UpdateWeightError:
    UpdateTaskWeight = False
End Function

Private Function CalculateNextRecurrence(ByVal BaseDate As Date, _
                                       ByVal RecurringType As String, _
                                       ByVal Interval As Integer) As Date
    Select Case UCase$(RecurringType)
        Case "DAILY"
            CalculateNextRecurrence = DateAdd("d", Interval, BaseDate)
        Case "WEEKLY"
            CalculateNextRecurrence = DateAdd("ww", Interval, BaseDate)
        Case "MONTHLY"
            CalculateNextRecurrence = DateAdd("m", Interval, BaseDate)
        Case "YEARLY"
            CalculateNextRecurrence = DateAdd("yyyy", Interval, BaseDate)
        Case Else
            CalculateNextRecurrence = BaseDate
    End Select
End Function

Public Function ProcessRecurringTasks() As Boolean
    On Error GoTo ProcessError
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Get all recurring tasks that are due
    rs.Open "SELECT * FROM Tasks WHERE IsRecurring = True " & _
            "AND IsDeleted = False AND NextRecurrence <= Now()", _
            cn, adOpenStatic, adLockOptimistic
            
    Do While Not rs.EOF
        ' Create new instance of recurring task
        Dim newTaskID As Long
        newTaskID = CreateRecurringTaskInstance(rs)
        
        If newTaskID > 0 Then
            ' Update next recurrence date
            rs("LastRecurrence") = Now
            rs("NextRecurrence") = CalculateNextRecurrence(Now, _
                                  rs("RecurringType"), rs("RecurringInterval"))
            rs.Update
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    ProcessRecurringTasks = True
    Exit Function
    
ProcessError:
    ProcessRecurringTasks = False
End Function

Private Function CreateRecurringTaskInstance(ByVal SourceTask As ADODB.Recordset) As Long
    On Error GoTo CreateError
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT * FROM Tasks", cn, adOpenStatic, adLockOptimistic
    rs.AddNew
    
    ' Copy task details
    rs("UserID") = SourceTask("UserID")
    rs("TeamID") = SourceTask("TeamID")
    rs("TaskText") = SourceTask("TaskText")
    rs("Priority") = SourceTask("Priority")
    rs("CategoryID") = SourceTask("CategoryID")
    rs("CreatedDate") = Now
    rs("Weight") = SourceTask("Weight")
    rs("Notes") = SourceTask("Notes")
    
    ' Calculate due date based on recurrence
    If Not IsNull(SourceTask("DueDate")) Then
        rs("DueDate") = DateAdd("d", _
                       DateDiff("d", SourceTask("CreatedDate"), SourceTask("DueDate")), _
                       Now)
    End If
    
    rs.Update
    CreateRecurringTaskInstance = rs("TaskID")
    rs.Close
    Exit Function
    
CreateError:
    CreateRecurringTaskInstance = 0
End Function
