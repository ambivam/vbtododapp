VERSION 5.00
Begin VB.Form frmEditTask 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Task"
   ClientHeight    =   3090
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboPriority 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cboCategory 
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtDueDate 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtTaskText 
      Height          =   765
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   180
      Width           =   3975
   End
   Begin VB.Label lblPriority 
      Caption         =   "Priority:"
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblCategory 
      Caption         =   "Category:"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblDueDate 
      Caption         =   "Due Date:"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
End
Attribute VB_Name = "frmEditTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TaskID As Long
Private mCancelled As Boolean

Private Sub Form_Load()
    CenterForm Me
    
    ' Initialize Priority ComboBox
    With cboPriority
        .Clear
        .AddItem "High"
        .AddItem "Medium"
        .AddItem "Low"
    End With
    
    ' Load categories
    LoadCategories
    
    ' Load task data if editing
    If TaskID > 0 Then
        LoadTaskData
    End If
End Sub

Private Sub LoadCategories()
    Dim rs As ADODB.Recordset
    Set rs = LoadCategories
    
    cboCategory.Clear
    Do While Not rs.EOF
        cboCategory.AddItem rs("CategoryName")
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub LoadTaskData()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT T.*, C.CategoryName FROM Tasks T " & _
            "LEFT JOIN Categories C ON T.CategoryID = C.CategoryID " & _
            "WHERE T.TaskID = " & TaskID, cn, adOpenStatic, adLockReadOnly
            
    If Not rs.EOF Then
        txtTaskText.Text = rs("TaskText")
        cboPriority.Text = rs("Priority")
        cboCategory.Text = rs("CategoryName")
        If Not IsNull(rs("DueDate")) Then
            txtDueDate.Text = Format$(rs("DueDate"), "dd/mm/yyyy")
        End If
    End If
    
    rs.Close
End Sub

Private Sub cmdSave_Click()
    If Trim(txtTaskText.Text) = "" Then
        MsgBox "Please enter a task description!", vbExclamation
        txtTaskText.SetFocus
        Exit Sub
    End If
    
    If Trim(txtDueDate.Text) <> "" Then
        If Not IsDate(txtDueDate.Text) Then
            MsgBox "Please enter a valid date in dd/mm/yyyy format!", vbExclamation
            txtDueDate.SetFocus
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
    
    ' Update task
    Dim sql As String
    If TaskID > 0 Then
        sql = "UPDATE Tasks SET " & _
              "TaskText = '" & Replace(txtTaskText.Text, "'", "''") & "', " & _
              "Priority = '" & cboPriority.Text & "', " & _
              "CategoryID = " & categoryID & ", " & _
              "DueDate = " & IIf(Trim(txtDueDate.Text) = "", "NULL", "#" & txtDueDate.Text & "#") & _
              " WHERE TaskID = " & TaskID
    Else
        sql = "INSERT INTO Tasks (TaskText, Priority, CategoryID, DueDate, IsCompleted, CreatedDate) " & _
              "VALUES ('" & Replace(txtTaskText.Text, "'", "''") & "', " & _
              "'" & cboPriority.Text & "', " & _
              categoryID & ", " & _
              IIf(Trim(txtDueDate.Text) = "", "NULL", "#" & txtDueDate.Text & "#") & ", " & _
              "False, Now())"
    End If
    
    On Error GoTo SaveError
    cn.Execute sql
    mCancelled = False
    Unload Me
    Exit Sub
    
SaveError:
    MsgBox "Error saving task: " & Err.Description, vbCritical
End Sub

Private Sub cmdCancel_Click()
    mCancelled = True
    Unload Me
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = mCancelled
End Property
