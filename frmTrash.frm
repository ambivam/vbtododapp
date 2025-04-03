VERSION 5.00
Begin VB.Form frmTrash 
   Caption         =   "Trash"
   ClientHeight    =   5235
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstTrashedTasks 
      Height          =   3765
      Left            =   120
      TabIndex        =   3
      Top            =   720
      Width           =   6735
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Forever"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Caption         =   "Items in trash will be permanently deleted after 30 days. You can restore items or delete them permanently."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top            =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmTrash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TrashedTask
    ID As Long
    Text As String
    DeletedDate As Date
End Type

Private TrashedTasks() As TrashedTask
Private ItemCount As Integer
Private mUserID As Long

Public Property Let UserID(ByVal Value As Long)
    mUserID = Value
    If Me.Visible Then LoadTrashedTasks
End Property

Private Sub Form_Load()
    ReDim TrashedTasks(0)
    ItemCount = 0
    LoadTrashedTasks
End Sub

Private Sub LoadTrashedTasks()
    Dim rs As ADODB.Recordset
    Set rs = LoadTasks(mUserID, True)  ' True to show deleted items
    
    ReDim TrashedTasks(0)
    ItemCount = 0
    lstTrashedTasks.Clear
    
    Do While Not rs.EOF
        ItemCount = ItemCount + 1
        ReDim Preserve TrashedTasks(ItemCount - 1)
        
        With TrashedTasks(ItemCount - 1)
            .ID = rs("TaskID")
            .Text = rs("TaskText")
            .DeletedDate = rs("DeletedDate")
            
            ' Add to listbox with deletion date
            lstTrashedTasks.AddItem .Text & " (Deleted: " & Format$(.DeletedDate, "dd/mm/yyyy") & ")"
        End With
        
        rs.MoveNext
    Loop
    
    rs.Close
End Sub

Private Sub cmdRestore_Click()
    If lstTrashedTasks.ListIndex < 0 Then
        MsgBox "Please select a task to restore!", vbExclamation
        Exit Sub
    End If
    
    If RestoreTaskFromTrash(TrashedTasks(lstTrashedTasks.ListIndex).ID) Then
        LoadTrashedTasks  ' Refresh the list
        MsgBox "Task restored successfully!", vbInformation
    Else
        MsgBox "Error restoring task!", vbCritical
    End If
End Sub

Private Sub cmdDelete_Click()
    If lstTrashedTasks.ListIndex < 0 Then
        MsgBox "Please select a task to delete!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to permanently delete this task?" & vbCrLf & _
              "This action cannot be undone!", vbQuestion + vbYesNo) = vbYes Then
        
        If PermanentlyDeleteTask(TrashedTasks(lstTrashedTasks.ListIndex).ID) Then
            LoadTrashedTasks  ' Refresh the list
            MsgBox "Task deleted permanently!", vbInformation
        Else
            MsgBox "Error deleting task!", vbCritical
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
