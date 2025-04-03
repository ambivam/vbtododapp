Attribute VB_Name = "modTodo"
Option Explicit

Public Function TrimString(ByVal strInput As String) As String
    TrimString = Trim$(strInput)
End Function

Public Function FormatDate(ByVal dateStr As String) As String
    If IsDate(dateStr) Then
        FormatDate = Format$(CDate(dateStr), "dd/mm/yyyy")
    Else
        FormatDate = ""
    End If
End Function

Public Function IsValidDate(ByVal dateStr As String) As Boolean
    On Error GoTo InvalidDate
    Dim d As Date
    d = CDate(dateStr)
    IsValidDate = True
    Exit Function
    
InvalidDate:
    IsValidDate = False
End Function

Public Function GetPriorityIcon(ByVal priority As String) As String
    Select Case LCase$(priority)
        Case "high"
            GetPriorityIcon = "!"
        Case "medium"
            GetPriorityIcon = "-"
        Case "low"
            GetPriorityIcon = "."
        Case Else
            GetPriorityIcon = "-"
    End Select
End Function

Public Sub CenterForm(frm As Form)
    frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 2
End Sub

Public Function IsDueSoon(ByVal dueDate As String) As Boolean
    If Trim$(dueDate) = "" Then Exit Function
    
    If Not IsDate(dueDate) Then Exit Function
    
    Dim daysUntilDue As Long
    daysUntilDue = DateDiff("d", Date, CDate(dueDate))
    
    IsDueSoon = (daysUntilDue >= 0 And daysUntilDue <= 3)
End Function
