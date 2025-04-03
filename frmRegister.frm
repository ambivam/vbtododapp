VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register New User"
   ClientHeight    =   3390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblPhone 
      Caption         =   "Phone:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblConfirmPassword 
      Caption         =   "Confirm:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRegistered As Boolean
Private mEmail As String
Private mPassword As String

Private Sub cmdRegister_Click()
    If Not ValidateInput Then Exit Sub
    
    ' Check if email already exists
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT UserID FROM Users WHERE Email = '" & _
            Replace(txtEmail.Text, "'", "''") & "'", cn, adOpenStatic, adLockReadOnly
            
    If Not rs.EOF Then
        MsgBox "This email is already registered!", vbExclamation
        rs.Close
        Exit Sub
    End If
    rs.Close
    
    ' Insert new user
    cn.Execute "INSERT INTO Users (Email, Password, Phone) VALUES ('" & _
               Replace(txtEmail.Text, "'", "''") & "', '" & _
               Replace(txtPassword.Text, "'", "''") & "', '" & _
               Replace(txtPhone.Text, "'", "''") & "')"
               
    mEmail = txtEmail.Text
    mPassword = txtPassword.Text
    mRegistered = True
    
    MsgBox "Registration successful! You can now login.", vbInformation
    Unload Me
End Sub

Private Function ValidateInput() As Boolean
    If Trim(txtEmail.Text) = "" Then
        MsgBox "Please enter your email!", vbExclamation
        txtEmail.SetFocus
        Exit Function
    End If
    
    If InStr(1, txtEmail.Text, "@") = 0 Or InStr(1, txtEmail.Text, ".") = 0 Then
        MsgBox "Please enter a valid email address!", vbExclamation
        txtEmail.SetFocus
        Exit Function
    End If
    
    If Len(txtPassword.Text) < 6 Then
        MsgBox "Password must be at least 6 characters long!", vbExclamation
        txtPassword.SetFocus
        Exit Function
    End If
    
    If txtPassword.Text <> txtConfirmPassword.Text Then
        MsgBox "Passwords do not match!", vbExclamation
        txtConfirmPassword.SetFocus
        Exit Function
    End If
    
    If Trim(txtPhone.Text) = "" Then
        MsgBox "Please enter your phone number!", vbExclamation
        txtPhone.SetFocus
        Exit Function
    End If
    
    ValidateInput = True
End Function

Private Sub cmdCancel_Click()
    mRegistered = False
    Unload Me
End Sub

Public Property Get Registered() As Boolean
    Registered = mRegistered
End Property

Public Property Get Email() As String
    Email = mEmail
End Property

Public Property Get Password() As String
    Password = mPassword
End Property
