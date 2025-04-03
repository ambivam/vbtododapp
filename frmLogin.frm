VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2385
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLoggedIn As Boolean
Private mUserID As Long
Private mUserEmail As String

Private Sub Form_Load()
    mLoggedIn = False
    mUserID = 0
End Sub

Private Sub cmdLogin_Click()
    If Trim(txtEmail.Text) = "" Or Trim(txtPassword.Text) = "" Then
        MsgBox "Please enter both email and password!", vbExclamation
        Exit Sub
    End If
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT UserID, Email, Password FROM Users WHERE Email = '" & _
            Replace(txtEmail.Text, "'", "''") & "'", cn, adOpenStatic, adLockReadOnly
            
    If rs.EOF Then
        MsgBox "Invalid email or password!", vbExclamation
        rs.Close
        Exit Sub
    End If
    
    ' Simple password check (in real app, use proper hashing)
    If rs("Password") <> txtPassword.Text Then
        MsgBox "Invalid email or password!", vbExclamation
        rs.Close
        Exit Sub
    End If
    
    mUserID = rs("UserID")
    mUserEmail = rs("Email")
    mLoggedIn = True
    rs.Close
    
    Unload Me
End Sub

Private Sub cmdRegister_Click()
    Dim frmReg As New frmRegister
    frmReg.Show vbModal
    
    If frmReg.Registered Then
        txtEmail.Text = frmReg.Email
        txtPassword.Text = frmReg.Password
    End If
End Sub

Private Sub cmdCancel_Click()
    mLoggedIn = False
    Unload Me
End Sub

Public Property Get LoggedIn() As Boolean
    LoggedIn = mLoggedIn
End Property

Public Property Get UserID() As Long
    UserID = mUserID
End Property

Public Property Get UserEmail() As String
    UserEmail = mUserEmail
End Property
