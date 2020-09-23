VERSION 5.00
Begin VB.Form frmUserAccount 
   Caption         =   "TDTRS-SECURITY TOOL"
   ClientHeight    =   6375
   ClientLeft      =   5505
   ClientTop       =   2370
   ClientWidth     =   5160
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5160
   Begin VB.CommandButton cmdCLose 
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   8
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpadate 
      Caption         =   "&UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   7
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   5520
      Width           =   1575
   End
   Begin VB.ComboBox cboUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form5.frx":0000
      Left            =   240
      List            =   "Form5.frx":0002
      TabIndex        =   5
      Text            =   "SELECT"
      Top             =   600
      Width           =   4695
   End
   Begin VB.TextBox txtVerifyPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   4440
      Width           =   4695
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3480
      Width           =   4695
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   4695
   End
   Begin VB.ComboBox cboUserType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form5.frx":0004
      Left            =   240
      List            =   "Form5.frx":000E
      TabIndex        =   0
      Text            =   "SELECT"
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label lblID 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "VERIFY USER PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   4695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "USER PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "USER TYPE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "USER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmUserAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboUser_Click()
    On Error Resume Next
    Call Sec
        SecRS.Find "USER_NAME='" + cboUser.Text + "'", 0, 1

            cboUserType.Text = ""
            txtUsername.Text = ""
            txtPassword.Text = ""
            lblID.Caption = ""

        With SecRS
            cboUserType.Text = !USER_TYPE
            txtUsername.Text = !USER_NAME
            txtPassword.Text = !PASSWORD
            lblID.Caption = !ID
        End With
        
    SecRS.Close
    Set SecRS = Nothing
End Sub

Private Sub cmdCLose_Click()
    Unload Me
    MDIMain.Show
End Sub

Private Sub cmdSave_Click()
    Call Sec
    
    If Not txtPassword.Text = txtVerifyPassword.Text Then
        MsgBox "Password doesn't match!", vbInformation
        txtPassword.SetFocus
        Exit Sub
    End If
    With SecRS
        .AddNew
            !USER_TYPE = cboUserType.Text
            !USER_NAME = txtUsername.Text
            !PASSWORD = txtPassword.Text
        .Update
    End With
    SecRS.Close
    Set SecRS = Nothing
    Call ShowUser
    MsgBox "New user has been successfully added", vbInformation
End Sub

Private Sub cmdUpadate_Click()
    Call Sec
    SecRS.Find "ID='" + lblID.Caption + "'", 0, 1
    With SecRS
            !USER_TYPE = cboUserType.Text
            !USER_NAME = txtUsername.Text
            !PASSWORD = txtPassword.Text
        .Update
    End With
    SecRS.Close
    Set SecRS = Nothing
    Call ShowUser
    MsgBox "User: " & txtUsername.Text & " has been successfully updated!", vbInformation
End Sub

Private Sub Form_Load()
    Call DBLOAD
    Call ShowUser
End Sub

Private Sub ShowUser()
    On Error Resume Next
    Call Sec
    
    
    cboUser.Clear
    SecRS.MoveFirst
    Do
            cboUser.AddItem SecRS("USER_NAME").Value
        SecRS.MoveNext
    Loop While Not SecRS.EOF
    
    SecRS.Close
    Set SecRS = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.Show
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtVerifyPassword.SetFocus
    End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPassword.SetFocus
    End If
End Sub

