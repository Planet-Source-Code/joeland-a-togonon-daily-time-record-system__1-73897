VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Login"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3720
      TabIndex        =   12
      Top             =   2320
      Width           =   1200
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   345
      Left            =   2520
      TabIndex        =   11
      Top             =   2320
      Width           =   1200
   End
   Begin VB.PictureBox CtlLiner2 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   5295
      TabIndex        =   8
      Top             =   790
      Width           =   5295
   End
   Begin VB.PictureBox MCLiner1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   7335
      TabIndex        =   6
      Top             =   2235
      Width           =   7335
      Begin VB.PictureBox CtlLiner1 
         Height          =   30
         Left            =   120
         ScaleHeight     =   30
         ScaleWidth      =   4815
         TabIndex        =   7
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.TextBox txtServer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   1800
      Width           =   2955
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "admin"
      Top             =   1440
      Width           =   2955
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Text            =   "admin"
      Top             =   1080
      Width           =   2955
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   4980
      TabIndex        =   9
      Top             =   0
      Width           =   4980
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Please type your username and password in the space provided bellow."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   840
         TabIndex        =   10
         Top             =   120
         Width           =   3675
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   0
         Picture         =   "frmLogin.frx":0E42
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Login Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   825
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Explicit
Public CloseMe              As Boolean

Private Sub Form_Load()
On Error GoTo err
CenterForm frmLogin

If CONNECT_TO_ACCESS = False Then CloseMe = True: Unload Me: Exit Sub

Exit Sub
err:
    MsgBox "Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation, Me.Caption
End Sub


Private Sub Form_Activate()
On Error Resume Next
If CloseMe = True Then Unload Me: Exit Sub
txtUsername.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
End
End If
End Sub

Private Sub cmdLogin_Click()
    Dim sql As String

If Trim(txtUsername.Text) = "" Then
MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation, Me.Caption
txtUsername.SetFocus
Exit Sub
End If

If Trim(txtPassword.Text) = "" Then
MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation, Me.Caption
txtPassword.SetFocus
Exit Sub
End If

If Not Trim(txtServer.Text) = strServer Then
MsgBox "Unable to connect to the server.Please check connection!", vbExclamation, Me.Caption
txtServer.SetFocus
Exit Sub
End If

Set rs_user = New ADODB.Recordset
If rs_user.State = adStateOpen Then rs_user.Close
rs_user.Open "SELECT * FROM USERS WHERE Username='" & txtUsername.Text & "' AND Password ='" & txtPassword.Text & "'", conn, adOpenStatic, adLockReadOnly

If rs_user.BOF Or rs_user.EOF = True Then
    MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation, Me.Caption
Exit Sub

ElseIf Not rs_user.Fields("Username") = txtUsername.Text Then
    MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation, Me.Caption
Exit Sub

ElseIf Not rs_user.Fields("Password") = txtPassword.Text Then
    MsgBox "Username and/or Password is incorrect.Try Again!", vbExclamation, Me.Caption
Exit Sub

Else

    Unload Me
    MDIMain.Show
    
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
Exit Sub
'frmDTR.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmLogin = Nothing

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdLogin_Click
End If
End Sub

Private Sub txtServer_GotFocus()
Highlight txtServer
End Sub

Private Sub txtUsername_GotFocus()
Highlight txtUsername
End Sub

Private Sub txtPassword_GotFocus()
Highlight txtPassword
End Sub
