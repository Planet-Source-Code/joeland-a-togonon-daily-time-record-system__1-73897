VERSION 5.00
Begin VB.Form frmComInfo 
   Caption         =   "TDTRS-SYSTEM SETTINGS"
   ClientHeight    =   1965
   ClientLeft      =   1560
   ClientTop       =   5145
   ClientWidth     =   10365
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   10365
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   7695
   End
   Begin VB.TextBox txtCompanyName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ADRESS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "COMPANY NAME:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmComInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DATA As Boolean

Private Sub cmdSave_Click()
'On Error Resume Next

    COMPANYNAME = ""
    ADDRESS = ""
    
    Call CompanySettings
    
    If DATA = False Then
        With CompanyRS
            .AddNew
            !COMPANY_NAME = txtCompanyName.Text
            !ADDRESS = txtAddress.Text
            .Update
        End With
        MsgBox "Current Settings has been successfully Added!", vbInformation
    Else
        With CompanyRS
            !COMPANY_NAME = txtCompanyName.Text
            !ADDRESS = txtAddress.Text
            .Update
        End With
        MsgBox "Current Settings has been successfully updated!", vbInformation
    End If
    
End Sub

Private Sub Form_Load()

    'On Error Resume Next
    Call DBLOAD
    Call CompanySettings
        With CompanyRS
            txtCompanyName.Text = !COMPANY_NAME
            txtAddress.Text = !ADDRESS
        End With
        If txtCompanyName.Text = "" Then
            DATA = False
        Else
            DATA = True
        End If
End Sub
