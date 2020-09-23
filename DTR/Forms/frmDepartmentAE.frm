VERSION 5.00
Begin VB.Form frmDepartmentAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Entry"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3240
      TabIndex        =   3
      Top             =   1920
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   1125
   End
   Begin VB.TextBox txtDescription 
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
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtUnitCD 
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
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.PictureBox CtlLiner1 
      Height          =   30
      Left            =   120
      ScaleHeight     =   30
      ScaleWidth      =   5775
      TabIndex        =   4
      Top             =   1800
      Width           =   5775
   End
   Begin VB.PictureBox CtlLiner2 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   7695
      TabIndex        =   5
      Top             =   795
      Width           =   7695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   -120
      ScaleHeight     =   855
      ScaleWidth      =   4815
      TabIndex        =   8
      Top             =   0
      Width           =   4815
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Please fill up all fields provided below. Add/Update."
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
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   3795
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   0
         Picture         =   "frmDepartmentAE.frx":0000
         Top             =   0
         Width           =   1155
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   210
      TabIndex        =   7
      Top             =   960
      Width           =   405
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   100
      TabIndex        =   6
      Top             =   1320
      Width           =   795
   End
End
Attribute VB_Name = "frmDepartmentAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public State                        As FORM_STATE
Public PK                           As String


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim obj As Control
            For Each obj In Me
            If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
                If obj.Text = "" Then
                    MsgBox obj.Name & " could not be left blank. Please complete the field.", vbExclamation, Me.Caption
                    obj.SetFocus
                    Exit Sub
                End If
            End If
            Next obj
            
            If State = AddStateMode Then
                Set RS_Department = New ADODB.Recordset
                SQL_INSERT "INSERT INTO DEPARTMENT (Name, Description) VALUES ('" & txtUnitCD.Text & "', '" & txtDescription.Text & "')"
                MsgBox "New Departments has been successfully saved!", vbInformation, Me.Caption
                Unload Me
            
            Else

                Set RS_Department = New ADODB.Recordset
                SQL_UPDATE "UPDATE DEPARTMENT SET Description= '" & txtDescription.Text & "' WHERE Name='" & txtUnitCD.Text & "'"
                MsgBox "Information saved successfully!", vbInformation, Me.Caption
                Unload Me
            
            End If
End Sub


Private Sub Form_Load()
On Error GoTo err
CenterForm frmDepartmentAE

If State = AddStateMode Then
    Me.Caption = "Create New Entry"
Else

    strSQL = "SELECT DEPARTMENT.* " & _
            "FROM DEPARTMENT " & _
            "WHERE (((DEPARTMENT.Name)='" & PK & "'))"

    Set RS_Department = New ADODB.Recordset
    If RS_Department.State = adStateOpen Then RS_Department.Close
    RS_Department.Open strSQL, conn, adOpenDynamic, adLockOptimistic
    
    Me.Caption = "Modify Entry"

    With RS_Department
        txtUnitCD.Text = .Fields("Name")
        txtDescription.Text = .Fields("Description")
    End With
    
End If

Exit Sub
err:
    MsgBox "Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDepartmentAE = Nothing
frmDepartment.CommandPass "Refresh"
End Sub


