VERSION 5.00
Begin VB.Form frmDepartmentAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Entry"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4680
      TabIndex        =   4
      Top             =   3000
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   3480
      TabIndex        =   3
      Top             =   3000
      Width           =   1125
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
      ScaleWidth      =   5970
      TabIndex        =   5
      Top             =   0
      Width           =   5970
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Please fill up all fields provided below. Add/Update Product/SKU Unit."
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
         Left            =   720
         TabIndex        =   6
         Top             =   120
         Width           =   3675
      End
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
      Left            =   1335
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtRemarks 
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
      Height          =   1035
      Left            =   1335
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1680
      Width           =   4560
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
      TabIndex        =   7
      Top             =   2880
      Width           =   5775
   End
   Begin VB.PictureBox CtlLiner2 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   7695
      TabIndex        =   8
      Top             =   795
      Width           =   7695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UnitCD"
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
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   495
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
      TabIndex        =   10
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      TabIndex        =   9
      Top             =   1680
      Width           =   615
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
                Set RS_UNIT = New ADODB.Recordset
                SQLExecute "INSERT INTO tblunit (UnitCD, Description, Remarks, DateCreated, CreatedBy) VALUES ('" & txtUnitCD.Text & "', '" & txtDescription.Text & "', '" & txtRemarks.Text & _
                "', '" & Format(Now, "M/d/yyyy") & "', '" & ACTIVE_USER.USER_NAME & "')"
                MsgBox "New product/sku unit has been successfully saved!", vbInformation, Me.Caption
                Unload Me
            
            Else

                Set RS_UNIT = New ADODB.Recordset
                SQLExecute "UPDATE tblunit SET Description= '" & txtDescription.Text & "', Remarks= '" & txtRemarks.Text & "', LastDateModified= '" & Now & _
                "', ModifiedBy= '" & ACTIVE_USER.USER_NAME & "' WHERE UnitCD='" & txtUnitCD.Text & "'"
                MsgBox "Information saved successfully!", vbInformation, Me.Caption
                Unload Me
            
            End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
txtDescription.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo err
CenterForm frmUnitAE

If State = AddStateMode Then
    Me.Caption = "Create New Entry"
    txtUnitCD.Text = Format(GenerateCD("tblunit"), "00000")
Else

    strSQL = "SELECT tblunit.* " & _
            "FROM tblunit " & _
            "WHERE (((tblunit.UnitCD)='" & PK & "'))"

    Set RS_UNIT = New ADODB.Recordset
    If RS_UNIT.State = adStateOpen Then RS_UNIT.Close
    RS_UNIT.Open strSQL, CN, adOpenDynamic, adLockOptimistic
    
    Me.Caption = "Modify Entry"

    With RS_UNIT
        txtUnitCD.Text = .Fields("UnitCD")
        txtDescription.Text = .Fields("Description")
        txtRemarks.Text = .Fields("Remarks")
    End With
    
End If

Exit Sub
err:
    MsgBox "Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmUnitAE = Nothing
frmUnit.CommandPass "Refresh"
End Sub


