VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEmpDetailsAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Details"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodesignation 
      Height          =   330
      Left            =   3120
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adocorporation 
      Height          =   330
      Left            =   1920
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo txtCorporation 
      Bindings        =   "Form2.frx":0000
      Height          =   315
      Left            =   1560
      TabIndex        =   28
      Top             =   2760
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Name"
      BoundColumn     =   "Name"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo txtDesignation 
      Bindings        =   "Form2.frx":001D
      Height          =   315
      Left            =   1560
      TabIndex        =   27
      Top             =   2400
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Name"
      BoundColumn     =   "Name"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboDept 
      Bindings        =   "Form2.frx":003A
      Height          =   315
      Left            =   1560
      TabIndex        =   26
      Top             =   2040
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Description"
      BoundColumn     =   "Description"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodepartment 
      Height          =   330
      Left            =   720
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   9660
      TabIndex        =   22
      Top             =   0
      Width           =   9660
      Begin VB.Image Image1 
         Height          =   720
         Left            =   50
         Picture         =   "Form2.frx":0056
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Please fill up all fields provided below. Add/Update Employee category."
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   840
         TabIndex        =   23
         Top             =   120
         Width           =   3675
      End
   End
   Begin VB.PictureBox PicBlank 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7560
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   20
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtImagePath 
      Height          =   315
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1395
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog dlgSelect 
      Left            =   6360
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdselect 
      Caption         =   "Picture"
      Height          =   345
      Left            =   8400
      TabIndex        =   8
      Top             =   3600
      Width           =   1125
   End
   Begin VB.TextBox txtLTime 
      Height          =   315
      Left            =   8040
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtFTime 
      Height          =   315
      Left            =   6000
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   315
      Left            =   8040
      TabIndex        =   6
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   6000
      TabIndex        =   4
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   345
      Left            =   7200
      TabIndex        =   11
      Top             =   3600
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   4800
      TabIndex        =   9
      Top             =   3600
      Width           =   1125
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   345
      Left            =   6000
      TabIndex        =   10
      Top             =   3600
      Width           =   1125
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   20
      PasswordChar    =   "|"
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox txtEM_ID 
      Height          =   315
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9480
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label10 
      Caption         =   "Corporation:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblEmID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image picEmployee 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   255
      Left            =   7680
      TabIndex        =   19
      Top             =   1035
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Schedule From:"
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblInfoID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Designation:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Department:"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee No:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmEmpDetailsAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                        As FORM_STATE
Public PK                           As String
Dim srcRecord               As String

Dim UPDATETRIGGER As Boolean


Private Sub cmdCLose_Click()
    Unload Me
    MDIMain.Show
End Sub

Private Sub cmdSave_Click()

    If CheckFields = True Then
        Exit Sub
    End If
    
    If CheckExistingID = True Then
        Exit Sub
    End If
    
        Call Info
        With InfoRS
            .AddNew
                !EM_ID = txtEM_ID.Text
                !Name = txtName.Text
                !Department = cboDept.Text
                !Designation = txtDesignation.Text
                !Corporation = txtCorporation.Text
                !PASSWORD = txtPassword.Text
                !WORK_TIMEB = MaskEdBox1 & " AM"
                !WORK_TIMEE = MaskEdBox2 & " PM"
            .Update
        End With
    
        If Not VerifyAndSave() Then
        End If
    
        MsgBox "Record has been saccessfully added!", vbInformation
        Call Clearfileds
        'Call Lockfileds
   ' End If
End Sub

Function VerifyAndSave() As Boolean

    On Error GoTo Errorhandler:
    Call Info
    Dim strMsg, ans As String
    Dim bytBLOB() As Byte
    Dim strImageTitle As String
    Dim strImagePath As String
    Dim intNum As Integer


    strImagePath = Trim$(txtImagePath.Text)
    InfoRS.Find "EM_ID='" + txtEM_ID.Text + "'", 0, 1
        With InfoRS
            If (txtImagePath.Text <> "") Then
                'Open the picture file
                intNum = FreeFile
                Open strImagePath For Binary As #intNum
                ReDim bytBLOB(FileLen(strImagePath))
                
                'Read the data and close the file
                Get #intNum, , bytBLOB
                Close #1
                .Fields("PICTURE").AppendChunk bytBLOB

           End If
        .Update
        End With

    Exit Function
Errorhandler:
       MsgBox "Error :" & " " & err.Number & err.Description, vbCritical
       err.Clear

End Function
Public Sub FillFields()

    Call Info
    InfoRS.Find "EM_ID='" + txtEM_ID.Text + "'", 0, 1
       ' If Trim$("" & InfoRS.Fields("PICTURE")) = "" Then 'Image saved as BLOB

            'The below two lines of code can be used as an alternative to
            'easily link the Image control to the BLOB field in the database
            'instead of loading the stored image into a byte array and saving
            'it to a temporary file.
            
            Dim lngImageSize As Long
            Dim lngOffset As Long
            Dim bytChunk() As Byte
            Dim intFile As Integer
            Dim strTempPic As String
            Const conChunkSize = 100

            'Make sure the temporary file does not already exist
            strTempPic = App.Path & "\TempPic.jpg"
            If Len(Dir(strTempPic)) > 0 Then
                Kill strTempPic
            End If

            'Open the temporary file to save the BLOB to
            intFile = FreeFile
            Open strTempPic For Binary As #intFile

            'Read the binary data into the byte variable array
            lngImageSize = InfoRS("PICTURE").ActualSize
            Do While lngOffset < lngImageSize
               bytChunk() = InfoRS("PICTURE").GetChunk(conChunkSize)
               Put #intFile, , bytChunk()
               lngOffset = lngOffset + conChunkSize
            Loop

            Close #intFile

            'After loading the image, get rid of the temporary file
            picEmployee.Picture = LoadPicture(strTempPic)
            Kill strTempPic
            InfoRS.Close
   
        PicBlank.Visible = False

End Sub

Private Sub cmdselect_Click()
On Error Resume Next
   ' On Error GoTo Err_CS

    'Select an image file to open
    With dlgSelect
        .CancelError = True
        .DialogTitle = "Select image..."
        .Filter = "Image Files (*.jpg, *.bmp, *.gif)|*.jpg;*.bmp;*.gif"
        .ShowOpen
        txtImagePath = .FileName
        If Trim$(txtImagePath.Text) <> "" Then
            LoadPointerImage (Trim$(txtImagePath.Text))
        End If
    End With

    Exit Sub
    
Err_CS:
    If err.Number <> 32755 Then 'User cancelled
        err.Raise err.Number, err.Source, err.Description
        err.Clear
    End If
End Sub
Sub LoadPointerImage(sImage As String)

    On Error GoTo Errorhandler:
    'Load the file pointer into the Image control
    If Len(Dir(sImage)) Then
        picEmployee.Picture = LoadPicture(sImage)
    Else
        picEmployee.Picture = LoadPicture()
    End If
    PicBlank.Visible = False
    Exit Sub
Errorhandler:
    err.Raise err.Number, err.Source, "File does not appear to be a valid image file..." & _
        vbCrLf & vbCrLf & err.Description
    err.Clear
End Sub
Private Sub cmdUpdate_Click()

'On Error GoTo Errorhandler:
    UPDATETRIGGER = True
    If CheckFields = True Then
        Exit Sub
    End If
    
    Call Info
    InfoRS.Find "EM_INFO_ID='" + lblInfoID.Caption + "'", 0, 1
    With InfoRS
            !Name = txtName.Text
            !Department = cboDept.Text
            !Designation = txtDesignation.Text
            !PASSWORD = txtPassword.Text
        
            If txtFTime.Visible = False Then
                !WORK_TIMEB = MaskEdBox1 & " AM"
            End If
            If txtLTime.Visible = False Then
                !WORK_TIMEE = MaskEdBox2 & " PM"
            End If
        .Update
    End With
    
            If Not VerifyAndSave() Then
                
            End If
            
    MsgBox "Record has been saccessfully updated!", vbInformation
    Call Clearfileds
    UPDATETRIGGER = False
    Exit Sub
Errorhandler:
            MsgBox err.Number & " : " & err.Description, vbInformation
End Sub

Private Sub txtEM_ID_Change()
    If Not txtEM_ID.Text = "" Then
        cmdUpdate.Enabled = True
    End If
End Sub

Private Sub txtEM_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error GoTo Errorhandler
        Call Info
        InfoRS.Find "EM_ID='" + txtEM_ID.Text + "'", 0, 1
        With InfoRS
            txtLTime.Visible = True
            txtFTime.Visible = True
            txtName.Text = !Name
            cboDept.Text = !Department
            txtDesignation.Text = !Designation
            txtCorporation.Text = !Corporation
            txtPassword.Text = !PASSWORD
            lblInfoID.Caption = !EM_INFO_ID
            MaskEdBox1 = !WORK_TIMEB
            MaskEdBox2 = !WORK_TIMEE
            txtFTime.Text = !WORK_TIMEB
            txtLTime.Text = !WORK_TIMEE
            txtImagePath.Text = !Picture
        End With
        Call FillFields
        Exit Sub
Errorhandler:
        If err.Number = 3021 Then
            MsgBox "EM.ID: " & txtEM_ID.Text & " has not found!", vbInformation
            txtName.SetFocus
        Else
            Resume Next
        End If
    End If
End Sub
Private Sub Form_Load()
    Call DBLOAD
    Call SQLDB(Adodepartment, "Select * from DEPARTMENT")
    Call SQLDB1(Adocorporation, "Select * from CORPORATION")
    Call SQLDB2(Adodesignation, "Select * from DESIGNATION")

    UPDATETRIGGER = False
End Sub
Private Sub Clearfileds()
     txtEM_ID.Text = ""
     txtName.Text = ""
     cboDept.Text = "- SELECT -"
     txtImagePath.Text = ""
     txtDesignation.Text = ""
     txtCorporation.Text = ""
     txtPassword.Text = ""
     txtFTime.Text = ""
     txtLTime.Text = ""
     MaskEdBox1 = "__:__:__"
     MaskEdBox2 = "__:__:__"
     PicBlank.Visible = True
End Sub
Private Sub Lockfileds()
     txtEM_ID.Locked = True
     txtName.Locked = True
     cboDept.Locked = True
     txtDesignation.Locked = True
     txtCorporation.Locked = True
     txtPassword.Locked = True
     txtFTime.Text = ""
     txtLTime.Text = ""
     PicBlank.Visible = True
End Sub
Private Sub UnLockfileds()
     txtEM_ID.Locked = False
     txtName.Locked = False
     cboDept.Locked = False
     txtDesignation.Locked = False
     txtCorporation.Locked = False
     txtPassword.Locked = False
     txtFTime.Text = ""
     txtLTime.Text = ""
End Sub
Private Sub txtFTime_Click()
    txtFTime.Visible = False
    MaskEdBox1.SetFocus
End Sub
Private Sub txtLTime_Click()
    txtLTime.Visible = False
    MaskEdBox2.SetFocus
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboDept.SetFocus
    End If
End Sub
Private Sub txtPosition_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPassword.SetFocus
    End If
End Sub
Private Function CheckFields() As Boolean
    CheckFields = False
     If txtEM_ID.Text = "" Then
        CheckFields = True
        MsgBox "Employees ID is Empty", vbInformation
        txtEM_ID.SetFocus
     ElseIf txtName.Text = "" Then
        CheckFields = True
        MsgBox "Employees Name is Empty", vbInformation
        txtName.SetFocus
     ElseIf cboDept.Text = "- SELECT -" Then
        CheckFields = True
        MsgBox "Department is Empty", vbInformation
        cboDept.SetFocus
     ElseIf txtDesignation.Text = "" Then
        CheckFields = True
        MsgBox "Designation is Empty", vbInformation
        txtDesignation.SetFocus
     ElseIf txtCorporation.Text = "" Then
        CheckFields = True
        MsgBox "Corporation is Empty", vbInformation
        txtCorporation.SetFocus
     ElseIf txtPassword.Text = "" Then
        CheckFields = True
        MsgBox "Position is Empty", vbInformation
        txtPassword.SetFocus
    End If
     If UPDATETRIGGER = False Then
        If MaskEdBox1 = "__:__:__" Then
            CheckFields = True
            MsgBox "Beginning Time is Empty", vbInformation
            MaskEdBox1.SetFocus
        ElseIf MaskEdBox2 = "__:__:__" Then
            CheckFields = True
            MsgBox "Ending time is Empty", vbInformation
            MaskEdBox2.SetFocus
        End If
     End If
End Function
Private Function CheckExistingID() As Boolean
    On Error Resume Next
    lblEmID.Caption = ""
    CheckExistingID = False
        Call Info
        InfoRS.Find "EM_ID='" + txtEM_ID.Text + "'", 0, 1
        With InfoRS
            lblEmID.Caption = !EM_ID
        End With
        If Not lblEmID.Caption = "" Then
            CheckExistingID = True
            MsgBox "Cannot be save, Employees No. " & txtEM_ID.Text & " is already existing...please try another one.", vbInformation
        End If
End Function
