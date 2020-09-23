VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDTR 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000C0&
      Height          =   5235
      Left            =   4680
      ScaleHeight     =   5175
      ScaleWidth      =   6075
      TabIndex        =   22
      Top             =   6240
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid msGrid 
         Height          =   4935
         Left            =   0
         TabIndex        =   23
         Top             =   240
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   8705
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         ForeColor       =   0
         BackColorFixed  =   -2147483630
         BackColorSel    =   255
         ForeColorSel    =   -2147483624
         BackColorBkg    =   -2147483633
         GridColor       =   -2147483634
         GridColorFixed  =   8421504
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Out"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5160
         TabIndex        =   28
         Top             =   0
         Width           =   765
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time In"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3960
         TabIndex        =   27
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Out"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2640
         TabIndex        =   26
         Top             =   0
         Width           =   765
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time In"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   25
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Emp. No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   45
         TabIndex        =   24
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdShowReport 
      BackColor       =   &H8000000E&
      Caption         =   " LOGGED HISTORY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   600
      MaxLength       =   10
      PasswordChar    =   "l"
      TabIndex        =   5
      Top             =   5040
      Width           =   4335
   End
   Begin VB.TextBox txtEmployeeID 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   3840
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   240
   End
   Begin VB.Label Label17 
      Height          =   375
      Left            =   480
      TabIndex        =   29
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press Esc to Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Label lblcorporation 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   20
      Top             =   4920
      Width           =   4935
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Corporation:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   19
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   3015
      Left            =   5400
      Top             =   3000
      Width           =   9855
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1335
      Left            =   8040
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1335
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H80000006&
      BorderWidth     =   2
      Height          =   2775
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   5055
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   14640
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   14760
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "PERSONAL INFORMATION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7680
      TabIndex        =   17
      Top             =   3120
      Width           =   9855
   End
   Begin VB.Image imgEmployee01 
      Height          =   2295
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000000&
      Height          =   2535
      Left            =   5520
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   " DAILY TIME RECORD SYSTEM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   15375
   End
   Begin VB.Label lblLTime 
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblFTime 
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      Height          =   1095
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      Height          =   1095
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label lblLogID 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   2880
      TabIndex        =   12
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER PASSWORD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   4815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER EMPLOYEES ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   10200
      TabIndex        =   9
      Top             =   3480
      Width           =   4935
   End
   Begin VB.Label lblDept 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   10200
      TabIndex        =   8
      Top             =   3960
      Width           =   4935
   End
   Begin VB.Label lblPosition 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   10200
      TabIndex        =   7
      Top             =   4440
      Width           =   4935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Designation:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Department:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   8160
      TabIndex        =   0
      Top             =   1320
      Width           =   4575
   End
End
Attribute VB_Name = "frmDTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim timein As Date, timeout As Date, time01 As Date, TimeDifferent As Date, TimeDifferent01 As Date, TimeDifferent02 As Date
Dim HourDiff As Integer, MinuteDiff As Integer, SecondDiff As Integer, HourDiff01 As Integer, MinuteDiff01 As Integer, SecondDiff01 As Integer, HourDiff02 As Integer, MinuteDiff02 As Integer, SecondDiff02 As Integer
Dim Result, Result01, t1, t2, tt1, tt2 As String
Dim login_am, logout_am, login_pm, logout_pm, COMPUTE_TOTAL, LATE, ONE_HOUR_DEDUC As Boolean


Private Sub cmdShowReport_Click()
    frmLogHistory.Show
    frmLogHistory.txtEmID.Text = txtEmployeeID.Text

End Sub

Private Sub Form_Activate()
'Me.SetFocus
End Sub

Private Sub Form_Load()
 init_Grid
    init_data
Call DBLOAD

    login_am = False
    login_pm = False
    logout_am = False
    logout_pm = False
    COMPUTE_TOTAL = False
    ONE_HOUR_DEDUC = False
    LATE = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblClose.ForeColor = vbWhite
End Sub

Private Sub Form_Terminate()
   ' Form1.Show
End Sub


Private Sub lblClose_Click()
    frmLogin.Show
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    lblClose.ForeColor = vbRed
End Sub

Private Sub Timer1_Timer()
 lblTime.Caption = Time
 lblDate.Caption = Date
End Sub

Private Sub txtEmployeeID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error GoTo Errorhandler
        Call Info
        InfoRS.Find "EM_ID='" + txtEmployeeID.Text + "'", 0, 1
        With InfoRS
        
            lblName.Caption = !Name
            lblDept.Caption = !Department
            lblPosition.Caption = !Designation
            lblcorporation.Caption = !Corporation
            Label17.Caption = !Status
            lblFTime.Caption = !WORK_TIMEB
            lblLTime.Caption = !WORK_TIMEE
            
        End With
        Call FillFieldsLogin
        Call LoadOject
        txtPassword.SetFocus
        Call extract_timein
        Exit Sub
Errorhandler:
        If err.Number = 3021 Then
            MsgBox "EMPLOYEES NUMBER.: " & txtEmployeeID.Text & " has not found!", vbInformation
            Call Timer1_Timer
            Call UnloadOject
        Else
            Resume Next
        End If
    End If
End Sub
Public Sub FillFieldsLogin()

 Call Info
    InfoRS.Find "EM_ID='" + txtEmployeeID.Text + "'", 0, 1

            Dim lngImageSize As Long
            Dim lngOffset As Long
            Dim bytChunk() As Byte
            Dim intFile As Integer
            Dim strTempPic As String
            Const conChunkSize = 100

            strTempPic = App.Path & "\TempPic.jpg"
            If Len(Dir(strTempPic)) > 0 Then
                Kill strTempPic
            End If

            intFile = FreeFile
            Open strTempPic For Binary As #intFile

            lngImageSize = InfoRS("PICTURE").ActualSize
            Do While lngOffset < lngImageSize
               bytChunk() = InfoRS("PICTURE").GetChunk(conChunkSize)
               Put #intFile, , bytChunk()
               lngOffset = lngOffset + conChunkSize
            Loop

            Close #intFile

            imgEmployee01.Picture = LoadPicture(strTempPic)
            Kill strTempPic
            InfoRS.Close
End Sub
Sub LoadPointerImageLogin(sImage As String)

    On Error GoTo Errorhandler:
    If Len(Dir(sImage)) Then
        imgEmployee01.Picture = LoadPicture(sImage)
    Else
        imgEmployee01.Picture = LoadPicture()
    End If

    Exit Sub
Errorhandler:
    err.Raise err.Number, err.Source, "File does not appear to be a valid image file..." & _
        vbCrLf & vbCrLf & err.Description
    err.Clear
End Sub
Private Sub LoadOject()
    imgEmployee01.Visible = True

End Sub
Private Sub UnloadOject()
    lblName.Caption = ""
    lblDept.Caption = ""
    lblPosition.Caption = ""
    lblcorporation.Caption = ""
imgEmployee01.Visible = False
    txtEmployeeID.Text = ""
    txtPassword.Text = ""
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Dim PASSWORD As String
    Dim correct As Boolean
Dim rs_user As New ADODB.Recordset
    
    If KeyAscii = 13 Then
        correct = False
        Call Info
        InfoRS.MoveFirst
        Do
            If InfoRS("EM_ID").Value = txtEmployeeID.Text Then
                If InfoRS("PASSWORD").Value = txtPassword.Text Then
                    correct = True
                Else
                    correct = False
                End If
            End If
            InfoRS.MoveNext
        Loop While Not InfoRS.EOF
    
        If correct = True Then
            Call checklog
            txtEmployeeID.SetFocus
        Else
            MsgBox "Wrong password!", vbInformation
            txtPassword.Text = Empty
            Call Timer1_Timer
        End If
    End If

        msGrid.Refresh
        Picture1.Refresh
                Exit Sub
End Sub
Private Sub computetime()
    Dim temptime, temptime1, temptime2, temptime3, t3 As String
    Dim timelenght, timelenght1, timelenght2, i As Integer
           
    timelenght2 = 0
    timelenght = 0
    temptime1 = ""
    temptime2 = ""
    temptime3 = ""
    temptime = ""
    t3 = ""
    
        Result = ""
    If COMPUTE_TOTAL = True Then
        timein = TimeValue(tt1)
        timeout = TimeValue(tt2)
        TimeDifferent = (timein + timeout)
        HourDiff = Hour(TimeDifferent)
        MinuteDiff = Minute(TimeDifferent)
        SecondDiff = Second(TimeDifferent)
        Result = HourDiff & ":" & MinuteDiff & ":" & SecondDiff
    Else
        If ShowFistLog = True Then
            If CheckPM_AM(lblTime.Caption) = "PM" Then
                
                ONE_HOUR_DEDUC = True
            End If
        End If
        
        If ONE_HOUR_DEDUC = False Then
            t2 = lblTime.Caption
            timein = TimeValue(t1)
            timeout = TimeValue(t2)
            TimeDifferent = (timein - timeout)
            HourDiff = Hour(TimeDifferent)
            MinuteDiff = Minute(TimeDifferent)
            SecondDiff = Second(TimeDifferent)
            Result = HourDiff & ":" & MinuteDiff & ":" & SecondDiff
            
        Else
            
            t2 = lblTime.Caption
            timelenght = Len(t2)
            MsgBox (t2)
            For i = 1 To timelenght
                temptime = Mid(t2, i, 1)
                If Not temptime = ":" Then
                    temptime1 = temptime1 & temptime
                    MsgBox ("time1: " & temptime1 & " time: " & temptime)
                Else
                    Exit For
                End If
                temptime = ""
            Next i
            
            
            timelenght2 = Len(temptime1)
            temptime2 = Mid(t2, timelenght2 + 1, 9)
            temptime3 = Val(temptime1) - 1
            t3 = temptime3 & temptime2
            MsgBox (temptime3 & " : " & temptime2)
            timein = TimeValue(t1)
            timeout = TimeValue(t3)
            TimeDifferent = (timein - timeout)
            HourDiff = Hour(TimeDifferent)
            MinuteDiff = Minute(TimeDifferent)
            SecondDiff = Second(TimeDifferent)
            Result = HourDiff & ":" & MinuteDiff & ":" & SecondDiff

            MsgBox (Result)
        End If
    End If
End Sub
Private Sub extract_timein()

        timein = TimeValue(lblTime.Caption)
        time01 = TimeValue(lblFTime.Caption)
        
        TimeDifferent01 = (timein)
        HourDiff01 = Hour(TimeDifferent01)
        MinuteDiff01 = Minute(TimeDifferent01)
        SecondDiff01 = Second(TimeDifferent01)
        
        TimeDifferent02 = (time01)
        HourDiff02 = Hour(TimeDifferent02)
        MinuteDiff02 = Minute(TimeDifferent02)
        SecondDiff02 = Second(TimeDifferent02)

End Sub
Private Sub computetime_late_and_undertime()
        Dim timein_l As Date, timeworkin_l As Date, TimeDifferent_l As Date
        Dim HourDiff_l As Integer, MinuteDiff_l As Integer, SecondDiff_l As Integer
        
        Result01 = ""
 
        timein_l = TimeValue(lblTime.Caption)
        timeworkin_l = TimeValue(lblFTime.Caption)
        TimeDifferent_l = (timein_l - timeworkin_l)
        HourDiff_l = Hour(TimeDifferent_l)
        MinuteDiff_l = Minute(TimeDifferent_l)
        SecondDiff_l = Second(TimeDifferent_l)
        Result01 = HourDiff_l & ":" & MinuteDiff_l & ":" & SecondDiff_l

End Sub
Private Sub checklog()
    Dim sql, RECTIME As String
    On Error GoTo Errorhandler
    sql = ""
    RECTIME = ""
    Result = ""
    Call LogIn
    
    LogInRS.Close
        sql = "SELECT * FROM TIMELOG WHERE EM_ID='" & txtEmployeeID.Text & "'and DATE_LOG like'" & lblDate.Caption & "'"
    LogInRS.Open sql, conn, adOpenDynamic, adLockOptimistic
    If Not LogInRS("IN_AM").Value = "" Then
        login_am = True
        
        With LogInRS
            t1 = !IN_AM
            lblLogID.Caption = !LOG_ID
        End With
        Call computetime
        
        If Not LogInRS("OUT_AM").Value = "00:00:00" Then
            logout_am = True
            With LogInRS
                lblLogID.Caption = !LOG_ID
            End With
            
            If Not LogInRS("IN_PM").Value = "00:00:00" Then
                login_pm = True
                With LogInRS
                    lblLogID.Caption = !LOG_ID
                    t1 = !IN_PM
                End With
                Call computetime
                
                If Not LogInRS("OUT_PM").Value = "00:00:00" Then
                    logout_pm = True
                    With LogInRS
                        lblLogID.Caption = !LOG_ID
                    End With
                Else
                    logout_pm = False
                End If
            Else
                login_pm = False
            End If
        Else
            logout_am = False
        End If
    Else
        login_am = False
    End If
    Call Savelog
    Exit Sub
Errorhandler:
        If err.Number = 3021 Then
            login_am = False
            Call Savelog
        Else
            Resume Next
        End If
End Sub
Private Function CheckPM_AM(ByVal CurrentTime As String) As String
    Dim timePMAM As String
    timePMAM = ""
    
    timePMAM = Right(CurrentTime, 2)
    
    CheckPM_AM = timePMAM
End Function

Private Function ShowFistLog() As Boolean
    Dim sql, Ttemp As String
    sql = ""
    Ttemp = ""
    On Error GoTo Errorhandler:
    ShowFistLog = False
    Call LogIn
        
        LogInRS.Close
            sql = "SELECT * FROM TIMELOG WHERE EM_ID='" & txtEmployeeID.Text & "'and DATE_LOG like'" & lblDate.Caption & "'"
        LogInRS.Open sql, conn, adOpenDynamic, adLockOptimistic
        
        With LogInRS
               Ttemp = !IN_AM
        End With
        If Not Ttemp = "" Then
            If CheckPM_AM(Ttemp) = "AM" Then
             ShowFistLog = True
            End If
        End If
Errorhandler:
        If err.Number = 3021 Then
        ShowFistLog = False
        Else
            Resume Next
        End If
End Function
Private Sub Savelog()
    Call LogIn
    'MsgBox "LOGIN AM: " & login_am & " LOGOUT AM: " & logout_am & " LOGIN PM: " & login_pm & "LOGOUT PM:" & logout_pm
      If login_am = False And logout_am = False And login_pm = False And logout_pm = False Then 'IN AM
       ' MsgBox "LOGIN HR: " & HourDiff01 & " : " & TimeDifferent01 '& " LOGIN SCH: " & HourDiff02 & " LOGIN MN: " & MinuteDiff01 & " LOGIN SCH MN: " & MinuteDiff02 & " LOGIN SEC: " & SecondDiff01 & " LOGIN SCH SEC: " & SecondDiff02
        If HourDiff01 > HourDiff02 Then
            LATE = True
            Call computetime_late_and_undertime
        ElseIf HourDiff01 = HourDiff02 Then
            If MinuteDiff01 > MinuteDiff02 Then
                LATE = True
                Call computetime_late_and_undertime
            ElseIf MinuteDiff01 = MinuteDiff02 Then
                If SecondDiff01 > SecondDiff02 Then
                    LATE = True
                    Call computetime_late_and_undertime
                End If
            End If
        End If
                   
        With LogInRS
            .AddNew
                !EM_ID = txtEmployeeID.Text
                !Name = lblName.Caption
                !DATE_LOG = lblDate.Caption
                !IN_AM = lblTime.Caption
                !OUT_AM = "00:00:00"
                !IN_PM = "00:00:00"
                !OUT_PM = "00:00:00"
                !TOTAL_AM = "00:00:00"
                !TOTAL_PM = "00:00:00"
                !GRAND_TOTAL = "00:00:00"
                If LATE = True Then
                !TOTAL_LATE = Result01
                !REMARKS = "LATE"
                Else
                !TOTAL_LATE = "00:00:00"
                !REMARKS = ""
                End If
            .Update
        End With
         
        MsgBox "Employees ID: " & txtEmployeeID.Text & " has been successfully LOG IN!", vbInformation
        Call Timer1_Timer
init_data
    ElseIf login_am = True And logout_am = False And login_pm = False And logout_pm = False Then 'OUT AM

        LogInRS.Find "LOG_ID='" + lblLogID.Caption + "'", 0, 1
        With LogInRS
                !OUT_AM = lblTime.Caption
                !TOTAL_AM = Result
                !GRAND_TOTAL = Result
            .Update
        End With
        
        MsgBox "Employees ID: " & txtEmployeeID.Text & " has been successfully LOG OUT!", vbInformation
        Call Timer1_Timer
init_data
    ElseIf login_am = True And logout_am = True And login_pm = False And logout_pm = False Then ' IN PM
        
        LogInRS.Find "LOG_ID='" + lblLogID.Caption + "'", 0, 1
        With LogInRS
                !IN_PM = lblTime.Caption
            .Update
        End With
        
        MsgBox "Employees ID: " & txtEmployeeID.Text & " has been successfully LOG IN!", vbInformation
        Call Timer1_Timer
init_data
    ElseIf login_am = True And logout_am = True And login_pm = True And logout_pm = False Then ' OUT PM
        
        LogInRS.Find "LOG_ID='" + lblLogID.Caption + "'", 0, 1
        With LogInRS
                !OUT_PM = lblTime.Caption
                !TOTAL_PM = Result
                !GRAND_TOTAL = Result
            .Update
        End With
        
        With LogInRS
                tt1 = !TOTAL_AM
                tt2 = !TOTAL_PM
        End With
        
        COMPUTE_TOTAL = True
        Call computetime
        With LogInRS
                !GRAND_TOTAL = Result
            .Update
        End With
        
        MsgBox "Employees ID: " & txtEmployeeID.Text & " has been successfully LOG OUT!", vbInformation
        Call Timer1_Timer
init_data
    ElseIf login_am = True And logout_am = True And login_pm = True And logout_pm = True Then
    
         MsgBox "Employees ID: " & txtEmployeeID.Text & " has already logged completely  for the day!", vbInformation
         txtEmployeeID.SetFocus
    Call Timer1_Timer
init_data
    End If
    
    COMPUTE_TOTAL = False
    Call UnloadOject
End Sub

Private Sub init_data()
    Dim lngCurrentRow As Long
    Dim intNumberOfRows As Integer
    
'On Error GoTo err:

    If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from TIMELOG where DATE_LOG = #" & FormatDateTime(Now, vbShortDate) & "#", conn, adOpenKeyset, adLockPessimistic
        intNumberOfRows = rs.RecordCount
        lngCurrentRow = 0
        
        'If rs.EOF Then rs.MoveFirst
        Do While rs.EOF = False
              With msGrid
                    .Rows = intNumberOfRows
                    .Row = lngCurrentRow
                    
                     .Col = 0: .Text = rs.Fields("EM_ID").Value
                    .Col = 1: .Text = rs.Fields("IN_AM").Value
                    .Col = 2: .Text = rs.Fields("OUT_AM").Value
                    .Col = 3: .Text = rs.Fields("IN_PM").Value
                    .Col = 4: .Text = rs.Fields("OUT_PM").Value
                    
                End With
                lngCurrentRow = lngCurrentRow + 1
        rs.MoveNext
        Loop
        Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    
End Sub

Private Sub init_Grid()
    With msGrid
        .Clear
        .Cols = 5
        .Rows = 1
       
        .ColWidth(0) = 1265: .ColWidth(1) = 1200: .ColWidth(2) = 1200: .ColWidth(3) = 1200: .ColWidth(4) = 1200
    End With
End Sub
