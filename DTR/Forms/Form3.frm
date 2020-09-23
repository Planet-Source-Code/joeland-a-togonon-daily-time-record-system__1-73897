VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmpRep 
   Caption         =   "Generate Report"
   ClientHeight    =   10620
   ClientLeft      =   2685
   ClientTop       =   1875
   ClientWidth     =   11430
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   10620
   ScaleWidth      =   11430
   Begin VB.Frame fraHeader 
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   14655
      Begin VB.CheckBox chkSummary 
         Caption         =   "Summary"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11040
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Query all"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkUsedDateRange 
         Caption         =   "Enable date range query"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtPosition 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtDept 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtEM_ID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         MaxLength       =   12
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Enabled         =   0   'False
         Height          =   375
         Left            =   12960
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPickerDateEnd 
         Height          =   315
         Left            =   12720
         TabIndex        =   25
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   50528257
         CurrentDate     =   36892
      End
      Begin MSComCtl2.DTPicker DTPickerDateBegin 
         Height          =   315
         Left            =   10800
         TabIndex        =   26
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   50528257
         CurrentDate     =   36892
      End
      Begin VB.Label Label5 
         Caption         =   "LOGGED FROM:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   28
         Top             =   645
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "TO:"
         Height          =   255
         Left            =   12360
         TabIndex        =   27
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Designation:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   24
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Department:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Employee ID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quick Search:"
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
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
      Begin VB.CommandButton cmdDTR 
         Caption         =   "DTR"
         Enabled         =   0   'False
         Height          =   405
         Left            =   75
         TabIndex        =   29
         Top             =   720
         Width           =   1200
      End
      Begin VB.CommandButton cmdCLose 
         Caption         =   "Close"
         Height          =   405
         Left            =   75
         TabIndex        =   7
         Top             =   1680
         Width           =   1200
      End
      Begin VB.CommandButton cmdPrintSummary 
         Caption         =   "Summary"
         Enabled         =   0   'False
         Height          =   405
         Left            =   75
         TabIndex        =   6
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton cmdPrintEmployeesLog 
         Caption         =   "Individual"
         Enabled         =   0   'False
         Height          =   405
         Left            =   75
         TabIndex        =   5
         Top             =   240
         Width           =   1200
      End
   End
   Begin MSComctlLib.ListView lvwSummary 
      Height          =   7935
      Left            =   1560
      TabIndex        =   9
      Top             =   1560
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   13996
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee No."
         Object.Width           =   3369
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   7003
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Total Work Hour"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Total Late Hour"
         Object.Width           =   3969
      EndProperty
   End
   Begin MSComctlLib.ListView lvwLog 
      Height          =   7935
      Left            =   1560
      TabIndex        =   8
      Top             =   1560
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   13996
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Employee No."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Login A.M."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Logout A.M."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Total log A.M."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Login P.M."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Logout P.M."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Total log P.M."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Total worked hour"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Total hour late"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Remarks"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.TextBox txtCorporation 
      Height          =   285
      Left            =   6120
      TabIndex        =   30
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Attendance Masterlist"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   195
      TabIndex        =   15
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL HOUR LATE:"
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   9600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL WORKED HOUR:"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   9600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblTotalHrWork 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   9600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblTotalHrLate 
      Alignment       =   2  'Center
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
      Left            =   12120
      TabIndex        =   10
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   120
      Top             =   120
      Width           =   7275
   End
End
Attribute VB_Name = "frmEmpRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New Recordset
Dim timein As Date, timeout As Date, TimeDifferent As Date, timein01 As Date, timeout01 As Date, TimeDifferent01 As Date
Dim HourDiff As Integer, MinuteDiff As Integer, SecondDiff As Integer, HourDiff01 As Integer, MinuteDiff01 As Integer, SecondDiff01 As Integer
Dim Result, Result01, EMID, COMPANYNAME, ADDRESS As String
Dim SINGLERECORDQUERY As Boolean
Dim INFO_JOIN_LOGRS As ADODB.Recordset
Private Sub QueryEmployeeInfo()
    Call Info
    
    On Error GoTo Errorhandler:
    
    InfoRS.Find "EM_ID='" + txtEM_ID.Text + "'", 0, 1
    
    With InfoRS
        txtName.Text = !Name
        txtDept.Text = !Department
        txtPosition.Text = !Designation
        txtCorporation.Text = !corporation
    End With
    InfoRS.Close
    Set InfoRS = Nothing
        Exit Sub
Errorhandler:
        If err.Number = 3021 Then
            MsgBox "Employee No.: " & txtEM_ID.Text & " has not found!", vbInformation
            txtEM_ID.Text = ""
            txtEM_ID.SetFocus
        Else
            Resume Next
        End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDTR_Click()
    Call PrintEmployeesLog2
End Sub

Private Sub cmdPrintEmployeesLog_Click()
    
    rptEmployeesLog.Orientation = rptOrientLandscape
End Sub

Private Sub cmdPrintSummary_Click()
    Call PrintSummary
      rptSummary.Orientation = rptOrientLandscape
End Sub

Private Sub Form_Load()

    Call DBLOAD
    SINGLERECORDQUERY = False
    
End Sub

Private Sub Form_Terminate()
    Form1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.Show
End Sub

Private Sub txtEM_ID_Change()
    If txtEM_ID.Text = "" And chkAll.Value = 0 Then
        cmdPrintEmployeesLog.Enabled = False
'        cmdDTR.Enabled = False
        txtName.Text = ""
        txtDept.Text = ""
        txtPosition.Text = ""
        txtCorporation.Text = ""
    Else
        cmdPrintEmployeesLog.Enabled = True
'        cmdDTR.Enabled = True
    End If
End Sub

Private Sub txtEM_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SINGLERECORDQUERY = True
         chkAll.Value = 0
         lvwSummary.Visible = False
         lvwLog.Visible = True
         cmdPrintEmployeesLog.Enabled = True
         cmdPrintSummary.Enabled = False
        Call QueryEmployeeInfo

    End If
End Sub


Private Sub LoadObjects()
    Label7.Visible = True
    lblTotalHrWork.Visible = True
    Label8.Visible = True
    lblTotalHrLate.Visible = True
End Sub

Private Sub UnLoadObjects()
    Label7.Visible = False
    lblTotalHrWork.Visible = False
    Label8.Visible = False
    lblTotalHrLate.Visible = False
End Sub
