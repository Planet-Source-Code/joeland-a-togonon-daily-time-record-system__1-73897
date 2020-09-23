VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmpDetails 
   Caption         =   "Employee's Masterlist"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9720
   ScaleWidth      =   10875
   Begin VB.Frame fraHeader 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   10095
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   345
         Left            =   6480
         TabIndex        =   2
         Top             =   230
         Width           =   1200
      End
      Begin VB.ComboBox cboFields 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtSearch 
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
         Left            =   3840
         TabIndex        =   1
         Top             =   240
         Width           =   2475
      End
      Begin VB.Image picArrow 
         Height          =   255
         Left            =   3480
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
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
         TabIndex        =   11
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame fraLeft 
      Height          =   8175
      Left            =   120
      TabIndex        =   12
      Top             =   1155
      Width           =   1335
      Begin VB.CommandButton Command3 
         Caption         =   "Export"
         Height          =   420
         Left            =   75
         TabIndex        =   6
         Top             =   1680
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New"
         Height          =   405
         Left            =   75
         TabIndex        =   3
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         Height          =   420
         Left            =   75
         TabIndex        =   4
         Top             =   720
         Width           =   1200
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Delete"
         Height          =   420
         Left            =   75
         TabIndex        =   5
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Refresh"
         Height          =   420
         Left            =   75
         TabIndex        =   7
         Top             =   7200
         Width           =   1200
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Close"
         Height          =   420
         Left            =   75
         TabIndex        =   8
         Top             =   7680
         Width           =   1200
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   7935
      Left            =   1560
      TabIndex        =   9
      Top             =   1200
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   13996
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   8880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpDetails.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRecSum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1560
      TabIndex        =   14
      Top             =   9120
      Width           =   585
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Masterlist"
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
      TabIndex        =   13
      Top             =   120
      Width           =   1935
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
Attribute VB_Name = "frmEmpDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim srcItem                        As ListItem
Dim srcRecord                      As String
Dim srcDetails                  As Variant
Dim srcSQL                         As String

Private Sub cmdSearch_Click()
On Error Resume Next

strSQL = "SELECT INFO.* " & _
      "FROM INFO " & _
      "WHERE (((" & cboFields.Text & ") Like '" & txtSearch.Text & "%'))"

Set RS_Details = New ADODB.Recordset
If RS_Details.State = adStateOpen Then RS_Details.Close
RS_Details.Open strSQL, conn, adOpenDynamic, adLockOptimistic
Call FillListview
End Sub

Private Sub Command1_Click()
CommandPass "New"
End Sub

Private Sub Command2_Click()
CommandPass "Update"
End Sub

Private Sub Command6_Click()
CommandPass "Delete"
End Sub

Private Sub Command7_Click()
CommandPass "Refresh"
End Sub

Private Sub Command8_Click()
CommandPass "Close"
End Sub

Private Sub Form_Activate()
On Error Resume Next
InitCombo
Me.WindowState = vbMaximized

End Sub

Private Sub Form_Load()
On Error GoTo err_trap
picArrow.Picture = MDIMain.i16x16.ListImages(6).Picture

Set lvList.SmallIcons = i16x16
Set lvList.Icons = i16x16



srcSQL = "SELECT INFO.* " & _
            " FROM INFO " & _
            " ORDER BY INFO.EM_ID ASC "

Set RS_Details = New ADODB.Recordset
If RS_Details.State = adStateOpen Then RS_Details.Close
RS_Details.Open srcSQL, conn, adOpenDynamic, adLockPessimistic

srcDetails = "NONE"
srcRecord = vbNullString

Call FillListview
Call RefreshRecSum

Exit Sub
err_trap:
    MsgBox "Error Number: " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set frmEmpDetails = Nothing
Set RS_Details = Nothing
End Sub

Private Sub lvList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvList.Sorted And _
        ColumnHeader.Index - 1 = lvList.SortKey Then
        lvList.SortOrder = 1 - lvList.SortOrder
    Else
        lvList.SortOrder = lvwAscending
        lvList.SortKey = ColumnHeader.Index - 1
    End If
    lvList.Sorted = True
End Sub

Private Sub lvList_DblClick()
On Error Resume Next
CommandPass "Update"
End Sub

Private Sub lvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
srcDetails = lvList.SelectedItem.Index
srcRecord = lvList.ListItems.Item(srcDetails).Text
Call RefreshRecSum
End Sub

Private Sub txtSearch_GotFocus()
Highlight txtSearch
End Sub


Private Sub txtSearch_KeyPress(KeyAscii As Integer)

On Error Resume Next

strSQL = "SELECT INFO.* " & _
      "FROM INFO " & _
      "WHERE (((" & cboFields.Text & ") Like '" & txtSearch.Text & "%'))"

Set RS_Details = New ADODB.Recordset
If RS_Details.State = adStateOpen Then RS_Details.Close
RS_Details.Open strSQL, conn, adOpenDynamic, adLockOptimistic
Call FillListview

If txtSearch.Text = Empty Then
    Call Command7_Click
End If

End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat
    Case "New" 'New
            With frmEmpDetailsAE
                .State = AddStateMode
                .Show vbModal
            End With
    Case "Update" 'Update
            If srcRecord = vbNullString Then
                MsgBox "Invalid selection.Can't proceed to this operation!", vbExclamation, Me.Caption
                Exit Sub
            Else
                With frmEmpDetailsAE
                    .State = EditStateMode
                    .PK = srcRecord
                    .Show vbModal
                End With
            End If
            
    Case "Delete" 'Delete
            If lvList.ListItems.Count < 1 Then
            MsgBox "There's no record to delete!", vbExclamation, Me.Caption
            Exit Sub
            End If
            
            If srcRecord = vbNullString Then
            MsgBox "Invalid selection.Can't proceed to this operation!", vbExclamation, Me.Caption
            Exit Sub
            End If
            
            If MsgBox("Are you sure you want to delete this record?", vbCritical + vbYesNo, Me.Caption) = vbYes Then
                SQL_DELETE "DELETE FROM INFO WHERE EM_ID='" & srcRecord & "'"
                MsgBox "Selected record successfully deleted!", vbInformation, Me.Caption
                Call ReloadListview
            Else
                Exit Sub
            End If
            
    Case "Refresh" 'Refresh
           Call ReloadListview
           
            
    Case "Close" 'Close
            Unload Me
End Select
Exit Sub
errPerformWhat:
     MsgBox "Error Number:" & err.Number & vbNewLine & _
            "Description:" & err.Description, vbExclamation, Me.Caption
End Sub

Public Sub FillListview()
On Error Resume Next
With lvList
    .View = lvwReport
    
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Employee No.", 1300
    .ColumnHeaders.Add , , "PASSWORD", 0
    .ColumnHeaders.Add , , "Name", 2800
    .ColumnHeaders.Add , , "Department", 2000
    .ColumnHeaders.Add , , "Designation", 2000
    .ColumnHeaders.Add , , "Corporation", 2400
    .ColumnHeaders.Add , , "Work Start", 1300
    .ColumnHeaders.Add , , "Work End", 1300
    .ColumnHeaders.Add , , "Photos", 0

    
    .ListItems.Clear
    Do While Not RS_Details.EOF
    Set srcItem = .ListItems.Add(, , RS_Details.Fields("EM_ID"), 1, 1)
        srcItem.SubItems(1) = RS_Details.Fields("Password")
        srcItem.SubItems(2) = RS_Details.Fields("Name")
        srcItem.SubItems(3) = RS_Details.Fields("Department")
        srcItem.SubItems(4) = RS_Details.Fields("Designation")
        srcItem.SubItems(5) = RS_Details.Fields("Corporation")
        srcItem.SubItems(6) = RS_Details.Fields("WORK_TIMEB")
        srcItem.SubItems(7) = RS_Details.Fields("WORK_TIMEE")
        srcItem.SubItems(8) = RS_Details.Fields("PICTURE")
    RS_Details.MoveNext
    Loop
End With
End Sub


Public Sub ReloadListview()
On Error Resume Next
srcSQL = " SELECT INFO.* " & _
            " FROM INFO " & _
            " ORDER BY INFO.EM_ID ASC"

Set RS_Details = New ADODB.Recordset
If RS_Details.State = adStateOpen Then RS_Details.Close
RS_Details.Open srcSQL, conn, adOpenDynamic, adLockOptimistic

srcDetails = "NONE"
srcRecord = vbNullString

Call FillListview
Call RefreshRecSum

End Sub

Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record: " & srcDetails & " of " & lvList.ListItems.Count
End Sub
Public Sub InitCombo()
On Error Resume Next
    With cboFields
        .Clear
        .AddItem "EM_ID"
        .AddItem "Name"
        .AddItem "Department"
        .AddItem "Designation"
        .AddItem "Corporation"
        .ListIndex = 0
    End With
End Sub


