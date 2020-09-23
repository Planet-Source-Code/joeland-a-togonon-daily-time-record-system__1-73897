VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUnit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
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
      Left            =   3155
      TabIndex        =   1
      Top             =   480
      Width           =   1875
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
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3375
      Left            =   780
      TabIndex        =   2
      Top             =   960
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
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
      Left            =   1275
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":1848
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnit.frx":4C56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Height          =   3300
      Left            =   45
      TabIndex        =   3
      Top             =   960
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   5821
      ButtonWidth     =   1191
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "i16x16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
            Key             =   "Update"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Export"
            Key             =   "Export"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "Close"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Product/SKU Units"
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
      Left            =   105
      TabIndex        =   5
      Top             =   120
      Width           =   2175
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
      Left            =   75
      TabIndex        =   4
      Top             =   480
      Width           =   1140
   End
   Begin VB.Image picArrow 
      Height          =   255
      Left            =   2805
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcRecord               As String
Dim srcUnitCD        As Variant


Private Sub Form_Activate()
On Error Resume Next
txtSearch.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo err
CenterForm frmUnit
InitCombo

Set lvList.Icons = MDIMain.i16x162
Set lvList.SmallIcons = MDIMain.i16x162

picArrow.Picture = MDIMain.i16x16.ListImages(7).Picture

strSQL = "SELECT tblunit.* " & _
         "FROM tblunit " & _
         "ORDER BY tblunit.UnitCD ASC"

Set RS_UNIT = New ADODB.Recordset
If RS_UNIT.State = adStateOpen Then RS_UNIT.Close
RS_UNIT.Open strSQL, CN, adOpenDynamic, adLockOptimistic


If RS_UNIT.RecordCount < 1 Then
    MsgBox "No Records on the database!", vbExclamation, Me.Caption
    Exit Sub
Else
    Call FillListview
End If

srcRecord = vbNullString

Exit Sub
err:
    MsgBox "Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbExclamation, "Connection Failure"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmUnit = Nothing
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
srcUnitCD = lvList.SelectedItem.Index
srcRecord = lvList.ListItems.Item(srcUnitCD).Text
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "New": CommandPass "New"
    Case "Update": CommandPass "Update"
    Case "Delete": CommandPass "Delete"
    Case "Refresh": CommandPass "Refresh"
    Case "Export": CommandPass "Export"
    Case "Close": CommandPass "Close"
End Select
End Sub

Private Sub txtSearch_Change()
On Error Resume Next

strSQL = "SELECT tblunit.* " & _
        "FROM tblunit " & _
        "WHERE (((" & cboFields.Text & ") Like '" & txtSearch.Text & "%'))"

Set RS_UNIT = New ADODB.Recordset
If RS_UNIT.State = adStateOpen Then RS_UNIT.Close
RS_UNIT.Open strSQL, CN, adOpenDynamic, adLockOptimistic
Call FillListview

End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat
    Case "New" 'New
            With frmUnitAE
                .State = AddStateMode
                .Show vbModal
            End With
    Case "Update" 'Edit
            If Trim(srcRecord) = vbNullString Then
                MsgBox "Can't proceed to the operation!No active recored selected.", vbExclamation, Me.Caption
                Exit Sub
            Else
                With frmUnitAE
                    .State = EditStateMode
                    .PK = srcRecord
                    .Show vbModal
                End With
            End If
            
    Case "Delete" 'Delete
            If lvList.ListItems.Count < 1 Then
            MsgBox "There's no record to modify or delete!", vbExclamation, Me.Caption
            Exit Sub
            End If
            
            If Trim(srcRecord) = vbNullString Then
            MsgBox "No selected record for deletion!Please check it!", vbExclamation, Me.Caption
            Exit Sub
            End If
            
            If MsgBox("Are you sure you want to delete" & Space(1) & lvList.SelectedItem.SubItems(1) & "?", vbCritical + vbYesNo, Me.Caption) = vbYes Then
                SQLExecute "DELETE FROM tblunit WHERE UnitCD='" & srcRecord & "'"
                MsgBox "Selected record successfully deleted!", vbInformation, Me.Caption
                Call Form_Load
            Else
                Exit Sub
            End If
    Case "Query" ' Query
            If lvList.ListItems.Count < 1 Then MsgBox "No record to search.", vbExclamation: Exit Sub
            txtSearch.SetFocus
            
    Case "Refresh" 'Refresh
           Form_Load
           
    Case "Export" 'Preview
            With lvList
                If .ListItems.Count = 0 Then
                    MsgBox "There's no records to export!Please check it.", vbExclamation, Me.Caption
                    Exit Sub
                End If
            End With
                         
            xlsFilename = ""
            
            With MDIMain.CDExporter
                .Filter = "Excel Workbook (.xls)"
                .ShowSave
            xlsFilename = .FileName
            End With
            
            If Trim(xlsFilename) = "" Then
            MsgBox "Action cancelled!Failed to export records!", vbExclamation, Me.Caption
            Exit Sub
            End If
            
            
            Call ExportListview(lvList, xlsFilename)
            MsgBox "Records successfully exported!", vbInformation, Me.Caption
            xlsFilename = ""
            Form_Load
            
    Case "Close" 'Close
            Unload Me
End Select
Exit Sub
errPerformWhat:
     MsgBox "Error:" & err.Number & vbNewLine & _
            "Description:" & err.Description, vbExclamation, Me.Caption
End Sub

Public Sub FillListview()
On Error Resume Next
With lvList
    .FullRowSelect = True
    .GridLines = True
    .View = lvwReport
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Name", 1500
    .ColumnHeaders.Add , , "Description", 2000

    
    .ListItems.Clear
    Do While Not RS_UNIT.EOF
        Set lvItem = .ListItems.Add(, , RS_UNIT.Fields("Name"), 1, 1)
            lvItem.SubItems(1) = RS_UNIT.Fields("Description")
        RS_UNIT.MoveNext
    Loop
End With
End Sub
Public Sub InitCombo()
On Error Resume Next
    With cboFields
        .Clear
        .AddItem "UnitCD"
        .AddItem "Description"
        .AddItem "Remarks"
        .ListIndex = 0
    End With
End Sub





