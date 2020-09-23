Attribute VB_Name = "ModVariables"
Option Explicit

Public strServer     As String
Public conn As ADODB.Connection
Public LogInRS As ADODB.Recordset
Public CompanyRS As ADODB.Recordset
Public InfoRS As ADODB.Recordset
Public SumRS As ADODB.Recordset
Public SecRS As ADODB.Recordset
Public DateRangeRS As ADODB.Recordset
Public RS_Details As ADODB.Recordset
Public RS_Department As ADODB.Recordset
Public RS_Designation As ADODB.Recordset
Public RS_Corporation As ADODB.Recordset
Public RS_LogInRS As ADODB.Recordset

Public RS_USER                              As New ADODB.Recordset

Public FORMNAME As String

Public COMMAND_DELETE                       As New ADODB.Command
Public COMMAND_INSERT                   As New ADODB.Command
Public COMMAND_UPDATE                   As New ADODB.Command

Public strSQL      As String


Public lvItem                           As Variant

