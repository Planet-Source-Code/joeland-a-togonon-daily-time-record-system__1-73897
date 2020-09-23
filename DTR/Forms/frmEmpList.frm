VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmpList 
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7935
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
      Begin VB.CommandButton Command5 
         Caption         =   "Close"
         Height          =   420
         Left            =   75
         TabIndex        =   6
         Top             =   2160
         Width           =   1200
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Refresh"
         Height          =   420
         Left            =   75
         TabIndex        =   5
         Top             =   1680
         Width           =   1200
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   420
         Left            =   75
         TabIndex        =   4
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit"
         Height          =   420
         Left            =   75
         TabIndex        =   3
         Top             =   720
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New"
         Height          =   420
         Left            =   75
         TabIndex        =   2
         Top             =   240
         Width           =   1200
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   2640
      TabIndex        =   0
      Top             =   1200
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7646
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmEmpList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
