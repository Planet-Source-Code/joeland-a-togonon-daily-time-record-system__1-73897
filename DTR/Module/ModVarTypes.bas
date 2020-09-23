Attribute VB_Name = "ModVarTypes"
Option Explicit


Public Enum FORM_STATE
    AddStateMode = 0
    EditStateMode = 1
End Enum


Public Type Employee
    idnumber As String
    Fname As String
    mname As String
    sname As String
    cnumber As String
    Picture As String
End Type

Global tymrec(0 To 31) As String

Global tymrecAMIN(0 To 31) As String
Global tymrecPMIN(0 To 31) As String
Global tymrecAMOUT(0 To 31) As String
Global tymrecPMOUT(0 To 31) As String

Public Type USER_INFO
    USERID                              As String
    USERNAME                            As String
    PASSWORD                            As String
    FULLNAME                            As String
    USERACCESSID                        As String
    ISONLINE                            As Boolean
End Type
