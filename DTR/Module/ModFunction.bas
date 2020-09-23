Attribute VB_Name = "ModFunction"
Option Explicit

Public Sub LoadForm(ByRef srcForm As Form)
    srcForm.Show
    srcForm.WindowState = vbMaximized
    srcForm.SetFocus
End Sub

Public Sub CenterForm(ByRef srcForm As Form)
On Error Resume Next
    With srcForm
    .Move (Screen.Width - srcForm.Width) / 2, (Screen.Height - srcForm.Height) / 2
    End With
End Sub


Public Sub Highlight(ByRef srcText)
On Error Resume Next
    With srcText
        .SelStart = 0
        .SelLength = Len(srcText.Text)
    End With
End Sub
