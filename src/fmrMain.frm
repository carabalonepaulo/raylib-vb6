VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   689
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CloseWindow
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Dim newWidth As Long
    Dim newHeight As Long
    
    newWidth = Me.ScaleWidth
    newHeight = Me.ScaleHeight
    
    If newWidth > 0 And newHeight > 0 Then
        UpdateViewport newWidth, newHeight
    End If
End Sub















