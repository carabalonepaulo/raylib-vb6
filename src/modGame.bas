Attribute VB_Name = "modGame"
Option Explicit

Sub Main()
    If InitWindow(frmMain, 800, 600, "Raylib") = False Then
        MsgBox "Failed to init OpenGL!", vbCritical
        Exit Sub
    End If
    SetExitKey KEY_NULL

    Do While WindowShouldClose() = False
        BeginDrawing
        Clear 0.1, 0.1, 0.1, 1#
        EndDrawing
    Loop

    CloseWindow
End Sub
