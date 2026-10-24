Attribute VB_Name = "modShaderDemo"
Option Explicit

' A small shader demo: the first frame of the RMXP-style 001-Fighter01.png
' sprite sheet, drawn with a single shader loaded from two files
' (assets/wave.vs + assets/tint.fs) - vertex wave plus a tint.

Sub Main()
    If InitWindow(frmMain, 480, 320, "Shader demo") = False Then
        MsgBox "Failed to init OpenGL!", vbCritical
        Exit Sub
    End If
    SetExitKey KEY_NULL

    Dim tex As Texture2D
    Dim sh As Shader
    Dim src As Rectangle
    Dim dst As Rectangle
    Dim origin As Vector2
    Dim tint As Vector4
    Dim t As Single

    tex = LoadTexture("..\..\assets\001-Fighter01.png")

    ' First frame of the 4x4 sheet: top-left cell.
    Rect src, 0#, 0#, 32#, 48#
    Rect dst, 208#, 112#, 64#, 96#
    Vec origin, 32#, 48#

    tint.x = 0.2
    tint.y = 1#
    tint.z = 0.4
    tint.W = 1#

    ' VS and FS live in assets/ as separate files, loaded together.
    sh = LoadShader("..\..\assets\wave.vs", "..\..\assets\tint.fs")
    If IsShaderValid(sh) = False Then
        MsgBox "Shader failed to load/compile", vbCritical
        CloseWindow
        Exit Sub
    End If
    SetShaderValueV4 sh, GetShaderLocation(sh, "tint"), tint

    Do While WindowShouldClose() = False
        PollInputEvents
        t = GetTime()

        BeginDrawing
        Clear 0.05, 0.05, 0.05, 1#

        BeginShaderMode sh
        SetShaderValueF sh, GetShaderLocation(sh, "time"), CSng(t)
        DrawTexturePro tex, src, dst, origin, 0, White
        EndShaderMode

        EndDrawing
    Loop

    UnloadShader sh
    UnloadTexture tex.Id
    CloseWindow
End Sub
