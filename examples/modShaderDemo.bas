Attribute VB_Name = "modShaderDemo"
Option Explicit

' A small shader demo: the first frame of the RMXP-style 001-Fighter01.png
' sprite sheet, drawn twice - on the left with a vertex shader that waves the
' vertices, on the right with a fragment shader that applies a tint.

Sub Main()
    If InitWindow(frmMain, 480, 320, "Shader demo") = False Then
        MsgBox "Failed to init OpenGL!", vbCritical
        Exit Sub
    End If
    SetExitKey KEY_NULL

    Dim tex As Texture2D
    Dim vsWave As String
    Dim fsTint As String
    Dim shWave As Shader
    Dim shTint As Shader
    Dim src As Rectangle
    Dim dstWave As Rectangle
    Dim dstTint As Rectangle
    Dim origin As Vector2
    Dim tint As Vector4
    Dim t As Single

    tex = LoadTexture("..\assets\001-Fighter01.png")

    ' First frame of the 4x4 sheet: top-left cell.
    Rect src, 0#, 0#, 32#, 48#
    Rect dstWave, 80#, 100#, 64#, 96#
    Rect dstTint, 280#, 100#, 64#, 96#
    Vec origin, 32#, 48#

    tint.x = 0.2
    tint.y = 1#
    tint.z = 0.4
    tint.W = 1#

    vsWave = "#version 110" & vbCrLf & _
             "uniform mat4 mvp;" & vbCrLf & _
             "uniform float time;" & vbCrLf & _
             "void main()" & vbCrLf & _
             "{" & vbCrLf & _
             "    vec4 p = gl_Vertex;" & vbCrLf & _
             "    p.x = p.x + sin(p.y*0.05 + time)*8.0;" & vbCrLf & _
             "    gl_Position = mvp*p;" & vbCrLf & _
             "    gl_FrontColor = gl_Color;" & vbCrLf & _
             "    gl_TexCoord[0] = gl_MultiTexCoord0;" & vbCrLf & _
             "}" & vbCrLf

    fsTint = "#version 110" & vbCrLf & _
             "uniform sampler2D texture0;" & vbCrLf & _
             "uniform vec4 tint;" & vbCrLf & _
             "void main()" & vbCrLf & _
             "{" & vbCrLf & _
             "    vec4 c = texture2D(texture0, gl_TexCoord[0].xy);" & vbCrLf & _
             "    gl_FragColor = c*tint;" & vbCrLf & _
             "}" & vbCrLf

    shWave = LoadShaderFromMemory(vsWave, "")
    shTint = LoadShaderFromMemory("", fsTint)
    SetShaderValueV4 shTint, GetShaderLocation(shTint, "tint"), tint

    Do While WindowShouldClose() = False
        PollInputEvents
        t = GetTime()

        BeginDrawing
        Clear 0.05, 0.05, 0.05, 1#

        BeginShaderMode shWave
        SetShaderValueF shWave, GetShaderLocation(shWave, "time"), CSng(t)
        DrawTexturePro tex, src, dstWave, origin, 0, White
        EndShaderMode

        BeginShaderMode shTint
        DrawTexturePro tex, src, dstTint, origin, 0, White
        EndShaderMode

        EndDrawing
    Loop

    UnloadShader shWave
    UnloadShader shTint
    UnloadTexture tex.Id
    CloseWindow
End Sub
