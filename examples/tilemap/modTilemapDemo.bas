Attribute VB_Name = "modTilemapDemo"
Option Explicit

' A small tilemap demo: loads map 0 (assets/maps/0.ini) and draws it with
' rlDrawVertexArray, the same path modMap uses internally. Shows camera
' movement (WASD / Space) and the vertex-array draw call.

Sub Main()
    If InitWindow(frmMain, 800, 600, "Tilemap demo") = False Then
        MsgBox "Failed to init OpenGL!", vbCritical
        Exit Sub
    End If
    SetExitKey KEY_NULL

    Dim map As Tilemap
    Dim camera As Camera2D
    Dim font As ImageFont
    Dim cameraSpeed As Single
    Dim dt As Single
    Dim targetX As Single
    Dim targetY As Single

    map = LoadTilemap(0, "..\..")
    font = LoadImageFont("..\..\assets\font_regular_16.png", "? ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()_+-=[]{}|;':"",./<>~`")
    DebugInit font
    camera.Zoom = 2!
    camera.Rot = 0!
    cameraSpeed = 1000

    Do While WindowShouldClose() = False
        dt = GetFrameTime()
        targetX = 0
        targetY = 0
        If IsKeyDown(KEY_W) Then targetY = targetY + cameraSpeed * dt
        If IsKeyDown(KEY_S) Then targetY = targetY - cameraSpeed * dt
        If IsKeyDown(KEY_A) Then targetX = targetX + cameraSpeed * dt
        If IsKeyDown(KEY_D) Then targetX = targetX - cameraSpeed * dt

        camera.Offset.x = camera.Offset.x + targetX
        camera.Offset.y = camera.Offset.y + targetY

        If IsKeyPressed(KEY_SPACE) Then
            camera.Offset.x = 0
            camera.Offset.y = 0
        End If

        BeginDrawing
        Clear 0.1, 0.1, 0.1, 1#

        BeginMode2D camera
        DrawMap map
        DrawRectangle 0, 0, map.TileWidth, map.TileHeight, Red
        DrawRectangle 0, map.Height * map.TileHeight - map.TileHeight, map.TileWidth, map.TileHeight, Red
        DrawRectangle map.Width * map.TileWidth - map.TileWidth, 0, map.TileWidth, map.TileHeight, Red
        DrawRectangle map.Width * map.TileWidth - map.TileWidth, map.Height * map.TileHeight - map.TileHeight, map.TileWidth, map.TileHeight, Red
        EndMode2D

        DrawText font, "FPS: " & CStr(GetFPS()), 10, 10, White
        DebugDrawOverlay
        EndDrawing
    Loop

    CloseWindow
End Sub
