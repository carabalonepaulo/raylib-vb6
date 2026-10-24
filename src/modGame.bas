Attribute VB_Name = "modGame"
Dim sprite As Texture2D
Dim src As Rectangle
Dim dst As Vector2
Dim dstrect As Rectangle
Dim origin As Vector2
Dim Rot As Single
Dim font As ImageFont
Dim sepRed As ColorRGBA
Dim mouseColor As ColorRGBA
Dim camera As Camera2D

Dim cameraMode As Byte
Dim cameraSpeed As Single

Sub Main()
    SetConfigFlags FLAG_VSYNC_HINT

    If InitWindow(frmMain, 1280, 768, "raylib") = False Then
        MsgBox "Failed to init OpenGL!", vbCritical
        Exit Sub
    End If

    SetExitKey KEY_NULL

    Game_Load

    Do While WindowShouldClose() = False
        Game_Update GetFrameTime()

        BeginDrawing
        Game_Draw GetFrameTime()
        EndDrawing
    Loop

    Game_Unload
    CloseWindow
End Sub

Sub Game_Load()
    Color mouseColor, 1, 0, 0, 0.3
    sprite = LoadTexture("assets/001-Fighter01.png")
    rect src, 0#, 0#, 32#, 48#
    Vec dst, 0, 32#
    rect dstrect, 300#, 100#, 32#, 48#

    font = LoadImageFont("assets/font_regular_16.png", "? ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()_+-=[]{}|;':"",./<>~`��������������������������������")

    cameraSpeed = 1000
    camera.Zoom = 2!
    camera.Rot = 0!

    DebugInit font
End Sub

Sub Game_Unload()
End Sub

Sub Game_Update(ByVal dt As Single)
    If IsKeyPressed(KEY_F1) Then DebugToggle

    Rot = Rot + (90 * dt)

    If IsMouseButtonDown(MOUSE_BUTTON_LEFT) Then
        mouseColor = ColorBrightness(mouseColor, 1 * dt)
    End If

    If IsMouseButtonDown(MOUSE_BUTTON_RIGHT) Then
        mouseColor = ColorBrightness(mouseColor, -1 * dt)
    End If

    Dim targetX As Single, targetY As Single
    If IsKeyDown(KEY_W) Then targetY = targetY + cameraSpeed * dt
    If IsKeyDown(KEY_S) Then targetY = targetY - cameraSpeed * dt
    If IsKeyDown(KEY_A) Then targetX = targetX + cameraSpeed * dt
    If IsKeyDown(KEY_D) Then targetX = targetX - cameraSpeed * dt

    camera.Offset.x = camera.Offset.x + targetX
    camera.Offset.y = camera.Offset.y + targetY

    If IsKeyPressed(KEY_DOWN) Then
        If cameraMode > 0 Then
            cameraMode = cameraMode - 1
        End If
    End If

    If IsKeyPressed(KEY_UP) Then
        If cameraMode < 2 Then
            cameraMode = cameraMode + 1
        End If
    End If

    If cameraMode <= 0 Then
        camera.Zoom = 0.5
    End If

    If cameraMode = 1 Then
        camera.Zoom = 1
    End If

    If cameraMode >= 2 Then
        camera.Zoom = 2
    End If

    If IsKeyPressed(KEY_SPACE) Then
        camera.Offset.x = 0
        camera.Offset.y = 0
    End If
End Sub

Sub Game_Draw(ByVal dt As Single)
    Clear 0.1, 0.1, 0.1, 1#

    BeginMode2D camera

    DrawTextureRec sprite, src, dst, White
    DrawTexturePro sprite, src, dstrect, origin, Rot, White

    Dim screenPos As Vector2
    screenPos = GetMousePos

    Dim worldPos As Vector2
    worldPos = GetScreenToWorld2D(screenPos, camera)

    DrawCircle worldPos.x, worldPos.y, 5#, mouseColor
    EndMode2D

    DrawText font, "FPS: " & CStr(GetFPS()), 10, 10, White

    DebugDrawOverlay
End Sub

