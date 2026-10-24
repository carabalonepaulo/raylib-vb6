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
Dim map As Tilemap

Sub Main()
    Run frmMain, 1280, 768, 60
End Sub

Sub Game_Load()
    SetConfigFlags FLAG_VSYNC_HINT

    Color mouseColor, 1, 0, 0, 0.3
    sprite = LoadTexture("assets/001-Fighter01.png")
    Rect src, 0#, 0#, 32#, 48#
    Vec dst, 0, 32#
    Rect dstrect, 300#, 100#, 32#, 48#

    font = LoadImageFont("assets/font2.png", " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()_+-=[]{}|;':"",./<>?~`çÇáéíóúÁÉÍÓÚãõÃÕâêôÂÊÔàèìòùÀÈÌÒÙ")

    cameraSpeed = 1000
    camera.Zoom = 2!
    camera.Rot = 0!

    map = LoadTilemap(0)

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

    camera.Offset.x = camera.Offset.x + CInt(targetX)
    camera.Offset.y = camera.Offset.y + CInt(targetY)

    ClampCamera camera, map

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
    BeginDrawing
    Clear 0.1, 0.1, 0.1, 1#

    BeginMode2D camera

    DrawTextureRec sprite, src, dst, White
    DrawTexturePro sprite, src, dstrect, origin, Rot, White

    DrawMap map

    DrawRectangle 0, 0, map.TileWidth, map.TileHeight, Red
    DrawRectangle 0, map.Height * map.TileHeight - map.TileHeight, map.TileWidth, map.TileHeight, Red
    DrawRectangle map.Width * map.TileWidth - map.TileWidth, 0, map.TileWidth, map.TileHeight, Red
    DrawRectangle map.Width * map.TileWidth - map.TileWidth, map.Height * map.TileHeight - map.TileHeight, map.TileWidth, map.TileHeight, Red

    Dim screenPos As Vector2
    screenPos = GetMousePos

    Dim worldPos As Vector2
    worldPos = GetScreenToWorld2D(screenPos, camera)

    DrawCircle worldPos.x, worldPos.y, 5#, mouseColor
    EndMode2D

    DrawText font, "FPS: " & CStr(GetFPS()), 10, 10, White

    DebugDrawOverlay
    EndDrawing
End Sub

Private Sub ClampCamera(ByRef c As Camera2D, ByRef m As Tilemap)
    Dim mapWidthPx As Single, mapHeightPx As Single
    mapWidthPx = CSng(m.Width) * CSng(m.TileWidth)
    mapHeightPx = CSng(m.Height) * CSng(m.TileHeight)

    Dim minOffsetX As Single, minOffsetY As Single
    minOffsetX = -((mapWidthPx * camera.Zoom) - screenWidth)
    minOffsetY = -((mapHeightPx * camera.Zoom) - screenHeight)

    If (mapWidthPx * c.Zoom) > screenWidth Then
        If c.Offset.x > 0 Then c.Offset.x = 0
        If c.Offset.x < minOffsetX Then c.Offset.x = minOffsetX
    Else
        c.Offset.x = 0
    End If

    If (mapHeightPx * c.Zoom) > screenHeight Then
        If c.Offset.y > 0 Then c.Offset.y = 0
        If c.Offset.y < minOffsetY Then c.Offset.y = minOffsetY
    Else
        c.Offset.y = 0
    End If
End Sub

Private Sub DrawMap(ByRef map As Tilemap)
    If map.TotalVertices <= 0 Then Exit Sub
    If map.Tileset.Id = 0 Then Exit Sub

    rlDrawRenderBatchActive

    glEnable GL_TEXTURE_2D
    glBindTexture GL_TEXTURE_2D, map.Tileset.Id

    glColor4f White.r, White.g, White.b, White.a

    glEnableClientState GL_VERTEX_ARRAY
    glEnableClientState GL_TEXTURE_COORD_ARRAY

    glVertexPointer 2, GL_FLOAT, 0, ByVal VarPtr(map.Vertices(0))
    glTexCoordPointer 2, GL_FLOAT, 0, ByVal VarPtr(map.TexCoords(0))

    rlDrawVertexArray RL_QUADS, 0, map.TotalVertices

    glDisableClientState GL_TEXTURE_COORD_ARRAY
    glDisableClientState GL_VERTEX_ARRAY
    glDisable GL_TEXTURE_2D
End Sub
