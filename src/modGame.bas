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

Sub Main()
    Run frmMain, InitialScreenWidth, InitialScreenHeight, TargetFPS
End Sub

Sub Game_Load()
    PlayMusic "assets/BGM 5 - Feel.wav"

    mouseColor = Red
    sprite = LoadTexture("assets/001-Fighter01.png")
    Rect src, 0#, 0#, 32#, 48#
    Vec dst, 0, 32#
    Rect dstrect, 300#, 100#, 32#, 48#

    font = LoadImageFont("assets/font2.png", " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()_+-=[]{}|;':"",./<>?~`Á«·ÈÌÛ˙¡…Õ”⁄„ı√’‚ÍÙ¬ ‘‡ËÏÚ˘¿»Ã“Ÿ")
    
    camera.Zoom = 2!
    camera.Rot = 0!
End Sub

Sub Game_Unload()
End Sub

Sub Game_Update(ByVal dt As Single)
    Rot = Rot + (90 * dt)

    If IsMouseButtonDown(MOUSE_BUTTON_LEFT) Then
        mouseColor = ColorBrightness(mouseColor, 1 * dt)
    End If

    If IsMouseButtonDown(MOUSE_BUTTON_RIGHT) Then
        mouseColor = ColorBrightness(mouseColor, -1 * dt)
    End If
End Sub

Sub Game_Draw(ByVal dt As Single)
    BeginDrawing
    Clear 0.1, 0.1, 0.1, 1#

    BeginMode2D camera

    'DrawRectangle 0#, 0#, 32#, 32#, White
    DrawTexture sprite, 100, 100, sprite.Width, sprite.Height, White
    DrawTextureRec sprite, src, dst, White
    DrawTexturePro sprite, src, dstrect, origin, Rot, White
    'DrawText font, "porra que sofrimento miser·vel", 0, 0, White
    'DrawRectangleRounded src, 20#, 16, White

    Dim screenPos As Vector2
    screenPos = GetMousePos
    
    Dim worldPos As Vector2
    worldPos = GetScreenToWorld2D(screenPos, camera)

    DrawCircle worldPos.X, worldPos.Y, 10#, mouseColor
    'DrawRectangle GetMouseX, GetMouseY, 10, 10, mouseColor
    EndMode2D
    
    DrawText font, "FPS: " & CStr(fps), 10, 10, White
    EndDrawing
End Sub

