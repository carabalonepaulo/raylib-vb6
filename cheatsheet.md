# raylib for VB6

Cheatsheet. Low-level batching/GL-emulation layer: [cheatsheet_rlgl.md](cheatsheet_rlgl.md).

- Coordinates in pixels, Y grows downward (same as raylib).
- `ColorRGBA` components are `Single` in the `0.0 .. 1.0` range.
- Text uses `ImageFont` (bitmap font). There is no TTF/font loading.
- Everything is drawn between `BeginDrawing()` and `EndDrawing()`.
- Struct helpers: `Vec v, x, y` / `Rect r, x, y, w, h` / `Color c, r, g, b, a`.
- Typical entry point:

```vb
Sub Main()
    Run frmMain, 1280, 768, 60
End Sub
```

`Run` loads the form, calls `InitWindow` (creates the GL context and initializes the
rlgl layer), calls `Game_Load`, loops `Game_Update(dt)` / `Game_Draw(dt)` each frame,
and finally calls `Game_Unload`.

# Core

```vb
' window / frame
Sub InitWindow(ByRef frmMain As Form, ByVal width As Long, ByVal height As Long)
Sub Run(ByRef TargetForm As Form, ByVal ScreenW As Long, ByVal ScreenH As Long, ByVal TargetFPS As Long)
Sub UpdateViewport(ByVal w As Long, ByVal h As Long)

' drawing
Sub BeginDrawing()
Sub EndDrawing()
Sub Clear(ByVal R As Single, ByVal G As Single, ByVal B As Single, ByVal A As Single)

' timing
Sub SetTargetFPS(ByVal target As Long)
Function GetFPS() As Long
Function GetFrameTime() As Single
Function GetTime() As Double
Sub WaitTime(ByVal seconds As Double)

' window state
Function WindowShouldClose() As Boolean
Sub CloseWindow()
Function IsWindowReady() As Boolean
Function IsWindowFocused() As Boolean
Function IsWindowResized() As Boolean
Function IsWindowFullscreen() As Boolean
Function IsWindowMinimized() As Boolean
Function IsWindowMaximized() As Boolean

' window setters
Sub SetWindowTitle(ByVal title As String)
Sub SetWindowSize(ByVal width As Long, ByVal height As Long)
Sub SetWindowPosition(ByVal x As Long, ByVal y As Long)
Sub SetWindowMinSize(ByVal width As Long, ByVal height As Long)
Sub SetWindowOpacity(ByVal opacity As Single)
Sub ToggleFullscreen()
Sub MaximizeWindow()
Sub MinimizeWindow()
Sub RestoreWindow()

' config flags / states
Sub SetConfigFlags(ByVal flags As Long)
Sub SetWindowState(ByVal flags As Long)
Sub ClearWindowState(ByVal flags As Long)
Function IsVSyncEnabled() As Boolean

' monitor
Function GetMonitorCount() As Long
Function GetMonitorWidth(ByVal monitor As Long) As Long
Function GetMonitorHeight(ByVal monitor As Long) As Long
Function GetMonitorRefreshRate(ByVal monitor As Long) As Long
Function GetCurrentMonitor() As Long
Function GetWindowScaleDPI() As Vector2

' clipboard
Function GetClipboardText() As String
Sub SetClipboardText(ByVal text As String)

' random
Function GetRandomValue(ByVal min As Long, ByVal max As Long) As Long
Sub SetRandomSeed(ByVal seed As Long)

' utils
Function GetFullPath(ByVal relativePath As String) As String
Sub TakeScreenshot(ByVal fileName As String)

' globals
screenWidth  As Long
screenHeight As Long
```

- `InitWindow` does NOT create a window: it receives the already-created form (`frmMain`) and only initializes what is needed - pixel format, GL context, viewport, blending, image loader, colors and the rlgl layer. `Run` calls it for you; you normally call `Run`, not `InitWindow`.
- `BeginDrawing` sets an ortho projection `(0, w, h, 0)` through rlgl and `EndDrawing` flushes the rlgl batch and swaps buffers. All the `Draw*` functions below record geometry into the rlgl batch instead of calling OpenGL directly. The map (`modGame.DrawMap`) is the exception: it uses its own world-space vertex array, so it flushes the batch and draws with `rlDrawVertexArray`.
- `Clear(r, g, b, a)` takes four components, not a `ColorRGBA`.
- FLAG list: `FLAG_VSYNC_HINT`, `FLAG_FULLSCREEN_MODE`, `FLAG_WINDOW_RESIZABLE`, `FLAG_WINDOW_UNDECORATED`, `FLAG_WINDOW_TRANSPARENT`, `FLAG_MSAA_4X_HINT`, `FLAG_WINDOW_HIDDEN`, `FLAG_WINDOW_ALWAYS_RUN`, `FLAG_WINDOW_MINIMIZED`, `FLAG_WINDOW_MAXIMIZED`, `FLAG_WINDOW_UNFOCUSED`, `FLAG_WINDOW_TOPMOST`, `FLAG_WINDOW_HIGHDPI`, `FLAG_WINDOW_MOUSE_PASSTHROUGH`, `FLAG_BORDERLESS_WINDOWED_MODE`, `FLAG_INTERLACED_HINT`.
- `SetConfigFlags` may be called before `Run` (flags applied at init) or during `Game_Load`. `FLAG_VSYNC_HINT` replaces the old `SetVSync`/`ToggleVSync`.

# 2D Camera mode

```vb
Sub BeginMode2D(ByRef camera As Camera2D)
Sub EndMode2D()
Function GetWorldToScreen2D(ByRef position As Vector2, ByRef camera As Camera2D) As Vector2
Function GetScreenToWorld2D(ByRef position As Vector2, ByRef camera As Camera2D) As Vector2
```

# Render texture mode (framebuffer)

```vb
Function LoadRenderTexture(ByVal Width As Long, ByVal Height As Long) As RenderTexture2D
Sub UnloadRenderTexture(ByRef target As RenderTexture2D)
Sub BeginTextureMode(ByRef target As RenderTexture2D)
Sub EndTextureMode()
```

- Inside `BeginTextureMode`/`EndTextureMode` coordinates stay Y-down like the screen: the texture is flipped into the framebuffer and `DrawTexture` flips it back on display, so no draw code has to change.
- `RenderTexture2D.Texture` is drawn with `DrawTexture` / `DrawTextureRec` / `DrawTexturePro`.
- `BeginTextureMode` flushes the pending batch before switching framebuffer, and `EndTextureMode` flushes what was drawn into the target; both save and restore projection, modelview and viewport.

# Mouse

```vb
Sub UpdateMouse(ByVal targetHwnd As Long)

Function IsMouseButtonPressed(ByVal button As MouseButton) As Boolean
Function IsMouseButtonDown(ByVal button As MouseButton) As Boolean
Function IsMouseButtonReleased(ByVal button As MouseButton) As Boolean
Function IsMouseButtonUp(ByVal button As MouseButton) As Boolean

Function GetMousePos() As Vector2
Function GetMouseX() As Long
Function GetMouseY() As Long
Function GetMouseDeltaX() As Single
Function GetMouseDeltaY() As Single

Sub SetMousePosition(ByVal x As Long, ByVal y As Long)
Sub SetMouseOffset(ByVal offsetX As Long, ByVal offsetY As Long)
Sub SetMouseScale(ByVal scaleX As Single, ByVal scaleY As Single)

Sub SetMouseWheel(ByVal value As Single)
Function GetMouseWheelMove() As Single
Function GetMouseWheelMoveV() As Vector2
```

# Cursor

```vb
Sub ShowCursor()
Sub HideCursor()
Function IsCursorHidden() As Boolean
Sub EnableCursor()
Sub DisableCursor()
Sub SetMouseCursor(ByVal cursor As MouseCursor)
```

# Keyboard

```vb
Sub UpdateInput()

Function IsKeyPressed(ByVal key As KeyboardKey) As Boolean
Function IsKeyPressedRepeat(ByVal key As KeyboardKey) As Boolean
Function IsKeyDown(ByVal key As KeyboardKey) As Boolean
Function IsKeyReleased(ByVal key As KeyboardKey) As Boolean
Function IsKeyUp(ByVal key As KeyboardKey) As Boolean

Function GetKeyPressed() As Long
Function GetCharPressed() As Long
Function GetKeyName(ByVal key As KeyboardKey) As String
```

# Shapes

```vb
' rectangles
Sub DrawRectangle(ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByRef Color As ColorRGBA)
Sub DrawRectangleRec(ByRef R As Rectangle, ByRef Color As ColorRGBA)
Sub DrawRectanglePro(ByRef R As Rectangle, ByRef origin As Vector2, ByVal Rot As Single, ByRef Color As ColorRGBA)
Sub DrawRectangleLines(ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByRef Color As ColorRGBA)
Sub DrawRectangleLinesEx(ByRef R As Rectangle, ByRef thick As Single, ByRef Color As ColorRGBA)
Sub DrawRectangleRounded(ByRef R As Rectangle, ByRef roundness As Single, ByVal segments As Long, ByRef Color As ColorRGBA)
Sub DrawRectangleRoundedLines(ByRef R As Rectangle, ByRef roundness As Single, ByVal segments As Long, ByRef Color As ColorRGBA)
Sub DrawRectangleRoundedLinesEx(ByRef R As Rectangle, ByRef roundness As Single, ByVal segments As Long, ByVal thick As Single, ByRef Color As ColorRGBA)

' circles
Sub DrawCircle(ByVal centerX As Long, ByVal centerY As Long, ByVal radius As Single, ByRef Color As ColorRGBA)
Sub DrawCircleV(ByRef center As Vector2, ByVal radius As Single, ByRef Color As ColorRGBA)
Sub DrawCircleGradient(ByRef center As Vector2, ByVal radius As Single, ByRef inner As ColorRGBA, ByRef outer As ColorRGBA)
Sub DrawCircleSector(ByRef center As Vector2, ByVal radius As Single, ByVal startAngle As Single, ByVal endAngle As Single, ByVal segments As Long, ByRef Color As ColorRGBA)
Sub DrawCircleSectorLines(ByRef center As Vector2, ByVal radius As Single, ByVal startAngle As Single, ByVal endAngle As Single, ByVal segments As Long, ByRef Color As ColorRGBA)
Sub DrawCircleSectorLinesEx(ByRef center As Vector2, ByVal radius As Single, ByVal startAngle As Single, ByVal endAngle As Single, ByVal segments As Long, ByVal thick As Single, ByRef Color As ColorRGBA)
Sub DrawCircleLines(ByVal centerX As Long, ByVal centerY As Long, ByVal radius As Single, ByRef Color As ColorRGBA)
Sub DrawCircleLinesV(ByRef center As Vector2, ByVal radius As Single, ByRef Color As ColorRGBA)
Sub DrawCircleLinesEx(ByRef center As Vector2, ByVal radius As Single, ByVal thick As Single, ByRef Color As ColorRGBA)

' triangles / lines
Sub DrawTriangle(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByRef Color As ColorRGBA)
Sub DrawLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef Color As ColorRGBA, ByVal thickness As Single)
```

# Textures

```vb
' loading
Function LoadTexture(ByVal filePath As String) As Texture2D
Function LoadTextureFromImage(ByRef img As Image) As Texture2D
Sub UnloadTexture(ByVal textureID As Long)

' drawing
Sub DrawTexture(ByRef tex As Texture2D, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByRef tint As ColorRGBA)
Sub DrawTextureRec(ByRef tex As Texture2D, ByRef sourceRec As Rectangle, ByRef pos As Vector2, ByRef tint As ColorRGBA)
Sub DrawTexturePro(ByRef tex As Texture2D, ByRef sourceRec As Rectangle, ByRef dstrect As Rectangle, ByRef origin As Vector2, ByVal Rot As Single, ByRef tint As ColorRGBA)
Sub DrawTextureNPatch(ByRef tex As Texture2D, ByRef info As NPatchInfo, ByRef dstrect As Rectangle, ByRef origin As Vector2, ByVal Rot As Single, ByRef Color As ColorRGBA)
```

- `DrawTexture` here takes explicit `Width`/`Height` (no separate tint-only overload and no origin/rotation; use `DrawTexturePro` for those).

# Text (ImageFont)

```vb
Function LoadImageFont(ByVal filePath As String, ByVal CharList As String) As ImageFont
Sub DrawText(ByRef font As ImageFont, ByVal text As String, ByVal x As Single, ByVal y As Single, ByRef tint As ColorRGBA)
```

# Images

```vb
' load / query
Function LoadImage(ByVal filePath As String) As Image
Function IsImageReady(ByRef img As Image) As Boolean

' draw primitives
Function GetImageColor(ByRef img As Image, ByVal x As Long, ByVal y As Long) As ColorRGBA
Sub ImageDrawPixel(ByRef dst As Image, ByVal x As Long, ByVal y As Long, ByRef color As ColorRGBA)
Sub ImageDrawPixelV(ByRef dst As Image, ByRef position As Vector2, ByRef color As ColorRGBA)
Sub ImageClearBackground(ByRef dst As Image, ByRef color As ColorRGBA)
Sub ImageDrawRectangle(ByRef dst As Image, ByVal posX As Long, ByVal posY As Long, ByVal width As Long, ByVal height As Long, ByRef color As ColorRGBA)
Sub ImageDrawRectangleV(ByRef dst As Image, ByRef position As Vector2, ByRef size As Vector2, ByRef color As ColorRGBA)
Sub ImageDrawRectangleRec(ByRef dst As Image, ByRef rec As Rectangle, ByRef color As ColorRGBA)
Sub ImageDraw(ByRef dst As Image, ByRef src As Image, ByRef srcRec As Rectangle, ByRef dstRec As Rectangle, ByRef tint As ColorRGBA)
Sub ImageDrawLine(ByRef dst As Image, ByVal startPosX As Long, ByVal startPosY As Long, ByVal endPosX As Long, ByVal endPosY As Long, ByRef color As ColorRGBA)
Sub ImageDrawLineV(ByRef dst As Image, ByRef start As Vector2, ByRef finish As Vector2, ByRef color As ColorRGBA)
Sub ImageDrawLineEx(ByRef dst As Image, ByRef start As Vector2, ByRef finish As Vector2, ByVal thick As Long, ByRef color As ColorRGBA)
Sub ImageDrawCircle(ByRef dst As Image, ByVal centerX As Long, ByVal centerY As Long, ByVal radius As Long, ByRef color As ColorRGBA)
Sub ImageDrawCircleV(ByRef dst As Image, ByRef center As Vector2, ByVal radius As Single, ByRef color As ColorRGBA)
Sub ImageDrawCircleLines(ByRef dst As Image, ByVal centerX As Long, ByVal centerY As Long, ByVal radius As Long, ByRef color As ColorRGBA)
Sub ImageDrawCircleLinesV(ByRef dst As Image, ByRef center As Vector2, ByVal radius As Single, ByRef color As ColorRGBA)
Sub ImageDrawRectangleLines(ByRef dst As Image, ByRef rec As Rectangle, ByVal thick As Long, ByRef color As ColorRGBA)
Sub ImageDrawTriangle(ByRef dst As Image, ByRef v1 As Vector2, ByRef v2 As Vector2, ByRef v3 As Vector2, ByRef color As ColorRGBA)
Sub ImageDrawTriangleEx(ByRef dst As Image, ByRef v1 As Vector2, ByRef v2 As Vector2, ByRef v3 As Vector2, ByRef c1 As ColorRGBA, ByRef c2 As ColorRGBA, ByRef c3 As ColorRGBA)
Sub ImageDrawTriangleLines(ByRef dst As Image, ByRef v1 As Vector2, ByRef v2 As Vector2, ByRef v3 As Vector2, ByRef color As ColorRGBA)
Sub ImageDrawTriangleFan(ByRef dst As Image, ByRef points() As Vector2, ByVal pointCount As Long, ByRef color As ColorRGBA)
Sub ImageDrawTriangleStrip(ByRef dst As Image, ByRef points() As Vector2, ByVal pointCount As Long, ByRef color As ColorRGBA)

' copy / transform
Function ImageCopy(ByRef image As Image) As Image
Function ImageFromImage(ByRef image As Image, ByRef rec As Rectangle) As Image
Sub ImageCrop(ByRef image As Image, ByRef crop As Rectangle)
Sub ImageFlipVertical(ByRef image As Image)
Sub ImageFlipHorizontal(ByRef image As Image)
Sub ImageRotateCW(ByRef image As Image)
Sub ImageRotateCCW(ByRef image As Image)
Sub ImageRotate(ByRef image As Image, ByVal degrees As Long)
Sub ImageResize(ByRef image As Image, ByVal newWidth As Long, ByVal newHeight As Long)
Sub ImageResizeNN(ByRef image As Image, ByVal newWidth As Long, ByVal newHeight As Long)
Sub ImageResizeCanvas(ByRef image As Image, ByVal newWidth As Long, ByVal newHeight As Long, ByVal offsetX As Long, ByVal offsetY As Long, ByRef fill As ColorRGBA)
Sub ImageToPOT(ByRef image As Image, ByRef fill As ColorRGBA)

' color / alpha / format
Sub ImageColorTint(ByRef image As Image, ByRef color As ColorRGBA)
Sub ImageColorInvert(ByRef image As Image)
Sub ImageColorGrayscale(ByRef image As Image)
Sub ImageColorContrast(ByRef image As Image, ByVal contrast As Single)
Sub ImageColorBrightness(ByRef image As Image, ByVal brightness As Long)
Sub ImageColorReplace(ByRef image As Image, ByRef color As ColorRGBA, ByRef replace As ColorRGBA)
Sub ImageAlphaClear(ByRef image As Image, ByRef color As ColorRGBA, ByVal threshold As Single)
Sub ImageAlphaMask(ByRef image As Image, ByRef alphaMask As Image)
Sub ImageAlphaPremultiply(ByRef image As Image)
Sub ImageFormat(ByRef image As Image, ByVal newFormat As Long)
Sub ImageMipmaps(ByRef image As Image)
Sub ImageBlurGaussian(ByRef image As Image, ByVal blurSize As Long)
Sub ImageDither(ByRef image As Image, ByVal rBpp As Long, ByVal gBpp As Long, ByVal bBpp As Long, ByVal aBpp As Long)

' load from GPU / export / read
Function LoadImageFromScreen() As Image
Function LoadImageFromTexture(ByRef texture As Texture2D) As Image
Function ExportImage(ByRef image As Image, ByVal fileName As String) As Boolean
Function ExportImageAsCode(ByRef image As Image, ByVal fileName As String) As Boolean
Function GetImageAlphaBorder(ByRef image As Image, ByVal threshold As Single) As Rectangle
Function LoadImageColors(ByRef image As Image) As ColorRGBA()
Function LoadImagePalette(ByRef image As Image, ByVal maxPaletteSize As Long, ByRef colorCount As Long) As ColorRGBA()

' generate
Function GenImageColor(ByVal width As Long, ByVal height As Long, ByRef color As ColorRGBA) As Image
Function GenImageGradientLinear(ByVal width As Long, ByVal height As Long, ByVal direction As Long, ByRef startColor As ColorRGBA, ByRef endColor As ColorRGBA) As Image
Function GenImageGradientRadial(ByVal width As Long, ByVal height As Long, ByVal density As Single, ByRef inner As ColorRGBA, ByRef outer As ColorRGBA) As Image
Function GenImageGradientSquare(ByVal width As Long, ByVal height As Long, ByVal density As Single, ByRef inner As ColorRGBA, ByRef outer As ColorRGBA) As Image
Function GenImageChecked(ByVal width As Long, ByVal height As Long, ByVal checksX As Long, ByVal checksY As Long, ByRef col1 As ColorRGBA, ByRef col2 As ColorRGBA) As Image
Function GenImageWhiteNoise(ByVal width As Long, ByVal height As Long, ByVal factor As Single) As Image
Function GenImagePerlinNoise(ByVal width As Long, ByVal height As Long, ByVal offsetX As Long, ByVal offsetY As Long, ByVal scaleFactor As Single) As Image
Function GenImageCellular(ByVal width As Long, ByVal height As Long, ByVal tileSize As Long) As Image
```

- `Image.Data` holds raw pixels in `Image.Format` (a `PixelFormat` value). The uncompressed formats are supported: GRAYSCALE, GRAY_ALPHA, R5G6B5, R8G8B8, R5G5B5A1, R4G4B4A4, R8G8B8A8 (R32 groups only define pixel size). `LoadImage` produces `PIXELFORMAT_UNCOMPRESSED_R8G8B8A8` and `Mipmaps = 1`. All the drawing and pixel functions are format-aware.
- There is no `UnloadImage`: VB6 frees the CPU-side dynamic array itself. Only GPU resources need an explicit unload.

# Colors

```vb
Function ColorTint(ByRef Color As ColorRGBA, ByRef tint As ColorRGBA) As ColorRGBA
Function ColorBrightness(ByRef Color As ColorRGBA, ByVal factor As Single) As ColorRGBA
Function ColorContrast(ByRef Color As ColorRGBA, ByVal contrast As Single) As ColorRGBA
Function ColorAlpha(ByRef Color As ColorRGBA, ByVal alpha As Single) As ColorRGBA
Function ColorAlphaBlend(ByRef dst As ColorRGBA, ByRef src As ColorRGBA, ByRef tint As ColorRGBA) As ColorRGBA
Function ColorLerp(ByRef color1 As ColorRGBA, ByRef color2 As ColorRGBA, ByVal factor As Single) As ColorRGBA
Function ColorIsEqual(ByRef col1 As ColorRGBA, ByRef col2 As ColorRGBA) As Boolean
Function ColorToInt(ByRef Color As ColorRGBA) As Long

' predefined colors (filled at startup)
' White, Black, Red, Green, Blue, Yellow, Gray, Blank
```

# Audio

```vb
Sub PlaySound(ByVal path As String)
Sub PlayMusic(ByVal path As String)
Sub StopMusic()
```

# Types

```vb
Type Vector2
    x As Single
    y As Single
End Type

Type Rectangle
    x As Single
    y As Single
    Width As Single
    Height As Single
End Type

Type ColorRGBA
    R As Single
    G As Single
    B As Single
    A As Single
End Type

Type Texture2D
    Id As Long
    Width As Long
    Height As Long
End Type

Type RenderTexture2D
    Id As Long
    Texture As Texture2D      ' color attachment, use it for drawing
    Depth As Texture2D        ' Depth.Id holds the depth renderbuffer
End Type

Type Image
    Data() As Byte            ' raw pixels in the format given by Format
    Width As Long
    Height As Long
    Mipmaps As Long           ' number of mipmap levels (1 = none)
    Format As Long            ' a PixelFormat value
End Type

Type ImageFont
    Texture As Texture2D
    Glyphs() As GlyphRect
    CharList As String
    Height As Long
End Type

Type GlyphRect
    Left As Single
    Width As Single
End Type

Type TexCoord2D
    u As Single
    v As Single
End Type

Type Camera2D
    Offset As Vector2
    Target As Vector2
    Rot As Single
    Zoom As Single
End Type

Type NPatchInfo
    Source As Rectangle
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    Layout As NPatchMode
End Type
```

# Enums

```vb
' PixelFormat
PIXELFORMAT_UNCOMPRESSED_GRAYSCALE  PIXELFORMAT_UNCOMPRESSED_GRAY_ALPHA
PIXELFORMAT_UNCOMPRESSED_R5G6B5     PIXELFORMAT_UNCOMPRESSED_R8G8B8
PIXELFORMAT_UNCOMPRESSED_R5G5B5A1   PIXELFORMAT_UNCOMPRESSED_R4G4B4A4
PIXELFORMAT_UNCOMPRESSED_R8G8B8A8
PIXELFORMAT_UNCOMPRESSED_R32        PIXELFORMAT_UNCOMPRESSED_R32G32B32
PIXELFORMAT_UNCOMPRESSED_R32G32B32A32
' compressed formats (DXT/ETC/PVRT/ASTC) are declared for compatibility only

' MouseButton
MOUSE_BUTTON_LEFT
MOUSE_BUTTON_RIGHT
MOUSE_BUTTON_MIDDLE

' MouseCursor
MOUSE_CURSOR_DEFAULT
MOUSE_CURSOR_ARROW
MOUSE_CURSOR_IBEAM
MOUSE_CURSOR_CROSSHAIR
MOUSE_CURSOR_POINTING_HAND
MOUSE_CURSOR_RESIZE_EW
MOUSE_CURSOR_RESIZE_NS
MOUSE_CURSOR_RESIZE_NWSE
MOUSE_CURSOR_RESIZE_NESW
MOUSE_CURSOR_RESIZE_ALL
MOUSE_CURSOR_NOT_ALLOWED

' NPatchMode
NinePatch
ThreePatchVertical
ThreePatchHorizontal

' KeyboardKey
KEY_NULL

' printable
KEY_APOSTROPHE  KEY_COMMA  KEY_MINUS  KEY_PERIOD  KEY_SLASH
KEY_ZERO .. KEY_NINE
KEY_SEMICOLON  KEY_EQUAL
KEY_A .. KEY_Z
KEY_LEFT_BRACKET  KEY_BACKSLASH  KEY_RIGHT_BRACKET  KEY_GRAVE

' control / navigation
KEY_SPACE  KEY_ESCAPE  KEY_ENTER  KEY_TAB  KEY_BACKSPACE
KEY_INSERT  KEY_DELETE
KEY_RIGHT  KEY_LEFT  KEY_DOWN  KEY_UP
KEY_PAGE_UP  KEY_PAGE_DOWN  KEY_HOME  KEY_END
KEY_CAPS_LOCK  KEY_SCROLL_LOCK  KEY_NUM_LOCK  KEY_PRINT_SCREEN  KEY_PAUSE

' function keys
KEY_F1 .. KEY_F12

' modifiers
KEY_LEFT_SHIFT  KEY_LEFT_CONTROL  KEY_LEFT_ALT
KEY_RIGHT_SHIFT  KEY_RIGHT_CONTROL  KEY_RIGHT_ALT
```
