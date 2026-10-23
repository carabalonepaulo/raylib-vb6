Attribute VB_Name = "modRaylib"
Option Explicit

' System

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Audio

Private Declare Function PlaySoundA Lib "winmm.dll" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Const SND_ASYNC = &H1
Const SND_FILENAME = &H2000
Const SND_LOOP = &H8
Const SND_NODEFAULT = &H2

' Mouse

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MouseButtonState
    IsDown As Boolean
    WasDown As Boolean
End Type

Public Enum MouseButton
    MOUSE_BUTTON_LEFT = 0
    MOUSE_BUTTON_RIGHT = 1
    MOUSE_BUTTON_MIDDLE = 2
End Enum

Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const VK_MBUTTON = &H4

' Keyboard

Private Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Private Declare Function VkKeyScanEx Lib "user32" Alias "VkKeyScanExA" (ByVal ch As Integer, ByVal dwhkl As Long) As Integer
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal idThread As Long) As Long
Private Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpString As String, ByVal nSize As Long) As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal uCode As Long, ByVal uMapType As Long) As Long

Public Enum KeyboardKey
    KEY_NULL = 0                ' Key pressed = 0, not used

    KEY_APOSTROPHE = 39         ' '
    KEY_COMMA = 44              ' ,
    KEY_MINUS = 45              ' -
    KEY_PERIOD = 46             ' .
    KEY_SLASH = 47              ' /
    KEY_ZERO = 48               ' 0
    KEY_ONE = 49                ' 1
    KEY_TWO = 50                ' 2
    KEY_THREE = 51              ' 3
    KEY_FOUR = 52               ' 4
    KEY_FIVE = 53               ' 5
    KEY_SIX = 54                ' 6
    KEY_SEVEN = 55              ' 7
    KEY_EIGHT = 56              ' 8
    KEY_NINE = 57               ' 9
    KEY_SEMICOLON = 59          ' ;
    KEY_EQUAL = 61              ' =

    KEY_A = 65                  ' A
    KEY_B = 66                  ' B
    KEY_C = 67                  ' C
    KEY_D = 68                  ' D
    KEY_E = 69                  ' E
    KEY_F = 70                  ' F
    KEY_G = 71                  ' G
    KEY_H = 72                  ' H
    KEY_I = 73                  ' I
    KEY_J = 74                  ' J
    KEY_K = 75                  ' K
    KEY_L = 76                  ' L
    KEY_M = 77                  ' M
    KEY_N = 78                  ' N
    KEY_O = 79                  ' O
    KEY_P = 80                  ' P
    KEY_Q = 81                  ' Q
    KEY_R = 82                  ' R
    KEY_S = 83                  ' S
    KEY_T = 84                  ' T
    KEY_U = 85                  ' U
    KEY_V = 86                  ' V
    KEY_W = 87                  ' W
    KEY_X = 88                  ' X
    KEY_Y = 89                  ' Y
    KEY_Z = 90                  ' Z

    KEY_LEFT_BRACKET = 91       ' [
    KEY_BACKSLASH = 92          ' \
    KEY_RIGHT_BRACKET = 93      ' ]
    KEY_GRAVE = 96              ' `

    KEY_SPACE = 32              ' Space
    KEY_ESCAPE = 27             ' Esc
    KEY_ENTER = 13              ' Enter
    KEY_TAB = 9                 ' Tab
    KEY_BACKSPACE = 8           ' Backspace
    KEY_INSERT = 45             ' Insert
    KEY_DELETE = 46             ' Delete
    KEY_RIGHT = 39              ' Right arrow
    KEY_LEFT = 37               ' Left arrow
    KEY_DOWN = 40               ' Down arrow
    KEY_UP = 38                 ' Up arrow
    KEY_PAGE_UP = 33            ' Page up
    KEY_PAGE_DOWN = 34          ' Page down
    KEY_HOME = 36               ' Home
    KEY_END = 35                ' End
    KEY_CAPS_LOCK = 20          ' Caps lock
    KEY_SCROLL_LOCK = 145       ' Scroll lock
    KEY_NUM_LOCK = 144          ' Num lock
    KEY_PRINT_SCREEN = 44       ' Print screen
    KEY_PAUSE = 19              ' Pause

    KEY_F1 = 112                ' F1
    KEY_F2 = 113                ' F2
    KEY_F3 = 114                ' F3
    KEY_F4 = 115                ' F4
    KEY_F5 = 116                ' F5
    KEY_F6 = 117                ' F6
    KEY_F7 = 118                ' F7
    KEY_F8 = 119                ' F8
    KEY_F9 = 120                ' F9
    KEY_F10 = 121               ' F10
    KEY_F11 = 122               ' F11
    KEY_F12 = 123               ' F12

    KEY_LEFT_SHIFT = 16         ' Shift esquerdo
    KEY_LEFT_CONTROL = 17       ' Ctrl esquerdo
    KEY_LEFT_ALT = 18           ' Alt esquerdo
    KEY_RIGHT_SHIFT = 16
    KEY_RIGHT_CONTROL = 17
    KEY_RIGHT_ALT = 18
End Enum

Private Type KeyState
    IsDown As Boolean
    WasDown As Boolean
    PressTime As Single
    FramesHeld As Long
End Type

' GDI+ (carregamento de imagens)

Private Type GsiStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GsiStartupInput, ByVal outputbuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As Long, Image As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As Long, Rect As Any, ByVal flags As Long, ByVal Format As Long, lockedBitmapData As Any) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As Long, lockedBitmapData As Any) As Long

Private Type BitmapData
    Width As Long
    Height As Long
    stride As Long
    PixelFormat As Long
    Scan0 As Long
    Reserved As Long
End Type

Private Const PixelFormat32bppARGB As Long = &H26200A
Private Const ImageLockModeRead As Long = &H1

Private Const vbByRef As Long = &H4000

' Types

Public Type TexCoord2D
    u As Single
    v As Single
End Type

Public Type Vector2
    x As Single
    y As Single
End Type

Public Type Rectangle
    x As Single
    y As Single
    Width As Single
    Height As Single
End Type

Public Type ColorRGBA
    R As Single
    G As Single
    B As Single
    A As Single
End Type

Public Type Texture2D
    Id As Long
    Width As Long
    Height As Long
End Type

' Render target (framebuffer). Texture guarda a cor; Depth.Id guarda o
' renderbuffer de profundidade (Depth.Width/Height nao sao usados).
Public Type RenderTexture2D
    Id As Long
    Texture As Texture2D
    Depth As Texture2D
End Type

Public Type Image
    Data() As Byte
    Width As Long
    Height As Long
End Type

Public Type GlyphRect
    Left As Single
    Width As Single
End Type

Public Type ImageFont
    Texture As Texture2D
    Glyphs() As GlyphRect
    CharList As String
    Height As Long
End Type

Public Enum NPatchMode
    NinePatch
    ThreePatchVertical
    ThreePatchHorizontal
End Enum

Public Type NPatchInfo
    Source As Rectangle
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    Layout As NPatchMode
End Type

Public Type Camera2D
    Offset As Vector2
    Target As Vector2
    Rot As Single
    Zoom As Single
End Type

' System

Public shouldQuit As Boolean
Public fps As Long

' Mouse

Private MouseButtons(0 To 2) As MouseButtonState
Private CurrentMouseX As Single
Private CurrentMouseY As Single
Private PrevMouseX As Single
Private PrevMouseY As Single
Private MouseWheelValue As Single

' Keyboard

Private Keys(0 To 255) As KeyState

Private KeyQueue(0 To 15) As Long
Private KeyQueueCount As Long
Private KeyQueueIndex As Long

Private CharQueue(0 To 15) As Long
Private CharQueueCount As Long
Private CharQueueIndex As Long

Private Const REPEAT_DELAY As Single = 0.4
Private Const REPEAT_RATE As Single = 0.05

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private LastTime As Single

' Window

Private m_hDC As Long
Private m_hRC As Long
Private mVSyncOn As Boolean

Public screenWidth As Long
Public screenHeight As Long

' Colors

Public White As ColorRGBA
Public Black As ColorRGBA
Public Red As ColorRGBA
Public Green As ColorRGBA
Public Blue As ColorRGBA
Public Yellow As ColorRGBA
Public Gray As ColorRGBA
Public Blank As ColorRGBA

' Render target (pilha de framebuffers/viewport ativos)

Private Const RENDER_STACK_MAX As Long = 16
Private mRenderFbo(0 To RENDER_STACK_MAX - 1) As Long
Private mRenderW(0 To RENDER_STACK_MAX - 1) As Long
Private mRenderH(0 To RENDER_STACK_MAX - 1) As Long
Private mRenderTop As Long

' Utils

Public Function GetFullPath(ByVal relativePath As String) As String
    Dim basePath As String
    basePath = app.path

    If Right$(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If

    If Left$(relativePath, 1) = "\" Then
        relativePath = Mid$(relativePath, 2)
    End If

    GetFullPath = basePath & relativePath
End Function

Private Function VBGetByte(ByVal Addr As Long) As Byte
    Dim B As Byte
    CopyMemory B, ByVal Addr, 1
    VBGetByte = B
End Function

Private Sub VBPutByte(ByVal Addr As Long, ByVal B As Byte)
    CopyMemory ByVal Addr, B, 1
End Sub

Sub Color(ByRef c As ColorRGBA, ByVal R As Single, ByVal G As Single, ByVal B As Single, ByVal A As Single)
    c.R = R
    c.G = G
    c.B = B
    c.A = A
End Sub

Sub Vec(ByRef v As Vector2, ByVal x As Single, ByVal y As Single)
    v.x = x
    v.y = y
End Sub

Sub Rect(ByRef R As Rectangle, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    R.x = x
    R.y = y
    R.Width = w
    R.Height = h
End Sub

Public Function GetWorldToScreen2D(ByRef position As Vector2, ByRef camera As Camera2D) As Vector2
    Dim result As Vector2

    result.x = position.x - camera.Target.x
    result.y = position.y - camera.Target.y

    If camera.Rot <> 0! Then
        Const PI As Single = 3.14159265
        Dim rad As Single
        rad = camera.Rot * (PI / 180!)

        Dim coss As Single, sinn As Single
        coss = Cos(rad)
        sinn = Sin(rad)

        Dim rx As Single, ry As Single
        rx = result.x * coss - result.y * sinn
        ry = result.x * sinn + result.y * coss
        result.x = rx
        result.y = ry
    End If

    result.x = (result.x * camera.Zoom) + camera.Offset.x
    result.y = (result.y * camera.Zoom) + camera.Offset.y

    GetWorldToScreen2D = result
End Function

Public Function GetScreenToWorld2D(ByRef position As Vector2, ByRef camera As Camera2D) As Vector2
    Dim result As Vector2

    result.x = (position.x - camera.Offset.x) / camera.Zoom
    result.y = (position.y - camera.Offset.y) / camera.Zoom

    If camera.Rot <> 0! Then
        Const PI As Single = 3.14159265
        Dim rad As Single
        rad = -camera.Rot * (PI / 180!)

        Dim coss As Single, sinn As Single
        coss = Cos(rad)
        sinn = Sin(rad)

        Dim rx As Single, ry As Single
        rx = result.x * coss - result.y * sinn
        ry = result.x * sinn + result.y * coss
        result.x = rx
        result.y = ry
    End If

    result.x = result.x + camera.Target.x
    result.y = result.y + camera.Target.y

    GetScreenToWorld2D = result
End Function

' Contexto OpenGL (todas as bindings ficam em modOpenGL)

Private Sub InitImageLoader()
    Dim token As Long
    Dim si As GsiStartupInput
    si.GdiplusVersion = 1
    GdiplusStartup token, si, 0
End Sub

Private Function InitOpenGL(ByVal hDC As Long, ByVal w As Long, ByVal h As Long) As Boolean
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim pf As Long

    With pfd
        .nSize = Len(pfd)
        .nVersion = 1
        .dwFlags = PFD_FLAG_DRAW_TO_WINDOW Or PFD_FLAG_SUPPORT_OPENGL Or PFD_FLAG_DUAL_BUFFERED
        .iPixelType = PFD_TYPE_RGBA
        .cColorBits = 24
        .cDepthBits = 16
    End With

    pf = ChoosePixelFormat(hDC, pfd)
    If pf = 0 Then Exit Function

    If SetPixelFormat(hDC, pf, pfd) = 0 Then Exit Function

    m_hRC = wglCreateContext(hDC)
    If m_hRC = 0 Then
        Debug.Print "wglCreateContext falhou"
        Exit Function
    End If

    m_hDC = hDC

    If wglMakeCurrent(hDC, m_hRC) = 0 Then
        Debug.Print "wglMakeCurrent falhou"
        Exit Function
    End If

    screenWidth = w
    screenHeight = h

    wglMakeCurrent hDC, m_hRC
    glViewport 0, 0, w, h

    glEnable GL_BLEND
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA

    InitImageLoader
    InitColors

    InitOpenGL = True
End Function

Private Sub InitColors()
    Color White, 1#, 1#, 1#, 1#
    Color Black, 0#, 0#, 0#, 1#
    Color Red, 0.9019, 0.1607, 0.2156, 1#
    Color Green, 0#, 0.8941, 0.1882, 1#
    Color Blue, 0#, 0.4745, 0.945, 1#
    Color Yellow, 0.9921, 0.9764, 0#, 1#
    Color Gray, 0.5098, 0.5098, 0.5098, 1#
    Color Blank, 0#, 0#, 0#, 0#
End Sub

Private Sub DeinitOpenGL()
    If m_hRC <> 0 Then
        wglMakeCurrent 0, 0
        wglDeleteContext m_hRC
        m_hRC = 0
    End If
End Sub

' Vsync (WGL_EXT_swap_control). Se o driver nao expuser a extensao, as
' chamadas viram no-op. Precisa de contexto corrente e double buffer.

Public Function IsVSyncEnabled() As Boolean
    IsVSyncEnabled = mVSyncOn
End Function

Public Sub SetVSync(ByVal enabled As Boolean)
    If GLVSyncSupported() = False Then Exit Sub
    mVSyncOn = enabled
    If enabled Then
        GLSetSwapInterval 1
    Else
        GLSetSwapInterval 0
    End If
End Sub

Public Sub ToggleVSync()
    If mVSyncOn Then
        SetVSync False
    Else
        SetVSync True
    End If
End Sub

Public Sub Run(ByRef TargetForm As Form, ByVal ScreenW As Long, ByVal ScreenH As Long, ByVal TargetFPS As Long)
    On Error GoTo ErrorHandler

    Load TargetForm
    TargetForm.ScaleMode = 3
    TargetForm.Width = (ScreenW * Screen.TwipsPerPixelX) + (TargetForm.Width - (TargetForm.ScaleWidth * Screen.TwipsPerPixelX))
    TargetForm.Height = (ScreenH * Screen.TwipsPerPixelY) + (TargetForm.Height - (TargetForm.ScaleHeight * Screen.TwipsPerPixelY))
    TargetForm.Left = (Screen.Width - TargetForm.Width) / 2
    TargetForm.Top = (Screen.Height - TargetForm.Height) / 2

    TargetForm.Show
    DoEvents

    If InitOpenGL(TargetForm.hDC, TargetForm.ScaleWidth, TargetForm.ScaleHeight) = False Then
        MsgBox "Failed to init OpenGL!", vbCritical
        Unload TargetForm
        Exit Sub
    End If

    Dim targetFrameTime As Single
    targetFrameTime = 1# / CSng(TargetFPS)

    Dim freq As Currency
    Dim counterStart As Currency
    Dim counterEnd As Currency
    Dim dt As Single
    Dim frameCount As Long
    Dim fpsTimer As Single

    QueryPerformanceFrequency freq
    QueryPerformanceCounter counterStart

    Game_Load

    Do While shouldQuit = False
        QueryPerformanceCounter counterEnd
        dt = CSng(counterEnd - counterStart) / CSng(freq)

        If dt < targetFrameTime Then
            Do
                QueryPerformanceCounter counterEnd
                dt = CSng(counterEnd - counterStart) / CSng(freq)
                DoEvents
            Loop While dt < targetFrameTime
        End If

        counterStart = counterEnd
        If dt > 0.25 Then dt = 0.25

        fpsTimer = fpsTimer + dt
        frameCount = frameCount + 1

        If fpsTimer >= 1# Then
            fps = frameCount
            frameCount = 0
            fpsTimer = fpsTimer - 1#
        End If

        UpdateInput
        UpdateMouse TargetForm.hwnd

        Game_Update (dt)
        Game_Draw (dt)

        DoEvents
    Loop

    Game_Unload
    Unload TargetForm
    StopMusic
    DeinitOpenGL
    Exit Sub

ErrorHandler:
    StopMusic
    DeinitOpenGL
    MsgBox "Error: " & Err.Description, vbCritical
    End
End Sub

' Audio

Public Sub PlaySound(ByVal path As String)
    PlaySoundA GetFullPath(path), 0, SND_ASYNC Or SND_FILENAME Or SND_NODEFAULT
End Sub

Public Sub PlayMusic(ByVal path As String)
    Dim fullPath As String
    Dim ext As String
    Dim mciType As String

    fullPath = GetFullPath(path)
    ext = LCase$(Mid$(fullPath, InStrRev(fullPath, ".") + 1))

    Select Case ext
    Case "mp3", "wav", "wma"
        mciType = "mpegvideo"
    Case "mid", "midi"
        mciType = "sequencer"
    Case Else
        mciType = "mpegvideo"
    End Select

    StopMusic

    mciSendString "open """ & fullPath & """ type " & mciType & " alias BGM", vbNullString, 0, 0
    mciSendString "play BGM repeat", vbNullString, 0, 0
End Sub

Public Sub StopMusic()
    mciSendString "stop BGM", vbNullString, 0, 0
    mciSendString "close BGM", vbNullString, 0, 0
End Sub

' Mouse

Public Sub UpdateMouse(ByVal targetHwnd As Long)
    Dim pt As POINTAPI

    GetCursorPos pt

    ScreenToClient targetHwnd, pt

    PrevMouseX = CurrentMouseX
    PrevMouseY = CurrentMouseY
    CurrentMouseX = pt.x
    CurrentMouseY = pt.y

    UpdateSingleButton 0, VK_LBUTTON
    UpdateSingleButton 1, VK_RBUTTON
    UpdateSingleButton 2, VK_MBUTTON

    MouseWheelValue = 0
End Sub

Private Sub UpdateSingleButton(ByVal index As Long, ByVal vk As Long)
    MouseButtons(index).WasDown = MouseButtons(index).IsDown
    If (GetAsyncKeyState(vk) And &H8000) Then
        MouseButtons(index).IsDown = True
    Else
        MouseButtons(index).IsDown = False
    End If
End Sub

Public Function IsMouseButtonPressed(ByVal button As MouseButton) As Boolean
    If button >= 0 And button <= 2 Then
        IsMouseButtonPressed = (MouseButtons(button).IsDown And Not MouseButtons(button).WasDown)
    End If
End Function

Public Function IsMouseButtonDown(ByVal button As MouseButton) As Boolean
    If button >= 0 And button <= 2 Then
        IsMouseButtonDown = MouseButtons(button).IsDown
    End If
End Function

Public Function IsMouseButtonReleased(ByVal button As MouseButton) As Boolean
    If button >= 0 And button <= 2 Then
        IsMouseButtonReleased = (Not MouseButtons(button).IsDown And MouseButtons(button).WasDown)
    End If
End Function

Public Function IsMouseButtonUp(ByVal button As MouseButton) As Boolean
    If button >= 0 And button <= 2 Then
        IsMouseButtonUp = Not MouseButtons(button).IsDown
    End If
End Function

Public Function GetMousePos() As Vector2
    Dim pos As Vector2
    pos.x = GetMouseX
    pos.y = GetMouseY
    GetMousePos = pos
End Function

Public Function GetMouseX() As Long
    GetMouseX = CLng(CurrentMouseX)
End Function

Public Function GetMouseY() As Long
    GetMouseY = CLng(CurrentMouseY)
End Function

Public Function GetMouseDeltaX() As Single
    GetMouseDeltaX = CurrentMouseX - PrevMouseX
End Function

Public Function GetMouseDeltaY() As Single
    GetMouseDeltaY = CurrentMouseY - PrevMouseY
End Function

Public Sub SetMouseWheel(ByVal value As Single)
    MouseWheelValue = value
End Sub

Public Function GetMouseWheelMove() As Single
    GetMouseWheelMove = MouseWheelValue
End Function

' Keyboard

Public Sub UpdateInput()
    Dim i As Long
    Dim currentTime As Single
    Dim dt As Single

    currentTime = GetTickCount() * 0.001
    If LastTime = 0 Then LastTime = currentTime
    dt = currentTime - LastTime
    LastTime = currentTime

    KeyQueueCount = 0
    KeyQueueIndex = 0
    CharQueueCount = 0
    CharQueueIndex = 0

    For i = 0 To 255
        Keys(i).WasDown = Keys(i).IsDown

        If (GetAsyncKeyState(i) And &H8000) Then
            Keys(i).IsDown = True
            If Not Keys(i).WasDown Then
                If KeyQueueCount < 16 Then
                    KeyQueue(KeyQueueCount) = i
                    KeyQueueCount = KeyQueueCount + 1
                End If
                Keys(i).PressTime = currentTime
                Keys(i).FramesHeld = 0
            Else
                Keys(i).FramesHeld = Keys(i).FramesHeld + 1
            End If
        Else
            Keys(i).IsDown = False
            Keys(i).PressTime = 0
            Keys(i).FramesHeld = 0
        End If
    Next i
End Sub

Public Function IsKeyPressed(ByVal key As KeyboardKey) As Boolean
    If key >= 0 And key <= 255 Then
        IsKeyPressed = (Keys(key).IsDown And Not Keys(key).WasDown)
    End If
End Function

Public Function IsKeyPressedRepeat(ByVal key As KeyboardKey) As Boolean
    If key < 0 Or key > 255 Then Exit Function

    If Keys(key).IsDown And Not Keys(key).WasDown Then
        IsKeyPressedRepeat = True
        Exit Function
    End If

    If Keys(key).IsDown Then
        Dim currentTime As Single
        currentTime = GetTickCount() * 0.001
        Dim heldTime As Single
        heldTime = currentTime - Keys(key).PressTime

        If heldTime >= REPEAT_DELAY Then
            Dim repeats As Long
            repeats = Int((heldTime - REPEAT_DELAY) / REPEAT_RATE)
            If (Keys(key).FramesHeld Mod 3) = 0 Then
                IsKeyPressedRepeat = True
            End If
        End If
    End If
End Function

Public Function IsKeyDown(ByVal key As KeyboardKey) As Boolean
    If key >= 0 And key <= 255 Then
        IsKeyDown = Keys(key).IsDown
    End If
End Function

Public Function IsKeyReleased(ByVal key As KeyboardKey) As Boolean
    If key >= 0 And key <= 255 Then
        IsKeyReleased = (Not Keys(key).IsDown And Keys(key).WasDown)
    End If
End Function

Public Function IsKeyUp(ByVal key As KeyboardKey) As Boolean
    If key >= 0 And key <= 255 Then
        IsKeyUp = Not Keys(key).IsDown
    End If
End Function

Public Function GetKeyPressed() As Long
    If KeyQueueIndex < KeyQueueCount Then
        GetKeyPressed = KeyQueue(KeyQueueIndex)
        KeyQueueIndex = KeyQueueIndex + 1
    Else
        GetKeyPressed = 0
    End If
End Function

Public Function GetCharPressed() As Long
    Dim i As Long
    For i = 65 To 90
        If IsKeyPressed(i) Then
            Dim isShift As Boolean
            isShift = (GetAsyncKeyState(&H10) And &H8000)
            If Not isShift Then
                GetCharPressed = i + 32
            Else
                GetCharPressed = i
            End If
            Exit Function
        End If
    Next i
    GetCharPressed = 0
End Function

Public Function GetKeyName(ByVal key As KeyboardKey) As String
    Dim scanCode As Long
    Dim result As Long
    Dim nameBuffer As String * 128

    scanCode = MapVirtualKey(key, 0)

    Dim lParam As Long
    lParam = scanCode * &H10000

    Select Case key
    Case &H25, &H26, &H27, &H28, &H2D, &H2E, &H24, &H23, &H21, &H22, &H5B, &H5C, &H5D
        lParam = lParam Or &H1000000
    End Select

    result = GetKeyNameText(lParam, nameBuffer, 128)
    If result > 0 Then
        GetKeyName = Left$(nameBuffer, InStr(nameBuffer, vbNullChar) - 1)
    Else
        GetKeyName = "UNKNOWN"
    End If
End Function

' Color

Private Sub RGBToHSV(ByVal R As Single, ByVal G As Single, ByVal B As Single, ByRef h As Single, ByRef s As Single, ByRef v As Single)
    Dim min As Single, max As Single, delta As Single

    min = R: If G < min Then min = G: If B < min Then min = B
    max = R: If G > max Then max = G: If B > max Then max = B

    v = max
    delta = max - min

    If max = 0! Then
        s = 0!
        h = 0!
        Exit Sub
    End If

    s = delta / max

    If delta = 0! Then
        h = 0!
        Exit Sub
    End If

    If R = max Then
        h = (G - B) / delta
    ElseIf G = max Then
        h = 2! + (B - R) / delta
    Else
        h = 4! + (R - G) / delta
    End If

    h = h * 60!
    If h < 0! Then h = h + 360!
End Sub

Private Sub HSVToRGB(ByVal h As Single, ByVal s As Single, ByVal v As Single, ByRef R As Single, ByRef G As Single, ByRef B As Single)
    If s = 0! Then
        R = v: G = v: B = v
        Exit Sub
    End If

    h = h / 60!
    Dim i As Long
    Dim f As Single, p As Single, q As Single, t As Single

    i = Fix(h)
    f = h - i
    p = v * (1! - s)
    q = v * (1! - s * f)
    t = v * (1! - s * (1! - f))

    Select Case i
    Case 0: R = v: G = t: B = p
    Case 1: R = q: G = v: B = p
    Case 2: R = p: G = v: B = t
    Case 3: R = p: G = q: B = v
    Case 4: R = t: G = p: B = v
    Case Else: R = v: G = p: B = q
    End Select
End Sub

Private Function Clamp01(ByVal value As Single) As Single
    If value < 0! Then
        Clamp01 = 0!
    ElseIf value > 1! Then
        Clamp01 = 1!
    Else
        Clamp01 = value
    End If
End Function

Public Function ColorTint(ByRef Color As ColorRGBA, ByRef tint As ColorRGBA) As ColorRGBA
    Dim result As ColorRGBA
    result.R = Clamp01(Color.R * tint.R)
    result.G = Clamp01(Color.G * tint.G)
    result.B = Clamp01(Color.B * tint.B)
    result.A = Clamp01(Color.A * tint.A)
    ColorTint = result
End Function

Public Function ColorBrightness(ByRef Color As ColorRGBA, ByVal factor As Single) As ColorRGBA
    Dim h As Single, s As Single, v As Single

    RGBToHSV Color.R, Color.G, Color.B, h, s, v

    v = v + factor
    If v < 0.1 Then v = 0.1
    If v > 1! Then v = 1!

    Dim result As ColorRGBA
    HSVToRGB h, s, v, result.R, result.G, result.B
    result.A = Color.A

    ColorBrightness = result
End Function

Public Function ColorContrast(ByRef Color As ColorRGBA, ByVal contrast As Single) As ColorRGBA
    Dim result As ColorRGBA
    contrast = contrast + 1!

    If contrast < 0! Then
        contrast = 0!
    End If

    result.R = Clamp01(((Color.R - 0.5) * contrast) + 0.5)
    result.G = Clamp01(((Color.G - 0.5) * contrast) + 0.5)
    result.B = Clamp01(((Color.B - 0.5) * contrast) + 0.5)
    result.A = Color.A
    ColorContrast = result
End Function

Public Function ColorAlpha(ByRef Color As ColorRGBA, ByVal alpha As Single) As ColorRGBA
    Dim result As ColorRGBA
    result.R = Color.R
    result.G = Color.G
    result.B = Color.B
    result.A = Clamp01(alpha)
    ColorAlpha = result
End Function

Public Function ColorAlphaBlend(ByRef dst As ColorRGBA, ByRef src As ColorRGBA, ByRef tint As ColorRGBA) As ColorRGBA
    Dim result As ColorRGBA
    Dim srcTint As ColorRGBA

    srcTint.R = src.R * tint.R
    srcTint.G = src.G * tint.G
    srcTint.B = src.B * tint.B
    srcTint.A = src.A * tint.A

    Dim alpha As Single
    alpha = srcTint.A

    result.R = Clamp01(srcTint.R * alpha + dst.R * (1! - alpha))
    result.G = Clamp01(srcTint.G * alpha + dst.G * (1! - alpha))
    result.B = Clamp01(srcTint.B * alpha + dst.B * (1! - alpha))
    result.A = Clamp01(alpha + dst.A * (1! - alpha))

    ColorAlphaBlend = result
End Function

Public Function ColorLerp(ByRef color1 As ColorRGBA, ByRef color2 As ColorRGBA, ByVal factor As Single) As ColorRGBA
    If factor < 0! Then factor = 0!
    If factor > 1! Then factor = 1!

    Dim h1 As Single, s1 As Single, v1 As Single
    Dim h2 As Single, s2 As Single, v2 As Single

    RGBToHSV color1.R, color1.G, color1.B, h1, s1, v1
    RGBToHSV color2.R, color2.G, color2.B, h2, s2, v2

    If Abs(h2 - h1) > 180! Then
        If h1 < h2 Then
            h1 = h1 + 360!
        Else
            h2 = h2 + 360!
        End If
    End If

    Dim hRes As Single, sRes As Single, vRes As Single, aRes As Single
    hRes = h1 + (h2 - h1) * factor
    If hRes >= 360! Then hRes = hRes - 360!
    If hRes < 0! Then hRes = hRes + 360!

    sRes = s1 + (s2 - s1) * factor
    vRes = v1 + (v2 - v1) * factor
    aRes = color1.A + (color2.A - color1.A) * factor

    Dim result As ColorRGBA
    HSVToRGB hRes, sRes, vRes, result.R, result.G, result.B
    result.A = aRes

    ColorLerp = result
End Function

Public Function ColorIsEqual(ByRef col1 As ColorRGBA, ByRef col2 As ColorRGBA) As Boolean
    Const EPSILON As Single = 0.00001
    ColorIsEqual = (Abs(col1.R - col2.R) < EPSILON) And _
                   (Abs(col1.G - col2.G) < EPSILON) And _
                   (Abs(col1.B - col2.B) < EPSILON) And _
                   (Abs(col1.A - col2.A) < EPSILON)
End Function

Public Function ColorToInt(ByRef Color As ColorRGBA) As Long
    Dim R As Long, G As Long, B As Long, A As Long
    R = CLng(Clamp01(Color.R) * 255!)
    G = CLng(Clamp01(Color.G) * 255!)
    B = CLng(Clamp01(Color.B) * 255!)
    A = CLng(Clamp01(Color.A) * 255!)

    ColorToInt = (R * &H1000000) Or (G * &H10000) Or (B * &H100) Or A
End Function

' Image

Public Function LoadImage(ByVal filePath As String) As Image
    Dim imgHandle As Long
    Dim w As Long, h As Long
    Dim bmpData As BitmapData
    Dim Rect(3) As Long
    Dim resultImage As Image

    resultImage.Width = 0
    resultImage.Height = 0

    If GdipLoadImageFromFile(StrPtr(GetFullPath(filePath)), imgHandle) <> 0 Then
        LoadImage = resultImage
        Exit Function
    End If

    GdipGetImageWidth imgHandle, w
    GdipGetImageHeight imgHandle, h

    Rect(0) = 0: Rect(1) = 0: Rect(2) = w: Rect(3) = h

    If GdipBitmapLockBits(imgHandle, Rect(0), 1, &H26200A, bmpData) = 0 Then
        Dim totalBytes As Long
        totalBytes = w * h * 4

        resultImage.Width = w
        resultImage.Height = h
        ReDim resultImage.Data(0 To totalBytes - 1)

        CopyMemory resultImage.Data(0), ByVal bmpData.Scan0, totalBytes

        Dim i As Long
        Dim bVal As Byte
        For i = 0 To (w * h) - 1
            Dim idx As Long
            idx = i * 4
            bVal = resultImage.Data(idx)
            resultImage.Data(idx) = resultImage.Data(idx + 2)
            resultImage.Data(idx + 2) = bVal
        Next i

        GdipBitmapUnlockBits imgHandle, bmpData
    End If

    GdipDisposeImage imgHandle
    LoadImage = resultImage
End Function

Public Function ImageGetPixel(ByRef img As Image, ByVal x As Long, ByVal y As Long) As ColorRGBA
    Dim col As ColorRGBA
    Color col, 0, 0, 0, 1#

    If x < 0 Or x >= img.Width Or y < 0 Or y >= img.Height Then Exit Function

    Dim idx As Long
    idx = (y * img.Width + x) * 4

    Color col, _
          CSng(img.Data(idx)) / 255#, _
          CSng(img.Data(idx + 1)) / 255#, _
          CSng(img.Data(idx + 2)) / 255#, _
          CSng(img.Data(idx + 3)) / 255#

    ImageGetPixel = col
End Function

Public Sub ImageSetPixel(ByRef img As Image, ByVal x As Long, ByVal y As Long, ByRef col As ColorRGBA)
    If x < 0 Or x >= img.Width Or y < 0 Or y >= img.Height Then Exit Sub

    Dim idx As Long
    idx = (y * img.Width + x) * 4

    img.Data(idx) = CByte(col.R * 255#)
    img.Data(idx + 1) = CByte(col.G * 255#)
    img.Data(idx + 2) = CByte(col.B * 255#)
    img.Data(idx + 3) = CByte(col.A * 255#)
End Sub

Public Sub ImageClear(ByRef img As Image, ByRef col As ColorRGBA)
    Dim x As Long, y As Long
    Dim R As Byte, G As Byte, B As Byte, A As Byte

    R = CByte(col.R * 255#)
    G = CByte(col.G * 255#)
    B = CByte(col.B * 255#)
    A = CByte(col.A * 255#)

    Dim rowBytes As Long
    rowBytes = img.Width * 4

    For y = 0 To img.Height - 1
        Dim rowStart As Long
        rowStart = y * rowBytes
        For x = 0 To img.Width - 1
            Dim idx As Long
            idx = rowStart + (x * 4)
            img.Data(idx) = R
            img.Data(idx + 1) = G
            img.Data(idx + 2) = B
            img.Data(idx + 3) = A
        Next x
    Next y
End Sub

Public Sub ImageDrawRect(ByRef img As Image, ByRef Rect As Rectangle, ByRef col As ColorRGBA)
    Dim rx As Long, ry As Long, rw As Long, rh As Long
    rx = CLng(Rect.x)
    ry = CLng(Rect.y)
    rw = CLng(Rect.Width)
    rh = CLng(Rect.Height)

    Dim R As Byte, G As Byte, B As Byte, A As Byte
    R = CByte(col.R * 255#)
    G = CByte(col.G * 255#)
    B = CByte(col.B * 255#)
    A = CByte(col.A * 255#)

    Dim x As Long, y As Long
    Dim startY As Long, endY As Long
    Dim startX As Long, endX As Long

    startY = IIf(ry < 0, 0, ry)
    endY = IIf((ry + rh) > img.Height, img.Height, ry + rh)

    startX = IIf(rx < 0, 0, rx)
    endX = IIf((rx + rw) > img.Width, img.Width, rx + rw)

    For y = startY To endY - 1
        Dim rowStart As Long
        rowStart = y * (img.Width * 4)
        For x = startX To endX - 1
            Dim idx As Long
            idx = rowStart + (x * 4)
            img.Data(idx) = R
            img.Data(idx + 1) = G
            img.Data(idx + 2) = B
            img.Data(idx + 3) = A
        Next x
    Next y
End Sub

Public Sub ImageDrawImage(ByRef imgDst As Image, ByRef Rect As Rectangle, ByRef imgSrc As Image, ByRef col As ColorRGBA)
    Dim dstX As Long, dstY As Long
    Dim srcW As Long, srcH As Long

    dstX = CLng(Rect.x)
    dstY = CLng(Rect.y)
    srcW = imgSrc.Width
    srcH = imgSrc.Height

    Dim px As Long, py As Long
    Dim targetX As Long, targetY As Long

    For py = 0 To srcH - 1
        targetY = dstY + py
        If targetY >= 0 And targetY < imgDst.Height Then
            For px = 0 To srcW - 1
                targetX = dstX + px
                If targetX >= 0 And targetX < imgDst.Width Then
                    Dim srcIdx As Long
                    Dim dstIdx As Long

                    srcIdx = (py * srcW + px) * 4
                    dstIdx = (targetY * imgDst.Width + targetX) * 4

                    imgDst.Data(dstIdx) = imgSrc.Data(srcIdx)
                    imgDst.Data(dstIdx + 1) = imgSrc.Data(srcIdx + 1)
                    imgDst.Data(dstIdx + 2) = imgSrc.Data(srcIdx + 2)
                    imgDst.Data(dstIdx + 3) = imgSrc.Data(srcIdx + 3)
                End If
            Next px
        End If
    Next py
End Sub

Public Function LoadTextureFromImage(ByRef img As Image) As Texture2D
    Dim tex As Texture2D
    Dim textureID As Long

    tex.Id = 0
    tex.Width = 0
    tex.Height = 0

    If img.Width <= 0 Or img.Height <= 0 Then
        LoadTextureFromImage = tex
        Exit Function
    End If

    glGenTextures 1, textureID
    glBindTexture GL_TEXTURE_2D, textureID

    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST

    glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, img.Width, img.Height, 0, GL_RGBA, GL_UNSIGNED_BYTE, ByVal VarPtr(img.Data(0))

    tex.Id = textureID
    tex.Width = img.Width
    tex.Height = img.Height

    LoadTextureFromImage = tex
End Function

' Texture

Public Function LoadTexture(ByVal filePath As String) As Texture2D
    Dim img As Image
    img = LoadImage(filePath)
    LoadTexture = LoadTextureFromImage(img)
End Function

Public Sub UnloadTexture(ByVal textureID As Long)
    If textureID <> 0 Then
        glDeleteTextures 1, textureID
    End If
End Sub

' Render target (framebuffer object)

Public Function LoadRenderTexture(ByVal Width As Long, ByVal Height As Long) As RenderTexture2D
    Dim rt As RenderTexture2D
    Dim prevFbo As Long
    Dim status As Long

    rt.Id = 0
    rt.Texture.Id = 0
    rt.Texture.Width = 0
    rt.Texture.Height = 0
    rt.Depth.Id = 0

    If GLFramebufferSupported() = False Then
        Debug.Print "FBO nao suportado por este driver"
        LoadRenderTexture = rt
        Exit Function
    End If

    If Width <= 0 Or Height <= 0 Then
        LoadRenderTexture = rt
        Exit Function
    End If

    ' guarda o framebuffer atual para restaurar no final
    glGetIntegerv GL_FRAMEBUFFER_BINDING, prevFbo

    ' framebuffer
    glGenFramebuffersEXT 1, rt.Id
    glBindFramebufferEXT GL_FRAMEBUFFER, rt.Id

    ' textura de cor
    glGenTextures 1, rt.Texture.Id
    glBindTexture GL_TEXTURE_2D, rt.Texture.Id
    glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, Width, Height, 0, GL_RGBA, GL_UNSIGNED_BYTE, 0
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE
    glFramebufferTexture2DEXT GL_FRAMEBUFFER, GL_COLOR_ATTACHMENT0, GL_TEXTURE_2D, rt.Texture.Id, 0

    ' renderbuffer de profundidade
    glGenRenderbuffersEXT 1, rt.Depth.Id
    glBindRenderbufferEXT GL_RENDERBUFFER, rt.Depth.Id
    glRenderbufferStorageEXT GL_RENDERBUFFER, GL_DEPTH_COMPONENT24, Width, Height
    glFramebufferRenderbufferEXT GL_FRAMEBUFFER, GL_DEPTH_ATTACHMENT, GL_RENDERBUFFER, rt.Depth.Id

    rt.Texture.Width = Width
    rt.Texture.Height = Height

    ' verifica enquanto o FBO ainda esta ligado
    status = glCheckFramebufferStatusEXT(GL_FRAMEBUFFER)

    ' desliga e restaura o alvo anterior
    glBindFramebufferEXT GL_FRAMEBUFFER, prevFbo
    glBindTexture GL_TEXTURE_2D, 0
    glBindRenderbufferEXT GL_RENDERBUFFER, 0

    If status <> GL_FRAMEBUFFER_COMPLETE Then
        Debug.Print "Framebuffer incompleto: &H" & Hex$(status)
        UnloadRenderTexture rt
        rt.Id = 0
        rt.Texture.Id = 0
        rt.Texture.Width = 0
        rt.Texture.Height = 0
        rt.Depth.Id = 0
    End If

    LoadRenderTexture = rt
End Function

Public Sub UnloadRenderTexture(ByRef target As RenderTexture2D)
    If target.Texture.Id <> 0 Then
        glDeleteTextures 1, target.Texture.Id
        target.Texture.Id = 0
    End If

    If target.Depth.Id <> 0 Then
        glDeleteRenderbuffersEXT 1, target.Depth.Id
        target.Depth.Id = 0
    End If

    If target.Id <> 0 Then
        glDeleteFramebuffersEXT 1, target.Id
        target.Id = 0
    End If
End Sub

' Font

Public Function LoadImageFont(ByVal filePath As String, ByVal CharList As String) As ImageFont
    Dim img As Image
    img = LoadImage(filePath)

    Dim fnt As ImageFont
    fnt.Height = img.Height
    fnt.CharList = CharList

    Dim sepCol As ColorRGBA
    sepCol = ImageGetPixel(img, 0, 0)

    ReDim fnt.Glyphs(Len(CharList) - 1)

    Dim charIndex As Long
    Dim x As Long, startX As Long
    Dim w As Long
    Dim vbCol As ColorRGBA

    charIndex = 0
    x = 0

    Do While x < img.Width And charIndex < Len(CharList)
        Do While x < img.Width
            vbCol = ImageGetPixel(img, x, 0)
            If Abs(vbCol.R - sepCol.R) < 0.05 And _
               Abs(vbCol.G - sepCol.G) < 0.05 And _
               Abs(vbCol.B - sepCol.B) < 0.05 Then
                x = x + 1
            Else
                Exit Do
            End If
        Loop

        If x >= img.Width Then Exit Do

        startX = x

        Do While x < img.Width
            vbCol = ImageGetPixel(img, x, 0)
            If Abs(vbCol.R - sepCol.R) < 0.05 And _
               Abs(vbCol.G - sepCol.G) < 0.05 And _
               Abs(vbCol.B - sepCol.B) < 0.05 Then
                Exit Do
            Else
                x = x + 1
            End If
        Loop

        w = x - startX
        If w <= 0 Then w = 1

        fnt.Glyphs(charIndex).Left = CSng(startX) / CSng(img.Width)
        fnt.Glyphs(charIndex).Width = CSng(w)

        charIndex = charIndex + 1
    Loop

    fnt.Texture = LoadTextureFromImage(img)
    LoadImageFont = fnt
End Function

' Draw

Public Sub UpdateViewport(ByVal w As Long, ByVal h As Long)
    screenWidth = w
    screenHeight = h
    glViewport 0, 0, screenWidth, screenHeight
End Sub

Public Sub BeginDrawing()
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    glOrtho 0, screenWidth, screenHeight, 0, -1, 1

    glMatrixMode GL_MODELVIEW
    glLoadIdentity
End Sub

' Render target: redireciona o desenho para um framebuffer.
'
' Dentro do BeginTextureMode as coordenadas continuam sendo y-para-baixo iguais
' as da tela: a projecao usa glOrtho(0, w, 0, h), que inverte a textura em
' relacao ao framebuffer, e o DrawTexture inverte de volta ao exibir. O efeito
' liquido e identidade, sem precisar mexer em nenhum codigo de desenho.
Public Sub BeginTextureMode(ByRef target As RenderTexture2D)
    Dim prevFbo As Long
    Dim vp(0 To 3) As Long

    If mRenderTop >= RENDER_STACK_MAX Then Exit Sub
    If target.Id = 0 Then Exit Sub

    glGetIntegerv GL_FRAMEBUFFER_BINDING, prevFbo
    glGetIntegerv GL_VIEWPORT, vp(0)

    mRenderFbo(mRenderTop) = prevFbo
    mRenderW(mRenderTop) = vp(2)
    mRenderH(mRenderTop) = vp(3)
    mRenderTop = mRenderTop + 1

    glBindFramebufferEXT GL_FRAMEBUFFER, target.Id
    glViewport 0, 0, target.Texture.Width, target.Texture.Height

    glMatrixMode GL_PROJECTION
    glPushMatrix
    glLoadIdentity
    glOrtho 0, target.Texture.Width, 0, target.Texture.Height, -1, 1

    glMatrixMode GL_MODELVIEW
    glPushMatrix
    glLoadIdentity
End Sub

Public Sub EndTextureMode()
    If mRenderTop <= 0 Then Exit Sub
    mRenderTop = mRenderTop - 1

    glMatrixMode GL_PROJECTION
    glPopMatrix

    glMatrixMode GL_MODELVIEW
    glPopMatrix

    glBindFramebufferEXT GL_FRAMEBUFFER, mRenderFbo(mRenderTop)
    glViewport 0, 0, mRenderW(mRenderTop), mRenderH(mRenderTop)
End Sub

Public Sub Clear(ByVal R As Single, ByVal G As Single, ByVal B As Single, ByVal A As Single)
    glClearColor R, G, B, A
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
End Sub

Public Sub EndDrawing()
    SwapBuffers m_hDC
End Sub

Public Sub BeginMode2D(ByRef camera As Camera2D)
    glMatrixMode GL_PROJECTION
    glPushMatrix

    glMatrixMode GL_MODELVIEW
    glPushMatrix

    glTranslatef camera.Offset.x, camera.Offset.y, 0!
    glRotatef camera.Rot, 0!, 0!, 1!
    glScalef camera.Zoom, camera.Zoom, 1!
    glTranslatef -camera.Target.x, -camera.Target.y, 0!
End Sub

Public Sub EndMode2D()
    glMatrixMode GL_PROJECTION
    glPopMatrix

    glMatrixMode GL_MODELVIEW
    glPopMatrix
End Sub

Public Sub DrawRectangle(ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByRef Color As ColorRGBA)
    glDisable GL_TEXTURE_2D

    glBegin GL_QUADS
    glColor4f Color.R, Color.G, Color.B, Color.A
    glVertex2f x, y
    glVertex2f x + Width, y
    glVertex2f x + Width, y + Height
    glVertex2f x, y + Height
    glEnd
End Sub

Public Sub DrawRectangleRec(ByRef R As Rectangle, ByRef Color As ColorRGBA)
    DrawRectangle R.x, R.y, R.Width, R.Height, Color
End Sub

Public Sub DrawRectanglePro(ByRef R As Rectangle, ByRef origin As Vector2, ByVal Rot As Single, ByRef Color As ColorRGBA)
    glPushMatrix
    glTranslatef R.x, R.y, 0!

    If Rot <> 0! Then
        glTranslatef origin.x, origin.y, 0!
        glRotatef Rot, 0!, 0!, 1!
        glTranslatef -origin.x, -origin.y, 0!
    End If

    glDisable GL_TEXTURE_2D
    glBegin GL_QUADS
    glColor4f Color.R, Color.G, Color.B, Color.A
    glVertex2f -origin.x, -origin.y
    glVertex2f -origin.x + R.Width, -origin.y
    glVertex2f -origin.x + R.Width, -origin.y + R.Height
    glVertex2f -origin.x, -origin.y + R.Height
    glEnd

    glPopMatrix
End Sub

Public Sub DrawRectangleLines(ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByRef Color As ColorRGBA)
    glDisable GL_TEXTURE_2D

    glBegin GL_LINE_LOOP
    glColor4f Color.R, Color.G, Color.B, Color.A
    glVertex2f x, y
    glVertex2f x + Width, y
    glVertex2f x + Width, y + Height
    glVertex2f x, y + Height
    glEnd
End Sub

Public Sub DrawRectangleLinesEx(ByRef R As Rectangle, ByRef thick As Single, ByRef Color As ColorRGBA)
    Dim halfThick As Single
    halfThick = thick * 0.5

    DrawRectangle R.x, R.y, R.Width, thick, Color
    DrawRectangle R.x, R.y + R.Height - thick, R.Width, thick, Color
    DrawRectangle R.x, R.y + thick, thick, R.Height - (thick * 2), Color
    DrawRectangle R.x + R.Width - thick, R.y + thick, thick, R.Height - (thick * 2), Color
End Sub

Public Sub DrawRectangleRounded(ByRef R As Rectangle, ByRef roundness As Single, ByVal segments As Long, ByRef Color As ColorRGBA)
    If roundness <= 0! Then
        DrawRectangleRec R, Color
        Exit Sub
    End If
    If roundness > 1! Then roundness = 1!

    Dim radius As Single
    If R.Width < R.Height Then
        radius = (R.Width * 0.5) * roundness
    Else
        radius = (R.Height * 0.5) * roundness
    End If

    If segments < 4 Then segments = 4

    DrawRectangle R.x + radius, R.y, R.Width - (radius * 2), R.Height, Color
    DrawRectangle R.x, R.y + radius, radius, R.Height - (radius * 2), Color
    DrawRectangle R.x + R.Width - radius, R.y + radius, radius, R.Height - (radius * 2), Color

    DrawCircleSectorSimple R.x + radius, R.y + radius, radius, 180, 270, segments, Color
    DrawCircleSectorSimple R.x + R.Width - radius, R.y + radius, radius, 270, 360, segments, Color
    DrawCircleSectorSimple R.x + radius, R.y + R.Height - radius, radius, 90, 180, segments, Color
    DrawCircleSectorSimple R.x + R.Width - radius, R.y + R.Height - radius, radius, 0, 90, segments, Color
End Sub

Public Sub DrawRectangleRoundedLines(ByRef R As Rectangle, ByRef roundness As Single, ByVal segments As Long, ByRef Color As ColorRGBA)
    DrawRectangleRoundedLinesEx R, roundness, segments, 1#, Color
End Sub

Public Sub DrawRectangleRoundedLinesEx(ByRef R As Rectangle, ByRef roundness As Single, ByVal segments As Long, ByVal thick As Single, ByRef Color As ColorRGBA)
    If roundness <= 0! Then
        DrawRectangleLinesEx R, thick, Color
        Exit Sub
    End If

    DrawRectangleLines R.x + roundness, R.y, R.Width - (roundness * 2), R.Height, Color
End Sub

Public Sub DrawTriangle(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByRef Color As ColorRGBA)
    glDisable GL_TEXTURE_2D

    glBegin GL_TRIANGLES
    glColor4f Color.R, Color.G, Color.B, Color.A
    glVertex2f x1, y1
    glVertex2f x2, y2
    glVertex2f x3, y3
    glEnd
End Sub

Public Sub DrawLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef Color As ColorRGBA, ByVal thickness As Single)
    glLineWidth thickness
    glBegin GL_LINES
    glColor4f Color.R, Color.G, Color.B, Color.A
    glVertex2f x1, y1
    glVertex2f x2, y2
    glEnd
End Sub

Public Sub DrawTexture(ByRef tex As Texture2D, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByRef tint As ColorRGBA)
    If tex.Id = 0 Then Exit Sub

    glEnable GL_TEXTURE_2D
    glBindTexture GL_TEXTURE_2D, tex.Id

    glColor4f tint.R, tint.G, tint.B, tint.A

    glBegin GL_QUADS
    glTexCoord2f 0#, 0#: glVertex2f x, y
    glTexCoord2f 1#, 0#: glVertex2f x + Width, y
    glTexCoord2f 1#, 1#: glVertex2f x + Width, y + Height
    glTexCoord2f 0#, 1#: glVertex2f x, y + Height
    glEnd

    glDisable GL_TEXTURE_2D
End Sub

Public Sub DrawTextureRec(ByRef tex As Texture2D, ByRef sourceRec As Rectangle, ByRef pos As Vector2, ByRef tint As ColorRGBA)
    If tex.Id = 0 Then Exit Sub

    glEnable GL_TEXTURE_2D
    glBindTexture GL_TEXTURE_2D, tex.Id

    glColor4f tint.R, tint.G, tint.B, tint.A

    Dim u1 As Single, v1 As Single, u2 As Single, v2 As Single

    u1 = sourceRec.x / tex.Width
    v1 = sourceRec.y / tex.Height
    u2 = (sourceRec.x + sourceRec.Width) / tex.Width
    v2 = (sourceRec.y + sourceRec.Height) / tex.Height

    glBegin GL_QUADS
    glTexCoord2f u1, v1: glVertex2f pos.x, pos.y
    glTexCoord2f u2, v1: glVertex2f pos.x + sourceRec.Width, pos.y
    glTexCoord2f u2, v2: glVertex2f pos.x + sourceRec.Width, pos.y + sourceRec.Height
    glTexCoord2f u1, v2: glVertex2f pos.x, pos.y + sourceRec.Height
    glEnd

    glDisable GL_TEXTURE_2D
End Sub

Public Sub DrawTexturePro(ByRef tex As Texture2D, ByRef sourceRec As Rectangle, ByRef dstrect As Rectangle, ByRef origin As Vector2, ByVal Rot As Single, ByRef tint As ColorRGBA)
    If tex.Id = 0 Then Exit Sub

    Const PI As Single = 3.14159265
    Dim u1 As Single, v1 As Single, u2 As Single, v2 As Single

    u1 = sourceRec.x / tex.Width
    v1 = sourceRec.y / tex.Height
    u2 = (sourceRec.x + sourceRec.Width) / tex.Width
    v2 = (sourceRec.y + sourceRec.Height) / tex.Height

    glEnable GL_TEXTURE_2D
    glBindTexture GL_TEXTURE_2D, tex.Id

    glPushMatrix

    glTranslatef dstrect.x, dstrect.y, 0!

    If Rot <> 0! Then
        glTranslatef origin.x, origin.y, 0!
        glRotatef Rot, 0!, 0!, 1!
        glTranslatef -origin.x, -origin.y, 0!
    End If

    glColor4f tint.R, tint.G, tint.B, tint.A

    glBegin GL_QUADS
    glTexCoord2f u1, v1: glVertex2f -origin.x, -origin.y
    glTexCoord2f u2, v1: glVertex2f -origin.x + dstrect.Width, -origin.y
    glTexCoord2f u2, v2: glVertex2f -origin.x + dstrect.Width, -origin.y + dstrect.Height
    glTexCoord2f u1, v2: glVertex2f -origin.x, -origin.y + dstrect.Height
    glEnd

    glPopMatrix
    glDisable GL_TEXTURE_2D
End Sub

Public Sub DrawTextureNPatch(ByRef tex As Texture2D, ByRef info As NPatchInfo, ByRef dstrect As Rectangle, ByRef origin As Vector2, ByVal Rot As Single, ByRef Color As ColorRGBA)
    If tex.Id = 0 Then Exit Sub

    Dim texWidth As Single
    Dim texHeight As Single
    texWidth = CSng(tex.Width)
    texHeight = CSng(tex.Height)

    Dim patchWidth As Single
    Dim patchHeight As Single
    patchWidth = IIf(dstrect.Width <= 0, 0!, CSng(dstrect.Width))
    patchHeight = IIf(dstrect.Height <= 0, 0!, CSng(dstrect.Height))

    Dim srcX As Single, srcY As Single, srcWidth As Single, srcHeight As Single
    srcX = CSng(info.Source.x)
    srcY = CSng(info.Source.y)
    srcWidth = CSng(info.Source.Width)
    srcHeight = CSng(info.Source.Height)

    If srcWidth < 0! Then
        srcX = srcX - srcWidth
        srcWidth = -srcWidth
    End If
    If srcHeight < 0! Then
        srcY = srcY - srcHeight
        srcHeight = -srcHeight
    End If

    If info.Layout = ThreePatchHorizontal Then patchHeight = srcHeight
    If info.Layout = ThreePatchVertical Then patchWidth = srcWidth

    Dim drawCenter As Boolean
    Dim drawMiddle As Boolean
    drawCenter = True
    drawMiddle = True

    Dim leftBorder As Single, topBorder As Single, rightBorder As Single, bottomBorder As Single
    leftBorder = CSng(info.Left)
    topBorder = CSng(info.Top)
    rightBorder = CSng(info.Right)
    bottomBorder = CSng(info.Bottom)

    If patchWidth <= (leftBorder + rightBorder) And info.Layout <> ThreePatchVertical Then
        drawCenter = False
        If (leftBorder + rightBorder) > 0! Then
            leftBorder = (leftBorder / (leftBorder + rightBorder)) * patchWidth
        End If
        rightBorder = patchWidth - leftBorder
    End If

    If patchHeight <= (topBorder + bottomBorder) And info.Layout <> ThreePatchHorizontal Then
        drawMiddle = False
        If (topBorder + bottomBorder) > 0! Then
            topBorder = (topBorder / (topBorder + bottomBorder)) * patchHeight
        End If
        bottomBorder = patchHeight - topBorder
    End If

    Dim vertAX As Single, vertAY As Single
    Dim vertBX As Single, vertBY As Single
    Dim vertCX As Single, vertCY As Single
    Dim vertDX As Single, vertDY As Single

    vertAX = -origin.x
    vertAY = -origin.y
    vertBX = -origin.x + leftBorder
    vertBY = -origin.y + topBorder
    vertCX = -origin.x + patchWidth - rightBorder
    vertCY = -origin.y + patchHeight - bottomBorder
    vertDX = -origin.x + patchWidth
    vertDY = -origin.y + patchHeight

    Dim coordAX As Single, coordAY As Single
    Dim coordBX As Single, coordBY As Single
    Dim coordCX As Single, coordCY As Single
    Dim coordDX As Single, coordDY As Single

    coordAX = srcX / texWidth
    coordAY = srcY / texHeight
    coordBX = (srcX + leftBorder) / texWidth
    coordBY = (srcY + topBorder) / texHeight
    coordCX = (srcX + srcWidth - rightBorder) / texWidth
    coordCY = (srcY + srcHeight - bottomBorder) / texHeight
    coordDX = (srcX + srcWidth) / texWidth
    coordDY = (srcY + srcHeight) / texHeight

    glEnable GL_TEXTURE_2D
    glBindTexture GL_TEXTURE_2D, tex.Id

    glPushMatrix
    glTranslatef CSng(dstrect.x), CSng(dstrect.y), 0!

    If Rot <> 0! Then
        glTranslatef origin.x, origin.y, 0!
        glRotatef Rot, 0!, 0!, 1!
        glTranslatef -origin.x, -origin.y, 0!
    End If

    glColor4f Color.R, Color.G, Color.B, Color.A

    glBegin GL_QUADS

    If info.Layout = NinePatch Then
        glTexCoord2f coordAX, coordBY: glVertex2f vertAX, vertBY
        glTexCoord2f coordBX, coordBY: glVertex2f vertBX, vertBY
        glTexCoord2f coordBX, coordAY: glVertex2f vertBX, vertAY
        glTexCoord2f coordAX, coordAY: glVertex2f vertAX, vertAY

        If drawCenter Then
            glTexCoord2f coordBX, coordBY: glVertex2f vertBX, vertBY
            glTexCoord2f coordCX, coordBY: glVertex2f vertCX, vertBY
            glTexCoord2f coordCX, coordAY: glVertex2f vertCX, vertAY
            glTexCoord2f coordBX, coordAY: glVertex2f vertBX, vertAY
        End If

        glTexCoord2f coordCX, coordBY: glVertex2f vertCX, vertBY
        glTexCoord2f coordDX, coordBY: glVertex2f vertDX, vertBY
        glTexCoord2f coordDX, coordAY: glVertex2f vertDX, vertAY
        glTexCoord2f coordCX, coordAY: glVertex2f vertCX, vertAY

        If drawMiddle Then
            glTexCoord2f coordAX, coordCY: glVertex2f vertAX, vertCY
            glTexCoord2f coordBX, coordCY: glVertex2f vertBX, vertCY
            glTexCoord2f coordBX, coordBY: glVertex2f vertBX, vertBY
            glTexCoord2f coordAX, coordBY: glVertex2f vertAX, vertBY

            If drawCenter Then
                glTexCoord2f coordBX, coordCY: glVertex2f vertBX, vertCY
                glTexCoord2f coordCX, coordCY: glVertex2f vertCX, vertCY
                glTexCoord2f coordCX, coordBY: glVertex2f vertCX, vertBY
                glTexCoord2f coordBX, coordBY: glVertex2f vertBX, vertBY
            End If

            glTexCoord2f coordCX, coordCY: glVertex2f vertCX, vertCY
            glTexCoord2f coordDX, coordCY: glVertex2f vertDX, vertCY
            glTexCoord2f coordDX, coordBY: glVertex2f vertDX, vertBY
            glTexCoord2f coordCX, coordBY: glVertex2f vertCX, vertBY
        End If

        glTexCoord2f coordAX, coordDY: glVertex2f vertAX, vertDY
        glTexCoord2f coordBX, coordDY: glVertex2f vertBX, vertDY
        glTexCoord2f coordBX, coordCY: glVertex2f vertBX, vertCY
        glTexCoord2f coordAX, coordCY: glVertex2f vertAX, vertCY

        If drawCenter Then
            glTexCoord2f coordBX, coordDY: glVertex2f vertBX, vertDY
            glTexCoord2f coordCX, coordDY: glVertex2f vertCX, vertDY
            glTexCoord2f coordCX, coordCY: glVertex2f vertCX, vertCY
            glTexCoord2f coordBX, coordCY: glVertex2f vertBX, vertCY
        End If

        glTexCoord2f coordCX, coordDY: glVertex2f vertCX, vertDY
        glTexCoord2f coordDX, coordDY: glVertex2f vertDX, vertDY
        glTexCoord2f coordDX, coordCY: glVertex2f vertDX, vertCY
        glTexCoord2f coordCX, coordCY: glVertex2f vertCX, vertCY

    ElseIf info.Layout = ThreePatchVertical Then
        glTexCoord2f coordAX, coordBY: glVertex2f vertAX, vertBY
        glTexCoord2f coordDX, coordBY: glVertex2f vertDX, vertBY
        glTexCoord2f coordDX, coordAY: glVertex2f vertDX, vertAY
        glTexCoord2f coordAX, coordAY: glVertex2f vertAX, vertAY

        If drawCenter Then
            glTexCoord2f coordAX, coordCY: glVertex2f vertAX, vertCY
            glTexCoord2f coordDX, coordCY: glVertex2f vertDX, vertCY
            glTexCoord2f coordDX, coordBY: glVertex2f vertDX, vertBY
            glTexCoord2f coordAX, coordBY: glVertex2f vertAX, vertBY
        End If

        glTexCoord2f coordAX, coordDY: glVertex2f vertAX, vertDY
        glTexCoord2f coordDX, coordDY: glVertex2f vertDX, vertDY
        glTexCoord2f coordDX, coordCY: glVertex2f vertDX, vertCY
        glTexCoord2f coordAX, coordCY: glVertex2f vertAX, vertCY

    ElseIf info.Layout = ThreePatchHorizontal Then
        glTexCoord2f coordAX, coordDY: glVertex2f vertAX, vertDY
        glTexCoord2f coordBX, coordDY: glVertex2f vertBX, vertDY
        glTexCoord2f coordBX, coordAY: glVertex2f vertBX, vertAY
        glTexCoord2f coordAX, coordAY: glVertex2f vertAX, vertAY

        If drawCenter Then
            glTexCoord2f coordBX, coordDY: glVertex2f vertBX, vertDY
            glTexCoord2f coordCX, coordDY: glVertex2f vertCX, vertDY
            glTexCoord2f coordCX, coordAY: glVertex2f vertCX, vertAY
            glTexCoord2f coordBX, coordAY: glVertex2f vertBX, vertAY
        End If

        glTexCoord2f coordCX, coordDY: glVertex2f vertCX, vertDY
        glTexCoord2f coordDX, coordDY: glVertex2f vertDX, vertDY
        glTexCoord2f coordDX, coordAY: glVertex2f vertDX, vertAY
        glTexCoord2f coordCX, coordAY: glVertex2f vertCX, vertAY
    End If

    glEnd

    glPopMatrix
    glDisable GL_TEXTURE_2D
End Sub

Public Sub DrawText(ByRef font As ImageFont, ByVal text As String, ByVal x As Single, ByVal y As Single, ByRef tint As ColorRGBA)
    Dim i As Long
    Dim curX As Single
    curX = x

    glEnable GL_TEXTURE_2D
    glBindTexture GL_TEXTURE_2D, font.Texture.Id

    glColor4f tint.R, tint.G, tint.B, tint.A

    glBegin GL_QUADS

    Dim texelWidth As Single
    texelWidth = 1! / CSng(font.Texture.Width)
    Dim fHeight As Single
    fHeight = CSng(font.Height)

    For i = 1 To Len(text)
        Dim charChar As String
        charChar = Mid$(text, i, 1)

        Dim charIdx As Long
        charIdx = InStr(1, font.CharList, charChar, vbBinaryCompare)

        If charIdx > 0 Then
            charIdx = charIdx - 1
        Else
            charIdx = 0
        End If

        Dim G As GlyphRect
        G = font.Glyphs(charIdx)

        Dim u1 As Single, u2 As Single
        u1 = G.Left + (texelWidth * 0.1)
        u2 = (G.Left + (G.Width / CSng(font.Texture.Width))) - (texelWidth * 0.1)

        glTexCoord2f u1, 1#: glVertex2f curX, y + fHeight
        glTexCoord2f u2, 1#: glVertex2f curX + G.Width, y + fHeight
        glTexCoord2f u2, 0#: glVertex2f curX + G.Width, y
        glTexCoord2f u1, 0#: glVertex2f curX, y

        curX = curX + G.Width
    Next i

    glEnd
    glDisable GL_TEXTURE_2D
End Sub

Public Sub DrawCircle(ByVal centerX As Long, ByVal centerY As Long, ByVal radius As Single, ByRef Color As ColorRGBA)
    DrawCircleSectorSimple centerX, centerY, radius, 0!, 360!, 32, Color
End Sub

Public Sub DrawCircleV(ByRef center As Vector2, ByVal radius As Single, ByRef Color As ColorRGBA)
    DrawCircleSectorSimple center.x, center.y, radius, 0!, 360!, 32, Color
End Sub

Public Sub DrawCircleGradient(ByRef center As Vector2, ByVal radius As Single, ByRef inner As ColorRGBA, ByRef outer As ColorRGBA)
    Const PI As Single = 3.14159265
    Dim i As Long
    Dim segments As Long
    Dim angleStep As Single

    segments = 32
    angleStep = 2! * PI / segments

    glDisable GL_TEXTURE_2D
    glBegin GL_TRIANGLE_FAN

    glColor4f inner.R, inner.G, inner.B, inner.A
    glVertex2f center.x, center.y

    For i = 0 To segments
        Dim currentAngle As Single
        currentAngle = i * angleStep

        glColor4f outer.R, outer.G, outer.B, outer.A
        glVertex2f center.x + Cos(currentAngle) * radius, center.y + Sin(currentAngle) * radius
    Next i
    glEnd
End Sub

Public Sub DrawCircleSector(ByRef center As Vector2, ByVal radius As Single, ByVal startAngle As Single, ByVal endAngle As Single, ByVal segments As Long, ByRef Color As ColorRGBA)
    DrawCircleSectorSimple center.x, center.y, radius, startAngle, endAngle, segments, Color
End Sub

Public Sub DrawCircleSectorLines(ByRef center As Vector2, ByVal radius As Single, ByVal startAngle As Single, ByVal endAngle As Single, ByVal segments As Long, ByRef Color As ColorRGBA)
    Const PI As Single = 3.14159265
    Dim i As Long
    Dim angleStep As Single

    If segments < 4 Then segments = 4
    angleStep = (endAngle - startAngle) * (PI / 180!) / segments

    glDisable GL_TEXTURE_2D
    glBegin GL_LINE_STRIP
    glColor4f Color.R, Color.G, Color.B, Color.A

    glVertex2f center.x, center.y

    For i = 0 To segments
        Dim currentAngle As Single
        currentAngle = (startAngle * (PI / 180!)) + (i * angleStep)
        glVertex2f center.x + Cos(currentAngle) * radius, center.y + Sin(currentAngle) * radius
    Next i

    glVertex2f center.x, center.y
    glEnd
End Sub

Public Sub DrawCircleSectorLinesEx(ByRef center As Vector2, ByVal radius As Single, ByVal startAngle As Single, ByVal endAngle As Single, ByVal segments As Long, ByVal thick As Single, ByRef Color As ColorRGBA)
    Dim t As Single
    For t = 0! To thick - 1! Step 1!
        DrawCircleSectorLines center, radius - t, startAngle, endAngle, segments, Color
    Next t
End Sub

Public Sub DrawCircleLines(ByVal centerX As Long, ByVal centerY As Long, ByVal radius As Single, ByRef Color As ColorRGBA)
    Const PI As Single = 3.14159265
    Dim i As Long
    Dim segments As Long
    Dim angleStep As Single

    segments = 32
    angleStep = 2! * PI / segments

    glDisable GL_TEXTURE_2D
    glBegin GL_LINE_LOOP
    glColor4f Color.R, Color.G, Color.B, Color.A

    For i = 0 To segments - 1
        Dim currentAngle As Single
        currentAngle = i * angleStep
        glVertex2f centerX + Cos(currentAngle) * radius, centerY + Sin(currentAngle) * radius
    Next i
    glEnd
End Sub

Public Sub DrawCircleLinesV(ByRef center As Vector2, ByVal radius As Single, ByRef Color As ColorRGBA)
    DrawCircleLines center.x, center.y, radius, Color
End Sub

Public Sub DrawCircleLinesEx(ByRef center As Vector2, ByVal radius As Single, ByVal thick As Single, ByRef Color As ColorRGBA)
    Dim t As Single
    For t = 0! To thick - 1! Step 1!
        DrawCircleLines center.x, center.y, radius - t, Color
    Next t
End Sub

Private Sub DrawCircleSectorSimple(ByVal centerX As Single, ByVal centerY As Single, ByVal radius As Single, ByVal startAngle As Single, ByVal endAngle As Single, ByVal segments As Long, ByRef Color As ColorRGBA)
    Const PI As Single = 3.14159265
    Dim i As Long
    Dim angleStep As Single

    angleStep = (endAngle - startAngle) * (PI / 180!) / segments

    glDisable GL_TEXTURE_2D
    glBegin GL_TRIANGLE_FAN
    glColor4f Color.R, Color.G, Color.B, Color.A
    glVertex2f centerX, centerY

    For i = 0 To segments
        Dim currentAngle As Single
        currentAngle = (startAngle * (PI / 180!)) + (i * angleStep)
        glVertex2f centerX + Cos(currentAngle) * radius, centerY + Sin(currentAngle) * radius
    Next i
    glEnd
End Sub
