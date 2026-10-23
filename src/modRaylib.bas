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
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal fileName As Long, Image As Long) As Long
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
    scan0 As Long
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
    r As Single
    g As Single
    b As Single
    a As Single
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
    texture As Texture2D
    depth As Texture2D
End Type

Public Type Image
    Data() As Byte
    Width As Long
    Height As Long
End Type

Public Type GlyphRect
    left As Single
    Width As Single
End Type

Public Type ImageFont
    texture As Texture2D
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
    left As Long
    top As Long
    right As Long
    bottom As Long
    Layout As NPatchMode
End Type

Public Type Camera2D
    Offset As Vector2
    target As Vector2
    Rot As Single
    Zoom As Single
End Type

' System

Private mShouldQuit As Boolean
Private mFps As Long

' Mouse

Private MouseButtons(0 To 2) As MouseButtonState
Private CurrentMouseX As Single
Private CurrentMouseY As Single
Private PrevMouseX As Single
Private PrevMouseY As Single
Private mMouseWheelVX As Single
Private mMouseWheelVY As Single
Private mMouseOffsetX As Long
Private mMouseOffsetY As Long
Private mMouseScaleX As Single
Private mMouseScaleY As Single
Private mCursorHidden As Boolean

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
Private m_hWnd As Long

Private mWindowReady As Boolean
Private mWindowResized As Boolean
Private mFullscreen As Boolean
Private mConfigFlags As Long

Private mMinSizeW As Long
Private mMinSizeH As Long
Private mSubclassed As Boolean
Private mOldWndProc As Long
Private mSavedStyle As Long
Private mSavedExStyle As Long
Private mSavedRect As Rect

Private mTargetFPS As Long
Private mFrameTime As Single
Private mTime As Double

Private mMonitors() As MONITORINFO
Private mMonitorHandles() As Long
Private mMonitorCount As Long
Private mMonitorsReady As Boolean

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

' Core: janela / monitor / clipboard / cursor (declaracoes Win32)

Private Type Rect
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As Rect
    rcWork As Rect
    dwFlags As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
    dmPanningWidth As Long
    dmPanningHeight As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' ConfigFlags (valores identicos aos da raylib)
Public Const FLAG_VSYNC_HINT As Long = &H40&
Public Const FLAG_FULLSCREEN_MODE As Long = &H2&
Public Const FLAG_WINDOW_RESIZABLE As Long = &H4&
Public Const FLAG_WINDOW_UNDECORATED As Long = &H8&
Public Const FLAG_WINDOW_TRANSPARENT As Long = &H10&
Public Const FLAG_MSAA_4X_HINT As Long = &H20&
Public Const FLAG_WINDOW_HIDDEN As Long = &H80&
Public Const FLAG_WINDOW_ALWAYS_RUN As Long = &H100&
Public Const FLAG_WINDOW_MINIMIZED As Long = &H200&
Public Const FLAG_WINDOW_MAXIMIZED As Long = &H400&
Public Const FLAG_WINDOW_UNFOCUSED As Long = &H800&
Public Const FLAG_WINDOW_TOPMOST As Long = &H1000&
Public Const FLAG_WINDOW_HIGHDPI As Long = &H2000&
Public Const FLAG_WINDOW_MOUSE_PASSTHROUGH As Long = &H4000&
Public Const FLAG_BORDERLESS_WINDOWED_MODE As Long = &H8000&
Public Const FLAG_INTERLACED_HINT As Long = &H10000

Public Enum MouseCursor
    MOUSE_CURSOR_DEFAULT = 0
    MOUSE_CURSOR_ARROW = 1
    MOUSE_CURSOR_IBEAM = 2
    MOUSE_CURSOR_CROSSHAIR = 3
    MOUSE_CURSOR_POINTING_HAND = 4
    MOUSE_CURSOR_RESIZE_EW = 5
    MOUSE_CURSOR_RESIZE_NS = 6
    MOUSE_CURSOR_RESIZE_NWSE = 7
    MOUSE_CURSOR_RESIZE_NESW = 8
    MOUSE_CURSOR_RESIZE_ALL = 9
    MOUSE_CURSOR_NOT_ALLOWED = 10
End Enum

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_WNDPROC As Long = -4

Private Const WS_CAPTION As Long = &HC00000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20
Private Const WS_EX_TOPMOST As Long = &H8

Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2

Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

Private Const SW_HIDE As Long = 0
Private Const SW_SHOW As Long = 5
Private Const SW_MINIMIZE As Long = 6
Private Const SW_MAXIMIZE As Long = 3
Private Const SW_RESTORE As Long = 9

Private Const LWA_ALPHA As Long = &H2
Private Const WM_GETMINMAXINFO As Long = &H24
Private Const MONITOR_DEFAULTTONEAREST As Long = 2
Private Const ENUM_CURRENT_SETTINGS As Long = -1
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Const CF_TEXT As Long = 1
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40

Private Const IDC_ARROW As Long = 32512
Private Const IDC_IBEAM As Long = 32513
Private Const IDC_WAIT As Long = 32514
Private Const IDC_CROSS As Long = 32515
Private Const IDC_SIZENWSE As Long = 32642
Private Const IDC_SIZENESW As Long = 32643
Private Const IDC_SIZEWE As Long = 32644
Private Const IDC_SIZENS As Long = 32645
Private Const IDC_SIZEALL As Long = 32646
Private Const IDC_NO As Long = 32648
Private Const IDC_HAND As Long = 32649

Private Declare Sub CopyMemoryPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Private Declare Function SetWindowTextA Lib "user32" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function AdjustWindowRectEx Lib "user32" (lpRect As Rect, ByVal dwStyle As Long, ByVal bMenu As Long, ByVal dwExStyle As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hDC As Long, ByVal lprcClip As Long, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long
Private Declare Function GetMonitorInfoA Lib "user32" (ByVal hMonitor As Long, lpmi As MONITORINFO) As Long
Private Declare Function MonitorFromWindow Lib "user32" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettingsA Lib "user32" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As DEVMODE) As Long

Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function apiShowCursor Lib "user32" Alias "ShowCursor" (ByVal bShow As Long) As Long
Private Declare Function apiSetCursor Lib "user32" Alias "SetCursor" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursorA Lib "user32" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByVal scan0 As Long, bitmap As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal fileName As Long, ByRef clsidEncoder As GUID, ByVal encoderParams As Long) As Long

' Utils

Public Function GetFullPath(ByVal relativePath As String) As String
    Dim basePath As String
    basePath = app.path

    If right$(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If

    If left$(relativePath, 1) = "\" Then
        relativePath = Mid$(relativePath, 2)
    End If

    GetFullPath = basePath & relativePath
End Function

Private Function VBGetByte(ByVal Addr As Long) As Byte
    Dim b As Byte
    CopyMemory b, ByVal Addr, 1
    VBGetByte = b
End Function

Private Sub VBPutByte(ByVal Addr As Long, ByVal b As Byte)
    CopyMemory ByVal Addr, b, 1
End Sub

Sub Color(ByRef c As ColorRGBA, ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    c.r = r
    c.g = g
    c.b = b
    c.a = a
End Sub

Sub Vec(ByRef v As Vector2, ByVal x As Single, ByVal y As Single)
    v.x = x
    v.y = y
End Sub

Sub Rect(ByRef r As Rectangle, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    r.x = x
    r.y = y
    r.Width = w
    r.Height = h
End Sub

Public Function GetWorldToScreen2D(ByRef position As Vector2, ByRef camera As Camera2D) As Vector2
    Dim result As Vector2

    result.x = position.x - camera.target.x
    result.y = position.y - camera.target.y

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

    result.x = result.x + camera.target.x
    result.y = result.y + camera.target.y

    GetScreenToWorld2D = result
End Function

' Contexto OpenGL (todas as bindings ficam em modOpenGL)

Private Sub InitImageLoader()
    Dim token As Long
    Dim si As GsiStartupInput
    si.GdiplusVersion = 1
    GdiplusStartup token, si, 0
End Sub

Public Function InitWindow(ByRef frmMain As Form, ByVal Width As Long, ByVal Height As Long) As Boolean
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim pf As Long
    Dim hDC As Long

    hDC = frmMain.hDC

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
    m_hWnd = frmMain.hwnd

    If wglMakeCurrent(hDC, m_hRC) = 0 Then
        Debug.Print "wglMakeCurrent falhou"
        Exit Function
    End If

    screenWidth = Width
    screenHeight = Height

    glViewport 0, 0, Width, Height

    glEnable GL_BLEND
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA

    InitImageLoader
    InitColors
    rlglInit

    InitWindow = True
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

' Interno: usado por SetConfigFlags/SetWindowState (FLAG_VSYNC_HINT). Se o
' driver nao expuser WGL_EXT_swap_control, a chamada vira no-op.
Private Sub SetVSyncInternal(ByVal enabled As Boolean)
    If GLVSyncSupported() = False Then Exit Sub
    mVSyncOn = enabled
    If enabled Then
        GLSetSwapInterval 1
    Else
        GLSetSwapInterval 0
    End If
End Sub

Public Sub Run(ByRef TargetForm As Form, ByVal ScreenW As Long, ByVal ScreenH As Long, ByVal TargetFPS As Long)
    On Error GoTo ErrorHandler

    Load TargetForm
    TargetForm.ScaleMode = 3
    TargetForm.Width = (ScreenW * Screen.TwipsPerPixelX) + (TargetForm.Width - (TargetForm.ScaleWidth * Screen.TwipsPerPixelX))
    TargetForm.Height = (ScreenH * Screen.TwipsPerPixelY) + (TargetForm.Height - (TargetForm.ScaleHeight * Screen.TwipsPerPixelY))
    TargetForm.left = (Screen.Width - TargetForm.Width) / 2
    TargetForm.top = (Screen.Height - TargetForm.Height) / 2

    TargetForm.Show
    DoEvents

    If InitWindow(TargetForm, ScreenW, ScreenH) = False Then
        MsgBox "Failed to init OpenGL!", vbCritical
        Unload TargetForm
        Exit Sub
    End If

    ResetCoreState
    SetTargetFPS TargetFPS
    mWindowReady = True
    ApplyConfigFlags

    Dim freq As Currency
    Dim counterStart As Currency
    Dim counterEnd As Currency
    Dim dt As Single
    Dim frameCount As Long
    Dim fpsTimer As Single

    QueryPerformanceFrequency freq
    QueryPerformanceCounter counterStart

    Game_Load

    Do While mShouldQuit = False
        QueryPerformanceCounter counterEnd
        dt = CSng(counterEnd - counterStart) / CSng(freq)

        If mTargetFPS > 0 Then
            Dim targetFrameTime As Single
            targetFrameTime = 1# / CSng(mTargetFPS)
            If dt < targetFrameTime Then
                Do
                    QueryPerformanceCounter counterEnd
                    dt = CSng(counterEnd - counterStart) / CSng(freq)
                    DoEvents
                Loop While dt < targetFrameTime
            End If
        Else
            DoEvents
        End If

        counterStart = counterEnd
        If dt > 0.25 Then dt = 0.25

        mFrameTime = dt
        mTime = mTime + dt

        fpsTimer = fpsTimer + dt
        frameCount = frameCount + 1

        If fpsTimer >= 1# Then
            mFps = frameCount
            frameCount = 0
            fpsTimer = fpsTimer - 1#
        End If

        UpdateInput
        UpdateMouse m_hWnd

        Game_Update (dt)
        Game_Draw (dt)

        mWindowResized = False
        DoEvents
    Loop

    DeinitWindowState
    Game_Unload
    Unload TargetForm
    StopMusic
    DeinitOpenGL
    mWindowReady = False
    Exit Sub

ErrorHandler:
    DeinitWindowState
    StopMusic
    DeinitOpenGL
    mWindowReady = False
    MsgBox "Error: " & Err.Description, vbCritical
    End
End Sub

' ============================================================================
' Core: timing / estado / janela / monitor / clipboard / cursor
' ============================================================================

Private Sub ResetCoreState()
    mMouseScaleX = 1!
    mMouseScaleY = 1!
    mMouseOffsetX = 0
    mMouseOffsetY = 0
    mMouseWheelVX = 0!
    mMouseWheelVY = 0!
    mCursorHidden = False
    mFullscreen = False
    mShouldQuit = False
    mMinSizeW = 0
    mMinSizeH = 0
    mFrameTime = 0!
    mTime = 0#
    mFps = 0
    mWindowResized = False
End Sub

' --- Timing ----------------------------------------------------------------

Public Sub SetTargetFPS(ByVal target As Long)
    If target < 0 Then target = 0
    mTargetFPS = target
End Sub

Public Function GetFPS() As Long
    GetFPS = mFps
End Function

Public Function GetFrameTime() As Single
    GetFrameTime = mFrameTime
End Function

Public Function GetTime() As Double
    GetTime = mTime
End Function

' Pausa a execucao por N segundos mantendo a fila de mensagens viva.
Public Sub WaitTime(ByVal seconds As Double)
    Dim freq As Currency, start As Currency, now As Currency
    If seconds <= 0# Then Exit Sub
    QueryPerformanceFrequency freq
    QueryPerformanceCounter start
    Do
        QueryPerformanceCounter now
        DoEvents
    Loop While (CSng(now - start) / CSng(freq)) < seconds
End Sub

' --- Estado da janela ------------------------------------------------------

Public Function WindowShouldClose() As Boolean
    WindowShouldClose = mShouldQuit
End Function

Public Sub CloseWindow()
    mShouldQuit = True
End Sub

Public Function IsWindowReady() As Boolean
    IsWindowReady = mWindowReady
End Function

Public Function IsWindowFocused() As Boolean
    If m_hWnd = 0 Then Exit Function
    IsWindowFocused = (GetForegroundWindow() = m_hWnd)
End Function

Public Function IsWindowResized() As Boolean
    IsWindowResized = mWindowResized
End Function

Public Function IsWindowFullscreen() As Boolean
    IsWindowFullscreen = mFullscreen
End Function

Public Function IsWindowMinimized() As Boolean
    If m_hWnd = 0 Then Exit Function
    IsWindowMinimized = (IsIconic(m_hWnd) <> 0)
End Function

Public Function IsWindowMaximized() As Boolean
    If m_hWnd = 0 Then Exit Function
    IsWindowMaximized = (IsZoomed(m_hWnd) <> 0)
End Function

' --- Janela (setters) ------------------------------------------------------

Public Sub SetWindowTitle(ByVal title As String)
    If m_hWnd = 0 Then Exit Sub
    SetWindowTextA m_hWnd, title
End Sub

Public Sub SetWindowSize(ByVal Width As Long, ByVal Height As Long)
    Dim r As Rect
    If m_hWnd = 0 Then Exit Sub
    If Width <= 0 Or Height <= 0 Then Exit Sub

    r.left = 0: r.top = 0: r.right = Width: r.bottom = Height
    AdjustWindowRectEx r, GetWindowLongA(m_hWnd, GWL_STYLE), 0, GetWindowLongA(m_hWnd, GWL_EXSTYLE)
    SetWindowPos m_hWnd, 0, 0, 0, r.right - r.left, r.bottom - r.top, SWP_NOMOVE Or SWP_NOZORDER
    UpdateViewport Width, Height
End Sub

Public Sub SetWindowPosition(ByVal x As Long, ByVal y As Long)
    If m_hWnd = 0 Then Exit Sub
    SetWindowPos m_hWnd, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER
End Sub

' Minimo de tamanho via subclass do WndProc (WM_GETMINMAXINFO).
Public Sub SetWindowMinSize(ByVal Width As Long, ByVal Height As Long)
    mMinSizeW = Width
    mMinSizeH = Height
    If m_hWnd <> 0 Then EnsureSubclass
End Sub

Public Sub SetWindowOpacity(ByVal opacity As Single)
    If m_hWnd = 0 Then Exit Sub
    If opacity < 0! Then opacity = 0!
    If opacity > 1! Then opacity = 1!
    SetWindowLongA m_hWnd, GWL_EXSTYLE, GetWindowLongA(m_hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes m_hWnd, 0, CByte(opacity * 255!), LWA_ALPHA
End Sub

Public Sub ToggleFullscreen()
    If mFullscreen Then
        LeaveFullscreen
    Else
        EnterFullscreen
    End If
End Sub

Public Sub MaximizeWindow()
    If m_hWnd <> 0 Then ShowWindow m_hWnd, SW_MAXIMIZE
End Sub

Public Sub MinimizeWindow()
    If m_hWnd <> 0 Then ShowWindow m_hWnd, SW_MINIMIZE
End Sub

Public Sub RestoreWindow()
    If m_hWnd <> 0 Then ShowWindow m_hWnd, SW_RESTORE
End Sub

' --- ConfigFlags / estados -------------------------------------------------

Public Sub SetConfigFlags(ByVal flags As Long)
    mConfigFlags = flags
    ApplyConfigFlags
End Sub

Public Sub SetWindowState(ByVal flags As Long)
    mConfigFlags = mConfigFlags Or flags
    If m_hWnd = 0 Then Exit Sub
    If (flags And FLAG_VSYNC_HINT) <> 0 Then SetVSyncInternal True
    If (flags And FLAG_FULLSCREEN_MODE) <> 0 Then EnterFullscreen
    If (flags And FLAG_WINDOW_TOPMOST) <> 0 Then SetTopmost True
    If (flags And FLAG_WINDOW_UNDECORATED) <> 0 Then SetUndecorated True
    If (flags And FLAG_WINDOW_TRANSPARENT) <> 0 Then SetTransparent True
    If (flags And FLAG_WINDOW_HIDDEN) <> 0 Then ShowWindow m_hWnd, SW_HIDE
    If (flags And FLAG_WINDOW_MINIMIZED) <> 0 Then ShowWindow m_hWnd, SW_MINIMIZE
    If (flags And FLAG_WINDOW_MAXIMIZED) <> 0 Then ShowWindow m_hWnd, SW_MAXIMIZE
End Sub

Public Sub ClearWindowState(ByVal flags As Long)
    mConfigFlags = mConfigFlags And Not flags
    If m_hWnd = 0 Then Exit Sub
    If (flags And FLAG_VSYNC_HINT) <> 0 Then SetVSyncInternal False
    If (flags And FLAG_FULLSCREEN_MODE) <> 0 Then LeaveFullscreen
    If (flags And FLAG_WINDOW_TOPMOST) <> 0 Then SetTopmost False
    If (flags And FLAG_WINDOW_UNDECORATED) <> 0 Then SetUndecorated False
    If (flags And FLAG_WINDOW_TRANSPARENT) <> 0 Then SetTransparent False
    If (flags And FLAG_WINDOW_HIDDEN) <> 0 Then ShowWindow m_hWnd, SW_SHOW
    If (flags And FLAG_WINDOW_MINIMIZED) <> 0 Then ShowWindow m_hWnd, SW_RESTORE
    If (flags And FLAG_WINDOW_MAXIMIZED) <> 0 Then ShowWindow m_hWnd, SW_RESTORE
End Sub

Private Sub ApplyConfigFlags()
    If m_hWnd = 0 Then Exit Sub
    If (mConfigFlags And FLAG_VSYNC_HINT) <> 0 Then SetVSyncInternal True
    If (mConfigFlags And FLAG_FULLSCREEN_MODE) <> 0 Then EnterFullscreen
    If (mConfigFlags And FLAG_WINDOW_TOPMOST) <> 0 Then SetTopmost True
    If (mConfigFlags And FLAG_WINDOW_UNDECORATED) <> 0 Then SetUndecorated True
    If (mConfigFlags And FLAG_WINDOW_TRANSPARENT) <> 0 Then SetTransparent True
    If (mConfigFlags And FLAG_WINDOW_HIDDEN) <> 0 Then ShowWindow m_hWnd, SW_HIDE
    If (mConfigFlags And FLAG_WINDOW_MINIMIZED) <> 0 Then ShowWindow m_hWnd, SW_MINIMIZE
    If (mConfigFlags And FLAG_WINDOW_MAXIMIZED) <> 0 Then ShowWindow m_hWnd, SW_MAXIMIZE
End Sub

Private Sub SetTopmost(ByVal enable As Boolean)
    If m_hWnd = 0 Then Exit Sub
    If enable Then
        SetWindowPos m_hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Private Sub SetUndecorated(ByVal enable As Boolean)
    Dim style As Long
    If m_hWnd = 0 Then Exit Sub
    style = GetWindowLongA(m_hWnd, GWL_STYLE)
    If enable Then
        style = style And Not (WS_CAPTION Or WS_THICKFRAME)
    Else
        style = style Or WS_CAPTION Or WS_THICKFRAME
    End If
    SetWindowLongA m_hWnd, GWL_STYLE, style
    SetWindowPos m_hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Private Sub SetTransparent(ByVal enable As Boolean)
    Dim exStyle As Long
    If m_hWnd = 0 Then Exit Sub
    exStyle = GetWindowLongA(m_hWnd, GWL_EXSTYLE)
    If enable Then
        exStyle = exStyle Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
    Else
        exStyle = exStyle And Not WS_EX_TRANSPARENT
    End If
    SetWindowLongA m_hWnd, GWL_EXSTYLE, exStyle
End Sub

' --- Fullscreen ------------------------------------------------------------

Private Sub EnterFullscreen()
    Dim mi As MONITORINFO
    If m_hWnd = 0 Then Exit Sub
    If mFullscreen Then Exit Sub

    mSavedStyle = GetWindowLongA(m_hWnd, GWL_STYLE)
    mSavedExStyle = GetWindowLongA(m_hWnd, GWL_EXSTYLE)
    GetWindowRect m_hWnd, mSavedRect

    mi.cbSize = Len(mi)
    If GetMonitorInfoA(MonitorFromWindow(m_hWnd, MONITOR_DEFAULTTONEAREST), mi) = 0 Then Exit Sub

    SetWindowLongA m_hWnd, GWL_STYLE, mSavedStyle And Not (WS_CAPTION Or WS_THICKFRAME)
    SetWindowPos m_hWnd, 0, mi.rcMonitor.left, mi.rcMonitor.top, _
        mi.rcMonitor.right - mi.rcMonitor.left, mi.rcMonitor.bottom - mi.rcMonitor.top, _
        SWP_FRAMECHANGED Or SWP_NOZORDER

    mFullscreen = True
End Sub

Private Sub LeaveFullscreen()
    If m_hWnd = 0 Then Exit Sub
    If mFullscreen = False Then Exit Sub

    SetWindowLongA m_hWnd, GWL_STYLE, mSavedStyle
    SetWindowLongA m_hWnd, GWL_EXSTYLE, mSavedExStyle
    SetWindowPos m_hWnd, 0, mSavedRect.left, mSavedRect.top, _
        mSavedRect.right - mSavedRect.left, mSavedRect.bottom - mSavedRect.top, _
        SWP_FRAMECHANGED Or SWP_NOZORDER

    mFullscreen = False
End Sub

' --- Subclass (min size) ---------------------------------------------------

Private Sub EnsureSubclass()
    If mSubclassed Then Exit Sub
    If m_hWnd = 0 Then Exit Sub
    mOldWndProc = SetWindowLongA(m_hWnd, GWL_WNDPROC, AddressOf WindowProc)
    mSubclassed = True
End Sub

Private Sub RemoveSubclass()
    If mSubclassed = False Then Exit Sub
    If m_hWnd <> 0 Then SetWindowLongA m_hWnd, GWL_WNDPROC, mOldWndProc
    mSubclassed = False
End Sub

Private Function WindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If msg = WM_GETMINMAXINFO Then
        Dim mmi As MINMAXINFO
        CopyMemoryPtr VarPtr(mmi), lParam, Len(mmi)
        If mMinSizeW > 0 Then mmi.ptMinTrackSize.x = mMinSizeW
        If mMinSizeH > 0 Then mmi.ptMinTrackSize.y = mMinSizeH
        CopyMemoryPtr lParam, VarPtr(mmi), Len(mmi)
        WindowProc = 0
        Exit Function
    End If
    WindowProc = CallWindowProcA(mOldWndProc, hwnd, msg, wParam, lParam)
End Function

Private Sub DeinitWindowState()
    RemoveSubclass
    If mFullscreen Then LeaveFullscreen
    m_hWnd = 0
End Sub

' --- Monitor ---------------------------------------------------------------

Private Sub EnsureMonitors()
    If mMonitorsReady Then Exit Sub
    mMonitorsReady = True
    mMonitorCount = 0
    ReDim mMonitors(0 To 15)
    ReDim mMonitorHandles(0 To 15)
    EnumDisplayMonitors 0, 0, AddressOf EnumMonitorsProc, 0

    ' o monitor primario (0,0) vai para o indice 0, como na raylib
    Dim i As Long
    Dim tmp As MONITORINFO
    Dim tmpH As Long
    For i = 1 To mMonitorCount - 1
        If mMonitors(i).rcMonitor.left = 0 And mMonitors(i).rcMonitor.top = 0 Then
            tmp = mMonitors(0)
            mMonitors(0) = mMonitors(i)
            mMonitors(i) = tmp
            tmpH = mMonitorHandles(0)
            mMonitorHandles(0) = mMonitorHandles(i)
            mMonitorHandles(i) = tmpH
            Exit For
        End If
    Next i
End Sub

Private Function EnumMonitorsProc(ByVal hMonitor As Long, ByVal hDC As Long, ByRef lprc As Rect, ByVal dwData As Long) As Long
    If mMonitorCount <= UBound(mMonitors) Then
        Dim mi As MONITORINFO
        mi.cbSize = Len(mi)
        If GetMonitorInfoA(hMonitor, mi) <> 0 Then
            mMonitors(mMonitorCount) = mi
            mMonitorHandles(mMonitorCount) = hMonitor
            mMonitorCount = mMonitorCount + 1
        End If
    End If
    EnumMonitorsProc = 1
End Function

Public Function GetMonitorCount() As Long
    EnsureMonitors
    GetMonitorCount = mMonitorCount
End Function

Public Function GetMonitorWidth(ByVal monitor As Long) As Long
    EnsureMonitors
    If mMonitorCount = 0 Then Exit Function
    If monitor < 0 Or monitor >= mMonitorCount Then monitor = 0
    GetMonitorWidth = mMonitors(monitor).rcMonitor.right - mMonitors(monitor).rcMonitor.left
End Function

Public Function GetMonitorHeight(ByVal monitor As Long) As Long
    EnsureMonitors
    If mMonitorCount = 0 Then Exit Function
    If monitor < 0 Or monitor >= mMonitorCount Then monitor = 0
    GetMonitorHeight = mMonitors(monitor).rcMonitor.bottom - mMonitors(monitor).rcMonitor.top
End Function

' Taxa de atualizacao do monitor atual (usa o dispositivo primario).
Public Function GetMonitorRefreshRate(ByVal monitor As Long) As Long
    Dim dm As DEVMODE
    dm.dmSize = Len(dm)
    If EnumDisplaySettingsA(0, ENUM_CURRENT_SETTINGS, dm) <> 0 Then
        GetMonitorRefreshRate = dm.dmDisplayFrequency
    End If
End Function

Public Function GetCurrentMonitor() As Long
    Dim hMon As Long, i As Long
    If m_hWnd = 0 Then Exit Function
    EnsureMonitors
    hMon = MonitorFromWindow(m_hWnd, MONITOR_DEFAULTTONEAREST)
    For i = 0 To mMonitorCount - 1
        If mMonitorHandles(i) = hMon Then
            GetCurrentMonitor = i
            Exit Function
        End If
    Next i
End Function

Public Function GetWindowScaleDPI() As Vector2
    Dim v As Vector2
    Dim hDC As Long
    v.x = 1!: v.y = 1!
    hDC = GetDC(0)
    If hDC <> 0 Then
        v.x = CSng(GetDeviceCaps(hDC, LOGPIXELSX)) / 96!
        v.y = CSng(GetDeviceCaps(hDC, LOGPIXELSY)) / 96!
        ReleaseDC 0, hDC
    End If
    GetWindowScaleDPI = v
End Function

' --- Clipboard -------------------------------------------------------------

Public Function GetClipboardText() As String
    Dim hData As Long, p As Long, n As Long
    Dim bytes() As Byte
    If OpenClipboard(0) = 0 Then Exit Function
    hData = GetClipboardData(CF_TEXT)
    If hData <> 0 Then
        p = GlobalLock(hData)
        If p <> 0 Then
            n = lstrlenA(p)
            If n > 0 Then
                ReDim bytes(0 To n - 1)
                CopyMemoryPtr VarPtr(bytes(0)), p, n
                GetClipboardText = StrConv(bytes, vbUnicode)
            End If
            GlobalUnlock hData
        End If
    End If
    CloseClipboard
End Function

Public Sub SetClipboardText(ByVal text As String)
    Dim hMem As Long, p As Long, n As Long
    Dim bytes() As Byte
    If OpenClipboard(0) = 0 Then Exit Sub
    EmptyClipboard
    If Len(text) > 0 Then
        bytes = StrConv(text, vbFromUnicode)
        n = UBound(bytes) - LBound(bytes) + 1
    Else
        n = 0
    End If
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, n + 1)
    If hMem <> 0 Then
        p = GlobalLock(hMem)
        If p <> 0 Then
            If n > 0 Then CopyMemoryPtr p, VarPtr(bytes(0)), n
            GlobalUnlock hMem
            SetClipboardData CF_TEXT, hMem
        End If
    End If
    CloseClipboard
End Sub

' --- Random ----------------------------------------------------------------

Public Function GetRandomValue(ByVal min As Long, ByVal max As Long) As Long
    Dim t As Long
    If min > max Then
        t = min: min = max: max = t
    End If
    GetRandomValue = min + Int(Rnd() * (max - min + 1))
    If GetRandomValue > max Then GetRandomValue = max
End Function

Public Sub SetRandomSeed(ByVal seed As Long)
    Randomize seed
End Sub

' --- Screenshot ------------------------------------------------------------

' Le o framebuffer atual e grava um PNG em fileName.
Public Sub TakeScreenshot(ByVal fileName As String)
    Dim w As Long, h As Long
    w = screenWidth
    h = screenHeight
    If w <= 0 Or h <= 0 Then Exit Sub

    Dim pixels() As Byte
    Dim flipped() As Byte
    ReDim pixels(0 To (w * h * 4) - 1)
    ReDim flipped(0 To (w * h * 4) - 1)

    ' Garante que tudo que foi acumulado no lote ja esta no framebuffer.
    rlDrawRenderBatchActive

    glReadPixels 0, 0, w, h, GL_RGBA, GL_UNSIGNED_BYTE, ByVal VarPtr(pixels(0))

    ' GL entrega de baixo para cima e em RGBA; GDI+ espera de cima para
    ' baixo e BGRA. Invertemos as linhas e trocamos R<->B.
    Dim rowBytes As Long
    rowBytes = w * 4
    Dim y As Long, x As Long
    Dim srcRow As Long, dstRow As Long
    For y = 0 To h - 1
        srcRow = (h - 1 - y) * rowBytes
        dstRow = y * rowBytes
        For x = 0 To w - 1
            flipped(dstRow + x * 4) = pixels(srcRow + x * 4 + 2)
            flipped(dstRow + x * 4 + 1) = pixels(srcRow + x * 4 + 1)
            flipped(dstRow + x * 4 + 2) = pixels(srcRow + x * 4)
            flipped(dstRow + x * 4 + 3) = pixels(srcRow + x * 4 + 3)
        Next x
    Next y

    SaveImageAsPng GetFullPath(fileName), flipped(), w, h
End Sub

Private Sub SaveImageAsPng(ByVal path As String, ByRef Data() As Byte, ByVal w As Long, ByVal h As Long)
    Dim bmp As Long
    Dim clsid As GUID
    If w <= 0 Or h <= 0 Then Exit Sub
    If GdipCreateBitmapFromScan0(w, h, w * 4, PixelFormat32bppARGB, VarPtr(Data(0)), bmp) <> 0 Then Exit Sub

    ' CLSID do encoder PNG: {557CF406-1A04-11D3-9A73-0000F81EF32E}
    clsid.Data1 = &H557CF406
    clsid.Data2 = &H1A04
    clsid.Data3 = &H11D3
    clsid.Data4(0) = &H9A: clsid.Data4(1) = &H73
    clsid.Data4(2) = &H0: clsid.Data4(3) = &H0
    clsid.Data4(4) = &HF8: clsid.Data4(5) = &H1E
    clsid.Data4(6) = &HF3: clsid.Data4(7) = &H2E

    GdipSaveImageToFile bmp, StrPtr(path), clsid, 0
    GdipDisposeImage bmp
End Sub

' --- Cursor ----------------------------------------------------------------

Public Sub ShowCursor()
    If mCursorHidden Then apiShowCursor True
    mCursorHidden = False
End Sub

Public Sub HideCursor()
    If mCursorHidden = False Then apiShowCursor False
    mCursorHidden = True
End Sub

Public Function IsCursorHidden() As Boolean
    IsCursorHidden = mCursorHidden
End Function

Public Sub EnableCursor()
    ShowCursor
End Sub

Public Sub DisableCursor()
    HideCursor
    SetMousePosition screenWidth \ 2, screenHeight \ 2
End Sub

Public Sub SetMouseCursor(ByVal cursor As MouseCursor)
    Dim Id As Long
    Select Case cursor
        Case MOUSE_CURSOR_IBEAM: Id = IDC_IBEAM
        Case MOUSE_CURSOR_CROSSHAIR: Id = IDC_CROSS
        Case MOUSE_CURSOR_POINTING_HAND: Id = IDC_HAND
        Case MOUSE_CURSOR_RESIZE_EW: Id = IDC_SIZEWE
        Case MOUSE_CURSOR_RESIZE_NS: Id = IDC_SIZENS
        Case MOUSE_CURSOR_RESIZE_NWSE: Id = IDC_SIZENWSE
        Case MOUSE_CURSOR_RESIZE_NESW: Id = IDC_SIZENESW
        Case MOUSE_CURSOR_RESIZE_ALL: Id = IDC_SIZEALL
        Case MOUSE_CURSOR_NOT_ALLOWED: Id = IDC_NO
        Case Else: Id = IDC_ARROW
    End Select
    apiSetCursor LoadCursorA(0, Id)
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
    CurrentMouseX = (pt.x * mMouseScaleX) + mMouseOffsetX
    CurrentMouseY = (pt.y * mMouseScaleY) + mMouseOffsetY

    UpdateSingleButton 0, VK_LBUTTON
    UpdateSingleButton 1, VK_RBUTTON
    UpdateSingleButton 2, VK_MBUTTON

    mMouseWheelVX = 0!
    mMouseWheelVY = 0!
End Sub

Public Sub SetMousePosition(ByVal x As Long, ByVal y As Long)
    Dim pt As POINTAPI
    If m_hWnd = 0 Then Exit Sub
    pt.x = x
    pt.y = y
    ClientToScreen m_hWnd, pt
    SetCursorPos pt.x, pt.y
End Sub

Public Sub SetMouseOffset(ByVal offsetX As Long, ByVal offsetY As Long)
    mMouseOffsetX = offsetX
    mMouseOffsetY = offsetY
End Sub

Public Sub SetMouseScale(ByVal scaleX As Single, ByVal scaleY As Single)
    mMouseScaleX = scaleX
    mMouseScaleY = scaleY
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
    mMouseWheelVY = value
End Sub

Public Function GetMouseWheelMove() As Single
    GetMouseWheelMove = mMouseWheelVY
End Function

Public Function GetMouseWheelMoveV() As Vector2
    Dim v As Vector2
    v.x = mMouseWheelVX
    v.y = mMouseWheelVY
    GetMouseWheelMoveV = v
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
        GetKeyName = left$(nameBuffer, InStr(nameBuffer, vbNullChar) - 1)
    Else
        GetKeyName = "UNKNOWN"
    End If
End Function

' Color

Private Sub RGBToHSV(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByRef h As Single, ByRef s As Single, ByRef v As Single)
    Dim min As Single, max As Single, delta As Single

    min = r: If g < min Then min = g: If b < min Then min = b
    max = r: If g > max Then max = g: If b > max Then max = b

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

    If r = max Then
        h = (g - b) / delta
    ElseIf g = max Then
        h = 2! + (b - r) / delta
    Else
        h = 4! + (r - g) / delta
    End If

    h = h * 60!
    If h < 0! Then h = h + 360!
End Sub

Private Sub HSVToRGB(ByVal h As Single, ByVal s As Single, ByVal v As Single, ByRef r As Single, ByRef g As Single, ByRef b As Single)
    If s = 0! Then
        r = v: g = v: b = v
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
    Case 0: r = v: g = t: b = p
    Case 1: r = q: g = v: b = p
    Case 2: r = p: g = v: b = t
    Case 3: r = p: g = q: b = v
    Case 4: r = t: g = p: b = v
    Case Else: r = v: g = p: b = q
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
    result.r = Clamp01(Color.r * tint.r)
    result.g = Clamp01(Color.g * tint.g)
    result.b = Clamp01(Color.b * tint.b)
    result.a = Clamp01(Color.a * tint.a)
    ColorTint = result
End Function

Public Function ColorBrightness(ByRef Color As ColorRGBA, ByVal factor As Single) As ColorRGBA
    Dim h As Single, s As Single, v As Single

    RGBToHSV Color.r, Color.g, Color.b, h, s, v

    v = v + factor
    If v < 0.1 Then v = 0.1
    If v > 1! Then v = 1!

    Dim result As ColorRGBA
    HSVToRGB h, s, v, result.r, result.g, result.b
    result.a = Color.a

    ColorBrightness = result
End Function

Public Function ColorContrast(ByRef Color As ColorRGBA, ByVal contrast As Single) As ColorRGBA
    Dim result As ColorRGBA
    contrast = contrast + 1!

    If contrast < 0! Then
        contrast = 0!
    End If

    result.r = Clamp01(((Color.r - 0.5) * contrast) + 0.5)
    result.g = Clamp01(((Color.g - 0.5) * contrast) + 0.5)
    result.b = Clamp01(((Color.b - 0.5) * contrast) + 0.5)
    result.a = Color.a
    ColorContrast = result
End Function

Public Function ColorAlpha(ByRef Color As ColorRGBA, ByVal alpha As Single) As ColorRGBA
    Dim result As ColorRGBA
    result.r = Color.r
    result.g = Color.g
    result.b = Color.b
    result.a = Clamp01(alpha)
    ColorAlpha = result
End Function

Public Function ColorAlphaBlend(ByRef dst As ColorRGBA, ByRef src As ColorRGBA, ByRef tint As ColorRGBA) As ColorRGBA
    Dim result As ColorRGBA
    Dim srcTint As ColorRGBA

    srcTint.r = src.r * tint.r
    srcTint.g = src.g * tint.g
    srcTint.b = src.b * tint.b
    srcTint.a = src.a * tint.a

    Dim alpha As Single
    alpha = srcTint.a

    result.r = Clamp01(srcTint.r * alpha + dst.r * (1! - alpha))
    result.g = Clamp01(srcTint.g * alpha + dst.g * (1! - alpha))
    result.b = Clamp01(srcTint.b * alpha + dst.b * (1! - alpha))
    result.a = Clamp01(alpha + dst.a * (1! - alpha))

    ColorAlphaBlend = result
End Function

Public Function ColorLerp(ByRef color1 As ColorRGBA, ByRef color2 As ColorRGBA, ByVal factor As Single) As ColorRGBA
    If factor < 0! Then factor = 0!
    If factor > 1! Then factor = 1!

    Dim h1 As Single, s1 As Single, v1 As Single
    Dim h2 As Single, s2 As Single, v2 As Single

    RGBToHSV color1.r, color1.g, color1.b, h1, s1, v1
    RGBToHSV color2.r, color2.g, color2.b, h2, s2, v2

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
    aRes = color1.a + (color2.a - color1.a) * factor

    Dim result As ColorRGBA
    HSVToRGB hRes, sRes, vRes, result.r, result.g, result.b
    result.a = aRes

    ColorLerp = result
End Function

Public Function ColorIsEqual(ByRef col1 As ColorRGBA, ByRef col2 As ColorRGBA) As Boolean
    Const EPSILON As Single = 0.00001
    ColorIsEqual = (Abs(col1.r - col2.r) < EPSILON) And _
                   (Abs(col1.g - col2.g) < EPSILON) And _
                   (Abs(col1.b - col2.b) < EPSILON) And _
                   (Abs(col1.a - col2.a) < EPSILON)
End Function

Public Function ColorToInt(ByRef Color As ColorRGBA) As Long
    Dim r As Long, g As Long, b As Long, a As Long
    r = CLng(Clamp01(Color.r) * 255!)
    g = CLng(Clamp01(Color.g) * 255!)
    b = CLng(Clamp01(Color.b) * 255!)
    a = CLng(Clamp01(Color.a) * 255!)

    ColorToInt = (r * &H1000000) Or (g * &H10000) Or (b * &H100) Or a
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

        CopyMemory resultImage.Data(0), ByVal bmpData.scan0, totalBytes

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

    img.Data(idx) = CByte(col.r * 255#)
    img.Data(idx + 1) = CByte(col.g * 255#)
    img.Data(idx + 2) = CByte(col.b * 255#)
    img.Data(idx + 3) = CByte(col.a * 255#)
End Sub

Public Sub ImageClear(ByRef img As Image, ByRef col As ColorRGBA)
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte

    r = CByte(col.r * 255#)
    g = CByte(col.g * 255#)
    b = CByte(col.b * 255#)
    a = CByte(col.a * 255#)

    Dim rowBytes As Long
    rowBytes = img.Width * 4

    For y = 0 To img.Height - 1
        Dim rowStart As Long
        rowStart = y * rowBytes
        For x = 0 To img.Width - 1
            Dim idx As Long
            idx = rowStart + (x * 4)
            img.Data(idx) = r
            img.Data(idx + 1) = g
            img.Data(idx + 2) = b
            img.Data(idx + 3) = a
        Next x
    Next y
End Sub

Public Sub ImageDrawRect(ByRef img As Image, ByRef Rect As Rectangle, ByRef col As ColorRGBA)
    Dim rx As Long, ry As Long, rw As Long, rh As Long
    rx = CLng(Rect.x)
    ry = CLng(Rect.y)
    rw = CLng(Rect.Width)
    rh = CLng(Rect.Height)

    Dim r As Byte, g As Byte, b As Byte, a As Byte
    r = CByte(col.r * 255#)
    g = CByte(col.g * 255#)
    b = CByte(col.b * 255#)
    a = CByte(col.a * 255#)

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
            img.Data(idx) = r
            img.Data(idx + 1) = g
            img.Data(idx + 2) = b
            img.Data(idx + 3) = a
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
    Dim textureId As Long

    tex.Id = 0
    tex.Width = 0
    tex.Height = 0

    If img.Width <= 0 Or img.Height <= 0 Then
        LoadTextureFromImage = tex
        Exit Function
    End If

    glGenTextures 1, textureId
    glBindTexture GL_TEXTURE_2D, textureId

    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST

    glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, img.Width, img.Height, 0, GL_RGBA, GL_UNSIGNED_BYTE, ByVal VarPtr(img.Data(0))

    tex.Id = textureId
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

Public Sub UnloadTexture(ByVal textureId As Long)
    If textureId <> 0 Then
        glDeleteTextures 1, textureId
    End If
End Sub

' Render target (framebuffer object)

Public Function LoadRenderTexture(ByVal Width As Long, ByVal Height As Long) As RenderTexture2D
    Dim rt As RenderTexture2D
    Dim prevFbo As Long
    Dim status As Long

    rt.Id = 0
    rt.texture.Id = 0
    rt.texture.Width = 0
    rt.texture.Height = 0
    rt.depth.Id = 0

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
    glGenTextures 1, rt.texture.Id
    glBindTexture GL_TEXTURE_2D, rt.texture.Id
    glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, Width, Height, 0, GL_RGBA, GL_UNSIGNED_BYTE, 0
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE
    glFramebufferTexture2DEXT GL_FRAMEBUFFER, GL_COLOR_ATTACHMENT0, GL_TEXTURE_2D, rt.texture.Id, 0

    ' renderbuffer de profundidade
    glGenRenderbuffersEXT 1, rt.depth.Id
    glBindRenderbufferEXT GL_RENDERBUFFER, rt.depth.Id
    glRenderbufferStorageEXT GL_RENDERBUFFER, GL_DEPTH_COMPONENT24, Width, Height
    glFramebufferRenderbufferEXT GL_FRAMEBUFFER, GL_DEPTH_ATTACHMENT, GL_RENDERBUFFER, rt.depth.Id

    rt.texture.Width = Width
    rt.texture.Height = Height

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
        rt.texture.Id = 0
        rt.texture.Width = 0
        rt.texture.Height = 0
        rt.depth.Id = 0
    End If

    LoadRenderTexture = rt
End Function

Public Sub UnloadRenderTexture(ByRef target As RenderTexture2D)
    If target.texture.Id <> 0 Then
        glDeleteTextures 1, target.texture.Id
        target.texture.Id = 0
    End If

    If target.depth.Id <> 0 Then
        glDeleteRenderbuffersEXT 1, target.depth.Id
        target.depth.Id = 0
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
            If Abs(vbCol.r - sepCol.r) < 0.05 And _
               Abs(vbCol.g - sepCol.g) < 0.05 And _
               Abs(vbCol.b - sepCol.b) < 0.05 Then
                x = x + 1
            Else
                Exit Do
            End If
        Loop

        If x >= img.Width Then Exit Do

        startX = x

        Do While x < img.Width
            vbCol = ImageGetPixel(img, x, 0)
            If Abs(vbCol.r - sepCol.r) < 0.05 And _
               Abs(vbCol.g - sepCol.g) < 0.05 And _
               Abs(vbCol.b - sepCol.b) < 0.05 Then
                Exit Do
            Else
                x = x + 1
            End If
        Loop

        w = x - startX
        If w <= 0 Then w = 1

        fnt.Glyphs(charIndex).left = CSng(startX) / CSng(img.Width)
        fnt.Glyphs(charIndex).Width = CSng(w)

        charIndex = charIndex + 1
    Loop

    fnt.texture = LoadTextureFromImage(img)
    LoadImageFont = fnt
End Function

' Draw

Public Sub UpdateViewport(ByVal w As Long, ByVal h As Long)
    screenWidth = w
    screenHeight = h
    rlViewport 0, 0, screenWidth, screenHeight
End Sub

Public Sub BeginDrawing()
    rlMatrixMode RL_PROJECTION
    rlLoadIdentity
    rlOrtho 0, screenWidth, screenHeight, 0, -1, 1

    rlMatrixMode RL_MODELVIEW
    rlLoadIdentity
End Sub

' Render target: redireciona o desenho para um framebuffer.
'
' Dentro do BeginTextureMode as coordenadas continuam sendo y-para-baixo iguais
' as da tela: a projecao usa rlOrtho(0, w, 0, h), que inverte a textura em
' relacao ao framebuffer, e o DrawTexture inverte de volta ao exibir. O efeito
' liquido e identidade, sem precisar mexer em nenhum codigo de desenho.
Public Sub BeginTextureMode(ByRef target As RenderTexture2D)
    Dim prevFbo As Long
    Dim vp(0 To 3) As Long

    If mRenderTop >= RENDER_STACK_MAX Then Exit Sub
    If target.Id = 0 Then Exit Sub

    ' O lote pendente pertence ao alvo anterior: precisa ser desenhado antes de
    ' trocar o framebuffer.
    rlDrawRenderBatchActive

    glGetIntegerv GL_FRAMEBUFFER_BINDING, prevFbo
    glGetIntegerv GL_VIEWPORT, vp(0)

    mRenderFbo(mRenderTop) = prevFbo
    mRenderW(mRenderTop) = vp(2)
    mRenderH(mRenderTop) = vp(3)
    mRenderTop = mRenderTop + 1

    glBindFramebufferEXT GL_FRAMEBUFFER, target.Id
    rlViewport 0, 0, target.texture.Width, target.texture.Height

    rlMatrixMode RL_PROJECTION
    rlPushMatrix
    rlLoadIdentity
    rlOrtho 0, target.texture.Width, 0, target.texture.Height, -1, 1

    rlMatrixMode RL_MODELVIEW
    rlPushMatrix
    rlLoadIdentity
End Sub

Public Sub EndTextureMode()
    If mRenderTop <= 0 Then Exit Sub

    ' Desenha o que foi acumulado dentro do alvo antes de voltar ao framebuffer
    ' anterior.
    rlDrawRenderBatchActive

    mRenderTop = mRenderTop - 1

    rlMatrixMode RL_PROJECTION
    rlPopMatrix

    rlMatrixMode RL_MODELVIEW
    rlPopMatrix

    glBindFramebufferEXT GL_FRAMEBUFFER, mRenderFbo(mRenderTop)
    rlViewport 0, 0, mRenderW(mRenderTop), mRenderH(mRenderTop)
End Sub

Public Sub Clear(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    rlClearColor r, g, b, a
    rlClearScreenBuffers
End Sub

Public Sub EndDrawing()
    rlDrawRenderBatchActive
    SwapBuffers m_hDC
End Sub

Public Sub BeginMode2D(ByRef camera As Camera2D)
    rlMatrixMode RL_PROJECTION
    rlPushMatrix

    rlMatrixMode RL_MODELVIEW
    rlPushMatrix

    rlTranslatef camera.Offset.x, camera.Offset.y, 0!
    rlRotatef camera.Rot, 0!, 0!, 1!
    rlScalef camera.Zoom, camera.Zoom, 1!
    rlTranslatef -camera.target.x, -camera.target.y, 0!
End Sub

Public Sub EndMode2D()
    rlMatrixMode RL_PROJECTION
    rlPopMatrix

    rlMatrixMode RL_MODELVIEW
    rlPopMatrix
End Sub

Public Sub DrawRectangle(ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByRef Color As ColorRGBA)
    rlDisableTexture

    rlBegin RL_QUADS
    rlColor4f Color.r, Color.g, Color.b, Color.a
    rlVertex2f x, y
    rlVertex2f x + Width, y
    rlVertex2f x + Width, y + Height
    rlVertex2f x, y + Height
    rlEnd
End Sub

Public Sub DrawRectangleRec(ByRef r As Rectangle, ByRef Color As ColorRGBA)
    DrawRectangle r.x, r.y, r.Width, r.Height, Color
End Sub

Public Sub DrawRectanglePro(ByRef r As Rectangle, ByRef origin As Vector2, ByVal Rot As Single, ByRef Color As ColorRGBA)
    rlPushMatrix
    rlTranslatef r.x, r.y, 0!

    If Rot <> 0! Then
        rlTranslatef origin.x, origin.y, 0!
        rlRotatef Rot, 0!, 0!, 1!
        rlTranslatef -origin.x, -origin.y, 0!
    End If

    rlDisableTexture
    rlBegin RL_QUADS
    rlColor4f Color.r, Color.g, Color.b, Color.a
    rlVertex2f -origin.x, -origin.y
    rlVertex2f -origin.x + r.Width, -origin.y
    rlVertex2f -origin.x + r.Width, -origin.y + r.Height
    rlVertex2f -origin.x, -origin.y + r.Height
    rlEnd

    rlPopMatrix
End Sub

Public Sub DrawRectangleLines(ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByRef Color As ColorRGBA)
    rlDisableTexture

    rlBegin RL_LINE_LOOP
    rlColor4f Color.r, Color.g, Color.b, Color.a
    rlVertex2f x, y
    rlVertex2f x + Width, y
    rlVertex2f x + Width, y + Height
    rlVertex2f x, y + Height
    rlEnd
End Sub

Public Sub DrawRectangleLinesEx(ByRef r As Rectangle, ByRef thick As Single, ByRef Color As ColorRGBA)
    Dim halfThick As Single
    halfThick = thick * 0.5

    DrawRectangle r.x, r.y, r.Width, thick, Color
    DrawRectangle r.x, r.y + r.Height - thick, r.Width, thick, Color
    DrawRectangle r.x, r.y + thick, thick, r.Height - (thick * 2), Color
    DrawRectangle r.x + r.Width - thick, r.y + thick, thick, r.Height - (thick * 2), Color
End Sub

Public Sub DrawRectangleRounded(ByRef r As Rectangle, ByRef roundness As Single, ByVal segments As Long, ByRef Color As ColorRGBA)
    If roundness <= 0! Then
        DrawRectangleRec r, Color
        Exit Sub
    End If
    If roundness > 1! Then roundness = 1!

    Dim radius As Single
    If r.Width < r.Height Then
        radius = (r.Width * 0.5) * roundness
    Else
        radius = (r.Height * 0.5) * roundness
    End If

    If segments < 4 Then segments = 4

    DrawRectangle r.x + radius, r.y, r.Width - (radius * 2), r.Height, Color
    DrawRectangle r.x, r.y + radius, radius, r.Height - (radius * 2), Color
    DrawRectangle r.x + r.Width - radius, r.y + radius, radius, r.Height - (radius * 2), Color

    DrawCircleSectorSimple r.x + radius, r.y + radius, radius, 180, 270, segments, Color
    DrawCircleSectorSimple r.x + r.Width - radius, r.y + radius, radius, 270, 360, segments, Color
    DrawCircleSectorSimple r.x + radius, r.y + r.Height - radius, radius, 90, 180, segments, Color
    DrawCircleSectorSimple r.x + r.Width - radius, r.y + r.Height - radius, radius, 0, 90, segments, Color
End Sub

Public Sub DrawRectangleRoundedLines(ByRef r As Rectangle, ByRef roundness As Single, ByVal segments As Long, ByRef Color As ColorRGBA)
    DrawRectangleRoundedLinesEx r, roundness, segments, 1#, Color
End Sub

Public Sub DrawRectangleRoundedLinesEx(ByRef r As Rectangle, ByRef roundness As Single, ByVal segments As Long, ByVal thick As Single, ByRef Color As ColorRGBA)
    If roundness <= 0! Then
        DrawRectangleLinesEx r, thick, Color
        Exit Sub
    End If

    DrawRectangleLines r.x + roundness, r.y, r.Width - (roundness * 2), r.Height, Color
End Sub

Public Sub DrawTriangle(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByRef Color As ColorRGBA)
    rlDisableTexture

    rlBegin RL_TRIANGLES
    rlColor4f Color.r, Color.g, Color.b, Color.a
    rlVertex2f x1, y1
    rlVertex2f x2, y2
    rlVertex2f x3, y3
    rlEnd
End Sub

Public Sub DrawLine(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef Color As ColorRGBA, ByVal thickness As Single)
    rlSetLineWidth thickness
    rlBegin RL_LINES
    rlColor4f Color.r, Color.g, Color.b, Color.a
    rlVertex2f x1, y1
    rlVertex2f x2, y2
    rlEnd
End Sub

Public Sub DrawTexture(ByRef tex As Texture2D, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByRef tint As ColorRGBA)
    If tex.Id = 0 Then Exit Sub

    rlSetTexture tex.Id

    rlColor4f tint.r, tint.g, tint.b, tint.a

    rlBegin RL_QUADS
    rlTexCoord2f 0#, 0#: rlVertex2f x, y
    rlTexCoord2f 1#, 0#: rlVertex2f x + Width, y
    rlTexCoord2f 1#, 1#: rlVertex2f x + Width, y + Height
    rlTexCoord2f 0#, 1#: rlVertex2f x, y + Height
    rlEnd

    rlDisableTexture
End Sub

Public Sub DrawTextureRec(ByRef tex As Texture2D, ByRef sourceRec As Rectangle, ByRef pos As Vector2, ByRef tint As ColorRGBA)
    If tex.Id = 0 Then Exit Sub

    rlSetTexture tex.Id

    rlColor4f tint.r, tint.g, tint.b, tint.a

    Dim u1 As Single, v1 As Single, u2 As Single, v2 As Single

    u1 = sourceRec.x / tex.Width
    v1 = sourceRec.y / tex.Height
    u2 = (sourceRec.x + sourceRec.Width) / tex.Width
    v2 = (sourceRec.y + sourceRec.Height) / tex.Height

    rlBegin RL_QUADS
    rlTexCoord2f u1, v1: rlVertex2f pos.x, pos.y
    rlTexCoord2f u2, v1: rlVertex2f pos.x + sourceRec.Width, pos.y
    rlTexCoord2f u2, v2: rlVertex2f pos.x + sourceRec.Width, pos.y + sourceRec.Height
    rlTexCoord2f u1, v2: rlVertex2f pos.x, pos.y + sourceRec.Height
    rlEnd

    rlDisableTexture
End Sub

Public Sub DrawTexturePro(ByRef tex As Texture2D, ByRef sourceRec As Rectangle, ByRef dstrect As Rectangle, ByRef origin As Vector2, ByVal Rot As Single, ByRef tint As ColorRGBA)
    If tex.Id = 0 Then Exit Sub

    Const PI As Single = 3.14159265
    Dim u1 As Single, v1 As Single, u2 As Single, v2 As Single

    u1 = sourceRec.x / tex.Width
    v1 = sourceRec.y / tex.Height
    u2 = (sourceRec.x + sourceRec.Width) / tex.Width
    v2 = (sourceRec.y + sourceRec.Height) / tex.Height

    rlSetTexture tex.Id

    rlPushMatrix

    rlTranslatef dstrect.x, dstrect.y, 0!

    If Rot <> 0! Then
        rlTranslatef origin.x, origin.y, 0!
        rlRotatef Rot, 0!, 0!, 1!
        rlTranslatef -origin.x, -origin.y, 0!
    End If

    rlColor4f tint.r, tint.g, tint.b, tint.a

    rlBegin RL_QUADS
    rlTexCoord2f u1, v1: rlVertex2f -origin.x, -origin.y
    rlTexCoord2f u2, v1: rlVertex2f -origin.x + dstrect.Width, -origin.y
    rlTexCoord2f u2, v2: rlVertex2f -origin.x + dstrect.Width, -origin.y + dstrect.Height
    rlTexCoord2f u1, v2: rlVertex2f -origin.x, -origin.y + dstrect.Height
    rlEnd

    rlPopMatrix
    rlDisableTexture
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
    leftBorder = CSng(info.left)
    topBorder = CSng(info.top)
    rightBorder = CSng(info.right)
    bottomBorder = CSng(info.bottom)

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

    rlSetTexture tex.Id

    rlPushMatrix
    rlTranslatef CSng(dstrect.x), CSng(dstrect.y), 0!

    If Rot <> 0! Then
        rlTranslatef origin.x, origin.y, 0!
        rlRotatef Rot, 0!, 0!, 1!
        rlTranslatef -origin.x, -origin.y, 0!
    End If

    rlColor4f Color.r, Color.g, Color.b, Color.a

    rlBegin RL_QUADS

    If info.Layout = NinePatch Then
        rlTexCoord2f coordAX, coordBY: rlVertex2f vertAX, vertBY
        rlTexCoord2f coordBX, coordBY: rlVertex2f vertBX, vertBY
        rlTexCoord2f coordBX, coordAY: rlVertex2f vertBX, vertAY
        rlTexCoord2f coordAX, coordAY: rlVertex2f vertAX, vertAY

        If drawCenter Then
            rlTexCoord2f coordBX, coordBY: rlVertex2f vertBX, vertBY
            rlTexCoord2f coordCX, coordBY: rlVertex2f vertCX, vertBY
            rlTexCoord2f coordCX, coordAY: rlVertex2f vertCX, vertAY
            rlTexCoord2f coordBX, coordAY: rlVertex2f vertBX, vertAY
        End If

        rlTexCoord2f coordCX, coordBY: rlVertex2f vertCX, vertBY
        rlTexCoord2f coordDX, coordBY: rlVertex2f vertDX, vertBY
        rlTexCoord2f coordDX, coordAY: rlVertex2f vertDX, vertAY
        rlTexCoord2f coordCX, coordAY: rlVertex2f vertCX, vertAY

        If drawMiddle Then
            rlTexCoord2f coordAX, coordCY: rlVertex2f vertAX, vertCY
            rlTexCoord2f coordBX, coordCY: rlVertex2f vertBX, vertCY
            rlTexCoord2f coordBX, coordBY: rlVertex2f vertBX, vertBY
            rlTexCoord2f coordAX, coordBY: rlVertex2f vertAX, vertBY

            If drawCenter Then
                rlTexCoord2f coordBX, coordCY: rlVertex2f vertBX, vertCY
                rlTexCoord2f coordCX, coordCY: rlVertex2f vertCX, vertCY
                rlTexCoord2f coordCX, coordBY: rlVertex2f vertCX, vertBY
                rlTexCoord2f coordBX, coordBY: rlVertex2f vertBX, vertBY
            End If

            rlTexCoord2f coordCX, coordCY: rlVertex2f vertCX, vertCY
            rlTexCoord2f coordDX, coordCY: rlVertex2f vertDX, vertCY
            rlTexCoord2f coordDX, coordBY: rlVertex2f vertDX, vertBY
            rlTexCoord2f coordCX, coordBY: rlVertex2f vertCX, vertBY
        End If

        rlTexCoord2f coordAX, coordDY: rlVertex2f vertAX, vertDY
        rlTexCoord2f coordBX, coordDY: rlVertex2f vertBX, vertDY
        rlTexCoord2f coordBX, coordCY: rlVertex2f vertBX, vertCY
        rlTexCoord2f coordAX, coordCY: rlVertex2f vertAX, vertCY

        If drawCenter Then
            rlTexCoord2f coordBX, coordDY: rlVertex2f vertBX, vertDY
            rlTexCoord2f coordCX, coordDY: rlVertex2f vertCX, vertDY
            rlTexCoord2f coordCX, coordCY: rlVertex2f vertCX, vertCY
            rlTexCoord2f coordBX, coordCY: rlVertex2f vertBX, vertCY
        End If

        rlTexCoord2f coordCX, coordDY: rlVertex2f vertCX, vertDY
        rlTexCoord2f coordDX, coordDY: rlVertex2f vertDX, vertDY
        rlTexCoord2f coordDX, coordCY: rlVertex2f vertDX, vertCY
        rlTexCoord2f coordCX, coordCY: rlVertex2f vertCX, vertCY

    ElseIf info.Layout = ThreePatchVertical Then
        rlTexCoord2f coordAX, coordBY: rlVertex2f vertAX, vertBY
        rlTexCoord2f coordDX, coordBY: rlVertex2f vertDX, vertBY
        rlTexCoord2f coordDX, coordAY: rlVertex2f vertDX, vertAY
        rlTexCoord2f coordAX, coordAY: rlVertex2f vertAX, vertAY

        If drawCenter Then
            rlTexCoord2f coordAX, coordCY: rlVertex2f vertAX, vertCY
            rlTexCoord2f coordDX, coordCY: rlVertex2f vertDX, vertCY
            rlTexCoord2f coordDX, coordBY: rlVertex2f vertDX, vertBY
            rlTexCoord2f coordAX, coordBY: rlVertex2f vertAX, vertBY
        End If

        rlTexCoord2f coordAX, coordDY: rlVertex2f vertAX, vertDY
        rlTexCoord2f coordDX, coordDY: rlVertex2f vertDX, vertDY
        rlTexCoord2f coordDX, coordCY: rlVertex2f vertDX, vertCY
        rlTexCoord2f coordAX, coordCY: rlVertex2f vertAX, vertCY

    ElseIf info.Layout = ThreePatchHorizontal Then
        rlTexCoord2f coordAX, coordDY: rlVertex2f vertAX, vertDY
        rlTexCoord2f coordBX, coordDY: rlVertex2f vertBX, vertDY
        rlTexCoord2f coordBX, coordAY: rlVertex2f vertBX, vertAY
        rlTexCoord2f coordAX, coordAY: rlVertex2f vertAX, vertAY

        If drawCenter Then
            rlTexCoord2f coordBX, coordDY: rlVertex2f vertBX, vertDY
            rlTexCoord2f coordCX, coordDY: rlVertex2f vertCX, vertDY
            rlTexCoord2f coordCX, coordAY: rlVertex2f vertCX, vertAY
            rlTexCoord2f coordBX, coordAY: rlVertex2f vertBX, vertAY
        End If

        rlTexCoord2f coordCX, coordDY: rlVertex2f vertCX, vertDY
        rlTexCoord2f coordDX, coordDY: rlVertex2f vertDX, vertDY
        rlTexCoord2f coordDX, coordAY: rlVertex2f vertDX, vertAY
        rlTexCoord2f coordCX, coordAY: rlVertex2f vertCX, vertAY
    End If

    rlEnd

    rlPopMatrix
    rlDisableTexture
End Sub

Public Sub DrawText(ByRef font As ImageFont, ByVal text As String, ByVal x As Single, ByVal y As Single, ByRef tint As ColorRGBA)
    Dim i As Long
    Dim curX As Single
    curX = x

    rlSetTexture font.texture.Id

    rlColor4f tint.r, tint.g, tint.b, tint.a

    rlBegin RL_QUADS

    Dim texelWidth As Single
    texelWidth = 1! / CSng(font.texture.Width)
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

        Dim g As GlyphRect
        g = font.Glyphs(charIdx)

        Dim u1 As Single, u2 As Single
        u1 = g.left + (texelWidth * 0.1)
        u2 = (g.left + (g.Width / CSng(font.texture.Width))) - (texelWidth * 0.1)

        rlTexCoord2f u1, 1#: rlVertex2f curX, y + fHeight
        rlTexCoord2f u2, 1#: rlVertex2f curX + g.Width, y + fHeight
        rlTexCoord2f u2, 0#: rlVertex2f curX + g.Width, y
        rlTexCoord2f u1, 0#: rlVertex2f curX, y

        curX = curX + g.Width
    Next i

    rlEnd
    rlDisableTexture
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

    rlDisableTexture
    rlBegin RL_TRIANGLE_FAN

    rlColor4f inner.r, inner.g, inner.b, inner.a
    rlVertex2f center.x, center.y

    For i = 0 To segments
        Dim currentAngle As Single
        currentAngle = i * angleStep

        rlColor4f outer.r, outer.g, outer.b, outer.a
        rlVertex2f center.x + Cos(currentAngle) * radius, center.y + Sin(currentAngle) * radius
    Next i
    rlEnd
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

    rlDisableTexture
    rlBegin RL_LINE_STRIP
    rlColor4f Color.r, Color.g, Color.b, Color.a

    rlVertex2f center.x, center.y

    For i = 0 To segments
        Dim currentAngle As Single
        currentAngle = (startAngle * (PI / 180!)) + (i * angleStep)
        rlVertex2f center.x + Cos(currentAngle) * radius, center.y + Sin(currentAngle) * radius
    Next i

    rlVertex2f center.x, center.y
    rlEnd
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

    rlDisableTexture
    rlBegin RL_LINE_LOOP
    rlColor4f Color.r, Color.g, Color.b, Color.a

    For i = 0 To segments - 1
        Dim currentAngle As Single
        currentAngle = i * angleStep
        rlVertex2f centerX + Cos(currentAngle) * radius, centerY + Sin(currentAngle) * radius
    Next i
    rlEnd
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

    rlDisableTexture
    rlBegin RL_TRIANGLE_FAN
    rlColor4f Color.r, Color.g, Color.b, Color.a
    rlVertex2f centerX, centerY

    For i = 0 To segments
        Dim currentAngle As Single
        currentAngle = (startAngle * (PI / 180!)) + (i * angleStep)
        rlVertex2f centerX + Cos(currentAngle) * radius, centerY + Sin(currentAngle) * radius
    Next i
    rlEnd
End Sub
