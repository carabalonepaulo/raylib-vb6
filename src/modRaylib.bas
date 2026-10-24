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
    KEY_NULL = 0

    KEY_APOSTROPHE = 39
    KEY_COMMA = 44
    KEY_MINUS = 45
    KEY_PERIOD = 46
    KEY_SLASH = 47
    KEY_ZERO = 48
    KEY_ONE = 49
    KEY_TWO = 50
    KEY_THREE = 51
    KEY_FOUR = 52
    KEY_FIVE = 53
    KEY_SIX = 54
    KEY_SEVEN = 55
    KEY_EIGHT = 56
    KEY_NINE = 57
    KEY_SEMICOLON = 59
    KEY_EQUAL = 61

    KEY_A = 65
    KEY_B = 66
    KEY_C = 67
    KEY_D = 68
    KEY_E = 69
    KEY_F = 70
    KEY_G = 71
    KEY_H = 72
    KEY_I = 73
    KEY_J = 74
    KEY_K = 75
    KEY_L = 76
    KEY_M = 77
    KEY_N = 78
    KEY_O = 79
    KEY_P = 80
    KEY_Q = 81
    KEY_R = 82
    KEY_S = 83
    KEY_T = 84
    KEY_U = 85
    KEY_V = 86
    KEY_W = 87
    KEY_X = 88
    KEY_Y = 89
    KEY_Z = 90

    KEY_LEFT_BRACKET = 91
    KEY_BACKSLASH = 92
    KEY_RIGHT_BRACKET = 93
    KEY_GRAVE = 96

    KEY_SPACE = 32
    KEY_ESCAPE = 27
    KEY_ENTER = 13
    KEY_TAB = 9
    KEY_BACKSPACE = 8
    KEY_INSERT = 45
    KEY_DELETE = 46
    KEY_RIGHT = 39
    KEY_LEFT = 37
    KEY_DOWN = 40
    KEY_UP = 38
    KEY_PAGE_UP = 33
    KEY_PAGE_DOWN = 34
    KEY_HOME = 36
    KEY_END = 35
    KEY_CAPS_LOCK = 20
    KEY_SCROLL_LOCK = 145
    KEY_NUM_LOCK = 144
    KEY_PRINT_SCREEN = 44
    KEY_PAUSE = 19

    KEY_F1 = 112
    KEY_F2 = 113
    KEY_F3 = 114
    KEY_F4 = 115
    KEY_F5 = 116
    KEY_F6 = 117
    KEY_F7 = 118
    KEY_F8 = 119
    KEY_F9 = 120
    KEY_F10 = 121
    KEY_F11 = 122
    KEY_F12 = 123

    KEY_LEFT_SHIFT = 16
    KEY_LEFT_CONTROL = 17
    KEY_LEFT_ALT = 18
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

' GDI+ (image loading)

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

' Render target type
Public Type RenderTexture2D
    Id As Long
    texture As Texture2D
    depth As Texture2D
End Type

' Image formats
Public Enum PixelFormat
    PIXELFORMAT_UNCOMPRESSED_GRAYSCALE = 1
    PIXELFORMAT_UNCOMPRESSED_GRAY_ALPHA
    PIXELFORMAT_UNCOMPRESSED_R5G6B5
    PIXELFORMAT_UNCOMPRESSED_R8G8B8
    PIXELFORMAT_UNCOMPRESSED_R5G5B5A1
    PIXELFORMAT_UNCOMPRESSED_R4G4B4A4
    PIXELFORMAT_UNCOMPRESSED_R8G8B8A8
    PIXELFORMAT_UNCOMPRESSED_R32
    PIXELFORMAT_UNCOMPRESSED_R32G32B32
    PIXELFORMAT_UNCOMPRESSED_R32G32B32A32
    PIXELFORMAT_COMPRESSED_DXT1_RGB
    PIXELFORMAT_COMPRESSED_DXT1_RGBA
    PIXELFORMAT_COMPRESSED_DXT3_RGBA
    PIXELFORMAT_COMPRESSED_DXT5_RGBA
    PIXELFORMAT_COMPRESSED_ETC1_RGB
    PIXELFORMAT_COMPRESSED_ETC2_RGB
    PIXELFORMAT_COMPRESSED_ETC2_EAC_RGBA
    PIXELFORMAT_COMPRESSED_PVRT_RGB
    PIXELFORMAT_COMPRESSED_PVRT_RGBA
    PIXELFORMAT_COMPRESSED_ASTC_4x4_RGBA
    PIXELFORMAT_COMPRESSED_ASTC_8x8_RGBA
End Enum

Public Type Image
    Data() As Byte
    Width As Long
    Height As Long
    Mipmaps As Long
    Format As Long
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

Private mShouldClose As Boolean
Private mFps As Long
Private mExitKey As Long
Private mEventWaiting As Boolean
Private mFpsFrames As Long
Private mFpsTimer As Single

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
Private mTargetForm As Form

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
Private mQpcFreq As Currency
Private mFrameMark As Currency

' Screen (logical, what you draw with) / render (framebuffer, screen * DPI) / DPI
Private mRenderWidth As Long
Private mRenderHeight As Long
Private mScreenScaleX As Single
Private mScreenScaleY As Single
Private mDpiScaleX As Single
Private mDpiScaleY As Single

Private mMonitors() As MONITORINFO
Private mMonitorHandles() As Long
Private mMonitorCount As Long
Private mMonitorsReady As Boolean

' Perlin noise
Private mPerm(0 To 511) As Long
Private mPermInit As Boolean

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

' Render target state

Private Const RENDER_STACK_MAX As Long = 16
Private mRenderFbo(0 To RENDER_STACK_MAX - 1) As Long
Private mRenderW(0 To RENDER_STACK_MAX - 1) As Long
Private mRenderH(0 To RENDER_STACK_MAX - 1) As Long
Private mRenderTop As Long

' Core declares (Win32)

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

' ConfigFlags
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
Private Const WM_CLOSE As Long = &H10
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
Private Declare Function WaitMessage Lib "user32" () As Long
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

' OpenGL context

Private Sub InitImageLoader()
    Dim token As Long
    Dim si As GsiStartupInput
    si.GdiplusVersion = 1
    GdiplusStartup token, si, 0
End Sub

Public Function InitWindow(ByRef frmMain As Form, ByVal Width As Long, ByVal Height As Long, ByVal title As String) As Boolean
    Set mTargetForm = frmMain

    Load frmMain
    frmMain.ScaleMode = 3
    frmMain.Width = (Width * Screen.TwipsPerPixelX) + (frmMain.Width - (frmMain.ScaleWidth * Screen.TwipsPerPixelX))
    frmMain.Height = (Height * Screen.TwipsPerPixelY) + (frmMain.Height - (frmMain.ScaleHeight * Screen.TwipsPerPixelY))
    frmMain.left = (Screen.Width - frmMain.Width) / 2
    frmMain.top = (Screen.Height - frmMain.Height) / 2
    If Len(title) > 0 Then frmMain.Caption = title

    frmMain.Show
    DoEvents

    If InitOpenGL(frmMain, Width, Height) = False Then
        Unload frmMain
        Exit Function
    End If

    QueryPerformanceFrequency mQpcFreq
    QueryPerformanceCounter mFrameMark

    ResetCoreState
    mWindowReady = True
    ApplyConfigFlags

    InitWindow = True
End Function

Private Function InitOpenGL(ByRef frmMain As Form, ByVal Width As Long, ByVal Height As Long) As Boolean
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

    EnsureSubclass

    If wglMakeCurrent(hDC, m_hRC) = 0 Then
        Debug.Print "wglMakeCurrent falhou"
        Exit Function
    End If

    screenWidth = Width
    screenHeight = Height
    UpdateRenderSize

    glViewport 0, 0, mRenderWidth, mRenderHeight

    glEnable GL_BLEND
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA

    InitImageLoader
    InitColors
    rlglInit

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

Public Function IsVSyncEnabled() As Boolean
    IsVSyncEnabled = mVSyncOn
End Function

Private Sub SetVSyncInternal(ByVal enabled As Boolean)
    If GLVSyncSupported() = False Then Exit Sub
    mVSyncOn = enabled
    If enabled Then
        GLSetSwapInterval 1
    Else
        GLSetSwapInterval 0
    End If
End Sub

' Core

Private Sub ResetCoreState()
    mMouseScaleX = 1!
    mMouseScaleY = 1!
    mMouseOffsetX = 0
    mMouseOffsetY = 0
    mMouseWheelVX = 0!
    mMouseWheelVY = 0!
    mCursorHidden = False
    mFullscreen = False
    mShouldClose = False
    mExitKey = KEY_ESCAPE
    mEventWaiting = False
    mMinSizeW = 0
    mMinSizeH = 0
    mFrameTime = 0!
    mTime = 0#
    mFps = 0
    mFpsFrames = 0
    mFpsTimer = 0!
    mWindowResized = False
End Sub

' Timing

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

' Window state

Public Function WindowShouldClose() As Boolean
    WindowShouldClose = mShouldClose
End Function

Public Sub CloseWindow()
    If mWindowReady = False Then
        Exit Sub
    End If
    mShouldClose = True

    DeinitWindowState
    StopMusic
    Unload mTargetForm
    Set mTargetForm = Nothing
    DeinitOpenGL
    mWindowReady = False
End Sub

Public Sub PollInputEvents()
    UpdateInput
    UpdateMouse m_hWnd

    If mExitKey >= 0 And mExitKey <= 255 Then
        If Keys(mExitKey).IsDown And Not Keys(mExitKey).WasDown Then mShouldClose = True
    End If

    mWindowResized = False

    If mEventWaiting Then WaitMessage
    DoEvents
End Sub

Public Sub SwapScreenBuffer()
    SwapBuffers m_hDC
End Sub

Public Sub EnableEventWaiting()
    mEventWaiting = True
End Sub

Public Sub DisableEventWaiting()
    mEventWaiting = False
End Sub

Public Sub SetExitKey(ByVal key As KeyboardKey)
    mExitKey = key
End Sub

Public Function IsWindowState(ByVal flag As Long) As Boolean
    IsWindowState = ((mConfigFlags And flag) <> 0)
End Function

Public Function GetScreenWidth() As Long
    GetScreenWidth = screenWidth
End Function

Public Function GetScreenHeight() As Long
    GetScreenHeight = screenHeight
End Function

Public Function GetRenderWidth() As Long
    GetRenderWidth = mRenderWidth
End Function

Public Function GetRenderHeight() As Long
    GetRenderHeight = mRenderHeight
End Function

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

' Window (setters)

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

' ConfigFlags / states

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
    If (flags And FLAG_WINDOW_HIGHDPI) <> 0 Then RefreshRenderSize
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
    If (flags And FLAG_WINDOW_HIGHDPI) <> 0 Then RefreshRenderSize
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

' Fullscreen

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

' Subclass (min size)

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
    If msg = WM_CLOSE Then
        ' Close request: let WindowShouldClose report it and keep the window
        ' alive until the user calls CloseWindow
        mShouldClose = True
        WindowProc = 0
        Exit Function
    End If

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

' Monitor

Private Sub EnsureMonitors()
    If mMonitorsReady Then Exit Sub
    mMonitorsReady = True
    mMonitorCount = 0
    ReDim mMonitors(0 To 15)
    ReDim mMonitorHandles(0 To 15)
    EnumDisplayMonitors 0, 0, AddressOf EnumMonitorsProc, 0

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

' Clipboard

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

' Random

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

' Screenshot

Public Sub TakeScreenshot(ByVal fileName As String)
    Dim w As Long, h As Long
    w = mRenderWidth
    h = mRenderHeight
    If w <= 0 Or h <= 0 Then Exit Sub

    Dim pixels() As Byte
    Dim flipped() As Byte
    ReDim pixels(0 To (w * h * 4) - 1)
    ReDim flipped(0 To (w * h * 4) - 1)

    rlDrawRenderBatchActive

    glReadPixels 0, 0, w, h, GL_RGBA, GL_UNSIGNED_BYTE, ByVal VarPtr(pixels(0))

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

' Cursor

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
    resultImage.Mipmaps = 1
    resultImage.Format = PIXELFORMAT_UNCOMPRESSED_R8G8B8A8

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

' Pixel access (format-aware)

Private Function GetPixelDataSize(ByVal format As Long) As Long
    Select Case format
        Case PIXELFORMAT_UNCOMPRESSED_GRAYSCALE: GetPixelDataSize = 1
        Case PIXELFORMAT_UNCOMPRESSED_GRAY_ALPHA: GetPixelDataSize = 2
        Case PIXELFORMAT_UNCOMPRESSED_R5G6B5: GetPixelDataSize = 2
        Case PIXELFORMAT_UNCOMPRESSED_R5G5B5A1: GetPixelDataSize = 2
        Case PIXELFORMAT_UNCOMPRESSED_R4G4B4A4: GetPixelDataSize = 2
        Case PIXELFORMAT_UNCOMPRESSED_R8G8B8: GetPixelDataSize = 3
        Case PIXELFORMAT_UNCOMPRESSED_R8G8B8A8: GetPixelDataSize = 4
        Case PIXELFORMAT_UNCOMPRESSED_R32: GetPixelDataSize = 4
        Case PIXELFORMAT_UNCOMPRESSED_R32G32B32: GetPixelDataSize = 12
        Case PIXELFORMAT_UNCOMPRESSED_R32G32B32A32: GetPixelDataSize = 16
        Case Else: GetPixelDataSize = 0
    End Select
End Function

Private Function ColorToByte(ByVal v As Single) As Byte
    If v <= 0! Then
        ColorToByte = 0
    ElseIf v >= 1! Then
        ColorToByte = 255
    Else
        ColorToByte = CByte(Int(v * 255# + 0.5))
    End If
End Function

Private Function GrayFromRGB(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte) As Byte
    Dim v As Long
    v = (CLng(r) * 299& + CLng(g) * 587& + CLng(b) * 114&) \ 1000&
    If v > 255& Then v = 255&
    GrayFromRGB = CByte(v)
End Function

Private Sub ReadPixelBytes(ByRef img As Image, ByVal x As Long, ByVal y As Long, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte, ByRef a As Byte)
    Dim idx As Long
    Dim v As Long

    r = 0: g = 0: b = 0: a = 255

    If x < 0 Or x >= img.Width Or y < 0 Or y >= img.Height Then Exit Sub

    idx = (y * img.Width + x) * GetPixelDataSize(img.Format)

    Select Case img.Format
        Case PIXELFORMAT_UNCOMPRESSED_GRAYSCALE
            r = img.Data(idx): g = r: b = r: a = 255
        Case PIXELFORMAT_UNCOMPRESSED_GRAY_ALPHA
            r = img.Data(idx): g = r: b = r: a = img.Data(idx + 1)
        Case PIXELFORMAT_UNCOMPRESSED_R5G6B5
            v = CLng(img.Data(idx)) Or (CLng(img.Data(idx + 1)) * 256&)
            r = CByte((((v \ 2048&) And &H1F) * 255&) \ 31&)
            g = CByte((((v \ 32&) And &H3F) * 255&) \ 63&)
            b = CByte(((v And &H1F) * 255&) \ 31&)
            a = 255
        Case PIXELFORMAT_UNCOMPRESSED_R8G8B8
            r = img.Data(idx): g = img.Data(idx + 1): b = img.Data(idx + 2): a = 255
        Case PIXELFORMAT_UNCOMPRESSED_R5G5B5A1
            v = CLng(img.Data(idx)) Or (CLng(img.Data(idx + 1)) * 256&)
            r = CByte((((v \ 2048&) And &H1F) * 255&) \ 31&)
            g = CByte((((v \ 64&) And &H1F) * 255&) \ 31&)
            b = CByte((((v \ 2&) And &H1F) * 255&) \ 31&)
            a = CByte((v And &H1&) * 255&)
        Case PIXELFORMAT_UNCOMPRESSED_R4G4B4A4
            v = CLng(img.Data(idx)) Or (CLng(img.Data(idx + 1)) * 256&)
            r = CByte(((v \ 4096&) And &HF&) * 17&)
            g = CByte(((v \ 256&) And &HF&) * 17&)
            b = CByte(((v \ 16&) And &HF&) * 17&)
            a = CByte((v And &HF&) * 17&)
        Case PIXELFORMAT_UNCOMPRESSED_R8G8B8A8
            r = img.Data(idx): g = img.Data(idx + 1): b = img.Data(idx + 2): a = img.Data(idx + 3)
    End Select
End Sub

Private Sub WritePixelBytes(ByRef img As Image, ByVal x As Long, ByVal y As Long, ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, ByVal a As Byte)
    Dim idx As Long
    Dim v As Long

    If x < 0 Or x >= img.Width Or y < 0 Or y >= img.Height Then Exit Sub

    idx = (y * img.Width + x) * GetPixelDataSize(img.Format)

    Select Case img.Format
        Case PIXELFORMAT_UNCOMPRESSED_GRAYSCALE
            img.Data(idx) = GrayFromRGB(r, g, b)
        Case PIXELFORMAT_UNCOMPRESSED_GRAY_ALPHA
            img.Data(idx) = GrayFromRGB(r, g, b)
            img.Data(idx + 1) = a
        Case PIXELFORMAT_UNCOMPRESSED_R5G6B5
            v = ((CLng(r) \ 8&) * 2048&) Or ((CLng(g) \ 4&) * 32&) Or (CLng(b) \ 8&)
            img.Data(idx) = CByte(v And &HFF&)
            img.Data(idx + 1) = CByte((v \ 256&) And &HFF&)
        Case PIXELFORMAT_UNCOMPRESSED_R8G8B8
            img.Data(idx) = r: img.Data(idx + 1) = g: img.Data(idx + 2) = b
        Case PIXELFORMAT_UNCOMPRESSED_R5G5B5A1
            v = ((CLng(r) \ 8&) * 2048&) Or ((CLng(g) \ 8&) * 64&) Or ((CLng(b) \ 8&) * 2&) Or (CLng(a) \ 128&)
            img.Data(idx) = CByte(v And &HFF&)
            img.Data(idx + 1) = CByte((v \ 256&) And &HFF&)
        Case PIXELFORMAT_UNCOMPRESSED_R4G4B4A4
            v = ((CLng(r) \ 16&) * 4096&) Or ((CLng(g) \ 16&) * 256&) Or ((CLng(b) \ 16&) * 16&) Or (CLng(a) \ 16&)
            img.Data(idx) = CByte(v And &HFF&)
            img.Data(idx + 1) = CByte((v \ 256&) And &HFF&)
        Case PIXELFORMAT_UNCOMPRESSED_R8G8B8A8
            img.Data(idx) = r: img.Data(idx + 1) = g: img.Data(idx + 2) = b: img.Data(idx + 3) = a
    End Select
End Sub

Private Sub GetImageRGBA8(ByRef img As Image, ByRef out() As Byte)
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte
    Dim i As Long

    If img.Width <= 0 Or img.Height <= 0 Then
        Erase out
        Exit Sub
    End If

    ReDim out(0 To img.Width * img.Height * 4 - 1)

    If img.Format = PIXELFORMAT_UNCOMPRESSED_R8G8B8A8 Then
        CopyMemory out(0), img.Data(0), img.Width * img.Height * 4
        Exit Sub
    End If

    i = 0
    For y = 0 To img.Height - 1
        For x = 0 To img.Width - 1
            ReadPixelBytes img, x, y, r, g, b, a
            out(i) = r: out(i + 1) = g: out(i + 2) = b: out(i + 3) = a
            i = i + 4
        Next x
    Next y
End Sub

' Image public API

Public Function IsImageReady(ByRef img As Image) As Boolean
    IsImageReady = (img.Width > 0) And (img.Height > 0) And (GetPixelDataSize(img.Format) > 0)
End Function

Public Function GetImageColor(ByRef img As Image, ByVal x As Long, ByVal y As Long) As ColorRGBA
    Dim r As Byte, g As Byte, b As Byte, a As Byte
    Dim col As ColorRGBA
    ReadPixelBytes img, x, y, r, g, b, a
    Color col, CSng(r) / 255#, CSng(g) / 255#, CSng(b) / 255#, CSng(a) / 255#
    GetImageColor = col
End Function

Public Sub ImageDrawPixel(ByRef dst As Image, ByVal x As Long, ByVal y As Long, ByRef color As ColorRGBA)
    WritePixelBytes dst, x, y, _
        ColorToByte(color.r), ColorToByte(color.g), ColorToByte(color.b), ColorToByte(color.a)
End Sub

Public Sub ImageDrawPixelV(ByRef dst As Image, ByRef position As Vector2, ByRef color As ColorRGBA)
    ImageDrawPixel dst, CLng(position.x), CLng(position.y), color
End Sub

Public Sub ImageClearBackground(ByRef dst As Image, ByRef color As ColorRGBA)
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte
    r = ColorToByte(color.r): g = ColorToByte(color.g)
    b = ColorToByte(color.b): a = ColorToByte(color.a)
    For y = 0 To dst.Height - 1
        For x = 0 To dst.Width - 1
            WritePixelBytes dst, x, y, r, g, b, a
        Next x
    Next y
End Sub

Public Sub ImageDrawRectangle(ByRef dst As Image, ByVal posX As Long, ByVal posY As Long, ByVal width As Long, ByVal height As Long, ByRef color As ColorRGBA)
    Dim x As Long, y As Long
    For y = posY To posY + height - 1
        For x = posX To posX + width - 1
            ImageDrawPixel dst, x, y, color
        Next x
    Next y
End Sub

Public Sub ImageDrawRectangleV(ByRef dst As Image, ByRef position As Vector2, ByRef size As Vector2, ByRef color As ColorRGBA)
    ImageDrawRectangle dst, CLng(position.x), CLng(position.y), CLng(size.x), CLng(size.y), color
End Sub

Public Sub ImageDrawRectangleRec(ByRef dst As Image, ByRef rec As Rectangle, ByRef color As ColorRGBA)
    ImageDrawRectangle dst, CLng(rec.x), CLng(rec.y), CLng(rec.Width), CLng(rec.Height), color
End Sub

Public Sub ImageDraw(ByRef dst As Image, ByRef src As Image, ByRef srcRec As Rectangle, ByRef dstRec As Rectangle, ByRef tint As ColorRGBA)
    Dim scaleX As Double, scaleY As Double
    Dim srcBaseX As Double, srcBaseY As Double
    Dim outBaseX As Long, outBaseY As Long
    Dim dstX As Long, dstY As Long
    Dim srcX As Long, srcY As Long
    Dim outX As Long, outY As Long
    Dim srcCol As ColorRGBA
    Dim dstCol As ColorRGBA
    Dim finalCol As ColorRGBA

    If srcRec.Width <= 0! Or srcRec.Height <= 0! Or dstRec.Width <= 0! Or dstRec.Height <= 0! Then Exit Sub
    If IsImageReady(src) = False Then Exit Sub

    scaleX = CDbl(srcRec.Width) / CDbl(dstRec.Width)
    scaleY = CDbl(srcRec.Height) / CDbl(dstRec.Height)
    srcBaseX = CDbl(srcRec.x)
    srcBaseY = CDbl(srcRec.y)
    outBaseX = CLng(Int(dstRec.x))
    outBaseY = CLng(Int(dstRec.y))

    For dstY = 0 To CLng(dstRec.Height) - 1
        srcY = CLng(Int(srcBaseY + dstY * scaleY))
        outY = outBaseY + dstY
        For dstX = 0 To CLng(dstRec.Width) - 1
            srcX = CLng(Int(srcBaseX + dstX * scaleX))
            outX = outBaseX + dstX

            srcCol = GetImageColor(src, srcX, srcY)
            dstCol = GetImageColor(dst, outX, outY)
            finalCol = ColorAlphaBlend(dstCol, srcCol, tint)
            ImageDrawPixel dst, outX, outY, finalCol
        Next dstX
    Next dstY
End Sub

' Direct 2D pixel drawing
Private Function Trunc(ByVal v As Single) As Long
    Trunc = CLng(Fix(v))
End Function

Private Function MinF3(ByVal a As Single, ByVal b As Single, ByVal c As Single) As Single
    Dim m As Single
    m = a
    If b < m Then m = b
    If c < m Then m = c
    MinF3 = m
End Function

Private Function MaxF3(ByVal a As Single, ByVal b As Single, ByVal c As Single) As Single
    Dim m As Single
    m = a
    If b > m Then m = b
    If c > m Then m = c
    MaxF3 = m
End Function

Public Sub ImageDrawLine(ByRef dst As Image, ByVal startPosX As Long, ByVal startPosY As Long, ByVal endPosX As Long, ByVal endPosY As Long, ByRef color As ColorRGBA)
    Dim x As Long, y As Long
    Dim dx As Long, dy As Long
    Dim sx As Long, sy As Long
    Dim err As Long, e2 As Long

    x = startPosX
    y = startPosY
    dx = Abs(endPosX - startPosX)
    sx = IIf(startPosX < endPosX, 1, -1)
    dy = -Abs(endPosY - startPosY)
    sy = IIf(startPosY < endPosY, 1, -1)
    err = dx + dy

    Do
        ImageDrawPixel dst, x, y, color
        If x = endPosX And y = endPosY Then Exit Do
        e2 = 2 * err
        If e2 >= dy Then
            err = err + dy
            x = x + sx
        End If
        If e2 <= dx Then
            err = err + dx
            y = y + sy
        End If
    Loop
End Sub

Public Sub ImageDrawLineV(ByRef dst As Image, ByRef start As Vector2, ByRef finish As Vector2, ByRef color As ColorRGBA)
    ImageDrawLine dst, Trunc(start.x), Trunc(start.y), Trunc(finish.x), Trunc(finish.y), color
End Sub

Public Sub ImageDrawLineEx(ByRef dst As Image, ByRef start As Vector2, ByRef finish As Vector2, ByVal thick As Long, ByRef color As ColorRGBA)
    Dim dx As Single, dy As Single, mag As Single
    Dim ox As Single, oy As Single
    Dim p1 As Vector2, p2 As Vector2, p3 As Vector2, p4 As Vector2

    dx = finish.x - start.x
    dy = finish.y - start.y
    mag = Sqr(dx * dx + dy * dy)

    If mag <= 0! Then
        ImageDrawPixel dst, Trunc(start.x), Trunc(start.y), color
        Exit Sub
    End If

    ox = (dy / mag) * CSng(thick) / 2!
    oy = -(dx / mag) * CSng(thick) / 2!

    Vec p1, start.x + ox, start.y + oy
    Vec p2, finish.x + ox, finish.y + oy
    Vec p3, finish.x - ox, finish.y - oy
    Vec p4, start.x - ox, start.y - oy

    ImageDrawTriangle dst, p1, p2, p3, color
    ImageDrawTriangle dst, p1, p3, p4, color
End Sub

Public Sub ImageDrawTriangle(ByRef dst As Image, ByRef v1 As Vector2, ByRef v2 As Vector2, ByRef v3 As Vector2, ByRef color As ColorRGBA)
    Dim minX As Long, maxX As Long, minY As Long, maxY As Long
    Dim x As Long, y As Long
    Dim det As Double
    Dim w1 As Double, w2 As Double, w3 As Double

    det = (CDbl(v2.y) - CDbl(v3.y)) * (CDbl(v1.x) - CDbl(v3.x)) + _
          (CDbl(v3.x) - CDbl(v2.x)) * (CDbl(v1.y) - CDbl(v3.y))
    If det = 0# Then Exit Sub

    minX = Trunc(MinF3(v1.x, v2.x, v3.x))
    maxX = Trunc(MaxF3(v1.x, v2.x, v3.x))
    minY = Trunc(MinF3(v1.y, v2.y, v3.y))
    maxY = Trunc(MaxF3(v1.y, v2.y, v3.y))

    For y = minY To maxY
        For x = minX To maxX
            w1 = ((CDbl(v2.y) - CDbl(v3.y)) * (CDbl(x) - CDbl(v3.x)) + _
                  (CDbl(v3.x) - CDbl(v2.x)) * (CDbl(y) - CDbl(v3.y))) / det
            w2 = ((CDbl(v3.y) - CDbl(v1.y)) * (CDbl(x) - CDbl(v3.x)) + _
                  (CDbl(v1.x) - CDbl(v3.x)) * (CDbl(y) - CDbl(v3.y))) / det
            w3 = 1# - w1 - w2
            If w1 >= 0# And w2 >= 0# And w3 >= 0# Then
                ImageDrawPixel dst, x, y, color
            End If
        Next x
    Next y
End Sub

Public Sub ImageDrawTriangleEx(ByRef dst As Image, ByRef v1 As Vector2, ByRef v2 As Vector2, ByRef v3 As Vector2, ByRef c1 As ColorRGBA, ByRef c2 As ColorRGBA, ByRef c3 As ColorRGBA)
    Dim minX As Long, maxX As Long, minY As Long, maxY As Long
    Dim x As Long, y As Long
    Dim det As Double
    Dim w1 As Double, w2 As Double, w3 As Double
    Dim col As ColorRGBA

    det = (CDbl(v2.y) - CDbl(v3.y)) * (CDbl(v1.x) - CDbl(v3.x)) + _
          (CDbl(v3.x) - CDbl(v2.x)) * (CDbl(v1.y) - CDbl(v3.y))
    If det = 0# Then Exit Sub

    minX = Trunc(MinF3(v1.x, v2.x, v3.x))
    maxX = Trunc(MaxF3(v1.x, v2.x, v3.x))
    minY = Trunc(MinF3(v1.y, v2.y, v3.y))
    maxY = Trunc(MaxF3(v1.y, v2.y, v3.y))

    For y = minY To maxY
        For x = minX To maxX
            w1 = ((CDbl(v2.y) - CDbl(v3.y)) * (CDbl(x) - CDbl(v3.x)) + _
                  (CDbl(v3.x) - CDbl(v2.x)) * (CDbl(y) - CDbl(v3.y))) / det
            w2 = ((CDbl(v3.y) - CDbl(v1.y)) * (CDbl(x) - CDbl(v3.x)) + _
                  (CDbl(v1.x) - CDbl(v3.x)) * (CDbl(y) - CDbl(v3.y))) / det
            w3 = 1# - w1 - w2
            If w1 >= 0# And w2 >= 0# And w3 >= 0# Then
                col.r = Clamp01(CSng(w1) * c1.r + CSng(w2) * c2.r + CSng(w3) * c3.r)
                col.g = Clamp01(CSng(w1) * c1.g + CSng(w2) * c2.g + CSng(w3) * c3.g)
                col.b = Clamp01(CSng(w1) * c1.b + CSng(w2) * c2.b + CSng(w3) * c3.b)
                col.a = Clamp01(CSng(w1) * c1.a + CSng(w2) * c2.a + CSng(w3) * c3.a)
                ImageDrawPixel dst, x, y, col
            End If
        Next x
    Next y
End Sub

Public Sub ImageDrawTriangleLines(ByRef dst As Image, ByRef v1 As Vector2, ByRef v2 As Vector2, ByRef v3 As Vector2, ByRef color As ColorRGBA)
    ImageDrawLineV dst, v1, v2, color
    ImageDrawLineV dst, v2, v3, color
    ImageDrawLineV dst, v3, v1, color
End Sub

Public Sub ImageDrawTriangleFan(ByRef dst As Image, ByRef points() As Vector2, ByVal pointCount As Long, ByRef color As ColorRGBA)
    Dim i As Long
    If pointCount < 3 Then Exit Sub
    For i = 1 To pointCount - 2
        ImageDrawTriangle dst, points(0), points(i), points(i + 1), color
    Next i
End Sub

Public Sub ImageDrawTriangleStrip(ByRef dst As Image, ByRef points() As Vector2, ByVal pointCount As Long, ByRef color As ColorRGBA)
    Dim i As Long
    If pointCount < 3 Then Exit Sub
    For i = 2 To pointCount - 1
        If (i Mod 2) = 0 Then
            ImageDrawTriangle dst, points(i - 2), points(i - 1), points(i), color
        Else
            ImageDrawTriangle dst, points(i - 1), points(i - 2), points(i), color
        End If
    Next i
End Sub

Public Sub ImageDrawCircle(ByRef dst As Image, ByVal centerX As Long, ByVal centerY As Long, ByVal radius As Long, ByRef color As ColorRGBA)
    Dim x As Long, y As Long, de As Long
    x = 0
    y = radius
    de = 3 - 2 * radius

    Do While y >= x
        ImageDrawLine dst, centerX - x, centerY + y, centerX + x, centerY + y, color
        ImageDrawLine dst, centerX - y, centerY + x, centerX + y, centerY + x, color
        ImageDrawLine dst, centerX - y, centerY - x, centerX + y, centerY - x, color
        ImageDrawLine dst, centerX - x, centerY - y, centerX + x, centerY - y, color
        x = x + 1
        If de <= 0 Then
            de = de + 4 * x + 6
        Else
            y = y - 1
            de = de + 4 * (x - y) + 10
        End If
    Loop
End Sub

Public Sub ImageDrawCircleV(ByRef dst As Image, ByRef center As Vector2, ByVal radius As Single, ByRef color As ColorRGBA)
    ImageDrawCircle dst, Trunc(center.x), Trunc(center.y), Trunc(radius), color
End Sub

Public Sub ImageDrawCircleLines(ByRef dst As Image, ByVal centerX As Long, ByVal centerY As Long, ByVal radius As Long, ByRef color As ColorRGBA)
    Dim x As Long, y As Long, de As Long
    x = 0
    y = radius
    de = 3 - 2 * radius

    Do While y >= x
        ImageDrawPixel dst, centerX + x, centerY + y, color
        ImageDrawPixel dst, centerX - x, centerY + y, color
        ImageDrawPixel dst, centerX + x, centerY - y, color
        ImageDrawPixel dst, centerX - x, centerY - y, color
        ImageDrawPixel dst, centerX + y, centerY + x, color
        ImageDrawPixel dst, centerX - y, centerY + x, color
        ImageDrawPixel dst, centerX + y, centerY - x, color
        ImageDrawPixel dst, centerX - y, centerY - x, color
        x = x + 1
        If de <= 0 Then
            de = de + 4 * x + 6
        Else
            y = y - 1
            de = de + 4 * (x - y) + 10
        End If
    Loop
End Sub

Public Sub ImageDrawCircleLinesV(ByRef dst As Image, ByRef center As Vector2, ByVal radius As Single, ByRef color As ColorRGBA)
    ImageDrawCircleLines dst, Trunc(center.x), Trunc(center.y), Trunc(radius), color
End Sub

Public Sub ImageDrawRectangleLines(ByRef dst As Image, ByRef rec As Rectangle, ByVal thick As Long, ByRef color As ColorRGBA)
    ImageDrawRectangle dst, Trunc(rec.x), Trunc(rec.y), Trunc(rec.Width), thick, color
    ImageDrawRectangle dst, Trunc(rec.x), Trunc(rec.y + rec.Height) - thick, Trunc(rec.Width), thick, color
    ImageDrawRectangle dst, Trunc(rec.x), Trunc(rec.y), thick, Trunc(rec.Height), color
    ImageDrawRectangle dst, Trunc(rec.x + rec.Width) - thick, Trunc(rec.y), thick, Trunc(rec.Height), color
End Sub

' Copy, crop, resize, transforms

Private Function NewImageLike(ByVal w As Long, ByVal h As Long, ByVal fmt As Long) As Image
    Dim img As Image
    Dim bpp As Long
    img.Width = w
    img.Height = h
    img.Format = fmt
    img.Mipmaps = 1
    bpp = GetPixelDataSize(fmt)
    If w > 0 And h > 0 And bpp > 0 Then
        ReDim img.Data(0 To w * h * bpp - 1)
    End If
    NewImageLike = img
End Function

Private Function LerpS(ByVal a As Single, ByVal b As Single, ByVal t As Single) As Single
    LerpS = a + (b - a) * t
End Function

Private Function NextPOT(ByVal v As Long) As Long
    Dim p As Long
    p = 1
    Do While p < v
        p = p * 2
    Loop
    NextPOT = p
End Function

Public Function ImageCopy(ByRef image As Image) As Image
    Dim out As Image
    out = NewImageLike(image.Width, image.Height, image.Format)
    out.Mipmaps = image.Mipmaps
    If out.Width > 0 And out.Height > 0 Then
        CopyMemory out.Data(0), image.Data(0), out.Width * out.Height * GetPixelDataSize(out.Format)
    End If
    ImageCopy = out
End Function

Public Function ImageFromImage(ByRef image As Image, ByRef rec As Rectangle) As Image
    Dim result As Image
    Dim cx As Long, cy As Long, cw As Long, ch As Long
    Dim x As Long, y As Long

    If IsImageReady(image) = False Then
        ImageFromImage = result
        Exit Function
    End If

    cx = Trunc(rec.x): cy = Trunc(rec.y)
    cw = Trunc(rec.Width): ch = Trunc(rec.Height)

    If cx >= image.Width Or cy >= image.Height Then
        ImageFromImage = result
        Exit Function
    End If
    If cx < 0 Then cx = 0
    If cy < 0 Then cy = 0
    If cx + cw > image.Width Then cw = image.Width - cx
    If cy + ch > image.Height Then ch = image.Height - cy
    If cw <= 0 Or ch <= 0 Then
        ImageFromImage = result
        Exit Function
    End If

    result = NewImageLike(cw, ch, image.Format)
    For y = 0 To ch - 1
        For x = 0 To cw - 1
            ImageDrawPixel result, x, y, GetImageColor(image, cx + x, cy + y)
        Next x
    Next y

    ImageFromImage = result
End Function

Public Sub ImageCrop(ByRef image As Image, ByRef crop As Rectangle)
    Dim tmp As Image
    tmp = ImageFromImage(image, crop)
    If tmp.Width > 0 And tmp.Height > 0 Then image = tmp
End Sub

Public Sub ImageFlipVertical(ByRef image As Image)
    Dim bpp As Long, rowBytes As Long
    Dim y As Long, top As Long, bot As Long
    Dim tmp() As Byte

    If IsImageReady(image) = False Then Exit Sub
    bpp = GetPixelDataSize(image.Format)
    rowBytes = image.Width * bpp
    ReDim tmp(0 To rowBytes - 1)

    For y = 0 To (image.Height \ 2) - 1
        top = y * rowBytes
        bot = (image.Height - 1 - y) * rowBytes
        CopyMemory tmp(0), image.Data(top), rowBytes
        CopyMemory image.Data(top), image.Data(bot), rowBytes
        CopyMemory image.Data(bot), tmp(0), rowBytes
    Next y
End Sub

Public Sub ImageFlipHorizontal(ByRef image As Image)
    Dim bpp As Long, rowBytes As Long
    Dim x As Long, y As Long, k As Long
    Dim a As Long, b As Long
    Dim t As Byte

    If IsImageReady(image) = False Then Exit Sub
    bpp = GetPixelDataSize(image.Format)
    rowBytes = image.Width * bpp

    For y = 0 To image.Height - 1
        For x = 0 To (image.Width \ 2) - 1
            a = y * rowBytes + x * bpp
            b = y * rowBytes + (image.Width - 1 - x) * bpp
            For k = 0 To bpp - 1
                t = image.Data(a + k)
                image.Data(a + k) = image.Data(b + k)
                image.Data(b + k) = t
            Next k
        Next x
    Next y
End Sub

Public Sub ImageRotateCW(ByRef image As Image)
    Dim bpp As Long, newW As Long, newH As Long
    Dim x As Long, y As Long
    Dim out() As Byte

    If IsImageReady(image) = False Then Exit Sub
    bpp = GetPixelDataSize(image.Format)
    newW = image.Height
    newH = image.Width
    ReDim out(0 To newW * newH * bpp - 1)

    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            CopyMemory out((x * newW + (image.Height - 1 - y)) * bpp), image.Data((y * image.Width + x) * bpp), bpp
        Next x
    Next y

    image.Width = newW
    image.Height = newH
    ReDim image.Data(0 To newW * newH * bpp - 1)
    CopyMemory image.Data(0), out(0), newW * newH * bpp
End Sub

Public Sub ImageRotateCCW(ByRef image As Image)
    Dim bpp As Long, newW As Long, newH As Long
    Dim x As Long, y As Long
    Dim out() As Byte

    If IsImageReady(image) = False Then Exit Sub
    bpp = GetPixelDataSize(image.Format)
    newW = image.Height
    newH = image.Width
    ReDim out(0 To newW * newH * bpp - 1)

    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            CopyMemory out(((image.Width - 1 - x) * newW + y) * bpp), image.Data((y * image.Width + x) * bpp), bpp
        Next x
    Next y

    image.Width = newW
    image.Height = newH
    ReDim image.Data(0 To newW * newH * bpp - 1)
    CopyMemory image.Data(0), out(0), newW * newH * bpp
End Sub

Public Sub ImageRotate(ByRef image As Image, ByVal degrees As Long)
    Dim d As Long
    d = degrees Mod 360
    If d < 0 Then d = d + 360

    Select Case d
        Case 0
        Case 90
            ImageRotateCW image
        Case 180
            ImageFlipVertical image
            ImageFlipHorizontal image
        Case 270
            ImageRotateCCW image
    End Select
End Sub

Public Sub ImageResizeNN(ByRef image As Image, ByVal newWidth As Long, ByVal newHeight As Long)
    Dim oldW As Long, oldH As Long
    Dim xRatio As Long, yRatio As Long
    Dim x As Long, y As Long, srcX As Long, srcY As Long
    Dim tmp As Image

    If IsImageReady(image) = False Then Exit Sub
    If newWidth <= 0 Or newHeight <= 0 Then Exit Sub
    If newWidth = image.Width And newHeight = image.Height Then Exit Sub

    oldW = image.Width
    oldH = image.Height
    xRatio = (oldW * 65536) \ newWidth + 1
    yRatio = (oldH * 65536) \ newHeight + 1

    tmp = NewImageLike(newWidth, newHeight, image.Format)
    For y = 0 To newHeight - 1
        srcY = (y * yRatio) \ 65536
        If srcY > oldH - 1 Then srcY = oldH - 1
        For x = 0 To newWidth - 1
            srcX = (x * xRatio) \ 65536
            If srcX > oldW - 1 Then srcX = oldW - 1
            ImageDrawPixel tmp, x, y, GetImageColor(image, srcX, srcY)
        Next x
    Next y

    image = tmp
End Sub

Public Sub ImageResize(ByRef image As Image, ByVal newWidth As Long, ByVal newHeight As Long)
    Dim oldW As Long, oldH As Long
    Dim scaleX As Double, scaleY As Double
    Dim x As Long, y As Long
    Dim srcXf As Double, srcYf As Double
    Dim x0 As Long, y0 As Long, x1 As Long, y1 As Long
    Dim fx As Single, fy As Single
    Dim c00 As ColorRGBA, c10 As ColorRGBA, c01 As ColorRGBA, c11 As ColorRGBA
    Dim col As ColorRGBA
    Dim tmp As Image

    If IsImageReady(image) = False Then Exit Sub
    If newWidth <= 0 Or newHeight <= 0 Then Exit Sub
    If newWidth = image.Width And newHeight = image.Height Then Exit Sub

    oldW = image.Width
    oldH = image.Height
    scaleX = CDbl(oldW) / CDbl(newWidth)
    scaleY = CDbl(oldH) / CDbl(newHeight)

    tmp = NewImageLike(newWidth, newHeight, image.Format)

    For y = 0 To newHeight - 1
        srcYf = (CDbl(y) + 0.5) * scaleY - 0.5
        If srcYf < 0# Then srcYf = 0#
        y0 = CLng(Int(srcYf))
        If y0 > oldH - 1 Then y0 = oldH - 1
        y1 = y0 + 1
        If y1 > oldH - 1 Then y1 = oldH - 1
        fy = CSng(srcYf - CDbl(y0))

        For x = 0 To newWidth - 1
            srcXf = (CDbl(x) + 0.5) * scaleX - 0.5
            If srcXf < 0# Then srcXf = 0#
            x0 = CLng(Int(srcXf))
            If x0 > oldW - 1 Then x0 = oldW - 1
            x1 = x0 + 1
            If x1 > oldW - 1 Then x1 = oldW - 1
            fx = CSng(srcXf - CDbl(x0))

            c00 = GetImageColor(image, x0, y0)
            c10 = GetImageColor(image, x1, y0)
            c01 = GetImageColor(image, x0, y1)
            c11 = GetImageColor(image, x1, y1)

            col.r = Clamp01(LerpS(LerpS(c00.r, c10.r, fx), LerpS(c01.r, c11.r, fx), fy))
            col.g = Clamp01(LerpS(LerpS(c00.g, c10.g, fx), LerpS(c01.g, c11.g, fx), fy))
            col.b = Clamp01(LerpS(LerpS(c00.b, c10.b, fx), LerpS(c01.b, c11.b, fx), fy))
            col.a = Clamp01(LerpS(LerpS(c00.a, c10.a, fx), LerpS(c01.a, c11.a, fx), fy))

            ImageDrawPixel tmp, x, y, col
        Next x
    Next y

    image = tmp
End Sub

Public Sub ImageResizeCanvas(ByRef image As Image, ByVal newWidth As Long, ByVal newHeight As Long, ByVal offsetX As Long, ByVal offsetY As Long, ByRef fill As ColorRGBA)
    Dim oldW As Long, oldH As Long
    Dim x As Long, y As Long
    Dim tmp As Image

    If IsImageReady(image) = False Then Exit Sub
    If newWidth <= 0 Or newHeight <= 0 Then Exit Sub

    oldW = image.Width
    oldH = image.Height

    tmp = NewImageLike(newWidth, newHeight, image.Format)
    ImageClearBackground tmp, fill

    For y = 0 To oldH - 1
        For x = 0 To oldW - 1
            ImageDrawPixel tmp, x + offsetX, y + offsetY, GetImageColor(image, x, y)
        Next x
    Next y

    image = tmp
End Sub

Public Sub ImageToPOT(ByRef image As Image, ByRef fill As ColorRGBA)
    Dim potW As Long, potH As Long
    If IsImageReady(image) = False Then Exit Sub

    potW = NextPOT(image.Width)
    potH = NextPOT(image.Height)

    If potW <> image.Width Or potH <> image.Height Then
        ImageResizeCanvas image, potW, potH, 0, 0, fill
    End If
End Sub

' Color, alpha and format ops

Private Function ClampByteL(ByVal v As Long) As Byte
    If v < 0 Then
        ClampByteL = 0
    ElseIf v > 255 Then
        ClampByteL = 255
    Else
        ClampByteL = CByte(v)
    End If
End Function

Private Function ClampByteD(ByVal v As Double) As Byte
    Dim n As Long
    n = CLng(Int(v + 0.5))
    ClampByteD = ClampByteL(n)
End Function

Private Function FloatToByte(ByVal v As Double) As Byte
    FloatToByte = ClampByteL(CLng(Int(v * 255#)))
End Function

Private Function MulByte(ByVal v As Byte, ByVal f As Single) As Byte
    MulByte = ClampByteL(CLng(Int(v * f)))
End Function

Private Function QuantBits(ByVal v As Long, ByVal bpp As Long) As Long
    Dim shift As Long
    Dim step As Long
    If bpp >= 8 Then
        QuantBits = v
        Exit Function
    End If
    If bpp < 1 Then
        QuantBits = 0
        Exit Function
    End If
    shift = 8 - bpp
    step = 2 ^ shift
    QuantBits = (v \ step) * step
End Function

Public Sub ImageColorTint(ByRef image As Image, ByRef color As ColorRGBA)
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte

    If IsImageReady(image) = False Then Exit Sub

    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            ReadPixelBytes image, x, y, r, g, b, a
            WritePixelBytes image, x, y, MulByte(r, color.r), MulByte(g, color.g), MulByte(b, color.b), MulByte(a, color.a)
        Next x
    Next y
End Sub

Public Sub ImageColorInvert(ByRef image As Image)
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte

    If IsImageReady(image) = False Then Exit Sub

    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            ReadPixelBytes image, x, y, r, g, b, a
            WritePixelBytes image, x, y, 255 - r, 255 - g, 255 - b, a
        Next x
    Next y
End Sub

Public Sub ImageColorGrayscale(ByRef image As Image)
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte
    Dim gray As Byte

    If IsImageReady(image) = False Then Exit Sub

    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            ReadPixelBytes image, x, y, r, g, b, a
            gray = GrayFromRGB(r, g, b)
            WritePixelBytes image, x, y, gray, gray, gray, a
        Next x
    Next y
End Sub

Public Sub ImageColorContrast(ByRef image As Image, ByVal contrast As Single)
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte
    Dim c As Double

    If IsImageReady(image) = False Then Exit Sub
    If contrast < -100! Then contrast = -100!
    If contrast > 100! Then contrast = 100!
    c = CDbl(contrast) / 100# + 1#

    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            ReadPixelBytes image, x, y, r, g, b, a
            WritePixelBytes image, x, y, _
                FloatToByte((r / 255# - 0.5) * c + 0.5), _
                FloatToByte((g / 255# - 0.5) * c + 0.5), _
                FloatToByte((b / 255# - 0.5) * c + 0.5), _
                a
        Next x
    Next y
End Sub

Public Sub ImageColorBrightness(ByRef image As Image, ByVal brightness As Long)
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte

    If IsImageReady(image) = False Then Exit Sub

    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            ReadPixelBytes image, x, y, r, g, b, a
            WritePixelBytes image, x, y, _
                ClampByteL(CLng(r) + brightness), ClampByteL(CLng(g) + brightness), _
                ClampByteL(CLng(b) + brightness), a
        Next x
    Next y
End Sub

Public Sub ImageColorReplace(ByRef image As Image, ByRef color As ColorRGBA, ByRef replace As ColorRGBA)
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte
    Dim cr As Byte, cg As Byte, cb As Byte, ca As Byte
    Dim rr As Byte, rg As Byte, rb As Byte, ra As Byte

    If IsImageReady(image) = False Then Exit Sub
    cr = ColorToByte(color.r): cg = ColorToByte(color.g)
    cb = ColorToByte(color.b): ca = ColorToByte(color.a)
    rr = ColorToByte(replace.r): rg = ColorToByte(replace.g)
    rb = ColorToByte(replace.b): ra = ColorToByte(replace.a)

    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            ReadPixelBytes image, x, y, r, g, b, a
            If r = cr And g = cg And b = cb And a = ca Then
                WritePixelBytes image, x, y, rr, rg, rb, ra
            End If
        Next x
    Next y
End Sub

Public Sub ImageAlphaClear(ByRef image As Image, ByRef color As ColorRGBA, ByVal threshold As Single)
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte
    Dim cr As Byte, cg As Byte, cb As Byte, ca As Byte
    Dim thr As Long

    If IsImageReady(image) = False Then Exit Sub
    cr = ColorToByte(color.r): cg = ColorToByte(color.g)
    cb = ColorToByte(color.b): ca = ColorToByte(color.a)
    thr = CLng(Int(threshold * 255#))

    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            ReadPixelBytes image, x, y, r, g, b, a
            If CLng(a) <= thr Then
                WritePixelBytes image, x, y, cr, cg, cb, ca
            End If
        Next x
    Next y
End Sub

Public Sub ImageAlphaMask(ByRef image As Image, ByRef alphaMask As Image)
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte
    Dim scaleX As Single, scaleY As Single
    Dim srcX As Long, srcY As Long
    Dim mc As ColorRGBA
    Dim gray As Long

    If IsImageReady(image) = False Then Exit Sub
    If IsImageReady(alphaMask) = False Then Exit Sub

    scaleX = CSng(image.Width) / CSng(alphaMask.Width)
    scaleY = CSng(image.Height) / CSng(alphaMask.Height)

    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            srcX = Trunc(CSng(x) * scaleX)
            srcY = Trunc(CSng(y) * scaleY)
            mc = GetImageColor(alphaMask, srcX, srcY)
            gray = CLng((mc.r * 0.299 + mc.g * 0.587 + mc.b * 0.114) * 255#)
            ReadPixelBytes image, x, y, r, g, b, a
            WritePixelBytes image, x, y, r, g, b, CByte((CLng(a) * gray) \ 255&)
        Next x
    Next y
End Sub

Public Sub ImageAlphaPremultiply(ByRef image As Image)
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte
    Dim alpha As Single

    If IsImageReady(image) = False Then Exit Sub

    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            ReadPixelBytes image, x, y, r, g, b, a
            alpha = CSng(a) / 255!
            WritePixelBytes image, x, y, MulByte(r, alpha), MulByte(g, alpha), MulByte(b, alpha), a
        Next x
    Next y
End Sub

Public Sub ImageFormat(ByRef image As Image, ByVal newFormat As Long)
    Dim tmp As Image
    Dim x As Long, y As Long

    If IsImageReady(image) = False Then Exit Sub
    If image.Format = newFormat Then Exit Sub
    If GetPixelDataSize(newFormat) = 0 Then Exit Sub

    tmp = NewImageLike(image.Width, image.Height, newFormat)
    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            ImageDrawPixel tmp, x, y, GetImageColor(image, x, y)
        Next x
    Next y

    image = tmp
End Sub

Private Function DownscaleHalf(ByRef src As Image) As Image
    Dim nw As Long, nh As Long
    Dim dst As Image
    Dim x As Long, y As Long
    Dim sx As Long, sy As Long
    Dim cnt As Long
    Dim r0 As Long, g0 As Long, b0 As Long, a0 As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte

    nw = src.Width \ 2: If nw < 1 Then nw = 1
    nh = src.Height \ 2: If nh < 1 Then nh = 1
    dst = NewImageLike(nw, nh, src.Format)

    For y = 0 To nh - 1
        For x = 0 To nw - 1
            r0 = 0: g0 = 0: b0 = 0: a0 = 0: cnt = 0
            For sy = y * 2 To y * 2 + 1
                For sx = x * 2 To x * 2 + 1
                    If sx < src.Width And sy < src.Height Then
                        ReadPixelBytes src, sx, sy, r, g, b, a
                        r0 = r0 + r: g0 = g0 + g: b0 = b0 + b: a0 = a0 + a
                        cnt = cnt + 1
                    End If
                Next sx
            Next sy
            If cnt > 0 Then
                WritePixelBytes dst, x, y, CByte(r0 \ cnt), CByte(g0 \ cnt), CByte(b0 \ cnt), CByte(a0 \ cnt)
            End If
        Next x
    Next y

    DownscaleHalf = dst
End Function

Public Sub ImageMipmaps(ByRef image As Image)
    Dim levels As Long
    Dim bpp As Long
    Dim w As Long, h As Long
    Dim total As Long
    Dim buffer() As Byte
    Dim off As Long
    Dim lvl As Long
    Dim cur As Image

    If IsImageReady(image) = False Then Exit Sub
    bpp = GetPixelDataSize(image.Format)

    w = image.Width: h = image.Height
    levels = 1
    Do While w > 1 Or h > 1
        w = w \ 2: If w < 1 Then w = 1
        h = h \ 2: If h < 1 Then h = 1
        levels = levels + 1
    Loop

    w = image.Width: h = image.Height
    total = w * h * bpp
    Do While w > 1 Or h > 1
        w = w \ 2: If w < 1 Then w = 1
        h = h \ 2: If h < 1 Then h = 1
        total = total + w * h * bpp
    Loop

    ReDim buffer(0 To total - 1)
    CopyMemory buffer(0), image.Data(0), image.Width * image.Height * bpp
    off = image.Width * image.Height * bpp

    cur = image
    For lvl = 2 To levels
        cur = DownscaleHalf(cur)
        CopyMemory buffer(off), cur.Data(0), cur.Width * cur.Height * bpp
        off = off + cur.Width * cur.Height * bpp
    Next lvl

    ReDim image.Data(0 To total - 1)
    CopyMemory image.Data(0), buffer(0), total
    image.Mipmaps = levels
End Sub

Public Sub ImageBlurGaussian(ByRef image As Image, ByVal blurSize As Long)
    Dim radius As Long, ksz As Long
    Dim sigma As Double
    Dim kernel() As Double
    Dim ksum As Double
    Dim i As Long, j As Long, x As Long, y As Long
    Dim w As Long, h As Long
    Dim accR As Double, accG As Double, accB As Double, accA As Double
    Dim r As Byte, g As Byte, b As Byte, a As Byte
    Dim hr() As Byte, hg() As Byte, hb() As Byte, ha() As Byte
    Dim d As Double
    Dim sx As Long, sy As Long

    If IsImageReady(image) = False Then Exit Sub
    If blurSize <= 1 Then Exit Sub

    radius = blurSize
    ksz = radius * 2 + 1
    sigma = CDbl(blurSize) / 3#
    If sigma < 0.0001 Then sigma = 1#

    ReDim kernel(0 To ksz - 1)
    ksum = 0#
    For i = 0 To ksz - 1
        d = CDbl(i - radius)
        kernel(i) = Exp(-(d * d) / (2# * sigma * sigma))
        ksum = ksum + kernel(i)
    Next i
    For i = 0 To ksz - 1
        kernel(i) = kernel(i) / ksum
    Next i

    w = image.Width: h = image.Height
    ReDim hr(0 To w * h - 1)
    ReDim hg(0 To w * h - 1)
    ReDim hb(0 To w * h - 1)
    ReDim ha(0 To w * h - 1)

    For y = 0 To h - 1
        For x = 0 To w - 1
            accR = 0#: accG = 0#: accB = 0#: accA = 0#
            For j = -radius To radius
                sx = x + j
                If sx < 0 Then sx = 0
                If sx > w - 1 Then sx = w - 1
                ReadPixelBytes image, sx, y, r, g, b, a
                accR = accR + CDbl(r) * kernel(j + radius)
                accG = accG + CDbl(g) * kernel(j + radius)
                accB = accB + CDbl(b) * kernel(j + radius)
                accA = accA + CDbl(a) * kernel(j + radius)
            Next j
            i = y * w + x
            hr(i) = ClampByteD(accR)
            hg(i) = ClampByteD(accG)
            hb(i) = ClampByteD(accB)
            ha(i) = ClampByteD(accA)
        Next x
    Next y

    For y = 0 To h - 1
        For x = 0 To w - 1
            accR = 0#: accG = 0#: accB = 0#: accA = 0#
            For j = -radius To radius
                sy = y + j
                If sy < 0 Then sy = 0
                If sy > h - 1 Then sy = h - 1
                i = sy * w + x
                accR = accR + CDbl(hr(i)) * kernel(j + radius)
                accG = accG + CDbl(hg(i)) * kernel(j + radius)
                accB = accB + CDbl(hb(i)) * kernel(j + radius)
                accA = accA + CDbl(ha(i)) * kernel(j + radius)
            Next j
            WritePixelBytes image, x, y, ClampByteD(accR), ClampByteD(accG), ClampByteD(accB), ClampByteD(accA)
        Next x
    Next y
End Sub

Public Sub ImageDither(ByRef image As Image, ByVal rBpp As Long, ByVal gBpp As Long, ByVal bBpp As Long, ByVal aBpp As Long)
    Dim w As Long, h As Long
    Dim x As Long, y As Long, i As Long
    Dim R() As Long, G() As Long, B() As Long, A() As Long
    Dim cur As ColorRGBA
    Dim orr As Long, ogg As Long, obb As Long, oaa As Long
    Dim qr As Long, qg As Long, qb As Long, qa As Long
    Dim re As Long, ge As Long, be As Long, ae As Long

    If IsImageReady(image) = False Then Exit Sub
    w = image.Width: h = image.Height
    ReDim R(0 To w * h - 1)
    ReDim G(0 To w * h - 1)
    ReDim B(0 To w * h - 1)
    ReDim A(0 To w * h - 1)

    For y = 0 To h - 1
        For x = 0 To w - 1
            cur = GetImageColor(image, x, y)
            i = y * w + x
            R(i) = CLng(Int(cur.r * 255# + 0.5))
            G(i) = CLng(Int(cur.g * 255# + 0.5))
            B(i) = CLng(Int(cur.b * 255# + 0.5))
            A(i) = CLng(Int(cur.a * 255# + 0.5))
        Next x
    Next y

    For y = 0 To h - 1
        For x = 0 To w - 1
            i = y * w + x
            orr = R(i): ogg = G(i): obb = B(i): oaa = A(i)

            qr = QuantBits(orr, rBpp)
            qg = QuantBits(ogg, gBpp)
            qb = QuantBits(obb, bBpp)
            qa = QuantBits(oaa, aBpp)
            R(i) = qr: G(i) = qg: B(i) = qb: A(i) = qa

            re = orr - qr: ge = ogg - qg: be = obb - qb: ae = oaa - qa

            If x + 1 < w Then
                R(i + 1) = ClampByteL(R(i + 1) + ((re * 7) \ 16))
                G(i + 1) = ClampByteL(G(i + 1) + ((ge * 7) \ 16))
                B(i + 1) = ClampByteL(B(i + 1) + ((be * 7) \ 16))
                A(i + 1) = ClampByteL(A(i + 1) + ((ae * 7) \ 16))
            End If
            If y + 1 < h Then
                If x > 0 Then
                    R(i + w - 1) = ClampByteL(R(i + w - 1) + ((re * 3) \ 16))
                    G(i + w - 1) = ClampByteL(G(i + w - 1) + ((ge * 3) \ 16))
                    B(i + w - 1) = ClampByteL(B(i + w - 1) + ((be * 3) \ 16))
                    A(i + w - 1) = ClampByteL(A(i + w - 1) + ((ae * 3) \ 16))
                End If
                R(i + w) = ClampByteL(R(i + w) + ((re * 5) \ 16))
                G(i + w) = ClampByteL(G(i + w) + ((ge * 5) \ 16))
                B(i + w) = ClampByteL(B(i + w) + ((be * 5) \ 16))
                A(i + w) = ClampByteL(A(i + w) + ((ae * 5) \ 16))
                If x + 1 < w Then
                    R(i + w + 1) = ClampByteL(R(i + w + 1) + ((re * 1) \ 16))
                    G(i + w + 1) = ClampByteL(G(i + w + 1) + ((ge * 1) \ 16))
                    B(i + w + 1) = ClampByteL(B(i + w + 1) + ((be * 1) \ 16))
                    A(i + w + 1) = ClampByteL(A(i + w + 1) + ((ae * 1) \ 16))
                End If
            End If
        Next x
    Next y

    For y = 0 To h - 1
        For x = 0 To w - 1
            i = y * w + x
            WritePixelBytes image, x, y, CByte(R(i)), CByte(G(i)), CByte(B(i)), CByte(A(i))
        Next x
    Next y
End Sub

' Screen/texture loading, export

Private Sub FlipRGBAVertical(ByRef data() As Byte, ByVal w As Long, ByVal h As Long)
    Dim rowBytes As Long
    Dim y As Long, top As Long, bot As Long
    Dim tmp() As Byte

    rowBytes = w * 4
    ReDim tmp(0 To rowBytes - 1)
    For y = 0 To (h \ 2) - 1
        top = y * rowBytes
        bot = (h - 1 - y) * rowBytes
        CopyMemory tmp(0), data(top), rowBytes
        CopyMemory data(top), data(bot), rowBytes
        CopyMemory data(bot), tmp(0), rowBytes
    Next y
End Sub

Public Function LoadImageFromScreen() As Image
    Dim img As Image

    img.Width = mRenderWidth
    img.Height = mRenderHeight
    img.Format = PIXELFORMAT_UNCOMPRESSED_R8G8B8A8
    img.Mipmaps = 1

    If img.Width <= 0 Or img.Height <= 0 Then
        img.Width = 0: img.Height = 0
        LoadImageFromScreen = img
        Exit Function
    End If

    ReDim img.Data(0 To img.Width * img.Height * 4 - 1)
    glReadPixels 0, 0, img.Width, img.Height, GL_RGBA, GL_UNSIGNED_BYTE, ByVal VarPtr(img.Data(0))
    FlipRGBAVertical img.Data, img.Width, img.Height

    LoadImageFromScreen = img
End Function

Public Function LoadImageFromTexture(ByRef texture As Texture2D) As Image
    Dim img As Image

    img.Width = texture.Width
    img.Height = texture.Height
    img.Format = PIXELFORMAT_UNCOMPRESSED_R8G8B8A8
    img.Mipmaps = 1

    If img.Width <= 0 Or img.Height <= 0 Then
        img.Width = 0: img.Height = 0
        LoadImageFromTexture = img
        Exit Function
    End If

    ReDim img.Data(0 To img.Width * img.Height * 4 - 1)
    glBindTexture GL_TEXTURE_2D, texture.Id
    glGetTexImage GL_TEXTURE_2D, 0, GL_RGBA, GL_UNSIGNED_BYTE, ByVal VarPtr(img.Data(0))

    LoadImageFromTexture = img
End Function

Private Sub SwapRgbaToBgra(ByRef data() As Byte)
    Dim i As Long
    Dim t As Byte
    For i = 0 To UBound(data) - 3 Step 4
        t = data(i)
        data(i) = data(i + 2)
        data(i + 2) = t
    Next i
End Sub

Public Function ExportImage(ByRef image As Image, ByVal fileName As String) As Boolean
    Dim px() As Byte
    If IsImageReady(image) = False Then Exit Function
    GetImageRGBA8 image, px
    SwapRgbaToBgra px
    SaveImageAsPng GetFullPath(fileName), px, image.Width, image.Height
    ExportImage = True
End Function

Private Function CodeIdentifier(ByVal fileName As String) As String
    Dim base As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    base = fileName
    i = InStrRev(base, "\")
    If i > 0 Then base = Mid$(base, i + 1)
    i = InStrRev(base, ".")
    If i > 0 Then base = Left$(base, i - 1)

    out = ""
    For i = 1 To Len(base)
        ch = Mid$(base, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            out = out & UCase$(ch)
        Else
            out = out & "_"
        End If
    Next i
    If Len(out) = 0 Then out = "IMAGE"
    CodeIdentifier = out
End Function

Private Function ByteArrayToCode(ByRef data() As Byte) As String
    Dim i As Long
    Dim s As String
    Dim out As String

    out = ""
    s = "    "
    For i = 0 To UBound(data)
        s = s & "0x" & Right$("0" & Hex$(data(i)), 2)
        If i < UBound(data) Then s = s & ", "
        If ((i + 1) Mod 12) = 0 Then
            out = out & s & vbCrLf
            s = "    "
        End If
    Next i
    If Len(s) > 4 Then out = out & s & vbCrLf
    ByteArrayToCode = out
End Function

Public Function ExportImageAsCode(ByRef image As Image, ByVal fileName As String) As Boolean
    Dim f As Integer
    Dim id As String

    If IsImageReady(image) = False Then Exit Function

    id = CodeIdentifier(fileName)
    f = FreeFile
    Open GetFullPath(fileName) For Output As #f

    Print #f, "// " & String(84, "/")
    Print #f, "//"
    Print #f, "// ImageAsCode exporter v1.0 - Image pixel data exported as an array of bytes"
    Print #f, "//"
    Print #f, "// " & String(84, "/")
    Print #f, ""
    Print #f, "// Image data information"
    Print #f, "#define " & id & "_WIDTH    " & CStr(image.Width)
    Print #f, "#define " & id & "_HEIGHT   " & CStr(image.Height)
    Print #f, "#define " & id & "_FORMAT   " & CStr(image.Format) & "          // raylib internal pixel format"
    Print #f, ""
    Print #f, "static unsigned char " & id & "_DATA[" & CStr(UBound(image.Data) + 1) & "] = {"
    Print #f, ByteArrayToCode(image.Data);
    Print #f, "};"
    Close #f

    ExportImageAsCode = True
End Function

Public Function GetImageAlphaBorder(ByRef image As Image, ByVal threshold As Single) As Rectangle
    Dim x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, a As Byte
    Dim xMin As Long, yMin As Long, xMax As Long, yMax As Long
    Dim rect As Rectangle

    xMin = 65536: yMin = 65536: xMax = 0: yMax = 0

    If IsImageReady(image) Then
        For y = 0 To image.Height - 1
            For x = 0 To image.Width - 1
                ReadPixelBytes image, x, y, r, g, b, a
                If CSng(a) / 255! > threshold Then
                    If x < xMin Then xMin = x
                    If x > xMax Then xMax = x
                    If y < yMin Then yMin = y
                    If y > yMax Then yMax = y
                End If
            Next x
        Next y
    End If

    rect.x = 0: rect.y = 0: rect.Width = 0: rect.Height = 0
    If xMin <= xMax Then
        rect.x = xMin
        rect.y = yMin
        rect.Width = xMax - xMin + 1
        rect.Height = yMax - yMin + 1
    End If

    GetImageAlphaBorder = rect
End Function

' Colors and palette

Private Function SameColor(ByRef a As ColorRGBA, ByRef b As ColorRGBA) As Boolean
    SameColor = (ColorToByte(a.r) = ColorToByte(b.r)) And _
                (ColorToByte(a.g) = ColorToByte(b.g)) And _
                (ColorToByte(a.b) = ColorToByte(b.b)) And _
                (ColorToByte(a.a) = ColorToByte(b.a))
End Function

Public Function LoadImageColors(ByRef image As Image) As ColorRGBA()
    Dim out() As ColorRGBA
    Dim x As Long, y As Long
    Dim i As Long

    If IsImageReady(image) = False Then
        LoadImageColors = out
        Exit Function
    End If

    ReDim out(0 To image.Width * image.Height - 1)
    i = 0
    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            out(i) = GetImageColor(image, x, y)
            i = i + 1
        Next x
    Next y

    LoadImageColors = out
End Function

Public Function LoadImagePalette(ByRef image As Image, ByVal maxPaletteSize As Long, ByRef colorCount As Long) As ColorRGBA()
    Dim palette() As ColorRGBA
    Dim counts() As Long
    Dim out() As ColorRGBA
    Dim x As Long, y As Long, i As Long, j As Long, n As Long
    Dim c As ColorRGBA
    Dim tc As ColorRGBA
    Dim tn As Long
    Dim found As Boolean

    colorCount = 0
    If IsImageReady(image) = False Then
        LoadImagePalette = out
        Exit Function
    End If
    If maxPaletteSize <= 0 Then maxPaletteSize = 256

    ReDim palette(0 To maxPaletteSize - 1)
    ReDim counts(0 To maxPaletteSize - 1)
    n = 0

    For y = 0 To image.Height - 1
        For x = 0 To image.Width - 1
            c = GetImageColor(image, x, y)
            found = False
            For i = 0 To n - 1
                If SameColor(palette(i), c) Then
                    counts(i) = counts(i) + 1
                    found = True
                    Exit For
                End If
            Next i
            If found = False And n < maxPaletteSize Then
                palette(n) = c
                counts(n) = 1
                n = n + 1
            End If
        Next x
    Next y

    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If counts(j) < counts(j + 1) Then
                tc = palette(j): palette(j) = palette(j + 1): palette(j + 1) = tc
                tn = counts(j): counts(j) = counts(j + 1): counts(j + 1) = tn
            End If
        Next j
    Next i

    If n = 0 Then
        LoadImagePalette = out
        Exit Function
    End If

    ReDim out(0 To n - 1)
    For i = 0 To n - 1
        out(i) = palette(i)
    Next i
    colorCount = n
    LoadImagePalette = out
End Function

' Image generation

Private Function MakeGray(ByVal gray As Byte) As ColorRGBA
    Dim c As ColorRGBA
    Color c, CSng(gray) / 255#, CSng(gray) / 255#, CSng(gray) / 255#, 1#
    MakeGray = c
End Function

Public Function GenImageColor(ByVal width As Long, ByVal height As Long, ByRef color As ColorRGBA) As Image
    Dim img As Image
    img = NewImageLike(width, height, PIXELFORMAT_UNCOMPRESSED_R8G8B8A8)
    ImageClearBackground img, color
    GenImageColor = img
End Function

Public Function GenImageGradientLinear(ByVal width As Long, ByVal height As Long, ByVal direction As Long, ByRef startColor As ColorRGBA, ByRef endColor As ColorRGBA) As Image
    Dim img As Image
    Dim x As Long, y As Long
    Dim factor As Single
    Dim rd As Single, gd As Single, bd As Single, ad As Single
    Dim col As ColorRGBA

    img = NewImageLike(width, height, PIXELFORMAT_UNCOMPRESSED_R8G8B8A8)
    rd = endColor.r - startColor.r
    gd = endColor.g - startColor.g
    bd = endColor.b - startColor.b
    ad = endColor.a - startColor.a

    For y = 0 To height - 1
        For x = 0 To width - 1
            Select Case direction
                Case 0
                    If height > 1 Then factor = CSng(y) / CSng(height - 1) Else factor = 0!
                Case 1
                    If width > 1 Then factor = CSng(x) / CSng(width - 1) Else factor = 0!
                Case Else
                    If (width + height) > 2 Then factor = CSng(x + y) / CSng(width + height - 2) Else factor = 0!
            End Select
            col.r = Clamp01(startColor.r + factor * rd)
            col.g = Clamp01(startColor.g + factor * gd)
            col.b = Clamp01(startColor.b + factor * bd)
            col.a = Clamp01(startColor.a + factor * ad)
            ImageDrawPixel img, x, y, col
        Next x
    Next y

    GenImageGradientLinear = img
End Function

Public Function GenImageGradientRadial(ByVal width As Long, ByVal height As Long, ByVal density As Single, ByRef inner As ColorRGBA, ByRef outer As ColorRGBA) As Image
    Dim img As Image
    Dim x As Long, y As Long
    Dim cx As Single, cy As Single
    Dim radius As Single
    Dim dx As Single, dy As Single
    Dim dist As Single
    Dim factor As Single
    Dim rd As Single, gd As Single, bd As Single, ad As Single
    Dim col As ColorRGBA

    img = NewImageLike(width, height, PIXELFORMAT_UNCOMPRESSED_R8G8B8A8)
    cx = CSng(width) / 2!
    cy = CSng(height) / 2!
    radius = CSng(IIf(width > height, width, height)) / 2!
    If radius <= 0! Then radius = 1!

    rd = outer.r - inner.r
    gd = outer.g - inner.g
    bd = outer.b - inner.b
    ad = outer.a - inner.a

    For y = 0 To height - 1
        For x = 0 To width - 1
            dx = CSng(x) - cx
            dy = CSng(y) - cy
            dist = Sqr(dx * dx + dy * dy)
            factor = dist / radius
            If density <> 1! And factor > 0! Then factor = CSng(factor ^ density)
            If factor > 1! Then factor = 1!
            If factor < 0! Then factor = 0!
            col.r = Clamp01(inner.r + factor * rd)
            col.g = Clamp01(inner.g + factor * gd)
            col.b = Clamp01(inner.b + factor * bd)
            col.a = Clamp01(inner.a + factor * ad)
            ImageDrawPixel img, x, y, col
        Next x
    Next y

    GenImageGradientRadial = img
End Function

Public Function GenImageGradientSquare(ByVal width As Long, ByVal height As Long, ByVal density As Single, ByRef inner As ColorRGBA, ByRef outer As ColorRGBA) As Image
    Dim img As Image
    Dim x As Long, y As Long
    Dim cx As Single, cy As Single
    Dim radius As Single
    Dim dx As Single, dy As Single
    Dim dist As Single
    Dim factor As Single
    Dim rd As Single, gd As Single, bd As Single, ad As Single
    Dim col As ColorRGBA

    img = NewImageLike(width, height, PIXELFORMAT_UNCOMPRESSED_R8G8B8A8)
    cx = CSng(width) / 2!
    cy = CSng(height) / 2!
    radius = CSng(IIf(width > height, width, height)) / 2!
    If radius <= 0! Then radius = 1!

    rd = outer.r - inner.r
    gd = outer.g - inner.g
    bd = outer.b - inner.b
    ad = outer.a - inner.a

    For y = 0 To height - 1
        For x = 0 To width - 1
            dx = Abs(CSng(x) - cx)
            dy = Abs(CSng(y) - cy)
            If dx > dy Then dist = dx Else dist = dy
            factor = dist / radius
            If density <> 1! And factor > 0! Then factor = CSng(factor ^ density)
            If factor > 1! Then factor = 1!
            If factor < 0! Then factor = 0!
            col.r = Clamp01(inner.r + factor * rd)
            col.g = Clamp01(inner.g + factor * gd)
            col.b = Clamp01(inner.b + factor * bd)
            col.a = Clamp01(inner.a + factor * ad)
            ImageDrawPixel img, x, y, col
        Next x
    Next y

    GenImageGradientSquare = img
End Function

Public Function GenImageChecked(ByVal width As Long, ByVal height As Long, ByVal checksX As Long, ByVal checksY As Long, ByRef col1 As ColorRGBA, ByRef col2 As ColorRGBA) As Image
    Dim img As Image
    Dim x As Long, y As Long

    img = NewImageLike(width, height, PIXELFORMAT_UNCOMPRESSED_R8G8B8A8)
    If checksX <= 0 Then checksX = 1
    If checksY <= 0 Then checksY = 1

    For y = 0 To height - 1
        For x = 0 To width - 1
            If ((x \ checksX) Mod 2) = ((y \ checksY) Mod 2) Then
                ImageDrawPixel img, x, y, col1
            Else
                ImageDrawPixel img, x, y, col2
            End If
        Next x
    Next y

    GenImageChecked = img
End Function

Public Function GenImageWhiteNoise(ByVal width As Long, ByVal height As Long, ByVal factor As Single) As Image
    Dim img As Image
    Dim x As Long, y As Long
    Dim gray As Byte

    img = NewImageLike(width, height, PIXELFORMAT_UNCOMPRESSED_R8G8B8A8)
    For y = 0 To height - 1
        For x = 0 To width - 1
            gray = ClampByteD(GetRandomValue(0, 255) * factor)
            ImageDrawPixel img, x, y, MakeGray(gray)
        Next x
    Next y

    GenImageWhiteNoise = img
End Function

' Noise (Perlin / cellular)

Private Sub InitPerlin()
    Dim i As Long, j As Long, t As Long
    Dim s As Long
    Dim p(0 To 255) As Long

    If mPermInit Then Exit Sub

    For i = 0 To 255
        p(i) = i
    Next i

    s = 12345
    For i = 255 To 1 Step -1
        s = (s * 1103 + 12345) Mod 32749
        j = s Mod (i + 1)
        t = p(i): p(i) = p(j): p(j) = t
    Next i

    For i = 0 To 255
        mPerm(i) = p(i)
        mPerm(i + 256) = p(i)
    Next i
    mPermInit = True
End Sub

Private Function PerlinFade(ByVal t As Single) As Single
    PerlinFade = t * t * t * (t * (t * 6! - 15!) + 10!)
End Function

Private Function PerlinGrad(ByVal hash As Long, ByVal x As Single, ByVal y As Single) As Single
    Dim h As Long
    Dim u As Single, v As Single

    h = hash And 7
    If h < 4 Then
        u = x: v = y
    Else
        u = y: v = x
    End If
    If (h And 1) <> 0 Then u = -u
    If (h And 2) <> 0 Then v = -v
    PerlinGrad = u + v
End Function

Private Function Perlin2D(ByVal x As Single, ByVal y As Single) As Single
    Dim xi As Long, yi As Long
    Dim xfl As Long, yfl As Long
    Dim xf As Single, yf As Single
    Dim u As Single, v As Single
    Dim aa As Long, ab As Long, ba As Long, bb As Long
    Dim x1 As Single, x2 As Single

    xfl = CLng(Int(CDbl(x)))
    yfl = CLng(Int(CDbl(y)))
    xi = xfl And 255
    yi = yfl And 255
    xf = x - xfl
    yf = y - yfl

    u = PerlinFade(xf)
    v = PerlinFade(yf)

    aa = mPerm(mPerm(xi) + yi)
    ab = mPerm(mPerm(xi) + yi + 1)
    ba = mPerm(mPerm(xi + 1) + yi)
    bb = mPerm(mPerm(xi + 1) + yi + 1)

    x1 = LerpS(PerlinGrad(aa, xf, yf), PerlinGrad(ba, xf - 1!, yf), u)
    x2 = LerpS(PerlinGrad(ab, xf, yf - 1!), PerlinGrad(bb, xf - 1!, yf - 1!), u)
    Perlin2D = LerpS(x1, x2, v)
End Function

Public Function GenImagePerlinNoise(ByVal width As Long, ByVal height As Long, ByVal offsetX As Long, ByVal offsetY As Long, ByVal scaleFactor As Single) As Image
    Dim img As Image
    Dim x As Long, y As Long
    Dim nx As Single, ny As Single
    Dim value As Single
    Dim gray As Byte
    Dim fw As Single, fh As Single

    InitPerlin
    img = NewImageLike(width, height, PIXELFORMAT_UNCOMPRESSED_R8G8B8A8)

    fw = CSng(img.Width)
    fh = CSng(img.Height)

    For y = 0 To img.Height - 1
        For x = 0 To img.Width - 1
            nx = CSng(x + offsetX) * scaleFactor / fw
            ny = CSng(y + offsetY) * scaleFactor / fh
            value = Perlin2D(nx, ny)
            gray = ClampByteD((CDbl(value) + 1#) * 127.5)
            ImageDrawPixel img, x, y, MakeGray(gray)
        Next x
    Next y

    GenImagePerlinNoise = img
End Function

Public Function GenImageCellular(ByVal width As Long, ByVal height As Long, ByVal tileSize As Long) As Image
    Dim img As Image
    Dim x As Long, y As Long, i As Long
    Dim seedsPerRow As Long, seedsPerCol As Long, seedCount As Long
    Dim sx() As Single, sy() As Single
    Dim minDist As Single
    Dim fx As Single, fy As Single
    Dim dx As Single, dy As Single
    Dim dist As Single
    Dim value As Byte
    Dim maxDist As Single

    If tileSize <= 0 Then tileSize = 16
    img = NewImageLike(width, height, PIXELFORMAT_UNCOMPRESSED_R8G8B8A8)

    seedsPerRow = width \ tileSize
    If seedsPerRow < 1 Then seedsPerRow = 1
    seedsPerCol = height \ tileSize
    If seedsPerCol < 1 Then seedsPerCol = 1
    seedCount = seedsPerRow * seedsPerCol

    ReDim sx(0 To seedCount - 1)
    ReDim sy(0 To seedCount - 1)

    For i = 0 To seedCount - 1
        sx(i) = CSng((i Mod seedsPerRow) * tileSize + GetRandomValue(0, tileSize - 1))
        sy(i) = CSng((i \ seedsPerRow) * tileSize + GetRandomValue(0, tileSize - 1))
    Next i

    maxDist = CSng(tileSize)
    For y = 0 To height - 1
        For x = 0 To width - 1
            fx = CSng(x)
            fy = CSng(y)
            minDist = 65536!
            For i = 0 To seedCount - 1
                dx = fx - sx(i)
                dy = fy - sy(i)
                dist = dx * dx + dy * dy
                If dist < minDist Then minDist = dist
            Next i
            value = ClampByteD(255# * (1# - Sqr(minDist) / CDbl(maxDist)))
            ImageDrawPixel img, x, y, MakeGray(value)
        Next x
    Next y

    GenImageCellular = img
End Function

Public Function LoadTextureFromImage(ByRef img As Image) As Texture2D
    Dim tex As Texture2D
    Dim textureId As Long
    Dim pixels() As Byte

    tex.Id = 0
    tex.Width = 0
    tex.Height = 0

    If IsImageReady(img) = False Then
        LoadTextureFromImage = tex
        Exit Function
    End If

    GetImageRGBA8 img, pixels

    glGenTextures 1, textureId
    glBindTexture GL_TEXTURE_2D, textureId

    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST

    glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, img.Width, img.Height, 0, GL_RGBA, GL_UNSIGNED_BYTE, ByVal VarPtr(pixels(0))

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

' Render target

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

    glGetIntegerv GL_FRAMEBUFFER_BINDING, prevFbo

    glGenFramebuffersEXT 1, rt.Id
    glBindFramebufferEXT GL_FRAMEBUFFER, rt.Id

    glGenTextures 1, rt.texture.Id
    glBindTexture GL_TEXTURE_2D, rt.texture.Id
    glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, Width, Height, 0, GL_RGBA, GL_UNSIGNED_BYTE, 0
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE
    glFramebufferTexture2DEXT GL_FRAMEBUFFER, GL_COLOR_ATTACHMENT0, GL_TEXTURE_2D, rt.texture.Id, 0

    glGenRenderbuffersEXT 1, rt.depth.Id
    glBindRenderbufferEXT GL_RENDERBUFFER, rt.depth.Id
    glRenderbufferStorageEXT GL_RENDERBUFFER, GL_DEPTH_COMPONENT24, Width, Height
    glFramebufferRenderbufferEXT GL_FRAMEBUFFER, GL_DEPTH_ATTACHMENT, GL_RENDERBUFFER, rt.depth.Id

    rt.texture.Width = Width
    rt.texture.Height = Height

    status = glCheckFramebufferStatusEXT(GL_FRAMEBUFFER)

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
    sepCol = GetImageColor(img, 0, 0)

    ' Read the font size (height) encoded in the first column: a marker pixel
    ' whose color differs from the separator color (magenta) at row = font size.
    Dim markerY As Long
    Dim yy As Long
    markerY = -1
    For yy = 0 To img.Height - 1
        Dim markerCol As ColorRGBA
        markerCol = GetImageColor(img, 0, yy)
        If Not (Abs(markerCol.r - sepCol.r) < 0.05 And _
                Abs(markerCol.g - sepCol.g) < 0.05 And _
                Abs(markerCol.b - sepCol.b) < 0.05) Then
            markerY = yy
            Exit For
        End If
    Next yy

    If markerY >= 0 Then
        fnt.Height = markerY
    Else
        fnt.Height = img.Height
    End If

    ReDim fnt.Glyphs(Len(CharList) - 1)

    Dim charIndex As Long
    Dim x As Long, startX As Long
    Dim w As Long
    Dim vbCol As ColorRGBA

    charIndex = 0
    x = 0

    Do While x < img.Width And charIndex < Len(CharList)
        Do While x < img.Width
            vbCol = GetImageColor(img, x, 0)
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
            vbCol = GetImageColor(img, x, 0)
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
    UpdateRenderSize
    mWindowResized = True
    rlViewport 0, 0, mRenderWidth, mRenderHeight
End Sub

' Screen (logical) is what the user draws with; render is the framebuffer and
' equals screen unless FLAG_WINDOW_HIGHDPI is set (then render = screen * DPI).
Private Sub UpdateRenderSize()
    Dim dpi As Vector2

    dpi = GetWindowScaleDPI()
    mDpiScaleX = dpi.x
    mDpiScaleY = dpi.y

    mScreenScaleX = 1!
    mScreenScaleY = 1!
    If (mConfigFlags And FLAG_WINDOW_HIGHDPI) <> 0 Then
        mScreenScaleX = mDpiScaleX
        mScreenScaleY = mDpiScaleY
    End If

    mRenderWidth = CLng(screenWidth * mScreenScaleX + 0.5)
    mRenderHeight = CLng(screenHeight * mScreenScaleY + 0.5)
    If mRenderWidth < 1 Then mRenderWidth = 1
    If mRenderHeight < 1 Then mRenderHeight = 1
End Sub

Private Sub RefreshRenderSize()
    UpdateRenderSize
    rlViewport 0, 0, mRenderWidth, mRenderHeight
End Sub

Public Sub BeginDrawing()
    DebugEndUpdate
    DebugBeginDraw

    rlMatrixMode RL_PROJECTION
    rlLoadIdentity
    rlOrtho 0, mRenderWidth, mRenderHeight, 0, -1, 1

    rlMatrixMode RL_MODELVIEW
    rlLoadIdentity
    If mScreenScaleX <> 1! Or mScreenScaleY <> 1! Then rlScalef mScreenScaleX, mScreenScaleY, 1!
End Sub

' Render target mode
Public Sub BeginTextureMode(ByRef target As RenderTexture2D)
    Dim prevFbo As Long
    Dim vp(0 To 3) As Long

    If mRenderTop >= RENDER_STACK_MAX Then Exit Sub
    If target.Id = 0 Then Exit Sub

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
    Dim now As Currency
    Dim dt As Single
    Dim targetFrameTime As Single

    DebugEndDraw

    QueryPerformanceCounter now
    If mQpcFreq > 0 Then dt = CSng(CDbl(now - mFrameMark) / CDbl(mQpcFreq))

    ' Frame limiter: busy-wait until the frame reaches the target period,
    ' re-measuring from mFrameMark so the wall clock (and GetFrameTime) is the
    ' full frame time. Only when vsync is off (with vsync the swap paces us).
    If mTargetFPS > 0 And IsVSyncEnabled() = False Then
        targetFrameTime = 1# / CSng(mTargetFPS)
        If dt < targetFrameTime Then
            Do
                QueryPerformanceCounter now
                If mQpcFreq > 0 Then dt = CSng(CDbl(now - mFrameMark) / CDbl(mQpcFreq))
                DoEvents
            Loop While dt < targetFrameTime
        End If
    End If

    mFrameMark = now
    If dt > 0.25 Then dt = 0.25

    DebugGpuBegin
    rlDrawRenderBatchActive
    DebugGpuEnd

    SwapScreenBuffer

    DebugEndFrame dt
    DebugBeginFrame

    mFrameTime = dt
    mTime = mTime + dt

    mFpsTimer = mFpsTimer + dt
    mFpsFrames = mFpsFrames + 1
    If mFpsTimer >= 1# Then
        mFps = mFpsFrames
        mFpsFrames = 0
        mFpsTimer = mFpsTimer - 1#
    End If

    PollInputEvents
    DebugBeginUpdate
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
    fHeight = CSng(font.texture.Height)

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

Public Function MeasureText(ByRef font As ImageFont, ByVal text As String) As Vector2
    Dim result As Vector2
    Dim i As Long
    Dim totalWidth As Single

    totalWidth = 0!

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

        totalWidth = totalWidth + g.Width
    Next i

    result.x = totalWidth
    result.y = CSng(font.Height)
    MeasureText = result
End Function

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
