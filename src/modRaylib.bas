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

Public Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hDC As Long, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function SetPixelFormat Lib "gdi32" (ByVal hDC As Long, ByVal PixelFormat As Long, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function SwapBuffers Lib "gdi32" (ByVal hDC As Long) As Long

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

' GDI

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

' ============================================================================
' OpenGL 1.1 - bindings estaticas (opengl32.dll)
'
' Convencoes impostas pelo VB6:
'   GLenum / GLbitfield / GLint / GLsizei / GLuint -> Long
'   GLboolean                                       -> Byte
'   GLfloat / GLclampf                              -> Single
'   GLdouble / GLclampd                             -> Double
'   GLvoid*                                         -> ByVal Long
'                                                      chamar: ByVal VarPtr(dados(0))
'   Vetores (GLfloat*, GLint*, GLdouble*...)         -> ByRef <tipo>
'                                                      chamar: dados(0)
'
' Ordem das secoes:
'   1. WGL / contexto        6. Texturas           11. Consultas de estado
'   2. Estado geral          7. Pixel ops          12. Pipeline obsoleto (GL 1.1)
'   3. Primitivas            8. Iluminacao/fog
'   4. Matrizes              9. Display lists
'   5. Rasterizacao         10. Vertex arrays
'
' Extensoes (FBO, shaders, VBO, WGL_ARB...) nao entram aqui.
' ============================================================================

' --- 1. WGL / contexto ------------------------------------------------------

Public Declare Function wglCreateContext Lib "opengl32" (ByVal hDC As Long) As Long
Public Declare Function wglMakeCurrent Lib "opengl32" (ByVal hDC As Long, ByVal hRC As Long) As Long
Public Declare Function wglDeleteContext Lib "opengl32" (ByVal hRC As Long) As Long
Public Declare Function wglGetCurrentContext Lib "opengl32" () As Long
Public Declare Function wglGetCurrentDC Lib "opengl32" () As Long
Public Declare Function wglShareLists Lib "opengl32" (ByVal hrc1 As Long, ByVal hrc2 As Long) As Long
' Reservado para a etapa de extensoes: carrega glGenFramebuffers, glCreateShader, etc.
Public Declare Function wglGetProcAddress Lib "opengl32" (ByVal lpszProc As String) As Long

' --- 2. Estado geral --------------------------------------------------------

Public Declare Sub glEnable Lib "opengl32" (ByVal cap As Long)
Public Declare Sub glDisable Lib "opengl32" (ByVal cap As Long)
Public Declare Function glIsEnabled Lib "opengl32" (ByVal cap As Long) As Byte
Public Declare Function glGetError Lib "opengl32" () As Long
Public Declare Function glGetString Lib "opengl32" (ByVal pname As Long) As Long
Public Declare Sub glHint Lib "opengl32" (ByVal Target As Long, ByVal mode As Long)
Public Declare Sub glFlush Lib "opengl32" ()
Public Declare Sub glFinish Lib "opengl32" ()
Public Declare Sub glPushAttrib Lib "opengl32" (ByVal mask As Long)
Public Declare Sub glPopAttrib Lib "opengl32" ()
Public Declare Sub glPushClientAttrib Lib "opengl32" (ByVal mask As Long)
Public Declare Sub glPopClientAttrib Lib "opengl32" ()

' --- 3. Primitivas ----------------------------------------------------------

Public Declare Sub glBegin Lib "opengl32" (ByVal mode As Long)
Public Declare Sub glEnd Lib "opengl32" ()

Public Declare Sub glColor3f Lib "opengl32" (ByVal Red As Single, ByVal Green As Single, ByVal Blue As Single)
Public Declare Sub glColor4f Lib "opengl32" (ByVal Red As Single, ByVal Green As Single, ByVal Blue As Single, ByVal alpha As Single)
Public Declare Sub glColor3d Lib "opengl32" (ByVal Red As Double, ByVal Green As Double, ByVal Blue As Double)
Public Declare Sub glColor4d Lib "opengl32" (ByVal Red As Double, ByVal Green As Double, ByVal Blue As Double, ByVal alpha As Double)
Public Declare Sub glColor3ub Lib "opengl32" (ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
Public Declare Sub glColor4ub Lib "opengl32" (ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, ByVal alpha As Byte)
Public Declare Sub glColor4fv Lib "opengl32" (ByRef v As Single)
Public Declare Sub glColor4ubv Lib "opengl32" (ByRef v As Byte)

Public Declare Sub glNormal3f Lib "opengl32" (ByVal nx As Single, ByVal ny As Single, ByVal nz As Single)
Public Declare Sub glNormal3d Lib "opengl32" (ByVal nx As Double, ByVal ny As Double, ByVal nz As Double)
Public Declare Sub glNormal3fv Lib "opengl32" (ByRef v As Single)

Public Declare Sub glTexCoord1f Lib "opengl32" (ByVal s As Single)
Public Declare Sub glTexCoord2f Lib "opengl32" (ByVal s As Single, ByVal t As Single)
Public Declare Sub glTexCoord3f Lib "opengl32" (ByVal s As Single, ByVal t As Single, ByVal r As Single)
Public Declare Sub glTexCoord4f Lib "opengl32" (ByVal s As Single, ByVal t As Single, ByVal r As Single, ByVal q As Single)
Public Declare Sub glTexCoord2d Lib "opengl32" (ByVal s As Double, ByVal t As Double)
Public Declare Sub glTexCoord3d Lib "opengl32" (ByVal s As Double, ByVal t As Double, ByVal r As Double)
Public Declare Sub glTexCoord4d Lib "opengl32" (ByVal s As Double, ByVal t As Double, ByVal r As Double, ByVal q As Double)
Public Declare Sub glTexCoord2fv Lib "opengl32" (ByRef v As Single)
Public Declare Sub glTexCoord3fv Lib "opengl32" (ByRef v As Single)

Public Declare Sub glVertex2f Lib "opengl32" (ByVal x As Single, ByVal y As Single)
Public Declare Sub glVertex3f Lib "opengl32" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
Public Declare Sub glVertex4f Lib "opengl32" (ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal w As Single)
Public Declare Sub glVertex2i Lib "opengl32" (ByVal x As Long, ByVal y As Long)
Public Declare Sub glVertex3i Lib "opengl32" (ByVal x As Long, ByVal y As Long, ByVal z As Long)
Public Declare Sub glVertex2d Lib "opengl32" (ByVal x As Double, ByVal y As Double)
Public Declare Sub glVertex3d Lib "opengl32" (ByVal x As Double, ByVal y As Double, ByVal z As Double)
Public Declare Sub glVertex4d Lib "opengl32" (ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal w As Double)
Public Declare Sub glVertex2fv Lib "opengl32" (ByRef v As Single)
Public Declare Sub glVertex3fv Lib "opengl32" (ByRef v As Single)

Public Declare Sub glRectf Lib "opengl32" (ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
Public Declare Sub glRecti Lib "opengl32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
Public Declare Sub glRects Lib "opengl32" (ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
Public Declare Sub glRectd Lib "opengl32" (ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)

' --- 4. Matrizes / transformacoes -------------------------------------------

Public Declare Sub glMatrixMode Lib "opengl32" (ByVal mode As Long)
Public Declare Sub glLoadIdentity Lib "opengl32" ()
Public Declare Sub glLoadMatrixf Lib "opengl32" (ByRef m As Single)
Public Declare Sub glLoadMatrixd Lib "opengl32" (ByRef m As Double)
Public Declare Sub glMultMatrixf Lib "opengl32" (ByRef m As Single)
Public Declare Sub glMultMatrixd Lib "opengl32" (ByRef m As Double)
Public Declare Sub glOrtho Lib "opengl32" (ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double, ByVal zNear As Double, ByVal zFar As Double)
Public Declare Sub glFrustum Lib "opengl32" (ByVal Left As Double, ByVal Right As Double, ByVal Bottom As Double, ByVal Top As Double, ByVal zNear As Double, ByVal zFar As Double)
Public Declare Sub glTranslatef Lib "opengl32" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
Public Declare Sub glTranslated Lib "opengl32" (ByVal x As Double, ByVal y As Double, ByVal z As Double)
Public Declare Sub glRotatef Lib "opengl32" (ByVal angle As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single)
Public Declare Sub glRotated Lib "opengl32" (ByVal angle As Double, ByVal x As Double, ByVal y As Double, ByVal z As Double)
Public Declare Sub glScalef Lib "opengl32" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
Public Declare Sub glScaled Lib "opengl32" (ByVal x As Double, ByVal y As Double, ByVal z As Double)
Public Declare Sub glPushMatrix Lib "opengl32" ()
Public Declare Sub glPopMatrix Lib "opengl32" ()

' --- 5. Rasterizacao / estado fixo ------------------------------------------

Public Declare Sub glViewport Lib "opengl32" (ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long)
Public Declare Sub glScissor Lib "opengl32" (ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long)
Public Declare Sub glClear Lib "opengl32" (ByVal mask As Long)
Public Declare Sub glClearColor Lib "opengl32" (ByVal Red As Single, ByVal Green As Single, ByVal Blue As Single, ByVal alpha As Single)
Public Declare Sub glClearDepth Lib "opengl32" (ByVal depth As Double)
Public Declare Sub glClearStencil Lib "opengl32" (ByVal s As Long)
Public Declare Sub glClearIndex Lib "opengl32" (ByVal c As Single)
Public Declare Sub glDrawBuffer Lib "opengl32" (ByVal mode As Long)
Public Declare Sub glReadBuffer Lib "opengl32" (ByVal mode As Long)

Public Declare Sub glColorMask Lib "opengl32" (ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, ByVal alpha As Byte)
Public Declare Sub glDepthFunc Lib "opengl32" (ByVal func As Long)
Public Declare Sub glDepthMask Lib "opengl32" (ByVal flag As Byte)
Public Declare Sub glDepthRange Lib "opengl32" (ByVal zNear As Double, ByVal zFar As Double)
Public Declare Sub glStencilFunc Lib "opengl32" (ByVal func As Long, ByVal refValue As Long, ByVal mask As Long)
Public Declare Sub glStencilMask Lib "opengl32" (ByVal mask As Long)
Public Declare Sub glStencilOp Lib "opengl32" (ByVal sfail As Long, ByVal dpfail As Long, ByVal dppass As Long)
Public Declare Sub glAlphaFunc Lib "opengl32" (ByVal func As Long, ByVal refValue As Single)
Public Declare Sub glBlendFunc Lib "opengl32" (ByVal sfactor As Long, ByVal dfactor As Long)
Public Declare Sub glLogicOp Lib "opengl32" (ByVal opcode As Long)
Public Declare Sub glCullFace Lib "opengl32" (ByVal mode As Long)
Public Declare Sub glFrontFace Lib "opengl32" (ByVal mode As Long)
Public Declare Sub glShadeModel Lib "opengl32" (ByVal mode As Long)
Public Declare Sub glPolygonMode Lib "opengl32" (ByVal face As Long, ByVal mode As Long)
Public Declare Sub glPolygonOffset Lib "opengl32" (ByVal factor As Single, ByVal units As Single)
Public Declare Sub glPolygonStipple Lib "opengl32" (ByRef mask As Byte)
Public Declare Sub glGetPolygonStipple Lib "opengl32" (ByRef mask As Byte)
Public Declare Sub glLineWidth Lib "opengl32" (ByVal Width As Single)
Public Declare Sub glLineStipple Lib "opengl32" (ByVal factor As Long, ByVal pattern As Integer)
Public Declare Sub glPointSize Lib "opengl32" (ByVal size As Single)

' --- 6. Vertex arrays / client state ----------------------------------------

Public Declare Sub glEnableClientState Lib "opengl32" (ByVal cap As Long)
Public Declare Sub glDisableClientState Lib "opengl32" (ByVal cap As Long)
Public Declare Sub glVertexPointer Lib "opengl32" (ByVal size As Long, ByVal ty As Long, ByVal stride As Long, ByVal pData As Long)
Public Declare Sub glNormalPointer Lib "opengl32" (ByVal ty As Long, ByVal stride As Long, ByVal pData As Long)
Public Declare Sub glColorPointer Lib "opengl32" (ByVal size As Long, ByVal ty As Long, ByVal stride As Long, ByVal pData As Long)
Public Declare Sub glIndexPointer Lib "opengl32" (ByVal ty As Long, ByVal stride As Long, ByVal pData As Long)
Public Declare Sub glTexCoordPointer Lib "opengl32" (ByVal size As Long, ByVal ty As Long, ByVal stride As Long, ByVal pData As Long)
Public Declare Sub glEdgeFlagPointer Lib "opengl32" (ByVal stride As Long, ByVal pData As Long)
Public Declare Sub glInterleavedArrays Lib "opengl32" (ByVal Format As Long, ByVal stride As Long, ByVal pData As Long)
Public Declare Sub glArrayElement Lib "opengl32" (ByVal i As Long)
Public Declare Sub glDrawArrays Lib "opengl32" (ByVal mode As Long, ByVal first As Long, ByVal count As Long)
Public Declare Sub glDrawElements Lib "opengl32" (ByVal mode As Long, ByVal count As Long, ByVal ty As Long, ByVal indices As Long)
Public Declare Sub glGetPointerv Lib "opengl32" (ByVal pname As Long, ByRef params As Long)

' --- 7. Texturas ------------------------------------------------------------

Public Declare Sub glGenTextures Lib "opengl32" (ByVal n As Long, ByRef textures As Long)
Public Declare Sub glDeleteTextures Lib "opengl32" (ByVal n As Long, ByRef textures As Long)
Public Declare Sub glBindTexture Lib "opengl32" (ByVal Target As Long, ByVal Texture As Long)
Public Declare Function glIsTexture Lib "opengl32" (ByVal Texture As Long) As Byte
Public Declare Function glAreTexturesResident Lib "opengl32" (ByVal n As Long, ByRef textures As Long, ByRef residences As Byte) As Byte
Public Declare Sub glPrioritizeTextures Lib "opengl32" (ByVal n As Long, ByRef textures As Long, ByRef priorities As Single)

Public Declare Sub glTexImage1D Lib "opengl32" (ByVal Target As Long, ByVal level As Long, ByVal internalformat As Long, ByVal Width As Long, ByVal border As Long, ByVal Format As Long, ByVal ty As Long, ByVal pixels As Long)
Public Declare Sub glTexImage2D Lib "opengl32" (ByVal Target As Long, ByVal level As Long, ByVal internalformat As Long, ByVal Width As Long, ByVal Height As Long, ByVal border As Long, ByVal Format As Long, ByVal ty As Long, ByVal pixels As Long)
Public Declare Sub glTexSubImage1D Lib "opengl32" (ByVal Target As Long, ByVal level As Long, ByVal xoffset As Long, ByVal Width As Long, ByVal Format As Long, ByVal ty As Long, ByVal pixels As Long)
Public Declare Sub glTexSubImage2D Lib "opengl32" (ByVal Target As Long, ByVal level As Long, ByVal xoffset As Long, ByVal yoffset As Long, ByVal Width As Long, ByVal Height As Long, ByVal Format As Long, ByVal ty As Long, ByVal pixels As Long)
Public Declare Sub glCopyTexImage1D Lib "opengl32" (ByVal Target As Long, ByVal level As Long, ByVal internalformat As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal border As Long)
Public Declare Sub glCopyTexImage2D Lib "opengl32" (ByVal Target As Long, ByVal level As Long, ByVal internalformat As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal border As Long)
Public Declare Sub glCopyTexSubImage1D Lib "opengl32" (ByVal Target As Long, ByVal level As Long, ByVal xoffset As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long)
Public Declare Sub glCopyTexSubImage2D Lib "opengl32" (ByVal Target As Long, ByVal level As Long, ByVal xoffset As Long, ByVal yoffset As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long)
Public Declare Sub glGetTexImage Lib "opengl32" (ByVal Target As Long, ByVal level As Long, ByVal Format As Long, ByVal ty As Long, ByVal pixels As Long)

Public Declare Sub glTexParameterf Lib "opengl32" (ByVal Target As Long, ByVal pname As Long, ByVal param As Single)
Public Declare Sub glTexParameteri Lib "opengl32" (ByVal Target As Long, ByVal pname As Long, ByVal param As Long)
Public Declare Sub glTexParameterfv Lib "opengl32" (ByVal Target As Long, ByVal pname As Long, ByRef params As Single)
Public Declare Sub glTexParameteriv Lib "opengl32" (ByVal Target As Long, ByVal pname As Long, ByRef params As Long)
Public Declare Sub glGetTexParameterfv Lib "opengl32" (ByVal Target As Long, ByVal pname As Long, ByRef params As Single)
Public Declare Sub glGetTexParameteriv Lib "opengl32" (ByVal Target As Long, ByVal pname As Long, ByRef params As Long)
Public Declare Sub glGetTexLevelParameterfv Lib "opengl32" (ByVal Target As Long, ByVal level As Long, ByVal pname As Long, ByRef params As Single)
Public Declare Sub glGetTexLevelParameteriv Lib "opengl32" (ByVal Target As Long, ByVal level As Long, ByVal pname As Long, ByRef params As Long)

Public Declare Sub glTexEnvf Lib "opengl32" (ByVal Target As Long, ByVal pname As Long, ByVal param As Single)
Public Declare Sub glTexEnvi Lib "opengl32" (ByVal Target As Long, ByVal pname As Long, ByVal param As Long)
Public Declare Sub glTexEnvfv Lib "opengl32" (ByVal Target As Long, ByVal pname As Long, ByRef params As Single)
Public Declare Sub glTexEnviv Lib "opengl32" (ByVal Target As Long, ByVal pname As Long, ByRef params As Long)
Public Declare Sub glGetTexEnvfv Lib "opengl32" (ByVal Target As Long, ByVal pname As Long, ByRef params As Single)
Public Declare Sub glGetTexEnviv Lib "opengl32" (ByVal Target As Long, ByVal pname As Long, ByRef params As Long)

Public Declare Sub glTexGend Lib "opengl32" (ByVal coord As Long, ByVal pname As Long, ByVal param As Double)
Public Declare Sub glTexGenf Lib "opengl32" (ByVal coord As Long, ByVal pname As Long, ByVal param As Single)
Public Declare Sub glTexGeni Lib "opengl32" (ByVal coord As Long, ByVal pname As Long, ByVal param As Long)
Public Declare Sub glTexGendv Lib "opengl32" (ByVal coord As Long, ByVal pname As Long, ByRef params As Double)
Public Declare Sub glTexGenfv Lib "opengl32" (ByVal coord As Long, ByVal pname As Long, ByRef params As Single)
Public Declare Sub glTexGeniv Lib "opengl32" (ByVal coord As Long, ByVal pname As Long, ByRef params As Long)
Public Declare Sub glGetTexGendv Lib "opengl32" (ByVal coord As Long, ByVal pname As Long, ByRef params As Double)
Public Declare Sub glGetTexGenfv Lib "opengl32" (ByVal coord As Long, ByVal pname As Long, ByRef params As Single)
Public Declare Sub glGetTexGeniv Lib "opengl32" (ByVal coord As Long, ByVal pname As Long, ByRef params As Long)

' --- 8. Pixel ops -----------------------------------------------------------

Public Declare Sub glPixelStoref Lib "opengl32" (ByVal pname As Long, ByVal param As Single)
Public Declare Sub glPixelStorei Lib "opengl32" (ByVal pname As Long, ByVal param As Long)
Public Declare Sub glPixelTransferf Lib "opengl32" (ByVal pname As Long, ByVal param As Single)
Public Declare Sub glPixelTransferi Lib "opengl32" (ByVal pname As Long, ByVal param As Long)
Public Declare Sub glPixelZoom Lib "opengl32" (ByVal xfactor As Single, ByVal yfactor As Single)
Public Declare Sub glPixelMapfv Lib "opengl32" (ByVal map As Long, ByVal mapsize As Long, ByRef values As Single)
Public Declare Sub glPixelMapuiv Lib "opengl32" (ByVal map As Long, ByVal mapsize As Long, ByRef values As Long)
Public Declare Sub glPixelMapusv Lib "opengl32" (ByVal map As Long, ByVal mapsize As Long, ByRef values As Integer)
Public Declare Sub glGetPixelMapfv Lib "opengl32" (ByVal map As Long, ByRef values As Single)
Public Declare Sub glGetPixelMapuiv Lib "opengl32" (ByVal map As Long, ByRef values As Long)
Public Declare Sub glGetPixelMapusv Lib "opengl32" (ByVal map As Long, ByRef values As Integer)

Public Declare Sub glReadPixels Lib "opengl32" (ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Format As Long, ByVal ty As Long, ByVal pixels As Long)
Public Declare Sub glDrawPixels Lib "opengl32" (ByVal Width As Long, ByVal Height As Long, ByVal Format As Long, ByVal ty As Long, ByVal pixels As Long)
Public Declare Sub glCopyPixels Lib "opengl32" (ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal ty As Long)
Public Declare Sub glBitmap Lib "opengl32" (ByVal Width As Long, ByVal Height As Long, ByVal xorig As Single, ByVal yorig As Single, ByVal xmove As Single, ByVal ymove As Single, ByVal bitmap As Long)

Public Declare Sub glRasterPos2d Lib "opengl32" (ByVal x As Double, ByVal y As Double)
Public Declare Sub glRasterPos2f Lib "opengl32" (ByVal x As Single, ByVal y As Single)
Public Declare Sub glRasterPos2i Lib "opengl32" (ByVal x As Long, ByVal y As Long)
Public Declare Sub glRasterPos2s Lib "opengl32" (ByVal x As Integer, ByVal y As Integer)
Public Declare Sub glRasterPos3d Lib "opengl32" (ByVal x As Double, ByVal y As Double, ByVal z As Double)
Public Declare Sub glRasterPos3f Lib "opengl32" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
Public Declare Sub glRasterPos3i Lib "opengl32" (ByVal x As Long, ByVal y As Long, ByVal z As Long)
Public Declare Sub glRasterPos3s Lib "opengl32" (ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
Public Declare Sub glRasterPos4d Lib "opengl32" (ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal w As Double)
Public Declare Sub glRasterPos4f Lib "opengl32" (ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal w As Single)
Public Declare Sub glRasterPos4i Lib "opengl32" (ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal w As Long)
Public Declare Sub glRasterPos4s Lib "opengl32" (ByVal x As Integer, ByVal y As Integer, ByVal z As Integer, ByVal w As Integer)

' --- 9. Iluminacao / material / fog / clip ----------------------------------

Public Declare Sub glLightModelf Lib "opengl32" (ByVal pname As Long, ByVal param As Single)
Public Declare Sub glLightModeli Lib "opengl32" (ByVal pname As Long, ByVal param As Long)
Public Declare Sub glLightModelfv Lib "opengl32" (ByVal pname As Long, ByRef params As Single)
Public Declare Sub glLightModeliv Lib "opengl32" (ByVal pname As Long, ByRef params As Long)
Public Declare Sub glLightf Lib "opengl32" (ByVal light As Long, ByVal pname As Long, ByVal param As Single)
Public Declare Sub glLighti Lib "opengl32" (ByVal light As Long, ByVal pname As Long, ByVal param As Long)
Public Declare Sub glLightfv Lib "opengl32" (ByVal light As Long, ByVal pname As Long, ByRef params As Single)
Public Declare Sub glLightiv Lib "opengl32" (ByVal light As Long, ByVal pname As Long, ByRef params As Long)
Public Declare Sub glGetLightfv Lib "opengl32" (ByVal light As Long, ByVal pname As Long, ByRef params As Single)
Public Declare Sub glGetLightiv Lib "opengl32" (ByVal light As Long, ByVal pname As Long, ByRef params As Long)
Public Declare Sub glMaterialf Lib "opengl32" (ByVal face As Long, ByVal pname As Long, ByVal param As Single)
Public Declare Sub glMateriali Lib "opengl32" (ByVal face As Long, ByVal pname As Long, ByVal param As Long)
Public Declare Sub glMaterialfv Lib "opengl32" (ByVal face As Long, ByVal pname As Long, ByRef params As Single)
Public Declare Sub glMaterialiv Lib "opengl32" (ByVal face As Long, ByVal pname As Long, ByRef params As Long)
Public Declare Sub glGetMaterialfv Lib "opengl32" (ByVal face As Long, ByVal pname As Long, ByRef params As Single)
Public Declare Sub glGetMaterialiv Lib "opengl32" (ByVal face As Long, ByVal pname As Long, ByRef params As Long)
Public Declare Sub glColorMaterial Lib "opengl32" (ByVal face As Long, ByVal mode As Long)
Public Declare Sub glFogf Lib "opengl32" (ByVal pname As Long, ByVal param As Single)
Public Declare Sub glFogi Lib "opengl32" (ByVal pname As Long, ByVal param As Long)
Public Declare Sub glFogfv Lib "opengl32" (ByVal pname As Long, ByRef params As Single)
Public Declare Sub glFogiv Lib "opengl32" (ByVal pname As Long, ByRef params As Long)
Public Declare Sub glClipPlane Lib "opengl32" (ByVal plane As Long, ByRef equation As Double)
Public Declare Sub glGetClipPlane Lib "opengl32" (ByVal plane As Long, ByRef equation As Double)

' --- 10. Display lists ------------------------------------------------------

Public Declare Sub glNewList Lib "opengl32" (ByVal list As Long, ByVal mode As Long)
Public Declare Sub glEndList Lib "opengl32" ()
Public Declare Sub glCallList Lib "opengl32" (ByVal list As Long)
Public Declare Sub glCallLists Lib "opengl32" (ByVal n As Long, ByVal ty As Long, ByVal lists As Long)
Public Declare Function glGenLists Lib "opengl32" (ByVal range As Long) As Long
Public Declare Sub glDeleteLists Lib "opengl32" (ByVal list As Long, ByVal range As Long)
Public Declare Sub glListBase Lib "opengl32" (ByVal base As Long)
Public Declare Function glIsList Lib "opengl32" (ByVal list As Long) As Byte

' --- 11. Consultas de estado ------------------------------------------------

Public Declare Sub glGetBooleanv Lib "opengl32" (ByVal pname As Long, ByRef params As Byte)
Public Declare Sub glGetDoublev Lib "opengl32" (ByVal pname As Long, ByRef params As Double)
Public Declare Sub glGetFloatv Lib "opengl32" (ByVal pname As Long, ByRef params As Single)
Public Declare Sub glGetIntegerv Lib "opengl32" (ByVal pname As Long, ByRef params As Long)

' --- 12. Pipeline obsoleto da GL 1.1 (mantido por completude) ---------------

' Acumulacao
Public Declare Sub glAccum Lib "opengl32" (ByVal op As Long, ByVal value As Single)
Public Declare Sub glClearAccum Lib "opengl32" (ByVal Red As Single, ByVal Green As Single, ByVal Blue As Single, ByVal alpha As Single)

' Modo indice de cor
Public Declare Sub glIndexd Lib "opengl32" (ByVal c As Double)
Public Declare Sub glIndexf Lib "opengl32" (ByVal c As Single)
Public Declare Sub glIndexi Lib "opengl32" (ByVal c As Long)
Public Declare Sub glIndexs Lib "opengl32" (ByVal c As Integer)
Public Declare Sub glIndexub Lib "opengl32" (ByVal c As Byte)
Public Declare Sub glIndexdv Lib "opengl32" (ByRef c As Double)
Public Declare Sub glIndexfv Lib "opengl32" (ByRef c As Single)
Public Declare Sub glIndexiv Lib "opengl32" (ByRef c As Long)
Public Declare Sub glIndexsv Lib "opengl32" (ByRef c As Integer)
Public Declare Sub glIndexubv Lib "opengl32" (ByRef c As Byte)
Public Declare Sub glIndexMask Lib "opengl32" (ByVal mask As Long)

' Evaluators
Public Declare Sub glEvalCoord1d Lib "opengl32" (ByVal u As Double)
Public Declare Sub glEvalCoord1f Lib "opengl32" (ByVal u As Single)
Public Declare Sub glEvalCoord1dv Lib "opengl32" (ByRef u As Double)
Public Declare Sub glEvalCoord1fv Lib "opengl32" (ByRef u As Single)
Public Declare Sub glEvalCoord2d Lib "opengl32" (ByVal u As Double, ByVal v As Double)
Public Declare Sub glEvalCoord2f Lib "opengl32" (ByVal u As Single, ByVal v As Single)
Public Declare Sub glEvalCoord2dv Lib "opengl32" (ByRef u As Double)
Public Declare Sub glEvalCoord2fv Lib "opengl32" (ByRef u As Single)
Public Declare Sub glMap1d Lib "opengl32" (ByVal Target As Long, ByVal u1 As Double, ByVal u2 As Double, ByVal stride As Long, ByVal order As Long, ByRef points As Double)
Public Declare Sub glMap1f Lib "opengl32" (ByVal Target As Long, ByVal u1 As Single, ByVal u2 As Single, ByVal stride As Long, ByVal order As Long, ByRef points As Single)
Public Declare Sub glMap2d Lib "opengl32" (ByVal Target As Long, ByVal u1 As Double, ByVal u2 As Double, ByVal ustride As Long, ByVal uorder As Long, ByVal v1 As Double, ByVal v2 As Double, ByVal vstride As Long, ByVal vorder As Long, ByRef points As Double)
Public Declare Sub glMap2f Lib "opengl32" (ByVal Target As Long, ByVal u1 As Single, ByVal u2 As Single, ByVal ustride As Long, ByVal uorder As Long, ByVal v1 As Single, ByVal v2 As Single, ByVal vstride As Long, ByVal vorder As Long, ByRef points As Single)
Public Declare Sub glMapGrid1d Lib "opengl32" (ByVal un As Long, ByVal u1 As Double, ByVal u2 As Double)
Public Declare Sub glMapGrid1f Lib "opengl32" (ByVal un As Long, ByVal u1 As Single, ByVal u2 As Single)
Public Declare Sub glMapGrid2d Lib "opengl32" (ByVal un As Long, ByVal u1 As Double, ByVal u2 As Double, ByVal vn As Long, ByVal v1 As Double, ByVal v2 As Double)
Public Declare Sub glMapGrid2f Lib "opengl32" (ByVal un As Long, ByVal u1 As Single, ByVal u2 As Single, ByVal vn As Long, ByVal v1 As Single, ByVal v2 As Single)
Public Declare Sub glEvalMesh1 Lib "opengl32" (ByVal mode As Long, ByVal i1 As Long, ByVal i2 As Long)
Public Declare Sub glEvalMesh2 Lib "opengl32" (ByVal mode As Long, ByVal i1 As Long, ByVal i2 As Long, ByVal j1 As Long, ByVal j2 As Long)
Public Declare Sub glEvalPoint1 Lib "opengl32" (ByVal i As Long)
Public Declare Sub glEvalPoint2 Lib "opengl32" (ByVal i As Long, ByVal j As Long)
Public Declare Sub glGetMapdv Lib "opengl32" (ByVal Target As Long, ByVal query As Long, ByRef v As Double)
Public Declare Sub glGetMapfv Lib "opengl32" (ByVal Target As Long, ByVal query As Long, ByRef v As Single)
Public Declare Sub glGetMapiv Lib "opengl32" (ByVal Target As Long, ByVal query As Long, ByRef v As Long)

' Selection / feedback
Public Declare Sub glFeedbackBuffer Lib "opengl32" (ByVal size As Long, ByVal ty As Long, ByRef buffer As Single)
Public Declare Sub glSelectBuffer Lib "opengl32" (ByVal size As Long, ByRef buffer As Long)
Public Declare Function glRenderMode Lib "opengl32" (ByVal mode As Long) As Long
Public Declare Sub glInitNames Lib "opengl32" ()
Public Declare Sub glPushName Lib "opengl32" (ByVal nameValue As Long)
Public Declare Sub glPopName Lib "opengl32" ()
Public Declare Sub glLoadName Lib "opengl32" (ByVal nameValue As Long)
Public Declare Sub glPassThrough Lib "opengl32" (ByVal token As Single)

' ============================================================================
' Constantes - agrupadas na mesma ordem das funcoes
' ============================================================================

' --- Booleanos / valores neutros --------------------------------------------
Public Const GL_FALSE As Long = &H0&
Public Const GL_TRUE As Long = &H1&
Public Const GL_NONE As Long = &H0&
Public Const GL_ZERO As Long = &H0&
Public Const GL_ONE As Long = &H1&
Public Const GL_NO_ERROR As Long = &H0&

' --- Buffers para glClear ----------------------------------------------
Public Const GL_COLOR_BUFFER_BIT As Long = &H4000&
Public Const GL_DEPTH_BUFFER_BIT As Long = &H100&
Public Const GL_STENCIL_BUFFER_BIT As Long = &H400&
Public Const GL_ACCUM_BUFFER_BIT As Long = &H200&

' --- Primitivas -------------------------------------------------------------
Public Const GL_POINTS As Long = &H0&
Public Const GL_LINES As Long = &H1&
Public Const GL_LINE_LOOP As Long = &H2&
Public Const GL_LINE_STRIP As Long = &H3&
Public Const GL_TRIANGLES As Long = &H4&
Public Const GL_TRIANGLE_STRIP As Long = &H5&
Public Const GL_TRIANGLE_FAN As Long = &H6&
Public Const GL_QUADS As Long = &H7&
Public Const GL_QUAD_STRIP As Long = &H8&
Public Const GL_POLYGON As Long = &H9&

' --- Erros (glGetError) -----------------------------------------------------
Public Const GL_INVALID_ENUM As Long = &H500&
Public Const GL_INVALID_VALUE As Long = &H501&
Public Const GL_INVALID_OPERATION As Long = &H502&
Public Const GL_STACK_OVERFLOW As Long = &H503&
Public Const GL_STACK_UNDERFLOW As Long = &H504&
Public Const GL_OUT_OF_MEMORY As Long = &H505&

' --- Tipos de dados ---------------------------------------------------------
Public Const GL_BYTE As Long = &H1400&
Public Const GL_UNSIGNED_BYTE As Long = &H1401&
Public Const GL_SHORT As Long = &H1402&
Public Const GL_UNSIGNED_SHORT As Long = &H1403&
Public Const GL_INT As Long = &H1404&
Public Const GL_UNSIGNED_INT As Long = &H1405&
Public Const GL_FLOAT As Long = &H1406&
Public Const GL_DOUBLE As Long = &H140A&
Public Const GL_2_BYTES As Long = &H1407&
Public Const GL_3_BYTES As Long = &H1408&
Public Const GL_4_BYTES As Long = &H1409&

' --- Funcoes de comparacao (depth/stencil/alpha) ----------------------------
Public Const GL_NEVER As Long = &H200&
Public Const GL_LESS As Long = &H201&
Public Const GL_EQUAL As Long = &H202&
Public Const GL_LEQUAL As Long = &H203&
Public Const GL_GREATER As Long = &H204&
Public Const GL_NOTEQUAL As Long = &H205&
Public Const GL_GEQUAL As Long = &H206&
Public Const GL_ALWAYS As Long = &H207&

' --- Blend ------------------------------------------------------------------
Public Const GL_BLEND As Long = &HBE2&
Public Const GL_BLEND_SRC As Long = &HBE1&
Public Const GL_BLEND_DST As Long = &HBE0&
Public Const GL_SRC_COLOR As Long = &H300&
Public Const GL_ONE_MINUS_SRC_COLOR As Long = &H301&
Public Const GL_SRC_ALPHA As Long = &H302&
Public Const GL_ONE_MINUS_SRC_ALPHA As Long = &H303&
Public Const GL_DST_ALPHA As Long = &H304&
Public Const GL_ONE_MINUS_DST_ALPHA As Long = &H305&
Public Const GL_DST_COLOR As Long = &H306&
Public Const GL_ONE_MINUS_DST_COLOR As Long = &H307&
Public Const GL_SRC_ALPHA_SATURATE As Long = &H308&

' --- Operacoes logicas ------------------------------------------------------
Public Const GL_LOGIC_OP As Long = &HBF1&
Public Const GL_LOGIC_OP_MODE As Long = &HBF0&
Public Const GL_CLEAR As Long = &H1500&
Public Const GL_AND As Long = &H1501&
Public Const GL_AND_REVERSE As Long = &H1502&
Public Const GL_COPY As Long = &H1503&
Public Const GL_AND_INVERTED As Long = &H1504&
Public Const GL_NOOP As Long = &H1505&
Public Const GL_XOR As Long = &H1506&
Public Const GL_OR As Long = &H1507&
Public Const GL_NOR As Long = &H1508&
Public Const GL_EQUIV As Long = &H1509&
Public Const GL_INVERT As Long = &H150A&
Public Const GL_OR_REVERSE As Long = &H150B&
Public Const GL_COPY_INVERTED As Long = &H150C&
Public Const GL_OR_INVERTED As Long = &H150D&
Public Const GL_NAND As Long = &H150E&
Public Const GL_SET As Long = &H150F&

' --- Depth buffer -----------------------------------------------------------
Public Const GL_DEPTH_TEST As Long = &HB71&
Public Const GL_DEPTH_FUNC As Long = &HB74&
Public Const GL_DEPTH_RANGE As Long = &HB70&
Public Const GL_DEPTH_WRITEMASK As Long = &HB72&
Public Const GL_DEPTH_CLEAR_VALUE As Long = &HB73&
Public Const GL_DEPTH_BITS As Long = &HD56&

' --- Stencil ----------------------------------------------------------------
Public Const GL_STENCIL_TEST As Long = &HB90&
Public Const GL_STENCIL_FUNC As Long = &HB92&
Public Const GL_STENCIL_VALUE_MASK As Long = &HB93&
Public Const GL_STENCIL_FAIL As Long = &HB94&
Public Const GL_STENCIL_PASS_DEPTH_FAIL As Long = &HB95&
Public Const GL_STENCIL_PASS_DEPTH_PASS As Long = &HB96&
Public Const GL_STENCIL_REF As Long = &HB97&
Public Const GL_STENCIL_WRITEMASK As Long = &HB98&
Public Const GL_STENCIL_BITS As Long = &HD57&
Public Const GL_KEEP As Long = &H1E00&
Public Const GL_REPLACE As Long = &H1E01&
Public Const GL_INCR As Long = &H1E02&
Public Const GL_DECR As Long = &H1E03&

' --- Alpha test / shade / faces ---------------------------------------------
Public Const GL_ALPHA_TEST As Long = &HBC0&
Public Const GL_ALPHA_TEST_FUNC As Long = &HBC1&
Public Const GL_ALPHA_TEST_REF As Long = &HBC2&
Public Const GL_SHADE_MODEL As Long = &HB54&
Public Const GL_FLAT As Long = &H1D00&
Public Const GL_SMOOTH As Long = &H1D01&
Public Const GL_FRONT As Long = &H404&
Public Const GL_BACK As Long = &H405&
Public Const GL_FRONT_AND_BACK As Long = &H408&
Public Const GL_CW As Long = &H900&
Public Const GL_CCW As Long = &H901&
Public Const GL_CULL_FACE As Long = &HB44&
Public Const GL_CULL_FACE_MODE As Long = &HB45&
Public Const GL_FRONT_FACE As Long = &HB46&

' --- Rasterizacao / pontos / linhas / poligonos -----------------------------
Public Const GL_POINT_SIZE As Long = &HB11&
Public Const GL_POINT_SIZE_RANGE As Long = &HB12&
Public Const GL_POINT_SIZE_GRANULARITY As Long = &HB13&
Public Const GL_POINT_SMOOTH As Long = &HB10&
Public Const GL_LINE_WIDTH As Long = &HB21&
Public Const GL_LINE_WIDTH_RANGE As Long = &HB22&
Public Const GL_LINE_WIDTH_GRANULARITY As Long = &HB23&
Public Const GL_LINE_SMOOTH As Long = &HB20&
Public Const GL_LINE_STIPPLE As Long = &HB24&
Public Const GL_LINE_STIPPLE_PATTERN As Long = &HB25&
Public Const GL_LINE_STIPPLE_REPEAT As Long = &HB26&
Public Const GL_POLYGON_MODE As Long = &HB40&
Public Const GL_POLYGON_SMOOTH As Long = &HB41&
Public Const GL_POLYGON_STIPPLE As Long = &HB42&
Public Const GL_POLYGON_OFFSET_FACTOR As Long = &H8038&
Public Const GL_POLYGON_OFFSET_UNITS As Long = &H2A00&
Public Const GL_POLYGON_OFFSET_POINT As Long = &H2A01&
Public Const GL_POLYGON_OFFSET_LINE As Long = &H2A02&
Public Const GL_POLYGON_OFFSET_FILL As Long = &H8037&
Public Const GL_POINT As Long = &HB00&
Public Const GL_LINE As Long = &HB01&
Public Const GL_FILL As Long = &HB02&

' --- Scissor / viewport / mascaras ------------------------------------------
Public Const GL_SCISSOR_TEST As Long = &HC11&
Public Const GL_SCISSOR_BOX As Long = &HC10&
Public Const GL_VIEWPORT As Long = &HBA2&
Public Const GL_COLOR_CLEAR_VALUE As Long = &HC22&
Public Const GL_COLOR_WRITEMASK As Long = &HC23&
Public Const GL_DITHER As Long = &HBD0&

' --- Hints ------------------------------------------------------------------
Public Const GL_PERSPECTIVE_CORRECTION_HINT As Long = &HC50&
Public Const GL_POINT_SMOOTH_HINT As Long = &HC51&
Public Const GL_LINE_SMOOTH_HINT As Long = &HC52&
Public Const GL_POLYGON_SMOOTH_HINT As Long = &HC53&
Public Const GL_FOG_HINT As Long = &HC54&
Public Const GL_DONT_CARE As Long = &H1100&
Public Const GL_FASTEST As Long = &H1101&
Public Const GL_NICEST As Long = &H1102&

' --- Pixel store / transfer ------------------------------------------------
Public Const GL_UNPACK_SWAP_BYTES As Long = &HCF0&
Public Const GL_UNPACK_LSB_FIRST As Long = &HCF1&
Public Const GL_UNPACK_ROW_LENGTH As Long = &HCF2&
Public Const GL_UNPACK_SKIP_ROWS As Long = &HCF3&
Public Const GL_UNPACK_SKIP_PIXELS As Long = &HCF4&
Public Const GL_UNPACK_ALIGNMENT As Long = &HCF5&
Public Const GL_PACK_SWAP_BYTES As Long = &HD00&
Public Const GL_PACK_LSB_FIRST As Long = &HD01&
Public Const GL_PACK_ROW_LENGTH As Long = &HD02&
Public Const GL_PACK_SKIP_ROWS As Long = &HD03&
Public Const GL_PACK_SKIP_PIXELS As Long = &HD04&
Public Const GL_PACK_ALIGNMENT As Long = &HD05&
Public Const GL_MAP_COLOR As Long = &HD10&
Public Const GL_MAP_STENCIL As Long = &HD11&
Public Const GL_INDEX_SHIFT As Long = &HD12&
Public Const GL_INDEX_OFFSET As Long = &HD13&
Public Const GL_RED_SCALE As Long = &HD14&
Public Const GL_RED_BIAS As Long = &HD15&
Public Const GL_GREEN_SCALE As Long = &HD18&
Public Const GL_GREEN_BIAS As Long = &HD19&
Public Const GL_BLUE_SCALE As Long = &HD1A&
Public Const GL_BLUE_BIAS As Long = &HD1B&
Public Const GL_ALPHA_SCALE As Long = &HD1C&
Public Const GL_ALPHA_BIAS As Long = &HD1D&
Public Const GL_DEPTH_SCALE As Long = &HD1E&
Public Const GL_DEPTH_BIAS As Long = &HD1F&
Public Const GL_ZOOM_X As Long = &HD16&
Public Const GL_ZOOM_Y As Long = &HD17&

' --- Pixel maps -------------------------------------------------------------
Public Const GL_PIXEL_MAP_I_TO_I As Long = &HC70&
Public Const GL_PIXEL_MAP_S_TO_S As Long = &HC71&
Public Const GL_PIXEL_MAP_I_TO_R As Long = &HC72&
Public Const GL_PIXEL_MAP_I_TO_G As Long = &HC73&
Public Const GL_PIXEL_MAP_I_TO_B As Long = &HC74&
Public Const GL_PIXEL_MAP_I_TO_A As Long = &HC75&
Public Const GL_PIXEL_MAP_R_TO_R As Long = &HC76&
Public Const GL_PIXEL_MAP_G_TO_G As Long = &HC77&
Public Const GL_PIXEL_MAP_B_TO_B As Long = &HC78&
Public Const GL_PIXEL_MAP_A_TO_A As Long = &HC79&

' --- Formatos de pixel ------------------------------------------------------
Public Const GL_COLOR_INDEX As Long = &H1900&
Public Const GL_STENCIL_INDEX As Long = &H1901&
Public Const GL_DEPTH_COMPONENT As Long = &H1902&
Public Const GL_RED As Long = &H1903&
Public Const GL_GREEN As Long = &H1904&
Public Const GL_BLUE As Long = &H1905&
Public Const GL_ALPHA As Long = &H1906&
Public Const GL_RGB As Long = &H1907&
Public Const GL_RGBA As Long = &H1908&
Public Const GL_LUMINANCE As Long = &H1909&
Public Const GL_LUMINANCE_ALPHA As Long = &H190A&
Public Const GL_BGR As Long = &H80E0&
Public Const GL_BGRA As Long = &H80E1&

' --- Texturas ---------------------------------------------------------------
Public Const GL_TEXTURE_1D As Long = &HDE0&
Public Const GL_TEXTURE_2D As Long = &HDE1&
Public Const GL_TEXTURE_WIDTH As Long = &H1000&
Public Const GL_TEXTURE_HEIGHT As Long = &H1001&
Public Const GL_TEXTURE_INTERNAL_FORMAT As Long = &H1003&
Public Const GL_TEXTURE_COMPONENTS As Long = &H1003&
Public Const GL_TEXTURE_BORDER_COLOR As Long = &H1004&
Public Const GL_TEXTURE_BORDER As Long = &H1005&
Public Const GL_TEXTURE_RED_SIZE As Long = &H805C&
Public Const GL_TEXTURE_GREEN_SIZE As Long = &H805D&
Public Const GL_TEXTURE_BLUE_SIZE As Long = &H805E&
Public Const GL_TEXTURE_ALPHA_SIZE As Long = &H805F&
Public Const GL_TEXTURE_LUMINANCE_SIZE As Long = &H8060&
Public Const GL_TEXTURE_INTENSITY_SIZE As Long = &H8061&
Public Const GL_TEXTURE_MAG_FILTER As Long = &H2800&
Public Const GL_TEXTURE_MIN_FILTER As Long = &H2801&
Public Const GL_TEXTURE_WRAP_S As Long = &H2802&
Public Const GL_TEXTURE_WRAP_T As Long = &H2803&
Public Const GL_TEXTURE_PRIORITY As Long = &H8066&
Public Const GL_TEXTURE_RESIDENT As Long = &H8067&
Public Const GL_TEXTURE_BINDING_1D As Long = &H8068&
Public Const GL_TEXTURE_BINDING_2D As Long = &H8069&
Public Const GL_NEAREST As Long = &H2600&
Public Const GL_LINEAR As Long = &H2601&
Public Const GL_NEAREST_MIPMAP_NEAREST As Long = &H2700&
Public Const GL_LINEAR_MIPMAP_NEAREST As Long = &H2701&
Public Const GL_NEAREST_MIPMAP_LINEAR As Long = &H2702&
Public Const GL_LINEAR_MIPMAP_LINEAR As Long = &H2703&
Public Const GL_REPEAT As Long = &H2901&
Public Const GL_CLAMP As Long = &H2900&
Public Const GL_CLAMP_TO_EDGE As Long = &H812F&
Public Const GL_CLAMP_TO_BORDER As Long = &H812D&
Public Const GL_TEXTURE_ENV As Long = &H2300&
Public Const GL_TEXTURE_ENV_MODE As Long = &H2200&
Public Const GL_TEXTURE_ENV_COLOR As Long = &H2201&
Public Const GL_MODULATE As Long = &H2100&
Public Const GL_DECAL As Long = &H2101&
Public Const GL_TEXTURE_GEN_MODE As Long = &H2500&
Public Const GL_OBJECT_LINEAR As Long = &H2401&
Public Const GL_EYE_LINEAR As Long = &H2400&
Public Const GL_SPHERE_MAP As Long = &H2402&
Public Const GL_TEXTURE_GEN_S As Long = &HC60&
Public Const GL_TEXTURE_GEN_T As Long = &HC61&
Public Const GL_TEXTURE_GEN_R As Long = &HC62&
Public Const GL_TEXTURE_GEN_Q As Long = &HC63&
Public Const GL_S As Long = &H2000&
Public Const GL_T As Long = &H2001&
Public Const GL_R As Long = &H2002&
Public Const GL_Q As Long = &H2003&
Public Const GL_MAX_TEXTURE_SIZE As Long = &HD33&

' --- Matrizes ---------------------------------------------------------------
Public Const GL_MATRIX_MODE As Long = &HBA0&
Public Const GL_MODELVIEW As Long = &H1700&
Public Const GL_PROJECTION As Long = &H1701&
Public Const GL_TEXTURE As Long = &H1702&
Public Const GL_NORMALIZE As Long = &HBA1&
Public Const GL_MODELVIEW_STACK_DEPTH As Long = &HBA3&
Public Const GL_PROJECTION_STACK_DEPTH As Long = &HBA4&
Public Const GL_TEXTURE_STACK_DEPTH As Long = &HBA5&
Public Const GL_MODELVIEW_MATRIX As Long = &HBA6&
Public Const GL_PROJECTION_MATRIX As Long = &HBA7&
Public Const GL_TEXTURE_MATRIX As Long = &HBA8&
Public Const GL_MAX_MODELVIEW_STACK_DEPTH As Long = &HD36&
Public Const GL_MAX_PROJECTION_STACK_DEPTH As Long = &HD38&
Public Const GL_MAX_TEXTURE_STACK_DEPTH As Long = &HD39&

' --- Buffers de destino -----------------------------------------------------
Public Const GL_DRAW_BUFFER As Long = &HC01&
Public Const GL_READ_BUFFER As Long = &HC02&
Public Const GL_AUX_BUFFERS As Long = &HC00&
Public Const GL_DOUBLEBUFFER As Long = &HC32&
Public Const GL_STEREO As Long = &HC33&
Public Const GL_INDEX_MODE As Long = &HC30&
Public Const GL_RGBA_MODE As Long = &HC31&
Public Const GL_INDEX_WRITEMASK As Long = &HC21&

' --- Iluminacao / materiais -------------------------------------------------
Public Const GL_LIGHTING As Long = &HB50&
Public Const GL_LIGHT0 As Long = &H4000&
Public Const GL_LIGHT1 As Long = &H4001&
Public Const GL_LIGHT2 As Long = &H4002&
Public Const GL_LIGHT3 As Long = &H4003&
Public Const GL_LIGHT4 As Long = &H4004&
Public Const GL_LIGHT5 As Long = &H4005&
Public Const GL_LIGHT6 As Long = &H4006&
Public Const GL_LIGHT7 As Long = &H4007&
Public Const GL_LIGHT_MODEL_LOCAL_VIEWER As Long = &HB51&
Public Const GL_LIGHT_MODEL_TWO_SIDE As Long = &HB52&
Public Const GL_LIGHT_MODEL_AMBIENT As Long = &HB53&
Public Const GL_COLOR_MATERIAL As Long = &HB57&
Public Const GL_COLOR_MATERIAL_FACE As Long = &HB55&
Public Const GL_COLOR_MATERIAL_PARAMETER As Long = &HB56&
Public Const GL_AMBIENT As Long = &H1200&
Public Const GL_DIFFUSE As Long = &H1201&
Public Const GL_SPECULAR As Long = &H1202&
Public Const GL_POSITION As Long = &H1203&
Public Const GL_SPOT_DIRECTION As Long = &H1204&
Public Const GL_SPOT_EXPONENT As Long = &H1205&
Public Const GL_SPOT_CUTOFF As Long = &H1206&
Public Const GL_CONSTANT_ATTENUATION As Long = &H1207&
Public Const GL_LINEAR_ATTENUATION As Long = &H1208&
Public Const GL_QUADRATIC_ATTENUATION As Long = &H1209&
Public Const GL_EMISSION As Long = &H1600&
Public Const GL_SHININESS As Long = &H1601&
Public Const GL_AMBIENT_AND_DIFFUSE As Long = &H1602&
Public Const GL_MAX_LIGHTS As Long = &HD31&

' --- Fog --------------------------------------------------------------------
Public Const GL_FOG As Long = &HB60&
Public Const GL_FOG_INDEX As Long = &HB61&
Public Const GL_FOG_DENSITY As Long = &HB62&
Public Const GL_FOG_START As Long = &HB63&
Public Const GL_FOG_END As Long = &HB64&
Public Const GL_FOG_MODE As Long = &HB65&
Public Const GL_FOG_COLOR As Long = &HB66&
Public Const GL_EXP As Long = &H800&
Public Const GL_EXP2 As Long = &H801&

' --- Planos de recorte ------------------------------------------------------
Public Const GL_CLIP_PLANE0 As Long = &H3000&
Public Const GL_CLIP_PLANE1 As Long = &H3001&
Public Const GL_CLIP_PLANE2 As Long = &H3002&
Public Const GL_CLIP_PLANE3 As Long = &H3003&
Public Const GL_CLIP_PLANE4 As Long = &H3004&
Public Const GL_CLIP_PLANE5 As Long = &H3005&
Public Const GL_MAX_CLIP_PLANES As Long = &HD32&

' --- Vertex arrays / client state -------------------------------------------
Public Const GL_VERTEX_ARRAY As Long = &H8074&
Public Const GL_NORMAL_ARRAY As Long = &H8075&
Public Const GL_COLOR_ARRAY As Long = &H8076&
Public Const GL_INDEX_ARRAY As Long = &H8077&
Public Const GL_TEXTURE_COORD_ARRAY As Long = &H8078&
Public Const GL_EDGE_FLAG_ARRAY As Long = &H8079&
Public Const GL_VERTEX_ARRAY_SIZE As Long = &H807A&
Public Const GL_VERTEX_ARRAY_TYPE As Long = &H807B&
Public Const GL_VERTEX_ARRAY_STRIDE As Long = &H807C&
Public Const GL_VERTEX_ARRAY_POINTER As Long = &H808E&
Public Const GL_NORMAL_ARRAY_TYPE As Long = &H807E&
Public Const GL_NORMAL_ARRAY_STRIDE As Long = &H807F&
Public Const GL_NORMAL_ARRAY_POINTER As Long = &H808F&
Public Const GL_COLOR_ARRAY_SIZE As Long = &H8081&
Public Const GL_COLOR_ARRAY_TYPE As Long = &H8082&
Public Const GL_COLOR_ARRAY_STRIDE As Long = &H8083&
Public Const GL_COLOR_ARRAY_POINTER As Long = &H8090&
Public Const GL_INDEX_ARRAY_TYPE As Long = &H8085&
Public Const GL_INDEX_ARRAY_STRIDE As Long = &H8086&
Public Const GL_INDEX_ARRAY_POINTER As Long = &H8091&
Public Const GL_TEXTURE_COORD_ARRAY_SIZE As Long = &H8088&
Public Const GL_TEXTURE_COORD_ARRAY_TYPE As Long = &H8089&
Public Const GL_TEXTURE_COORD_ARRAY_STRIDE As Long = &H808A&
Public Const GL_TEXTURE_COORD_ARRAY_POINTER As Long = &H8092&
Public Const GL_EDGE_FLAG_ARRAY_POINTER As Long = &H8093&
Public Const GL_V2F As Long = &H2A20&
Public Const GL_V3F As Long = &H2A21&
Public Const GL_C4UB_V2F As Long = &H2A22&
Public Const GL_C4UB_V3F As Long = &H2A23&
Public Const GL_C3F_V3F As Long = &H2A24&
Public Const GL_N3F_V3F As Long = &H2A25&
Public Const GL_C4F_N3F_V3F As Long = &H2A26&
Public Const GL_T2F_V3F As Long = &H2A27&
Public Const GL_T4F_V4F As Long = &H2A28&
Public Const GL_T2F_C4UB_V3F As Long = &H2A29&
Public Const GL_T2F_C3F_V3F As Long = &H2A2A&
Public Const GL_T2F_N3F_V3F As Long = &H2A2B&
Public Const GL_T2F_C4F_N3F_V3F As Long = &H2A2C&
Public Const GL_T4F_C4F_N3F_V4F As Long = &H2A2D&

' --- Display lists ----------------------------------------------------------
Public Const GL_LIST_BASE As Long = &HB32&
Public Const GL_LIST_INDEX As Long = &HB33&
Public Const GL_LIST_MODE As Long = &HB30&
Public Const GL_COMPILE As Long = &H1300&
Public Const GL_COMPILE_AND_EXECUTE As Long = &H1301&
Public Const GL_MAX_LIST_NESTING As Long = &HB31&

' --- Bits de glPushAttrib / glPushClientAttrib ------------------------------
Public Const GL_CURRENT_BIT As Long = &H1&
Public Const GL_POINT_BIT As Long = &H2&
Public Const GL_LINE_BIT As Long = &H4&
Public Const GL_POLYGON_BIT As Long = &H8&
Public Const GL_POLYGON_STIPPLE_BIT As Long = &H10&
Public Const GL_PIXEL_MODE_BIT As Long = &H20&
Public Const GL_LIGHTING_BIT As Long = &H40&
Public Const GL_FOG_BIT As Long = &H80&
Public Const GL_VIEWPORT_BIT As Long = &H800&
Public Const GL_TRANSFORM_BIT As Long = &H1000&
Public Const GL_ENABLE_BIT As Long = &H2000&
Public Const GL_HINT_BIT As Long = &H8000&
Public Const GL_EVAL_BIT As Long = &H10000&
Public Const GL_LIST_BIT As Long = &H20000&
Public Const GL_TEXTURE_BIT As Long = &H40000&
Public Const GL_SCISSOR_BIT As Long = &H80000&
Public Const GL_ALL_ATTRIB_BITS As Long = &HFFFFF&
Public Const GL_CLIENT_PIXEL_STORE_BIT As Long = &H1&
Public Const GL_CLIENT_VERTEX_ARRAY_BIT As Long = &H2&
Public Const GL_CLIENT_ALL_ATTRIB_BITS As Long = -1&

' --- Consultas gerais -------------------------------------------------------
Public Const GL_SUBPIXEL_BITS As Long = &HD50&
Public Const GL_INDEX_BITS As Long = &HD51&
Public Const GL_RED_BITS As Long = &HD52&
Public Const GL_GREEN_BITS As Long = &HD53&
Public Const GL_BLUE_BITS As Long = &HD54&
Public Const GL_ALPHA_BITS As Long = &HD55&
Public Const GL_ACCUM_BITS As Long = &HD5B&
Public Const GL_MAX_EVAL_ORDER As Long = &HD30&
Public Const GL_MAX_PIXEL_MAP_TABLE As Long = &HD34&
Public Const GL_MAX_ATTRIB_STACK_DEPTH As Long = &HD35&
Public Const GL_MAX_NAME_STACK_DEPTH As Long = &HD37&
Public Const GL_MAX_VIEWPORT_DIMS As Long = &HD3A&
Public Const GL_MAX_CLIENT_ATTRIB_STACK_DEPTH As Long = &HD3B&
Public Const GL_ATTRIB_STACK_DEPTH As Long = &HBB0&
Public Const GL_CLIENT_ATTRIB_STACK_DEPTH As Long = &HBB1&
Public Const GL_NAME_STACK_DEPTH As Long = &HD70&
Public Const GL_CURRENT_COLOR As Long = &HB00&
Public Const GL_CURRENT_INDEX As Long = &HB01&
Public Const GL_CURRENT_NORMAL As Long = &HB02&
Public Const GL_CURRENT_TEXTURE_COORDS As Long = &HB03&
Public Const GL_CURRENT_RASTER_COLOR As Long = &HB04&
Public Const GL_CURRENT_RASTER_INDEX As Long = &HB05&
Public Const GL_CURRENT_RASTER_TEXTURE_COORDS As Long = &HB06&
Public Const GL_CURRENT_RASTER_POSITION As Long = &HB07&
Public Const GL_CURRENT_RASTER_POSITION_VALID As Long = &HB08&
Public Const GL_CURRENT_RASTER_DISTANCE As Long = &HB09&
Public Const GL_EDGE_FLAG As Long = &HB43&
Public Const GL_VENDOR As Long = &H1F00&
Public Const GL_RENDERER As Long = &H1F01&
Public Const GL_VERSION As Long = &H1F02&
Public Const GL_EXTENSIONS As Long = &H1F03&

' --- Acumulacao -------------------------------------------------------------
Public Const GL_ACCUM As Long = &H100&
Public Const GL_LOAD As Long = &H101&
Public Const GL_RETURN As Long = &H102&
Public Const GL_MULT As Long = &H103&
Public Const GL_ADD As Long = &H104&
Public Const GL_ACCUM_CLEAR_VALUE As Long = &HB80&

' --- Evaluators -------------------------------------------------------------
Public Const GL_AUTO_NORMAL As Long = &HD80&
Public Const GL_COEFF As Long = &HA00&
Public Const GL_ORDER As Long = &HA01&
Public Const GL_DOMAIN As Long = &HA02&
Public Const GL_MAP1_COLOR_4 As Long = &HD90&
Public Const GL_MAP1_INDEX As Long = &HD91&
Public Const GL_MAP1_NORMAL As Long = &HD92&
Public Const GL_MAP1_TEXTURE_COORD_1 As Long = &HD93&
Public Const GL_MAP1_TEXTURE_COORD_2 As Long = &HD94&
Public Const GL_MAP1_TEXTURE_COORD_3 As Long = &HD95&
Public Const GL_MAP1_TEXTURE_COORD_4 As Long = &HD96&
Public Const GL_MAP1_VERTEX_3 As Long = &HD97&
Public Const GL_MAP1_VERTEX_4 As Long = &HD98&
Public Const GL_MAP2_COLOR_4 As Long = &HDB0&
Public Const GL_MAP2_INDEX As Long = &HDB1&
Public Const GL_MAP2_NORMAL As Long = &HDB2&
Public Const GL_MAP2_TEXTURE_COORD_1 As Long = &HDB3&
Public Const GL_MAP2_TEXTURE_COORD_2 As Long = &HDB4&
Public Const GL_MAP2_TEXTURE_COORD_3 As Long = &HDB5&
Public Const GL_MAP2_TEXTURE_COORD_4 As Long = &HDB6&
Public Const GL_MAP2_VERTEX_3 As Long = &HDB7&
Public Const GL_MAP2_VERTEX_4 As Long = &HDB8&
Public Const GL_MAP1_GRID_DOMAIN As Long = &HDD0&
Public Const GL_MAP1_GRID_SEGMENTS As Long = &HDD1&
Public Const GL_MAP2_GRID_DOMAIN As Long = &HDD2&
Public Const GL_MAP2_GRID_SEGMENTS As Long = &HDD3&

' --- Selection / feedback ---------------------------------------------------
Public Const GL_RENDER As Long = &H1C00&
Public Const GL_SELECT As Long = &H1C01&
Public Const GL_FEEDBACK As Long = &H1C02&
Public Const GL_RENDER_MODE As Long = &HC40&
Public Const GL_FEEDBACK_BUFFER_POINTER As Long = &HDF0&
Public Const GL_FEEDBACK_BUFFER_SIZE As Long = &HDF1&
Public Const GL_FEEDBACK_BUFFER_TYPE As Long = &HDF2&
Public Const GL_SELECTION_BUFFER_POINTER As Long = &HDF3&
Public Const GL_SELECTION_BUFFER_SIZE As Long = &HDF4&
Public Const GL_2D As Long = &H600&
Public Const GL_3D As Long = &H601&
Public Const GL_3D_COLOR As Long = &H602&
Public Const GL_3D_COLOR_TEXTURE As Long = &H603&
Public Const GL_4D_COLOR_TEXTURE As Long = &H604&
Public Const GL_PASS_THROUGH_TOKEN As Long = &H700&
Public Const GL_POINT_TOKEN As Long = &H701&
Public Const GL_LINE_TOKEN As Long = &H702&
Public Const GL_POLYGON_TOKEN As Long = &H703&
Public Const GL_BITMAP_TOKEN As Long = &H704&
Public Const GL_DRAW_PIXEL_TOKEN As Long = &H705&
Public Const GL_COPY_PIXEL_TOKEN As Long = &H706&
Public Const GL_LINE_RESET_TOKEN As Long = &H707&

Public Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlphaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type

Public Const PFD_TYPE_RGBA As Byte = 0
Public Const PFD_FLAG_DUAL_BUFFERED As Long = &H1
Public Const PFD_FLAG_DRAW_TO_WINDOW As Long = &H4
Public Const PFD_FLAG_SUPPORT_OPENGL As Long = &H20

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

' OpenGL

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
