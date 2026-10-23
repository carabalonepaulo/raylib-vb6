Attribute VB_Name = "modOpenGL"
Option Explicit

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
'   5. Rasterizacao         10. Vertex arrays         13. Pixel format (GDI)
'
' A secao 14 carrega extensoes (FBO/RBO) em runtime. Shaders e VBO ficam para
' depois, mas o mesmo mecanismo (GLLoadProc) ja serve para qualquer extensao.
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

' --- 13. Pixel format / bootstrap de contexto (GDI) -----------------------

Public Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hDC As Long, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function SetPixelFormat Lib "gdi32" (ByVal hDC As Long, ByVal PixelFormat As Long, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function SwapBuffers Lib "gdi32" (ByVal hDC As Long) As Long

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

' ============================================================================
' 14. Extensoes OpenGL (carregamento dinamico)
'
' O VB6 nao tem ponteiros de funcao. As extensoes (FBO/RBO, e futuramente
' shaders/VBO) sao resolvidas em runtime com wglGetProcAddress e chamadas por
' ponteiro via DispCallFunc (oleaut32).
'
' DispCallFunc recebe em prgpvarg um array de PONTEIROS para VARIANT (um por
' argumento) e em prgvt o VarType de cada um. O endereco da funcao vai em oVft
' com pvInstance = 0 e cc = CC_STDCALL.
'
' Uso tipico:
'   Dim fbo As Long
'   glGenFramebuffersEXT 1, fbo
'   glBindFramebufferEXT GL_FRAMEBUFFER, fbo
' ============================================================================

' --- Chamada de ponteiro de funcao -----------------------------------------

Private Declare Function DispCallFunc Lib "oleaut32" Alias "DispCallFunc" ( _
    ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, _
    ByVal vtReturn As Integer, ByVal cActuals As Long, _
    ByRef prgvt As Integer, ByRef prgpvarg As Long, ByRef pvargResult As Variant) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Const CC_STDCALL As Long = 4

' VARTYPEs usados como tipo de argumento/retorno
Private Const VT_EMPTY As Integer = 0
Private Const VT_I2 As Integer = 2
Private Const VT_I4 As Integer = 3
Private Const VT_R4 As Integer = 4
Private Const VT_R8 As Integer = 5
Private Const VT_UI1 As Integer = 17

' --- Constantes FBO/RBO ----------------------------------------------------
' Os valores sao identicos entre EXT/ARB/core; os dois nomes sao fornecidos
' para que o codigo funcione com qualquer estilo de nomenclatura.

' Framebuffer / renderbuffer
Public Const GL_FRAMEBUFFER As Long = &H8D40&
Public Const GL_FRAMEBUFFER_EXT As Long = &H8D40&
Public Const GL_RENDERBUFFER As Long = &H8D41&
Public Const GL_RENDERBUFFER_EXT As Long = &H8D41&
Public Const GL_READ_FRAMEBUFFER As Long = &H8CA8&
Public Const GL_READ_FRAMEBUFFER_EXT As Long = &H8CA8&
Public Const GL_DRAW_FRAMEBUFFER As Long = &H8CA9&
Public Const GL_DRAW_FRAMEBUFFER_EXT As Long = &H8CA9&
Public Const GL_FRAMEBUFFER_BINDING As Long = &H8CA6&
Public Const GL_FRAMEBUFFER_BINDING_EXT As Long = &H8CA6&
Public Const GL_RENDERBUFFER_BINDING As Long = &H8CA7&
Public Const GL_RENDERBUFFER_BINDING_EXT As Long = &H8CA7&

' Anexos (attachments)
Public Const GL_COLOR_ATTACHMENT0 As Long = &H8CE0&
Public Const GL_COLOR_ATTACHMENT0_EXT As Long = &H8CE0&
Public Const GL_COLOR_ATTACHMENT1 As Long = &H8CE1&
Public Const GL_COLOR_ATTACHMENT2 As Long = &H8CE2&
Public Const GL_COLOR_ATTACHMENT3 As Long = &H8CE3&
Public Const GL_COLOR_ATTACHMENT4 As Long = &H8CE4&
Public Const GL_COLOR_ATTACHMENT5 As Long = &H8CE5&
Public Const GL_COLOR_ATTACHMENT6 As Long = &H8CE6&
Public Const GL_COLOR_ATTACHMENT7 As Long = &H8CE7&
Public Const GL_COLOR_ATTACHMENT8 As Long = &H8CE8&
Public Const GL_COLOR_ATTACHMENT9 As Long = &H8CE9&
Public Const GL_COLOR_ATTACHMENT10 As Long = &H8CEA&
Public Const GL_COLOR_ATTACHMENT11 As Long = &H8CEB&
Public Const GL_COLOR_ATTACHMENT12 As Long = &H8CEC&
Public Const GL_COLOR_ATTACHMENT13 As Long = &H8CED&
Public Const GL_COLOR_ATTACHMENT14 As Long = &H8CEE&
Public Const GL_COLOR_ATTACHMENT15 As Long = &H8CEF&
Public Const GL_DEPTH_ATTACHMENT As Long = &H8D00&
Public Const GL_DEPTH_ATTACHMENT_EXT As Long = &H8D00&
Public Const GL_STENCIL_ATTACHMENT As Long = &H8D20&
Public Const GL_STENCIL_ATTACHMENT_EXT As Long = &H8D20&
Public Const GL_DEPTH_STENCIL_ATTACHMENT As Long = &H821A&
Public Const GL_DEPTH_STENCIL_ATTACHMENT_EXT As Long = &H821A&

' Status do framebuffer
Public Const GL_FRAMEBUFFER_COMPLETE As Long = &H8CD5&
Public Const GL_FRAMEBUFFER_COMPLETE_EXT As Long = &H8CD5&
Public Const GL_FRAMEBUFFER_INCOMPLETE_ATTACHMENT As Long = &H8CD6&
Public Const GL_FRAMEBUFFER_INCOMPLETE_ATTACHMENT_EXT As Long = &H8CD6&
Public Const GL_FRAMEBUFFER_INCOMPLETE_MISSING_ATTACHMENT As Long = &H8CD7&
Public Const GL_FRAMEBUFFER_INCOMPLETE_MISSING_ATTACHMENT_EXT As Long = &H8CD7&
Public Const GL_FRAMEBUFFER_INCOMPLETE_DRAW_BUFFER As Long = &H8CDB&
Public Const GL_FRAMEBUFFER_INCOMPLETE_DRAW_BUFFER_EXT As Long = &H8CDB&
Public Const GL_FRAMEBUFFER_INCOMPLETE_READ_BUFFER As Long = &H8CDC&
Public Const GL_FRAMEBUFFER_INCOMPLETE_READ_BUFFER_EXT As Long = &H8CDC&
Public Const GL_FRAMEBUFFER_UNSUPPORTED As Long = &H8CDD&
Public Const GL_FRAMEBUFFER_UNSUPPORTED_EXT As Long = &H8CDD&
Public Const GL_FRAMEBUFFER_INCOMPLETE_MULTISAMPLE As Long = &H8D56&
Public Const GL_FRAMEBUFFER_INCOMPLETE_MULTISAMPLE_EXT As Long = &H8D56&
Public Const GL_INVALID_FRAMEBUFFER_OPERATION As Long = &H506&
Public Const GL_INVALID_FRAMEBUFFER_OPERATION_EXT As Long = &H506&

' Consultas de framebuffer/attachment
Public Const GL_FRAMEBUFFER_ATTACHMENT_OBJECT_TYPE As Long = &H8CD0&
Public Const GL_FRAMEBUFFER_ATTACHMENT_OBJECT_NAME As Long = &H8CD1&
Public Const GL_FRAMEBUFFER_ATTACHMENT_TEXTURE_LEVEL As Long = &H8CD2&
Public Const GL_FRAMEBUFFER_ATTACHMENT_TEXTURE_CUBE_MAP_FACE As Long = &H8CD3&
Public Const GL_FRAMEBUFFER_ATTACHMENT_TEXTURE_LAYER As Long = &H8CD4&
Public Const GL_MAX_COLOR_ATTACHMENTS As Long = &H8CDF&
Public Const GL_MAX_COLOR_ATTACHMENTS_EXT As Long = &H8CDF&
Public Const GL_MAX_RENDERBUFFER_SIZE As Long = &H84E8&
Public Const GL_MAX_RENDERBUFFER_SIZE_EXT As Long = &H84E8&
Public Const GL_MAX_SAMPLES As Long = &H8D57&
Public Const GL_MAX_SAMPLES_EXT As Long = &H8D57&

' Consultas de renderbuffer
Public Const GL_RENDERBUFFER_WIDTH As Long = &H8D42&
Public Const GL_RENDERBUFFER_HEIGHT As Long = &H8D43&
Public Const GL_RENDERBUFFER_INTERNAL_FORMAT As Long = &H8D44&
Public Const GL_RENDERBUFFER_RED_SIZE As Long = &H8D50&
Public Const GL_RENDERBUFFER_GREEN_SIZE As Long = &H8D51&
Public Const GL_RENDERBUFFER_BLUE_SIZE As Long = &H8D52&
Public Const GL_RENDERBUFFER_ALPHA_SIZE As Long = &H8D53&
Public Const GL_RENDERBUFFER_DEPTH_SIZE As Long = &H8D54&
Public Const GL_RENDERBUFFER_STENCIL_SIZE As Long = &H8D55&

' Formatos internos usados como anexo de profundidade/estencil
Public Const GL_DEPTH_COMPONENT16 As Long = &H81A5&
Public Const GL_DEPTH_COMPONENT24 As Long = &H81A6&
Public Const GL_DEPTH_COMPONENT32 As Long = &H81A7&
Public Const GL_DEPTH_STENCIL As Long = &H84F9&
Public Const GL_DEPTH24_STENCIL8 As Long = &H88F0&
Public Const GL_STENCIL_INDEX1 As Long = &H8D46&
Public Const GL_STENCIL_INDEX4 As Long = &H8D47&
Public Const GL_STENCIL_INDEX8 As Long = &H8D48&
Public Const GL_STENCIL_INDEX16 As Long = &H8D49&

' --- Estado do carregador de extensoes --------------------------------------
' (no VB6 as variaveis de modulo precisam vir antes da primeira Sub/Function)

' Buffer de argumentos (reutilizado entre chamadas; VB6 aqui e single-thread)
Private mArgType(0 To 15) As Integer
Private mArgPtr(0 To 15) As Long
Private mArgVal(0 To 15) As Variant
Private mArgCount As Long
Private mResult As Variant

' Handle de opengl32.dll
Private mGLModule As Long

' Cache da lista de extensoes
Private mExtensions As String
Private mExtensionsLoaded As Boolean

' Ponteiros das extensoes FBO/RBO (carregados sob demanda)
Private mGenFramebuffers As Long
Private mBindFramebuffer As Long
Private mDeleteFramebuffers As Long
Private mFramebufferTexture2D As Long
Private mFramebufferRenderbuffer As Long
Private mCheckFramebufferStatus As Long
Private mGenRenderbuffers As Long
Private mBindRenderbuffer As Long
Private mDeleteRenderbuffers As Long
Private mRenderbufferStorage As Long
Private mGenerateMipmap As Long

Private mFboLoaded As Boolean
Private mFboReady As Boolean

Private Sub ArgsReset()
    mArgCount = 0
End Sub

' Empilha um argumento de 4 bytes (inteiro ou ponteiro).
Private Sub ArgL(ByVal value As Long)
    mArgVal(mArgCount) = value
    mArgType(mArgCount) = VT_I4
    mArgPtr(mArgCount) = VarPtr(mArgVal(mArgCount))
    mArgCount = mArgCount + 1
End Sub

' Empilha um ponteiro (VarPtr/StrPtr). Mesma representacao de ArgL, mas deixa
' a intencao explicita na leitura do codigo.
Private Sub ArgPtr(ByVal address As Long)
    ArgL address
End Sub

Private Function Invoke(ByVal pfn As Long, ByVal vtReturn As Integer) As Long
    mResult = Empty
    If pfn = 0 Then
        Invoke = 0
        Exit Function
    End If
    If DispCallFunc(0&, pfn, CC_STDCALL, vtReturn, mArgCount, mArgType(0), mArgPtr(0), mResult) <> 0 Then
        Invoke = 0
        Exit Function
    End If
    If IsEmpty(mResult) Then
        Invoke = 0
    Else
        Invoke = mResult
    End If
End Function

' --- Resolucao de enderecos de extensao ------------------------------------

' Resolve uma extensao: tenta wglGetProcAddress e cai para GetProcAddress em
' opengl32.dll (algumas funcoes ARB/core sao exportadas diretamente).
Public Function GLLoadProc(ByVal procName As String) As Long
    Dim p As Long
    If mGLModule = 0 Then mGLModule = GetModuleHandleA("opengl32.dll")
    p = wglGetProcAddress(procName)
    ' wglGetProcAddress retorna 1, 2, 3 ou -1 quando a funcao nao existe
    If p = 0 Or p = 1 Or p = 2 Or p = 3 Or p = -1 Then
        p = GetProcAddress(mGLModule, procName)
    End If
    GLLoadProc = p
End Function

' Tenta o nome EXT e depois o nome core (ex.: glGenFramebuffersEXT/glGenFramebuffers).
Public Function GLLoadProc2(ByVal extName As String, ByVal coreName As String) As Long
    Dim p As Long
    p = GLLoadProc(extName)
    If p = 0 Then p = GLLoadProc(coreName)
    GLLoadProc2 = p
End Function

' --- Strings da GL ---------------------------------------------------------

' Copia uma string C (char*) para uma String do VB6 (ANSI -> Unicode).
Private Function CStrFromPtr(ByVal p As Long) As String
    Dim n As Long
    Dim b() As Byte
    If p = 0 Then Exit Function
    n = lstrlenA(p)
    If n <= 0 Then Exit Function
    ReDim b(0 To n - 1)
    CopyMemory ByVal VarPtr(b(0)), ByVal p, n
    CStrFromPtr = StrConv(b, vbUnicode)
End Function

Public Function GLString(ByVal pname As Long) As String
    GLString = CStrFromPtr(glGetString(pname))
End Function

Public Function GLExtensionSupported(ByVal extension As String) As Boolean
    If mExtensionsLoaded = False Then
        mExtensions = " " & GLString(GL_EXTENSIONS) & " "
        mExtensionsLoaded = True
    End If
    GLExtensionSupported = (InStr(mExtensions, " " & extension & " ") > 0)
End Function

' --- Carregamento sob demanda das extensoes FBO/RBO -------------------------

Private Function LoadFramebufferExtensions() As Boolean
    If mFboLoaded Then
        LoadFramebufferExtensions = mFboReady
        Exit Function
    End If
    mFboLoaded = True

    mGenFramebuffers = GLLoadProc2("glGenFramebuffersEXT", "glGenFramebuffers")
    mBindFramebuffer = GLLoadProc2("glBindFramebufferEXT", "glBindFramebuffer")
    mDeleteFramebuffers = GLLoadProc2("glDeleteFramebuffersEXT", "glDeleteFramebuffers")
    mFramebufferTexture2D = GLLoadProc2("glFramebufferTexture2DEXT", "glFramebufferTexture2D")
    mFramebufferRenderbuffer = GLLoadProc2("glFramebufferRenderbufferEXT", "glFramebufferRenderbuffer")
    mCheckFramebufferStatus = GLLoadProc2("glCheckFramebufferStatusEXT", "glCheckFramebufferStatus")
    mGenRenderbuffers = GLLoadProc2("glGenRenderbuffersEXT", "glGenRenderbuffers")
    mBindRenderbuffer = GLLoadProc2("glBindRenderbufferEXT", "glBindRenderbuffer")
    mDeleteRenderbuffers = GLLoadProc2("glDeleteRenderbuffersEXT", "glDeleteRenderbuffers")
    mRenderbufferStorage = GLLoadProc2("glRenderbufferStorageEXT", "glRenderbufferStorage")
    mGenerateMipmap = GLLoadProc2("glGenerateMipmapEXT", "glGenerateMipmap")

    mFboReady = (mGenFramebuffers <> 0 And mBindFramebuffer <> 0 And mDeleteFramebuffers <> 0 And _
                 mFramebufferTexture2D <> 0 And mFramebufferRenderbuffer <> 0 And _
                 mCheckFramebufferStatus <> 0 And mGenRenderbuffers <> 0 And _
                 mBindRenderbuffer <> 0 And mDeleteRenderbuffers <> 0 And mRenderbufferStorage <> 0)

    LoadFramebufferExtensions = mFboReady
End Function

Public Function GLFramebufferSupported() As Boolean
    GLFramebufferSupported = LoadFramebufferExtensions()
End Function

' --- Framebuffers ----------------------------------------------------------

Public Sub glGenFramebuffersEXT(ByVal n As Long, ByRef framebuffers As Long)
    If LoadFramebufferExtensions() = False Then Exit Sub
    ArgsReset
    ArgL n
    ArgPtr VarPtr(framebuffers)
    Invoke mGenFramebuffers, VT_EMPTY
End Sub

Public Sub glBindFramebufferEXT(ByVal target As Long, ByVal framebuffer As Long)
    If LoadFramebufferExtensions() = False Then Exit Sub
    ArgsReset
    ArgL target
    ArgL framebuffer
    Invoke mBindFramebuffer, VT_EMPTY
End Sub

Public Sub glDeleteFramebuffersEXT(ByVal n As Long, ByRef framebuffers As Long)
    If LoadFramebufferExtensions() = False Then Exit Sub
    ArgsReset
    ArgL n
    ArgPtr VarPtr(framebuffers)
    Invoke mDeleteFramebuffers, VT_EMPTY
End Sub

Public Sub glFramebufferTexture2DEXT(ByVal target As Long, ByVal attachment As Long, ByVal textarget As Long, ByVal texture As Long, ByVal level As Long)
    If LoadFramebufferExtensions() = False Then Exit Sub
    ArgsReset
    ArgL target
    ArgL attachment
    ArgL textarget
    ArgL texture
    ArgL level
    Invoke mFramebufferTexture2D, VT_EMPTY
End Sub

Public Sub glFramebufferRenderbufferEXT(ByVal target As Long, ByVal attachment As Long, ByVal rbtarget As Long, ByVal renderbuffer As Long)
    If LoadFramebufferExtensions() = False Then Exit Sub
    ArgsReset
    ArgL target
    ArgL attachment
    ArgL rbtarget
    ArgL renderbuffer
    Invoke mFramebufferRenderbuffer, VT_EMPTY
End Sub

Public Function glCheckFramebufferStatusEXT(ByVal target As Long) As Long
    If LoadFramebufferExtensions() = False Then Exit Function
    ArgsReset
    ArgL target
    glCheckFramebufferStatusEXT = Invoke(mCheckFramebufferStatus, VT_I4)
End Function

' --- Renderbuffers ---------------------------------------------------------

Public Sub glGenRenderbuffersEXT(ByVal n As Long, ByRef renderbuffers As Long)
    If LoadFramebufferExtensions() = False Then Exit Sub
    ArgsReset
    ArgL n
    ArgPtr VarPtr(renderbuffers)
    Invoke mGenRenderbuffers, VT_EMPTY
End Sub

Public Sub glBindRenderbufferEXT(ByVal target As Long, ByVal renderbuffer As Long)
    If LoadFramebufferExtensions() = False Then Exit Sub
    ArgsReset
    ArgL target
    ArgL renderbuffer
    Invoke mBindRenderbuffer, VT_EMPTY
End Sub

Public Sub glDeleteRenderbuffersEXT(ByVal n As Long, ByRef renderbuffers As Long)
    If LoadFramebufferExtensions() = False Then Exit Sub
    ArgsReset
    ArgL n
    ArgPtr VarPtr(renderbuffers)
    Invoke mDeleteRenderbuffers, VT_EMPTY
End Sub

Public Sub glRenderbufferStorageEXT(ByVal target As Long, ByVal internalformat As Long, ByVal Width As Long, ByVal Height As Long)
    If LoadFramebufferExtensions() = False Then Exit Sub
    ArgsReset
    ArgL target
    ArgL internalformat
    ArgL Width
    ArgL Height
    Invoke mRenderbufferStorage, VT_EMPTY
End Sub

Public Sub glGenerateMipmapEXT(ByVal target As Long)
    If LoadFramebufferExtensions() = False Then Exit Sub
    ArgsReset
    ArgL target
    Invoke mGenerateMipmap, VT_EMPTY
End Sub

