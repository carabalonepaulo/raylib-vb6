Attribute VB_Name = "modRLGL"
Option Explicit

' ============================================================================
' modRLGL - camada de batching no estilo do rlgl da raylib (GL 1.1)
'
' O rlgl da raylib existe para normalizar varias versoes do OpenGL. Aqui so
' temos a 1.1, entao o foco e o outro grande beneficio: BATCHING.
'
' Ideia:
'   - As chamadas de desenho (rlBegin/rlVertex/rlColor/rlTexCoord/rlEnd) NAO
'     chamam o OpenGL na hora. Elas acumulam vertices num buffer interno.
'   - Quando o buffer enche, ou quando voce pede (rlDrawRenderBatchActive), o
'     lote inteiro e enviado de uma vez com glDrawArrays (vertex arrays).
'   - Cada "draw call" guarda o modo (linhas/triangulos/quads) e a textura,
'     entao um mesmo lote pode conter varias texturas/modos.
'
' Diferenca importante em relacao a raylib: as transformacoes (rlTranslatef,
' rlRotatef, rlScalef) sao aplicadas nos vertices na CPU, em ordem IDENTICA ao
' pipeline fixo do OpenGL (multiplicacao a direita: M = M * op). Assim
' rlTranslatef/rlRotatef/rlScalef compoem igual a glTranslatef/glRotatef/
' glScalef. No flush, as matrizes do OpenGL sao zeradas (identity), pois os
' vertices ja vem prontos em espaco de recorte.
'
' Uso tipico:
'   rlMatrixMode RL_PROJECTION: rlLoadIdentity
'   rlOrtho 0, screenWidth, screenHeight, 0, -1, 1
'   rlMatrixMode RL_MODELVIEW: rlLoadIdentity
'
'   rlSetTexture tex.Id
'   rlBegin RL_QUADS
'       rlColor4f 1, 1, 1, 1
'       rlTexCoord2f 0, 0: rlVertex2f x, y
'       rlTexCoord2f 1, 0: rlVertex2f x + w, y
'       rlTexCoord2f 1, 1: rlVertex2f x + w, y + h
'       rlTexCoord2f 0, 1: rlVertex2f x, y + h
'   rlEnd
' ============================================================================

' --- Modos de primitiva (valores iguais aos da raylib) ---------------------

Public Const RL_LINES As Long = &H1
Public Const RL_TRIANGLES As Long = &H4
Public Const RL_QUADS As Long = &H7

' --- Modos de matriz -------------------------------------------------------

Public Const RL_MODELVIEW As Long = &H1700
Public Const RL_PROJECTION As Long = &H1701

' --- Limites ---------------------------------------------------------------

Private Const MATRIX_STACK_MAX As Long = 32
Private Const DEFAULT_MAX_VERTICES As Long = 8192
Private Const DEFAULT_MAX_DRAWS As Long = 2048

' --- Tipos -----------------------------------------------------------------

' Matriz 4x4 em ordem por coluna (m(col * 4 + row)), igual a raylib.
Public Type rlMatrix
    m(0 To 15) As Single
End Type

Private Type rlDrawCall
    mode As Long
    vertexCount As Long
    vertexOffset As Long
    textureId As Long
End Type

Private Type rlVertexBuffer
    elementCount As Long
    vertices() As Single       ' 3 floats por vertice
    texcoords() As Single      ' 2 floats por vertice
    colors() As Byte           ' 4 bytes por vertice (RGBA)
    normals() As Single        ' 3 floats por vertice
End Type

Private Type rlRenderBatch
    bufferCount As Long
    currentBuffer As Long
    vertexBuffer() As rlVertexBuffer
    draws() As rlDrawCall
    drawCounter As Long
    maxDraws As Long
    currentDepth As Single
End Type

' --- Estado do modulo ------------------------------------------------------

Private mBatch As rlRenderBatch
Private mMaxVertices As Long
Private mInitialized As Boolean

Private mCurrentMode As Long
Private mInBeginEnd As Boolean
Private mCurrentTextureId As Long

Private mColorR As Byte
Private mColorG As Byte
Private mColorB As Byte
Private mColorA As Byte
Private mTexU As Single
Private mTexV As Single
Private mNormalX As Single
Private mNormalY As Single
Private mNormalZ As Single

Private mModelview As rlMatrix
Private mProjection As rlMatrix
Private mCurrentMatrixMode As Long
Private mModelviewStack(0 To MATRIX_STACK_MAX - 1) As rlMatrix
Private mProjectionStack(0 To MATRIX_STACK_MAX - 1) As rlMatrix
Private mStackTop As Long

' ============================================================================
' Ciclo de vida
' ============================================================================

Public Sub rlglInit()
    If mInitialized Then Exit Sub
    rlLoadRenderBatch 1, DEFAULT_MAX_VERTICES
End Sub

' Chamado no inicio de toda funcao publica que mexe em estado. rlglInit zera
' matrizes/cores/textura, entao ele PRECISA rodar antes de o usuario configurar
' qualquer coisa (ex.: antes de rlOrtho). Sem isso, o primeiro rlBegin
' apagaria a projecao ja configurada.
Private Sub EnsureInit()
    If mInitialized = False Then rlglInit
End Sub

' Aloca numBuffers buffers de maxVertices vertices cada. Na pratica usamos um
' buffer ativo por vez; ao encher, ele e desenhado e reaproveitado.
Public Sub rlLoadRenderBatch(ByVal numBuffers As Long, ByVal maxVertices As Long)
    Dim i As Long

    If numBuffers < 1 Then numBuffers = 1
    If maxVertices < 4 Then maxVertices = 4

    mMaxVertices = maxVertices
    mBatch.bufferCount = numBuffers
    mBatch.currentBuffer = 0
    mBatch.drawCounter = 0
    mBatch.currentDepth = 0!
    mBatch.maxDraws = DEFAULT_MAX_DRAWS

    ReDim mBatch.vertexBuffer(0 To numBuffers - 1)
    For i = 0 To numBuffers - 1
        ReDim mBatch.vertexBuffer(i).vertices(0 To maxVertices * 3 - 1)
        ReDim mBatch.vertexBuffer(i).texcoords(0 To maxVertices * 2 - 1)
        ReDim mBatch.vertexBuffer(i).colors(0 To maxVertices * 4 - 1)
        ReDim mBatch.vertexBuffer(i).normals(0 To maxVertices * 3 - 1)
        mBatch.vertexBuffer(i).elementCount = 0
    Next i
    ReDim mBatch.draws(0 To mBatch.maxDraws - 1)

    mModelview = MatrixIdentity()
    mProjection = MatrixIdentity()
    mCurrentMatrixMode = RL_MODELVIEW
    mStackTop = 0

    mCurrentTextureId = 0
    mInBeginEnd = False
    mColorR = 255: mColorG = 255: mColorB = 255: mColorA = 255
    mTexU = 0!: mTexV = 0!
    mNormalX = 0!: mNormalY = 0!: mNormalZ = 1!

    mInitialized = True
End Sub

Public Sub rlUnloadRenderBatch()
    mBatch.drawCounter = 0
    mInitialized = False
End Sub

' ============================================================================
' Primitivas (acumulam no buffer)
' ============================================================================

Public Sub rlBegin(ByVal mode As Long)
    EnsureInit
    mCurrentMode = mode
    mInBeginEnd = True
End Sub

Public Sub rlEnd()
    mInBeginEnd = False
End Sub

Public Sub rlVertex2f(ByVal x As Single, ByVal y As Single)
    rlVertex3f x, y, 0!
End Sub

Public Sub rlVertex2i(ByVal x As Long, ByVal y As Long)
    rlVertex3f CSng(x), CSng(y), 0!
End Sub

Public Sub rlVertex3f(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Dim b As Long
    Dim i As Long
    Dim dc As Long
    Dim tx As Single, ty As Single, tz As Single

    EnsureInit
    If mInBeginEnd = False Then Exit Sub

    ' garante espaco (pode disparar um flush se o buffer estiver cheio)
    rlCheckRenderBatchLimit 1

    ' Cria/seleciona o draw call ANTES de gravar o vertice: o offset e o
    ' elementCount no momento da criacao, que ainda aponta para este vertice.
    ' (Se chamasse depois de gravar, o offset ficaria adiantado em 1.)
    EnsureDrawCall

    b = mBatch.currentBuffer
    i = mBatch.vertexBuffer(b).elementCount

    TransformVertex x, y, z, tx, ty, tz
    mBatch.vertexBuffer(b).vertices(i * 3) = tx
    mBatch.vertexBuffer(b).vertices(i * 3 + 1) = ty
    mBatch.vertexBuffer(b).vertices(i * 3 + 2) = tz
    mBatch.vertexBuffer(b).texcoords(i * 2) = mTexU
    mBatch.vertexBuffer(b).texcoords(i * 2 + 1) = mTexV
    mBatch.vertexBuffer(b).colors(i * 4) = mColorR
    mBatch.vertexBuffer(b).colors(i * 4 + 1) = mColorG
    mBatch.vertexBuffer(b).colors(i * 4 + 2) = mColorB
    mBatch.vertexBuffer(b).colors(i * 4 + 3) = mColorA
    mBatch.vertexBuffer(b).normals(i * 3) = mNormalX
    mBatch.vertexBuffer(b).normals(i * 3 + 1) = mNormalY
    mBatch.vertexBuffer(b).normals(i * 3 + 2) = mNormalZ
    mBatch.vertexBuffer(b).elementCount = i + 1

    dc = mBatch.drawCounter - 1
    If dc >= 0 Then mBatch.draws(dc).vertexCount = mBatch.draws(dc).vertexCount + 1
End Sub

Public Sub rlVertex3i(ByVal x As Long, ByVal y As Long, ByVal z As Long)
    rlVertex3f CSng(x), CSng(y), CSng(z)
End Sub

' --- Estado por vertice ----------------------------------------------------

Public Sub rlColor4f(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    EnsureInit
    mColorR = ToByte(r)
    mColorG = ToByte(g)
    mColorB = ToByte(b)
    mColorA = ToByte(a)
End Sub

Public Sub rlColor3f(ByVal r As Single, ByVal g As Single, ByVal b As Single)
    rlColor4f r, g, b, 1!
End Sub

Public Sub rlColor4ub(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, ByVal a As Byte)
    EnsureInit
    mColorR = r: mColorG = g: mColorB = b: mColorA = a
End Sub

Public Sub rlTexCoord2f(ByVal u As Single, ByVal v As Single)
    EnsureInit
    mTexU = u: mTexV = v
End Sub

Public Sub rlTexCoord2i(ByVal u As Long, ByVal v As Long)
    EnsureInit
    mTexU = CSng(u): mTexV = CSng(v)
End Sub

Public Sub rlNormal3f(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    EnsureInit
    mNormalX = x: mNormalY = y: mNormalZ = z
End Sub

' --- Textura do lote -------------------------------------------------------

Public Sub rlSetTexture(ByVal id As Long)
    EnsureInit
    mCurrentTextureId = id
End Sub

Public Sub rlEnableTexture(ByVal id As Long)
    EnsureInit
    mCurrentTextureId = id
End Sub

Public Sub rlDisableTexture()
    EnsureInit
    mCurrentTextureId = 0
End Sub

Public Function rlGetTextureId() As Long
    rlGetTextureId = mCurrentTextureId
End Function

' ============================================================================
' Matrizes (aplicadas na CPU, ordem identica ao pipeline fixo do OpenGL)
' ============================================================================

Public Sub rlMatrixMode(ByVal mode As Long)
    EnsureInit
    If mode = RL_MODELVIEW Or mode = RL_PROJECTION Then mCurrentMatrixMode = mode
End Sub

Public Sub rlPushMatrix()
    EnsureInit
    If mStackTop >= MATRIX_STACK_MAX Then Exit Sub
    If mCurrentMatrixMode = RL_PROJECTION Then
        mProjectionStack(mStackTop) = mProjection
    Else
        mModelviewStack(mStackTop) = mModelview
    End If
    mStackTop = mStackTop + 1
End Sub

Public Sub rlPopMatrix()
    EnsureInit
    If mStackTop <= 0 Then Exit Sub
    mStackTop = mStackTop - 1
    If mCurrentMatrixMode = RL_PROJECTION Then
        mProjection = mProjectionStack(mStackTop)
    Else
        mModelview = mModelviewStack(mStackTop)
    End If
End Sub

Public Sub rlLoadIdentity()
    EnsureInit
    If mCurrentMatrixMode = RL_PROJECTION Then
        mProjection = MatrixIdentity()
    Else
        mModelview = MatrixIdentity()
    End If
End Sub

Public Sub rlTranslatef(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Dim t As rlMatrix
    EnsureInit
    t = MatrixTranslate(x, y, z)
    ApplyMatrix t
End Sub

Public Sub rlRotatef(ByVal angle As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Dim t As rlMatrix
    EnsureInit
    t = MatrixRotate(angle, x, y, z)
    ApplyMatrix t
End Sub

Public Sub rlScalef(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Dim t As rlMatrix
    EnsureInit
    t = MatrixScale(x, y, z)
    ApplyMatrix t
End Sub

Public Sub rlMultMatrixf(ByRef mat As rlMatrix)
    EnsureInit
    ApplyMatrix mat
End Sub

Public Sub rlOrtho(ByVal left As Double, ByVal right As Double, ByVal bottom As Double, ByVal top As Double, ByVal nearPlane As Double, ByVal farPlane As Double)
    EnsureInit
    mProjection = MatrixOrtho(left, right, bottom, top, nearPlane, farPlane)
End Sub

Public Function rlGetMatrixModelview() As rlMatrix
    EnsureInit
    rlGetMatrixModelview = mModelview
End Function

Public Function rlGetMatrixProjection() As rlMatrix
    EnsureInit
    rlGetMatrixProjection = mProjection
End Function

' ============================================================================
' Flush / limites do lote
' ============================================================================

' Envia o lote acumulado para a GPU e reinicia os contadores.
Public Sub rlDrawRenderBatchActive()
    If mInitialized = False Then Exit Sub
    If mBatch.drawCounter = 0 Then Exit Sub
    FlushBatch
End Sub

' Garante espaco para vCount vertices. Se nao couber, faz flush e retorna
' False (o chamador continua desenhando num lote novo).
Public Function rlCheckRenderBatchLimit(ByVal vCount As Long) As Boolean
    Dim b As Long
    EnsureInit
    b = mBatch.currentBuffer
    If mBatch.vertexBuffer(b).elementCount + vCount > mMaxVertices Then
        rlDrawRenderBatchActive
        rlCheckRenderBatchLimit = False
        Exit Function
    End If
    rlCheckRenderBatchLimit = True
End Function

Public Function rlGetBatchVertexCount() As Long
    If mInitialized = False Then Exit Function
    rlGetBatchVertexCount = mBatch.vertexBuffer(mBatch.currentBuffer).elementCount
End Function

Public Function rlGetBatchDrawCallCount() As Long
    If mInitialized = False Then Exit Function
    rlGetBatchDrawCallCount = mBatch.drawCounter
End Function

Public Function rlGetBatchMaxVertices() As Long
    rlGetBatchMaxVertices = mMaxVertices
End Function

Public Function rlIsBatchActive() As Boolean
    rlIsBatchActive = mInBeginEnd
End Function

' ============================================================================
' Emulacao direta de estado/funcoes do OpenGL
' ============================================================================

Public Sub rlClearColor(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    glClearColor r, g, b, a
End Sub

Public Sub rlClearScreenBuffers()
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
End Sub

Public Sub rlViewport(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    glViewport x, y, width, height
End Sub

' Passthrough de glEnable/glDisable. Nao faz flush: assim como na raylib, o
' estado vale para o lote inteiro que ainda vai ser desenhado.
Public Sub rlEnable(ByVal state As Long)
    glEnable state
End Sub

Public Sub rlDisable(ByVal state As Long)
    glDisable state
End Sub

Public Sub rlSetLineWidth(ByVal width As Single)
    glLineWidth width
End Sub

Public Sub rlSetPointSize(ByVal size As Single)
    glPointSize size
End Sub

' Desenho cru: faz flush do lote e chama glDrawArrays com os arrays que voce
' mesmo configurou (glVertexPointer etc).
Public Sub rlDrawVertexArray(ByVal mode As Long, ByVal first As Long, ByVal count As Long)
    rlDrawRenderBatchActive
    glDrawArrays mode, first, count
End Sub

Public Sub rlDrawVertexArrayElements(ByVal mode As Long, ByVal count As Long, ByVal indices As Long)
    rlDrawRenderBatchActive
    glDrawElements mode, count, GL_UNSIGNED_INT, indices
End Sub

' ============================================================================
' Internos
' ============================================================================

Private Sub FlushBatch()
    Dim b As Long
    Dim i As Long
    Dim dc As rlDrawCall
    Dim texOn As Boolean

    b = mBatch.currentBuffer

    ' Os vertices ja vem em espaco de recorte; zera as matrizes do OpenGL.
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    glMatrixMode GL_MODELVIEW
    glLoadIdentity

    glEnableClientState GL_VERTEX_ARRAY
    glEnableClientState GL_COLOR_ARRAY
    glEnableClientState GL_NORMAL_ARRAY
    glVertexPointer 3, GL_FLOAT, 0, ByVal VarPtr(mBatch.vertexBuffer(b).vertices(0))
    glColorPointer 4, GL_UNSIGNED_BYTE, 0, ByVal VarPtr(mBatch.vertexBuffer(b).colors(0))
    glNormalPointer GL_FLOAT, 0, ByVal VarPtr(mBatch.vertexBuffer(b).normals(0))

    texOn = False

    For i = 0 To mBatch.drawCounter - 1
        dc = mBatch.draws(i)
        If dc.textureId <> 0 Then
            If texOn = False Then
                glEnableClientState GL_TEXTURE_COORD_ARRAY
                glTexCoordPointer 2, GL_FLOAT, 0, ByVal VarPtr(mBatch.vertexBuffer(b).texcoords(0))
                glEnable GL_TEXTURE_2D
                texOn = True
            End If
            glBindTexture GL_TEXTURE_2D, dc.textureId
        Else
            If texOn Then
                glDisableClientState GL_TEXTURE_COORD_ARRAY
                glDisable GL_TEXTURE_2D
                texOn = False
            End If
        End If
        glDrawArrays dc.mode, dc.vertexOffset, dc.vertexCount
    Next i

    glDisableClientState GL_COLOR_ARRAY
    glDisableClientState GL_NORMAL_ARRAY
    glDisableClientState GL_VERTEX_ARRAY
    If texOn Then
        glDisableClientState GL_TEXTURE_COORD_ARRAY
        glDisable GL_TEXTURE_2D
    End If

    mBatch.vertexBuffer(b).elementCount = 0
    mBatch.drawCounter = 0
End Sub

' Cria um novo draw call quando o modo ou a textura mudam; caso contrario
' reaproveita o ultimo (assim quads consecutivos viram um unico draw call).
Private Sub EnsureDrawCall()
    Dim b As Long
    Dim dc As Long

    b = mBatch.currentBuffer

    If mBatch.drawCounter > 0 Then
        dc = mBatch.drawCounter - 1
        If mBatch.draws(dc).mode = mCurrentMode And mBatch.draws(dc).textureId = mCurrentTextureId Then Exit Sub
    End If

    If mBatch.drawCounter >= mBatch.maxDraws Then rlDrawRenderBatchActive

    dc = mBatch.drawCounter
    mBatch.draws(dc).mode = mCurrentMode
    mBatch.draws(dc).textureId = mCurrentTextureId
    mBatch.draws(dc).vertexOffset = mBatch.vertexBuffer(b).elementCount
    mBatch.draws(dc).vertexCount = 0
    mBatch.drawCounter = dc + 1
End Sub

Private Sub TransformVertex(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByRef ox As Single, ByRef oy As Single, ByRef oz As Single)
    Dim mx As Single, my As Single, mz As Single, mw As Single

    mx = mModelview.m(0) * x + mModelview.m(4) * y + mModelview.m(8) * z + mModelview.m(12)
    my = mModelview.m(1) * x + mModelview.m(5) * y + mModelview.m(9) * z + mModelview.m(13)
    mz = mModelview.m(2) * x + mModelview.m(6) * y + mModelview.m(10) * z + mModelview.m(14)
    mw = mModelview.m(3) * x + mModelview.m(7) * y + mModelview.m(11) * z + mModelview.m(15)

    ox = mProjection.m(0) * mx + mProjection.m(4) * my + mProjection.m(8) * mz + mProjection.m(12) * mw
    oy = mProjection.m(1) * mx + mProjection.m(5) * my + mProjection.m(9) * mz + mProjection.m(13) * mw
    oz = mProjection.m(2) * mx + mProjection.m(6) * my + mProjection.m(10) * mz + mProjection.m(14) * mw
End Sub

Private Sub ApplyMatrix(ByRef op As rlMatrix)
    Dim cur As rlMatrix
    Dim res As rlMatrix
    cur = MatrixGet()
    MatrixMultiplyInto res, cur, op
    MatrixSet res
End Sub

Private Sub MatrixSet(ByRef mat As rlMatrix)
    If mCurrentMatrixMode = RL_PROJECTION Then
        mProjection = mat
    Else
        mModelview = mat
    End If
End Sub

Private Function MatrixGet() As rlMatrix
    If mCurrentMatrixMode = RL_PROJECTION Then
        MatrixGet = mProjection
    Else
        MatrixGet = mModelview
    End If
End Function

Private Function ToByte(ByVal v As Single) As Byte
    If v <= 0! Then
        ToByte = 0
    ElseIf v >= 1! Then
        ToByte = 255
    Else
        ToByte = CByte(v * 255!)
    End If
End Function

' --- Matematica de matrizes ------------------------------------------------

Private Function MatrixIdentity() As rlMatrix
    Dim r As rlMatrix
    r.m(0) = 1!: r.m(5) = 1!: r.m(10) = 1!: r.m(15) = 1!
    MatrixIdentity = r
End Function

Private Function MatrixTranslate(ByVal x As Single, ByVal y As Single, ByVal z As Single) As rlMatrix
    Dim r As rlMatrix
    r = MatrixIdentity()
    r.m(12) = x: r.m(13) = y: r.m(14) = z
    MatrixTranslate = r
End Function

Private Function MatrixScale(ByVal x As Single, ByVal y As Single, ByVal z As Single) As rlMatrix
    Dim r As rlMatrix
    r = MatrixIdentity()
    r.m(0) = x: r.m(5) = y: r.m(10) = z
    MatrixScale = r
End Function

Private Function MatrixRotate(ByVal angle As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single) As rlMatrix
    Dim r As rlMatrix
    Dim a As Single, c As Single, s As Single, t As Single
    Dim d As Single

    d = Sqr(x * x + y * y + z * z)
    If d = 0! Then
        MatrixRotate = MatrixIdentity()
        Exit Function
    End If
    x = x / d: y = y / d: z = z / d

    Const PI As Single = 3.14159265
    a = angle * (PI / 180!)
    c = Cos(a): s = Sin(a): t = 1! - c

    r.m(0) = t * x * x + c
    r.m(1) = t * x * y + s * z
    r.m(2) = t * x * z - s * y
    r.m(3) = 0!
    r.m(4) = t * x * y - s * z
    r.m(5) = t * y * y + c
    r.m(6) = t * y * z + s * x
    r.m(7) = 0!
    r.m(8) = t * x * z + s * y
    r.m(9) = t * y * z - s * x
    r.m(10) = t * z * z + c
    r.m(11) = 0!
    r.m(12) = 0!: r.m(13) = 0!: r.m(14) = 0!
    r.m(15) = 1!

    MatrixRotate = r
End Function

Private Function MatrixOrtho(ByVal left As Double, ByVal right As Double, ByVal bottom As Double, ByVal top As Double, ByVal nearPlane As Double, ByVal farPlane As Double) As rlMatrix
    Dim r As rlMatrix
    r.m(0) = CSng(2# / (right - left))
    r.m(5) = CSng(2# / (top - bottom))
    r.m(10) = CSng(-2# / (farPlane - nearPlane))
    r.m(12) = CSng(-(right + left) / (right - left))
    r.m(13) = CSng(-(top + bottom) / (top - bottom))
    r.m(14) = CSng(-(farPlane + nearPlane) / (farPlane - nearPlane))
    r.m(15) = 1!
    MatrixOrtho = r
End Function

Private Sub MatrixMultiplyInto(ByRef out As rlMatrix, ByRef a As rlMatrix, ByRef b As rlMatrix)
    Dim row As Long, col As Long, k As Long
    Dim sum As Single
    For col = 0 To 3
        For row = 0 To 3
            sum = 0!
            For k = 0 To 3
                sum = sum + a.m(k * 4 + row) * b.m(col * 4 + k)
            Next k
            out.m(col * 4 + row) = sum
        Next row
    Next col
End Sub
