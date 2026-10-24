Attribute VB_Name = "modRLGL"
Option Explicit

' Primitive modes

Public Const RL_POINTS As Long = &H0
Public Const RL_LINES As Long = &H1
Public Const RL_LINE_LOOP As Long = &H2
Public Const RL_LINE_STRIP As Long = &H3
Public Const RL_TRIANGLES As Long = &H4
Public Const RL_TRIANGLE_STRIP As Long = &H5
Public Const RL_TRIANGLE_FAN As Long = &H6
Public Const RL_QUADS As Long = &H7

' Matrix modes

Public Const RL_MODELVIEW As Long = &H1700
Public Const RL_PROJECTION As Long = &H1701

' Limits

Private Const MATRIX_STACK_MAX As Long = 32
Private Const DEFAULT_MAX_VERTICES As Long = 8192
Private Const DEFAULT_MAX_DRAWS As Long = 2048

Public Const RL_MAX_SHADER_LOCATIONS As Long = 32

' Shader types

Public Const RL_SHADER_VERTEX As Long = &H8B31&
Public Const RL_SHADER_FRAGMENT As Long = &H8B30&

Private Const LOC_VERTEX_POSITION As Long = 0
Private Const LOC_VERTEX_TEXCOORD01 As Long = 1
Private Const LOC_VERTEX_NORMAL As Long = 3
Private Const LOC_VERTEX_COLOR As Long = 5
Private Const LOC_MATRIX_MVP As Long = 6
Private Const LOC_COLOR_DIFFUSE As Long = 12
Private Const LOC_MAP_ALBEDO As Long = 15

Private Const UT_FLOAT As Long = 0
Private Const UT_VEC2 As Long = 1
Private Const UT_VEC3 As Long = 2
Private Const UT_VEC4 As Long = 3
Private Const UT_INT As Long = 4
Private Const UT_IVEC2 As Long = 5
Private Const UT_IVEC3 As Long = 6
Private Const UT_IVEC4 As Long = 7
Private Const UT_UINT As Long = 8
Private Const UT_UIVEC2 As Long = 9
Private Const UT_UIVEC3 As Long = 10
Private Const UT_UIVEC4 As Long = 11
Private Const UT_SAMPLER2D As Long = 12

' Types

Public Type rlMatrix
    m(0 To 15) As Single
End Type

Private Type rlDrawCall
    mode As Long
    vertexCount As Long
    vertexOffset As Long
    textureId As Long
    mvp As rlMatrix
End Type

Private Type rlVertexBuffer
    elementCount As Long
    vertices() As Single
    texcoords() As Single
    colors() As Byte
    normals() As Single
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

' Module state

Private mBatch As rlRenderBatch
Private mMaxVertices As Long
Private mLineWidth As Single
Private mPointSize As Single

' Frame statistics
Private mFrameDrawCalls As Long
Private mFrameVertexCount As Long

Private mCurrentMode As Long
Private mInBeginEnd As Boolean
Private mCurrentTextureId As Long

' Shader state
Private mBatchShaderId As Long
Private mCurrentShaderId As Long
Private mShaderLocs(0 To RL_MAX_SHADER_LOCATIONS - 1) As Long
Private mDefaultTextureId As Long
Private mMatrixDirty As Boolean
Private mForceNewDrawCall As Boolean

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
Private mModelviewStackTop As Long
Private mProjectionStackTop As Long

' Lifecycle

Public Sub rlglInit()
    Dim i As Long

    mBatchShaderId = 0
    mCurrentShaderId = 0
    mForceNewDrawCall = False
    mMatrixDirty = False
    For i = 0 To RL_MAX_SHADER_LOCATIONS - 1
        mShaderLocs(i) = -1
    Next i

    rlLoadRenderBatch 1, DEFAULT_MAX_VERTICES
    LoadDefaultTexture
End Sub

Public Sub rlglClose()
    If GLShaderSupported() Then glUseProgramARB 0

    If mDefaultTextureId <> 0 Then
        glDeleteTextures 1, mDefaultTextureId
        mDefaultTextureId = 0
    End If

    mCurrentShaderId = 0
    mBatchShaderId = 0
    mForceNewDrawCall = False
    mBatch.drawCounter = 0
    If mBatch.bufferCount > 0 Then mBatch.vertexBuffer(mBatch.currentBuffer).elementCount = 0
End Sub

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
    mFrameDrawCalls = 0
    mFrameVertexCount = 0

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
    mModelviewStackTop = 0
    mProjectionStackTop = 0

    mCurrentTextureId = 0
    mInBeginEnd = False
    mColorR = 255: mColorG = 255: mColorB = 255: mColorA = 255
    mTexU = 0!: mTexV = 0!
    mNormalX = 0!: mNormalY = 0!: mNormalZ = 1!
    mLineWidth = 1!: mPointSize = 1!
    mBatchShaderId = 0
    mForceNewDrawCall = False
    mMatrixDirty = False
End Sub

Public Sub rlUnloadRenderBatch()
    mBatch.drawCounter = 0
End Sub

' Primitives (buffered)

Public Sub rlBegin(ByVal mode As Long)
    If mBatch.drawCounter = 0 Then mBatchShaderId = mCurrentShaderId
    If mBatchShaderId <> 0 And mMatrixDirty Then
        mForceNewDrawCall = True
        mMatrixDirty = False
    End If
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

    If mInBeginEnd = False Then Exit Sub

    rlCheckRenderBatchLimit 1

    EnsureDrawCall

    b = mBatch.currentBuffer
    i = mBatch.vertexBuffer(b).elementCount

    If mBatchShaderId <> 0 Then
        tx = x
        ty = y
        tz = z
    Else
        TransformVertex x, y, z, tx, ty, tz
    End If
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
    mBatch.draws(dc).vertexCount = mBatch.draws(dc).vertexCount + 1
End Sub

Public Sub rlVertex3i(ByVal x As Long, ByVal y As Long, ByVal z As Long)
    rlVertex3f CSng(x), CSng(y), CSng(z)
End Sub

' Per-vertex state

Public Sub rlColor4f(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    mColorR = ToByte(r)
    mColorG = ToByte(g)
    mColorB = ToByte(b)
    mColorA = ToByte(a)
End Sub

Public Sub rlColor3f(ByVal r As Single, ByVal g As Single, ByVal b As Single)
    rlColor4f r, g, b, 1!
End Sub

Public Sub rlColor4ub(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, ByVal a As Byte)
    mColorR = r: mColorG = g: mColorB = b: mColorA = a
End Sub

Public Sub rlTexCoord2f(ByVal u As Single, ByVal v As Single)
    mTexU = u: mTexV = v
End Sub

Public Sub rlTexCoord2i(ByVal u As Long, ByVal v As Long)
    mTexU = CSng(u): mTexV = CSng(v)
End Sub

Public Sub rlNormal3f(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    mNormalX = x: mNormalY = y: mNormalZ = z
End Sub

' Batch texture

Public Sub rlSetTexture(ByVal id As Long)
    mCurrentTextureId = id
End Sub

Public Sub rlEnableTexture(ByVal id As Long)
    mCurrentTextureId = id
End Sub

Public Sub rlDisableTexture()
    mCurrentTextureId = 0
End Sub

Public Function rlGetTextureId() As Long
    rlGetTextureId = mCurrentTextureId
End Function

' Matrices (CPU-side)

Public Sub rlMatrixMode(ByVal mode As Long)
    If mode = RL_MODELVIEW Or mode = RL_PROJECTION Then mCurrentMatrixMode = mode
End Sub

Public Sub rlPushMatrix()
    If mCurrentMatrixMode = RL_PROJECTION Then
        If mProjectionStackTop >= MATRIX_STACK_MAX Then Exit Sub
        mProjectionStack(mProjectionStackTop) = mProjection
        mProjectionStackTop = mProjectionStackTop + 1
    Else
        If mModelviewStackTop >= MATRIX_STACK_MAX Then Exit Sub
        mModelviewStack(mModelviewStackTop) = mModelview
        mModelviewStackTop = mModelviewStackTop + 1
    End If
    mMatrixDirty = True
End Sub

Public Sub rlPopMatrix()
    If mCurrentMatrixMode = RL_PROJECTION Then
        If mProjectionStackTop <= 0 Then Exit Sub
        mProjectionStackTop = mProjectionStackTop - 1
        mProjection = mProjectionStack(mProjectionStackTop)
    Else
        If mModelviewStackTop <= 0 Then Exit Sub
        mModelviewStackTop = mModelviewStackTop - 1
        mModelview = mModelviewStack(mModelviewStackTop)
    End If
    mMatrixDirty = True
End Sub

Public Sub rlLoadIdentity()
    If mCurrentMatrixMode = RL_PROJECTION Then
        mProjection = MatrixIdentity()
    Else
        mModelview = MatrixIdentity()
    End If
    mMatrixDirty = True
End Sub

Public Sub rlTranslatef(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Dim t As rlMatrix
    t = MatrixTranslate(x, y, z)
    ApplyMatrix t
End Sub

Public Sub rlRotatef(ByVal angle As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Dim t As rlMatrix
    t = MatrixRotate(angle, x, y, z)
    ApplyMatrix t
End Sub

Public Sub rlScalef(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Dim t As rlMatrix
    t = MatrixScale(x, y, z)
    ApplyMatrix t
End Sub

Public Sub rlMultMatrixf(ByRef mat As rlMatrix)
    ApplyMatrix mat
End Sub

Public Sub rlOrtho(ByVal left As Double, ByVal right As Double, ByVal bottom As Double, ByVal top As Double, ByVal nearPlane As Double, ByVal farPlane As Double)
    mProjection = MatrixOrtho(left, right, bottom, top, nearPlane, farPlane)
    mMatrixDirty = True
End Sub

Public Function rlGetMatrixModelview() As rlMatrix
    rlGetMatrixModelview = mModelview
End Function

Public Function rlGetMatrixProjection() As rlMatrix
    rlGetMatrixProjection = mProjection
End Function

' Flush / batch limits

Public Sub rlDrawRenderBatchActive()
    If mBatch.drawCounter = 0 Then Exit Sub
    FlushBatch
End Sub

Public Function rlCheckRenderBatchLimit(ByVal vCount As Long) As Boolean
    Dim b As Long
    b = mBatch.currentBuffer
    If mBatch.vertexBuffer(b).elementCount + vCount > mMaxVertices Then
        rlDrawRenderBatchActive
        rlCheckRenderBatchLimit = False
        Exit Function
    End If
    rlCheckRenderBatchLimit = True
End Function

Public Function rlGetBatchVertexCount() As Long
    rlGetBatchVertexCount = mBatch.vertexBuffer(mBatch.currentBuffer).elementCount
End Function

Public Function rlGetBatchDrawCallCount() As Long
    rlGetBatchDrawCallCount = mBatch.drawCounter
End Function

Public Sub rlResetFrameStats()
    mFrameDrawCalls = 0
    mFrameVertexCount = 0
End Sub

Public Function rlGetFrameDrawCallCount() As Long
    rlGetFrameDrawCallCount = mFrameDrawCalls
End Function

Public Function rlGetFrameVertexCount() As Long
    rlGetFrameVertexCount = mFrameVertexCount
End Function

Public Function rlGetBatchMaxVertices() As Long
    rlGetBatchMaxVertices = mMaxVertices
End Function

Public Function rlIsBatchActive() As Boolean
    rlIsBatchActive = mInBeginEnd
End Function

' Direct OpenGL state emulation

Public Sub rlClearColor(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
    glClearColor r, g, b, a
End Sub

Public Sub rlClearScreenBuffers()
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
End Sub

Public Sub rlViewport(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    glViewport x, y, width, height
End Sub

Public Sub rlEnable(ByVal state As Long)
    glEnable state
End Sub

Public Sub rlDisable(ByVal state As Long)
    glDisable state
End Sub

Public Sub rlSetLineWidth(ByVal width As Single)
    If width <> mLineWidth Then
        rlDrawRenderBatchActive
        mLineWidth = width
        glLineWidth width
    End If
End Sub

Public Sub rlSetPointSize(ByVal size As Single)
    If size <> mPointSize Then
        rlDrawRenderBatchActive
        mPointSize = size
        glPointSize size
    End If
End Sub

Public Sub rlDrawVertexArray(ByVal mode As Long, ByVal first As Long, ByVal count As Long)
    Dim useShader As Boolean

    mFrameDrawCalls = mFrameDrawCalls + 1
    mFrameVertexCount = mFrameVertexCount + count

    useShader = (mCurrentShaderId <> 0)
    If useShader Then
        BindShaderForRawDraw
    Else
        ApplyMatricesToGL
    End If

    glDrawArrays mode, first, count

    If useShader Then
        glUseProgramARB 0
    Else
        ResetGLMatrices
    End If
End Sub

Public Sub rlDrawVertexArrayElements(ByVal mode As Long, ByVal count As Long, ByVal indices As Long)
    Dim useShader As Boolean

    mFrameDrawCalls = mFrameDrawCalls + 1
    mFrameVertexCount = mFrameVertexCount + count

    useShader = (mCurrentShaderId <> 0)
    If useShader Then
        BindShaderForRawDraw
    Else
        ApplyMatricesToGL
    End If

    glDrawElements mode, count, GL_UNSIGNED_INT, indices

    If useShader Then
        glUseProgramARB 0
    Else
        ResetGLMatrices
    End If
End Sub

' Shaders

Public Function rlIsShaderReady(ByVal id As Long) As Boolean
    rlIsShaderReady = (id <> 0 And GLShaderSupported())
End Function

Public Function rlGetShaderIdDefault() As Long
    rlGetShaderIdDefault = 0
End Function

Public Function rlGetTextureIdDefault() As Long
    rlGetTextureIdDefault = mDefaultTextureId
End Function

Public Sub rlGetShaderLocsDefault(ByRef locs() As Long)
    Dim i As Long
    ReDim locs(0 To RL_MAX_SHADER_LOCATIONS - 1)
    For i = 0 To RL_MAX_SHADER_LOCATIONS - 1
        locs(i) = -1
    Next i
End Sub

Public Function rlCompileShader(ByVal code As String, ByVal shaderType As Long) As Long
    Dim shader As Long
    Dim b() As Byte
    Dim ptrs(0 To 0) As Long
    Dim lens(0 To 0) As Long
    Dim ok As Long

    If Len(code) = 0 Then Exit Function
    If GLShaderSupported() = False Then Exit Function

    glCreateShaderARB shaderType, shader
    If shader = 0 Then Exit Function

    b = AnsiZ(code)
    ptrs(0) = VarPtr(b(0))
    lens(0) = UBound(b)

    glShaderSourceARB shader, 1, VarPtr(ptrs(0)), VarPtr(lens(0))
    glCompileShaderARB shader

    ok = 0
    glGetShaderivARB shader, GL_OBJECT_COMPILE_STATUS_ARB, ok
    If ok = 0 Then
        LogShaderError shader
        glDeleteShaderARB shader
        Exit Function
    End If

    rlCompileShader = shader
End Function

Public Function rlLoadShaderCode(ByVal vsCode As String, ByVal fsCode As String) As Long
    Dim vs As Long
    Dim fs As Long
    Dim program As Long
    Dim ok As Long

    If Len(vsCode) = 0 Or Len(fsCode) = 0 Then Exit Function
    If GLShaderSupported() = False Then Exit Function

    vs = rlCompileShader(vsCode, RL_SHADER_VERTEX)
    If vs = 0 Then Exit Function

    fs = rlCompileShader(fsCode, RL_SHADER_FRAGMENT)
    If fs = 0 Then
        glDeleteShaderARB vs
        Exit Function
    End If

    glCreateProgramARB program
    If program = 0 Then
        glDeleteShaderARB vs
        glDeleteShaderARB fs
        Exit Function
    End If

    glAttachShaderARB program, vs
    glAttachShaderARB program, fs
    BindDefaultAttribs program
    glLinkProgramARB program

    ok = 0
    glGetShaderivARB program, GL_OBJECT_LINK_STATUS_ARB, ok
    If ok = 0 Then
        LogShaderError program
        glDeleteProgramARB program
        glDeleteShaderARB vs
        glDeleteShaderARB fs
        Exit Function
    End If

    glDetachShaderARB program, vs
    glDetachShaderARB program, fs
    glDeleteShaderARB vs
    glDeleteShaderARB fs

    rlLoadShaderCode = program
End Function

Public Sub rlUnloadShader(ByVal id As Long)
    If id = 0 Then Exit Sub
    glDeleteProgramARB id
End Sub

Public Function rlGetShaderLocation(ByVal program As Long, ByVal name As String) As Long
    Dim b() As Byte

    rlGetShaderLocation = -1
    If program = 0 Then Exit Function
    If Len(name) = 0 Then Exit Function
    If GLShaderSupported() = False Then Exit Function

    b = AnsiZ(name)
    rlGetShaderLocation = glGetUniformLocationARB(program, VarPtr(b(0)))
End Function

Public Function rlGetShaderLocationAttrib(ByVal program As Long, ByVal name As String) As Long
    Dim b() As Byte

    rlGetShaderLocationAttrib = -1
    If program = 0 Then Exit Function
    If Len(name) = 0 Then Exit Function
    If GLShaderSupported() = False Then Exit Function

    b = AnsiZ(name)
    rlGetShaderLocationAttrib = glGetAttribLocationARB(program, VarPtr(b(0)))
End Function

Public Sub rlSetShader(ByVal id As Long, ByRef locs() As Long)
    Dim i As Long
    Dim upper As Long

    If id <> 0 Then
        If rlIsShaderReady(id) = False Then Exit Sub
    End If
    If id = mCurrentShaderId Then Exit Sub

    rlDrawRenderBatchActive

    mCurrentShaderId = id

    upper = LocsUpper(locs)
    For i = 0 To RL_MAX_SHADER_LOCATIONS - 1
        If i <= upper And id <> 0 Then
            mShaderLocs(i) = locs(i)
        Else
            mShaderLocs(i) = -1
        End If
    Next i
End Sub

Public Sub rlEnableShader(ByVal id As Long)
    If id = 0 Then
        rlDisableShader
        Exit Sub
    End If
    If rlIsShaderReady(id) = False Then Exit Sub
    If id = mCurrentShaderId Then Exit Sub

    rlDrawRenderBatchActive
    mCurrentShaderId = id
    DisableShaderLocs
End Sub

Public Sub rlDisableShader()
    If mCurrentShaderId = 0 Then Exit Sub
    rlDrawRenderBatchActive
    mCurrentShaderId = 0
    DisableShaderLocs
End Sub

Public Sub rlUseShaderProgram(ByVal id As Long)
    glUseProgramARB id
End Sub

Public Sub rlSetUniform1f(ByVal programId As Long, ByVal locIndex As Long, ByVal value As Single)
    If locIndex < 0 Then Exit Sub
    If BindUniformProgram(programId) = False Then Exit Sub
    glUniform1fARB locIndex, value
    glUseProgramARB 0
End Sub

Public Sub rlSetUniform2f(ByVal programId As Long, ByVal locIndex As Long, ByVal v0 As Single, ByVal v1 As Single)
    If locIndex < 0 Then Exit Sub
    If BindUniformProgram(programId) = False Then Exit Sub
    glUniform2fARB locIndex, v0, v1
    glUseProgramARB 0
End Sub

Public Sub rlSetUniform3f(ByVal programId As Long, ByVal locIndex As Long, ByVal v0 As Single, ByVal v1 As Single, ByVal v2 As Single)
    If locIndex < 0 Then Exit Sub
    If BindUniformProgram(programId) = False Then Exit Sub
    glUniform3fARB locIndex, v0, v1, v2
    glUseProgramARB 0
End Sub

Public Sub rlSetUniform4f(ByVal programId As Long, ByVal locIndex As Long, ByVal v0 As Single, ByVal v1 As Single, ByVal v2 As Single, ByVal v3 As Single)
    If locIndex < 0 Then Exit Sub
    If BindUniformProgram(programId) = False Then Exit Sub
    glUniform4fARB locIndex, v0, v1, v2, v3
    glUseProgramARB 0
End Sub

Public Sub rlSetUniform1i(ByVal programId As Long, ByVal locIndex As Long, ByVal value As Long)
    If locIndex < 0 Then Exit Sub
    If BindUniformProgram(programId) = False Then Exit Sub
    glUniform1iARB locIndex, value
    glUseProgramARB 0
End Sub

Public Sub rlSetUniform1fv(ByVal programId As Long, ByVal locIndex As Long, ByRef values() As Single, ByVal count As Long)
    If count < 1 Then Exit Sub
    rlSetUniformPtr programId, locIndex, VarPtr(values(LBound(values))), UT_FLOAT, count
End Sub

Public Sub rlSetUniform2fv(ByVal programId As Long, ByVal locIndex As Long, ByRef values() As Single, ByVal count As Long)
    If count < 1 Then Exit Sub
    rlSetUniformPtr programId, locIndex, VarPtr(values(LBound(values))), UT_VEC2, count
End Sub

Public Sub rlSetUniform3fv(ByVal programId As Long, ByVal locIndex As Long, ByRef values() As Single, ByVal count As Long)
    If count < 1 Then Exit Sub
    rlSetUniformPtr programId, locIndex, VarPtr(values(LBound(values))), UT_VEC3, count
End Sub

Public Sub rlSetUniform4fv(ByVal programId As Long, ByVal locIndex As Long, ByRef values() As Single, ByVal count As Long)
    If count < 1 Then Exit Sub
    rlSetUniformPtr programId, locIndex, VarPtr(values(LBound(values))), UT_VEC4, count
End Sub

Public Sub rlSetUniform1iv(ByVal programId As Long, ByVal locIndex As Long, ByRef values() As Long, ByVal count As Long)
    If count < 1 Then Exit Sub
    rlSetUniformPtr programId, locIndex, VarPtr(values(LBound(values))), UT_INT, count
End Sub

Public Sub rlSetUniform2iv(ByVal programId As Long, ByVal locIndex As Long, ByRef values() As Long, ByVal count As Long)
    If count < 1 Then Exit Sub
    rlSetUniformPtr programId, locIndex, VarPtr(values(LBound(values))), UT_IVEC2, count
End Sub

Public Sub rlSetUniform3iv(ByVal programId As Long, ByVal locIndex As Long, ByRef values() As Long, ByVal count As Long)
    If count < 1 Then Exit Sub
    rlSetUniformPtr programId, locIndex, VarPtr(values(LBound(values))), UT_IVEC3, count
End Sub

Public Sub rlSetUniform4iv(ByVal programId As Long, ByVal locIndex As Long, ByRef values() As Long, ByVal count As Long)
    If count < 1 Then Exit Sub
    rlSetUniformPtr programId, locIndex, VarPtr(values(LBound(values))), UT_IVEC4, count
End Sub

Public Sub rlSetUniformPtr(ByVal programId As Long, ByVal locIndex As Long, ByVal valuePtr As Long, ByVal uniformType As Long, ByVal count As Long)
    If locIndex < 0 Then Exit Sub
    If count < 1 Then Exit Sub
    If BindUniformProgram(programId) = False Then Exit Sub

    Select Case uniformType
        Case UT_FLOAT
            glUniform1fvARB locIndex, count, valuePtr
        Case UT_VEC2
            glUniform2fvARB locIndex, count, valuePtr
        Case UT_VEC3
            glUniform3fvARB locIndex, count, valuePtr
        Case UT_VEC4
            glUniform4fvARB locIndex, count, valuePtr
        Case UT_INT
            glUniform1ivARB locIndex, count, valuePtr
        Case UT_IVEC2
            glUniform2ivARB locIndex, count, valuePtr
        Case UT_IVEC3
            glUniform3ivARB locIndex, count, valuePtr
        Case UT_IVEC4
            glUniform4ivARB locIndex, count, valuePtr
        Case UT_UINT
            glUniform1ivARB locIndex, count, valuePtr
        Case UT_UIVEC2
            glUniform2ivARB locIndex, count, valuePtr
        Case UT_UIVEC3
            glUniform3ivARB locIndex, count, valuePtr
        Case UT_UIVEC4
            glUniform4ivARB locIndex, count, valuePtr
        Case UT_SAMPLER2D
            glUniform1ivARB locIndex, count, valuePtr
    End Select

    glUseProgramARB 0
End Sub

Public Sub rlSetUniformMatrix4fv(ByVal programId As Long, ByVal locIndex As Long, ByRef mat As rlMatrix, ByVal count As Long)
    If locIndex < 0 Then Exit Sub
    If count < 1 Then Exit Sub
    If BindUniformProgram(programId) = False Then Exit Sub
    glUniformMatrix4fvARB locIndex, count, VarPtr(mat.m(0))
    glUseProgramARB 0
End Sub

Public Sub rlSetUniformSampler(ByVal programId As Long, ByVal locIndex As Long, ByVal textureId As Long, ByVal unit As Long)
    If locIndex < 0 Then Exit Sub
    If BindUniformProgram(programId) = False Then Exit Sub

    If unit < 0 Then unit = 0
    If unit > 3 Then unit = 3

    If unit > 0 Then
        If GLActiveTextureSupported() = False Then
            glUseProgramARB 0
            Exit Sub
        End If
    End If

    If GLActiveTextureSupported() Then
        glActiveTextureARB GL_TEXTURE0_ARB + unit
        glBindTexture GL_TEXTURE_2D, textureId
        glActiveTextureARB GL_TEXTURE0_ARB
    Else
        glBindTexture GL_TEXTURE_2D, textureId
    End If

    glUniform1iARB locIndex, unit
    glUseProgramARB 0
End Sub

' Shader internals

Private Function BindUniformProgram(ByVal programId As Long) As Boolean
    If programId = 0 Then Exit Function
    If GLShaderSupported() = False Then Exit Function
    glUseProgramARB programId
    BindUniformProgram = True
End Function

Private Function LocsUpper(ByRef locs() As Long) As Long
    On Error GoTo none
    LocsUpper = UBound(locs)
    Exit Function
none:
    LocsUpper = -1
End Function

Private Sub DisableShaderLocs()
    Dim i As Long
    For i = 0 To RL_MAX_SHADER_LOCATIONS - 1
        mShaderLocs(i) = -1
    Next i
End Sub

Private Function AnsiZ(ByVal s As String) As Byte()
    Dim b() As Byte

    If Len(s) = 0 Then
        ReDim b(0 To 0)
        b(0) = 0
        AnsiZ = b
        Exit Function
    End If

    b = StrConv(s, vbFromUnicode)
    ReDim Preserve b(0 To UBound(b) + 1)
    b(UBound(b)) = 0
    AnsiZ = b
End Function

Private Function AnsiToStr(ByRef b() As Byte) As String
    Dim s As String
    Dim p As Long

    If UBound(b) < 0 Then Exit Function
    s = StrConv(b, vbUnicode)
    p = InStr(s, vbNullChar)
    If p > 0 Then s = left$(s, p - 1)
    AnsiToStr = s
End Function

Private Sub LogShaderError(ByVal objectId As Long)
    Dim n As Long
    Dim got As Long
    Dim buf() As Byte

    n = 0
    glGetShaderivARB objectId, GL_OBJECT_INFO_LOG_LENGTH_ARB, n
    If n <= 1 Then Exit Sub

    ReDim buf(0 To n - 1)
    glGetShaderInfoLogARB objectId, n, got, buf(0)
    Debug.Print "Shader " & objectId & ": " & AnsiToStr(buf)
End Sub

Private Sub BindDefaultAttribs(ByVal program As Long)
    Dim b() As Byte

    b = AnsiZ("vertexPosition")
    glBindAttribLocationARB program, 0, VarPtr(b(0))
    b = AnsiZ("vertexTexCoord")
    glBindAttribLocationARB program, 1, VarPtr(b(0))
    b = AnsiZ("vertexNormal")
    glBindAttribLocationARB program, 2, VarPtr(b(0))
    b = AnsiZ("vertexColor")
    glBindAttribLocationARB program, 3, VarPtr(b(0))
End Sub

Private Sub LoadDefaultTexture()
    Dim px(0 To 3) As Byte

    If mDefaultTextureId <> 0 Then Exit Sub

    px(0) = 255
    px(1) = 255
    px(2) = 255
    px(3) = 255

    glGenTextures 1, mDefaultTextureId
    glBindTexture GL_TEXTURE_2D, mDefaultTextureId
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
    glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST
    glTexImage2D GL_TEXTURE_2D, 0, GL_RGBA, 1, 1, 0, GL_RGBA, GL_UNSIGNED_BYTE, ByVal VarPtr(px(0))
    glBindTexture GL_TEXTURE_2D, 0
End Sub

' Internals

Private Sub FlushBatch()
    Dim b As Long

    b = mBatch.currentBuffer
    If mBatch.drawCounter = 0 Then Exit Sub

    If mBatchShaderId <> 0 Then
        FlushBatchShader b
    Else
        FlushBatchFixed b
    End If

    mFrameDrawCalls = mFrameDrawCalls + mBatch.drawCounter
    mFrameVertexCount = mFrameVertexCount + mBatch.vertexBuffer(b).elementCount

    mBatch.vertexBuffer(b).elementCount = 0
    mBatch.drawCounter = 0
    mBatchShaderId = mCurrentShaderId
    mForceNewDrawCall = False
End Sub

Private Sub FlushBatchFixed(ByVal b As Long)
    Dim i As Long
    Dim dc As rlDrawCall
    Dim texOn As Boolean

    ResetGLMatrices

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
End Sub

Private Sub FlushBatchShader(ByVal b As Long)
    Dim i As Long
    Dim dc As rlDrawCall
    Dim texId As Long
    Dim useAttribs As Boolean

    If GLShaderSupported() = False Then Exit Sub

    glUseProgramARB mBatchShaderId

    If mShaderLocs(LOC_MAP_ALBEDO) >= 0 Then glUniform1iARB mShaderLocs(LOC_MAP_ALBEDO), 0

    useAttribs = (mShaderLocs(LOC_VERTEX_POSITION) >= 0)

    If useAttribs Then
        SetGenericArrays b, 1
    Else
        glEnableClientState GL_VERTEX_ARRAY
        glEnableClientState GL_COLOR_ARRAY
        glEnableClientState GL_NORMAL_ARRAY
        glEnableClientState GL_TEXTURE_COORD_ARRAY
        glVertexPointer 3, GL_FLOAT, 0, ByVal VarPtr(mBatch.vertexBuffer(b).vertices(0))
        glColorPointer 4, GL_UNSIGNED_BYTE, 0, ByVal VarPtr(mBatch.vertexBuffer(b).colors(0))
        glNormalPointer GL_FLOAT, 0, ByVal VarPtr(mBatch.vertexBuffer(b).normals(0))
        glTexCoordPointer 2, GL_FLOAT, 0, ByVal VarPtr(mBatch.vertexBuffer(b).texcoords(0))
    End If

    For i = 0 To mBatch.drawCounter - 1
        dc = mBatch.draws(i)

        If mShaderLocs(LOC_MATRIX_MVP) >= 0 Then
            glUniformMatrix4fvARB mShaderLocs(LOC_MATRIX_MVP), 1, VarPtr(dc.mvp.m(0))
        End If

        texId = dc.textureId
        If texId = 0 Then texId = mDefaultTextureId
        glBindTexture GL_TEXTURE_2D, texId
        glDrawArrays dc.mode, dc.vertexOffset, dc.vertexCount
    Next i

    If useAttribs Then
        SetGenericArrays b, 0
    Else
        glDisableClientState GL_TEXTURE_COORD_ARRAY
        glDisableClientState GL_COLOR_ARRAY
        glDisableClientState GL_NORMAL_ARRAY
        glDisableClientState GL_VERTEX_ARRAY
    End If

    glUseProgramARB 0
End Sub

Private Sub SetGenericArrays(ByVal b As Long, ByVal enable As Boolean)
    Dim locPos As Long
    Dim locUv As Long
    Dim locNrm As Long
    Dim locCol As Long

    locPos = mShaderLocs(LOC_VERTEX_POSITION)
    locUv = mShaderLocs(LOC_VERTEX_TEXCOORD01)
    locNrm = mShaderLocs(LOC_VERTEX_NORMAL)
    locCol = mShaderLocs(LOC_VERTEX_COLOR)

    If enable Then
        If locPos >= 0 Then glVertexAttribPointerARB locPos, 3, GL_FLOAT, 0, 0, VarPtr(mBatch.vertexBuffer(b).vertices(0))
        If locUv >= 0 Then glVertexAttribPointerARB locUv, 2, GL_FLOAT, 0, 0, VarPtr(mBatch.vertexBuffer(b).texcoords(0))
        If locNrm >= 0 Then glVertexAttribPointerARB locNrm, 3, GL_FLOAT, 0, 0, VarPtr(mBatch.vertexBuffer(b).normals(0))
        If locCol >= 0 Then glVertexAttribPointerARB locCol, 4, GL_UNSIGNED_BYTE, 1, 0, VarPtr(mBatch.vertexBuffer(b).colors(0))
        If locPos >= 0 Then glEnableVertexAttribArrayARB locPos
        If locUv >= 0 Then glEnableVertexAttribArrayARB locUv
        If locNrm >= 0 Then glEnableVertexAttribArrayARB locNrm
        If locCol >= 0 Then glEnableVertexAttribArrayARB locCol
    Else
        If locPos >= 0 Then glDisableVertexAttribArrayARB locPos
        If locUv >= 0 Then glDisableVertexAttribArrayARB locUv
        If locNrm >= 0 Then glDisableVertexAttribArrayARB locNrm
        If locCol >= 0 Then glDisableVertexAttribArrayARB locCol
    End If
End Sub

Private Sub BindShaderForRawDraw()
    Dim mvp As rlMatrix

    glUseProgramARB mCurrentShaderId

    If mShaderLocs(LOC_MAP_ALBEDO) >= 0 Then glUniform1iARB mShaderLocs(LOC_MAP_ALBEDO), 0

    If mShaderLocs(LOC_MATRIX_MVP) >= 0 Then
        mvp = MatrixMultiply(mProjection, mModelview)
        glUniformMatrix4fvARB mShaderLocs(LOC_MATRIX_MVP), 1, VarPtr(mvp.m(0))
    End If
End Sub

Private Sub ApplyMatricesToGL()
    glMatrixMode GL_PROJECTION
    glLoadMatrixf mProjection.m(0)
    glMatrixMode GL_MODELVIEW
    glLoadMatrixf mModelview.m(0)
End Sub

Private Sub ResetGLMatrices()
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    glMatrixMode GL_MODELVIEW
    glLoadIdentity
End Sub

Private Sub EnsureDrawCall()
    Dim b As Long
    Dim dc As Long

    b = mBatch.currentBuffer

    If mBatch.drawCounter > 0 And mForceNewDrawCall = False Then
        dc = mBatch.drawCounter - 1
        If mBatch.draws(dc).mode = mCurrentMode And mBatch.draws(dc).textureId = mCurrentTextureId Then Exit Sub
    End If

    If mBatch.drawCounter >= mBatch.maxDraws Then rlDrawRenderBatchActive

    dc = mBatch.drawCounter
    mBatch.draws(dc).mode = mCurrentMode
    mBatch.draws(dc).textureId = mCurrentTextureId
    mBatch.draws(dc).vertexOffset = mBatch.vertexBuffer(b).elementCount
    mBatch.draws(dc).vertexCount = 0
    If mBatchShaderId <> 0 Then mBatch.draws(dc).mvp = MatrixMultiply(mProjection, mModelview)
    mBatch.drawCounter = dc + 1
    mForceNewDrawCall = False
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
    Dim res As rlMatrix
    If mCurrentMatrixMode = RL_PROJECTION Then
        MatrixMultiplyInto res, mProjection, op
        mProjection = res
    Else
        MatrixMultiplyInto res, mModelview, op
        mModelview = res
    End If
    mMatrixDirty = True
End Sub

Private Function ToByte(ByVal v As Single) As Byte
    If v <= 0! Then
        ToByte = 0
    ElseIf v >= 1! Then
        ToByte = 255
    Else
        ToByte = CByte(v * 255!)
    End If
End Function

' Matrix math

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

Private Function MatrixMultiply(ByRef a As rlMatrix, ByRef b As rlMatrix) As rlMatrix
    Dim r As rlMatrix
    MatrixMultiplyInto r, a, b
    MatrixMultiply = r
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
