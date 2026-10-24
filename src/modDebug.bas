Attribute VB_Name = "modDebug"
Option Explicit

' --- Declares de baixo nivel (timer query da GPU) --------------------------

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As Long, ByVal prgpvarg As Long, ByVal pvargResult As Long) As Long

' ============================================================================
' modDebug - diagnostic overlay (rolling plot) for the game loop.
'
' Not part of the raylib API: it is a project convenience. Each frame it
' collects frame time, CPU time spent in Game_Update and Game_Draw, GPU time
' (via timer query, when the driver exposes it) and the rlgl statistics
' (draw calls / vertices). The core calls the timing markers; the overlay is
' drawn and toggled by the user, who owns the font.
'
' Usage:
'   Game_Load:    DebugInit font
'   Game_Update:  If IsKeyPressed(KEY_F1) Then DebugToggle
'   Game_Draw:    DebugDrawOverlay   (before EndDrawing)
' ============================================================================

' Metricas disponiveis no plot (valores em milissegundos).
Public Enum DebugMetric
    DEBUGM_FRAME = 0
    DEBUGM_UPDATE = 1
    DEBUGM_DRAW = 2
    DEBUGM_GPU = 3
End Enum

Private Const DEBUG_DEFAULT_SAMPLES As Long = 240

' Constantes do timer query (ARB/EXT_timer_query). O sufixo & forca o literal
' a Long: sem ele o VB6 interpreta hex com o bit alto ligado como Integer negativo.
Private Const GL_EXTENSIONS As Long = &H1F03&
Private Const GL_TIME_ELAPSED As Long = &H88BF&
Private Const GL_QUERY_RESULT As Long = &H8866&
Private Const GL_QUERY_RESULT_AVAILABLE As Long = &H8867&

' --- Estado do modulo ------------------------------------------------------

Private mCap As Long
Private mCount As Long
Private mHead As Long
Private mVisible As Boolean
Private mFont As ImageFont
Private mFontSet As Boolean

' Ring buffers (um valor por frame; mCount <= mCap).
Private mFrameMs() As Single
Private mUpdateMs() As Single
Private mDrawMs() As Single
Private mGpuMs() As Single
Private mDrawCalls() As Long
Private mVertices() As Long

' Valores do frame atual.
Private mCurFrameMs As Single
Private mCurUpdateMs As Single
Private mCurDrawMs As Single
Private mCurGpuMs As Single
Private mCurDrawCalls As Long
Private mCurVertices As Long

Private mT0 As Currency
Private mFreq As Currency

' Overlay.
Private mOverlayMs As Single
Private mGpuAvailable As Boolean

' Timer query da GPU (GL_ARB_timer_query / GL_EXT_timer_query).
Private mGpuUse64 As Boolean
Private mFnGpuGen As Long
Private mFnGpuBegin As Long
Private mFnGpuEnd As Long
Private mFnGpuGetIv As Long
Private mFnGpuGetU64 As Long
Private mGpuQuery(0 To 1) As Long
Private mGpuQIdx As Long
Private mGpuQUsed As Long
Private mGpuActive As Boolean
Private mGpuStatus As String

' ============================================================================
' Ciclo de vida / controle
' ============================================================================

' Guarda a fonte usada pelo overlay e aloca o historico. Chamar uma vez (ex. no
' Game_Load). Sem a fonte o overlay nao tem como desenhar texto.
Public Sub DebugInit(ByRef font As ImageFont, Optional ByVal maxSamples As Long = DEBUG_DEFAULT_SAMPLES)
    If maxSamples < 8 Then maxSamples = 8
    mCap = maxSamples
    mCount = 0
    mHead = 0

    ReDim mFrameMs(0 To mCap - 1)
    ReDim mUpdateMs(0 To mCap - 1)
    ReDim mDrawMs(0 To mCap - 1)
    ReDim mGpuMs(0 To mCap - 1)
    ReDim mDrawCalls(0 To mCap - 1)
    ReDim mVertices(0 To mCap - 1)

    mFont = font
    mFontSet = True
    mVisible = True

    QueryPerformanceFrequency mFreq

    GpuInit
End Sub

Public Sub DebugToggle()
    mVisible = Not mVisible
End Sub

Public Function DebugIsVisible() As Boolean
    DebugIsVisible = mVisible
End Function

Private Function DebugReady() As Boolean
    DebugReady = (mCap > 0)
End Function

' ============================================================================
' Marcadores do frame (chamados pelo Run)
' ============================================================================

Public Sub DebugBeginFrame()
    If DebugReady() = False Then Exit Sub
    rlResetFrameStats
End Sub

Public Sub DebugBeginUpdate()
    If DebugReady() = False Then Exit Sub
    QueryPerformanceCounter mT0
End Sub

Public Sub DebugEndUpdate()
    If DebugReady() = False Then Exit Sub
    mCurUpdateMs = ElapsedMs()
End Sub

Public Sub DebugBeginDraw()
    If DebugReady() = False Then Exit Sub
    QueryPerformanceCounter mT0
End Sub

Public Sub DebugEndDraw()
    If DebugReady() = False Then Exit Sub
    ' Desconta o custo do proprio overlay (ele roda dentro de Game_Draw).
    mCurDrawMs = ElapsedMs() - mOverlayMs
    If mCurDrawMs < 0! Then mCurDrawMs = 0!
End Sub

Private Function ElapsedMs() As Single
    Dim t As Currency
    QueryPerformanceCounter t
    If mFreq = 0 Then Exit Function
    ElapsedMs = CSng(CDbl(t - mT0) * 1000# / CDbl(mFreq))
End Function

' Fecha o frame: registra o tempo de frame (wall clock) e empurra a amostra.
Public Sub DebugEndFrame(ByVal frameDt As Single)
    If DebugReady() = False Then Exit Sub

    mCurFrameMs = frameDt * 1000!
    If mCurFrameMs < 0! Then mCurFrameMs = 0!
    mCurDrawCalls = rlGetFrameDrawCallCount()
    mCurVertices = rlGetFrameVertexCount()

    mFrameMs(mHead) = mCurFrameMs
    mUpdateMs(mHead) = mCurUpdateMs
    mDrawMs(mHead) = mCurDrawMs
    mGpuMs(mHead) = mCurGpuMs
    mDrawCalls(mHead) = mCurDrawCalls
    mVertices(mHead) = mCurVertices

    mHead = (mHead + 1) Mod mCap
    If mCount < mCap Then mCount = mCount + 1
End Sub

' ============================================================================
' Consulta (usado pelo overlay e por testes)
' ============================================================================

Public Function DebugGetCount() As Long
    DebugGetCount = mCount
End Function

Public Function DebugCurrent(ByVal metric As DebugMetric) As Single
    Select Case metric
        Case DEBUGM_FRAME: DebugCurrent = mCurFrameMs
        Case DEBUGM_UPDATE: DebugCurrent = mCurUpdateMs
        Case DEBUGM_DRAW: DebugCurrent = mCurDrawMs
        Case DEBUGM_GPU: DebugCurrent = mCurGpuMs
    End Select
End Function

Public Function DebugMaximum(ByVal metric As DebugMetric) As Single
    Dim i As Long, v As Single, mx As Single
    For i = 0 To mCount - 1
        v = MetricAt(metric, i)
        If i = 0 Or v > mx Then mx = v
    Next i
    DebugMaximum = mx
End Function

Public Function DebugAverage(ByVal metric As DebugMetric) As Single
    Dim i As Long, s As Single
    If mCount = 0 Then Exit Function
    For i = 0 To mCount - 1
        s = s + MetricAt(metric, i)
    Next i
    DebugAverage = s / CSng(mCount)
End Function

' Media dos piores 1% dos frames ("1% low"), em ms. Usa selecao parcial, sem
' ordenar o buffer inteiro (k e no maximo ~1% de mCap).
Public Function DebugLow1(ByVal metric As DebugMetric) As Single
    Dim i As Long, j As Long, k As Long, s As Single
    Dim worst As Single
    If mCount = 0 Then Exit Function

    k = mCount \ 100
    If k < 1 Then k = 1

    ' Acumula os k maiores por insercao num buffer local de tamanho k.
    Dim buf() As Single
    ReDim buf(0 To k - 1)
    Dim filled As Long
    filled = 0
    For i = 0 To mCount - 1
        worst = MetricAt(metric, i)
        If filled < k Then
            j = filled - 1
            Do While j >= 0
                If buf(j) >= worst Then Exit Do
                buf(j + 1) = buf(j)
                j = j - 1
            Loop
            buf(j + 1) = worst
            filled = filled + 1
        ElseIf worst > buf(k - 1) Then
            j = k - 2
            Do While j >= 0
                If buf(j) >= worst Then Exit Do
                buf(j + 1) = buf(j)
                j = j - 1
            Loop
            buf(j + 1) = worst
        End If
    Next i

    For i = 0 To k - 1
        s = s + buf(i)
    Next i
    DebugLow1 = s / CSng(k)
End Function

' i = 0 devolve a amostra mais antiga do historico.
Private Function MetricAt(ByVal metric As DebugMetric, ByVal i As Long) As Single
    Dim idx As Long
    If mCount < mCap Then
        idx = i
    Else
        idx = (mHead + i) Mod mCap
    End If
    Select Case metric
        Case DEBUGM_FRAME: MetricAt = mFrameMs(idx)
        Case DEBUGM_UPDATE: MetricAt = mUpdateMs(idx)
        Case DEBUGM_DRAW: MetricAt = mDrawMs(idx)
        Case DEBUGM_GPU: MetricAt = mGpuMs(idx)
    End Select
End Function

' i = 0 devolve a amostra mais antiga (mais a esquerda no plot).
Public Function DebugSample(ByVal metric As DebugMetric, ByVal i As Long) As Single
    If i < 0 Or i >= mCount Then Exit Function
    DebugSample = MetricAt(metric, i)
End Function

Public Function DebugCurrentFps() As Single
    If mCurFrameMs <= 0! Then Exit Function
    DebugCurrentFps = 1000! / mCurFrameMs
End Function

Public Function DebugLow1Fps() As Single
    Dim ms As Single
    ms = DebugLow1(DEBUGM_FRAME)
    If ms <= 0! Then Exit Function
    DebugLow1Fps = 1000! / ms
End Function

Public Function DebugDrawCallCount() As Long
    DebugDrawCallCount = mCurDrawCalls
End Function

Public Function DebugVertexCount() As Long
    DebugVertexCount = mCurVertices
End Function

' ============================================================================
' Timer query da GPU (GL_ARB_timer_query / GL_EXT_timer_query)
'
' OpenGL 1.1 nao tem query de tempo; usamos a extensao quando o driver expoe,
' carregando os ponteiros via wglGetProcAddress e chamando-os pelo DispCallFunc.
' Se a extensao nao existir, mGpuAvailable fica False e o overlay mostra "n/a".
' ============================================================================

' Chamada generica stdcall de ate 3 argumentos Long (ponteiros entram como
' ByVal Long). O DispCallFunc espera, para cada argumento, um ponteiro para um
' VARIANT (nao o valor cru), e o pvargResult tem que apontar para um VARIANT
' valido mesmo quando a funcao retorna void.
Private Function CallGL3(ByVal addr As Long, ByVal a1 As Long, ByVal a2 As Long, ByVal a3 As Long, ByVal nargs As Long) As Long
    Dim vt(0 To 2) As Integer
    Dim av(0 To 2) As Long
    Dim vv(0 To 2) As Variant
    Dim res As Variant
    Dim i As Long

    If nargs > 0 Then vv(0) = a1
    If nargs > 1 Then vv(1) = a2
    If nargs > 2 Then vv(2) = a3

    For i = 0 To nargs - 1
        vt(i) = 3                      ' VT_I4
        av(i) = VarPtr(vv(i))
    Next i

    res = 0&
    DispCallFunc 0&, addr, 4&, 0, nargs, VarPtr(vt(0)), VarPtr(av(0)), VarPtr(res)
    CallGL3 = 0
End Function

Public Function DebugGpuIsAvailable() As Boolean
    DebugGpuIsAvailable = mGpuAvailable
End Function

Public Function DebugGpuStatus() As String
    DebugGpuStatus = mGpuStatus
End Function

Private Sub GpuInit()
    Dim p As Long, n As Long
    Dim s As String
    Dim isArb As Boolean, isExt As Boolean
    Dim prefix As String
    Dim ids(0 To 1) As Long
    Dim b() As Byte

    mGpuStatus = "no-gl"
    p = glGetString(GL_EXTENSIONS)
    If p = 0 Then Exit Sub
    n = lstrlenA(p)
    mGpuStatus = "no-ext"
    If n <= 0 Then Exit Sub

    ReDim b(0 To n - 1)
    CopyMemory b(0), ByVal p, n
    s = StrConv(b, vbUnicode)

    isArb = (InStr(1, s, "GL_ARB_timer_query", vbBinaryCompare) > 0)
    isExt = (InStr(1, s, "GL_EXT_timer_query", vbBinaryCompare) > 0)
    If isArb Then
        prefix = "ARB"
    ElseIf isExt Then
        prefix = "EXT"
    Else
        mGpuStatus = "not-listed"
        Exit Sub
    End If

    mGpuStatus = prefix

    mFnGpuGen = wglGetProcAddress("glGenQueries" & prefix)
    mFnGpuBegin = wglGetProcAddress("glBeginQuery" & prefix)
    mFnGpuEnd = wglGetProcAddress("glEndQuery" & prefix)
    mFnGpuGetIv = wglGetProcAddress("glGetQueryObjectiv" & prefix)
    If isArb Then
        mFnGpuGetU64 = wglGetProcAddress("glGetQueryObjectui64v" & prefix)
        mGpuUse64 = (mFnGpuGetU64 <> 0)
    End If

    If mFnGpuGen = 0 Or mFnGpuBegin = 0 Or mFnGpuEnd = 0 Or mFnGpuGetIv = 0 Then
        mGpuStatus = prefix & "-load-fail"
        Exit Sub
    End If
    If isArb And mGpuUse64 = False Then
        mGpuStatus = prefix & "-no64"
        Exit Sub
    End If

    CallGL3 mFnGpuGen, 2, VarPtr(ids(0)), 0, 2
    mGpuQuery(0) = ids(0)
    mGpuQuery(1) = ids(1)
    If mGpuQuery(0) = 0 Or mGpuQuery(1) = 0 Then
        mGpuStatus = prefix & "-gen-fail"
        Exit Sub
    End If

    mGpuQIdx = 0
    mGpuQUsed = 0
    mGpuAvailable = True
    mGpuStatus = prefix & "-ok"
End Sub

' Inicia a query antes do flush do lote.
Public Sub DebugGpuBegin()
    If mGpuAvailable = False Then Exit Sub
    mGpuActive = False
    CallGL3 mFnGpuBegin, GL_TIME_ELAPSED, mGpuQuery(mGpuQIdx), 0, 2
    mGpuActive = True
End Sub

' Termina a query depois do flush e le o resultado do frame anterior (sem stall).
Public Sub DebugGpuEnd()
    Dim other As Long
    Dim avail As Long
    Dim v64(0 To 1) As Long
    Dim iv As Long
    Dim dlo As Double, dhi As Double

    If mGpuActive = False Then Exit Sub
    CallGL3 mFnGpuEnd, GL_TIME_ELAPSED, 0, 0, 1
    mGpuActive = False
    mGpuQUsed = mGpuQUsed + 1

    other = 1 - mGpuQIdx

    If mGpuQUsed > 1 Then
        CallGL3 mFnGpuGetIv, mGpuQuery(other), GL_QUERY_RESULT_AVAILABLE, VarPtr(avail), 3
        If avail <> 0 Then
            If mGpuUse64 Then
                CallGL3 mFnGpuGetU64, mGpuQuery(other), GL_QUERY_RESULT, VarPtr(v64(0)), 3
                dlo = CDbl(v64(0))
                If dlo < 0# Then dlo = dlo + 4294967296#
                dhi = CDbl(v64(1))
                If dhi < 0# Then dhi = dhi + 4294967296#
                mCurGpuMs = CSng((dhi * 4294967296# + dlo) / 1000000#)
            Else
                CallGL3 mFnGpuGetIv, mGpuQuery(other), GL_QUERY_RESULT, VarPtr(iv), 3
                If iv > 0 Then mCurGpuMs = CSng(CDbl(iv) / 1000000#)
            End If
        End If
    End If

    mGpuQIdx = other
End Sub

' ============================================================================
' Overlay (drawn by the user, already in screen space)
' ============================================================================

Public Sub DebugDrawOverlay()
    Dim t0 As Currency, t1 As Currency

    If mVisible = False Then Exit Sub
    If mFontSet = False Then Exit Sub
    If mCount = 0 Then Exit Sub

    QueryPerformanceCounter t0
    DebugPlot 10!, 10!, 360!, 300!
    QueryPerformanceCounter t1
    If mFreq <> 0 Then mOverlayMs = CSng(CDbl(t1 - t0) * 1000# / CDbl(mFreq))
End Sub

Private Sub DebugPlot(ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim pad As Single, fh As Single, topY As Single, rowH As Single
    Dim i As Long
    Dim bg As ColorRGBA
    Dim r As Rectangle
    Dim one As Single
    Dim s1 As String, s2 As String

    pad = 6!
    fh = CSng(mFont.Height)
    If fh <= 0! Then fh = 8!

    rlSetLineWidth 1!

    bg.r = 0!: bg.g = 0!: bg.b = 0!: bg.a = 0.65
    r.x = x: r.y = y: r.Width = w: r.Height = h
    DrawRectangleRec r, bg

    s1 = "FPS " & Format$(DebugCurrentFps(), "0.0") & "  1% " & Format$(DebugLow1Fps(), "0.0") & "  n " & CStr(mCount)
    s2 = "calls " & CStr(mCurDrawCalls) & "  verts " & CStr(mCurVertices)
    DrawText mFont, s1, x + pad, y + pad, White
    DrawText mFont, s2, x + pad, y + pad + fh, White

    topY = y + pad + fh * 2! + pad
    rowH = (y + h - pad - topY) / 4!
    If rowH < fh + 6! Then rowH = fh + 6!

    ' 1) paineis e rotulos
    For i = 0 To 3
        DebugRowPanel x + pad, topY + CSng(i) * rowH, w - pad * 2!, rowH, i
    Next i

    ' 2) marcadores de 1 segundo
    DebugGrid x + pad, topY, w - pad * 2!, rowH * 4!

    ' 3) linhas do plot
    For i = 0 To 3
        DebugRowLine x + pad, topY + CSng(i) * rowH, w - pad * 2!, rowH, i
    Next i

    one = 1!
    r.x = x: r.y = y: r.Width = w: r.Height = h
    DrawRectangleLinesEx r, one, White
End Sub

Private Sub DebugRowPanel(ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, ByVal idx As Long)
    Dim fh As Single, plotY As Single, plotH As Single
    Dim name As String, col As ColorRGBA
    Dim metric As DebugMetric
    Dim panel As ColorRGBA
    Dim r As Rectangle
    Dim label As String

    fh = CSng(mFont.Height)
    If fh <= 0! Then fh = 8!
    plotY = y + fh
    plotH = h - fh - 2!
    If plotH < 4! Then plotH = 4!

    metric = idx
    GetMetricStyle idx, name, col

    If idx = 3 And mGpuAvailable = False Then
        label = name & " n/a"
    Else
        label = name & " " & Format$(DebugCurrent(metric), "0.0") & "  max " & Format$(DebugMaximum(metric), "0.0") & "  avg " & Format$(DebugAverage(metric), "0.0")
    End If
    DrawText mFont, label, x, y, col

    If idx Mod 2 = 0 Then
        panel.r = 1!: panel.g = 1!: panel.b = 1!: panel.a = 0.06
    Else
        panel.r = 1!: panel.g = 1!: panel.b = 1!: panel.a = 0.02
    End If
    r.x = x: r.y = plotY: r.Width = w: r.Height = plotH
    DrawRectangleRec r, panel
End Sub

Private Sub DebugRowLine(ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, ByVal idx As Long)
    Dim fh As Single, plotY As Single, plotH As Single
    Dim name As String, col As ColorRGBA
    Dim metric As DebugMetric
    Dim mx As Single
    Dim ref As ColorRGBA

    If idx = 3 And mGpuAvailable = False Then Exit Sub

    fh = CSng(mFont.Height)
    If fh <= 0! Then fh = 8!
    plotY = y + fh
    plotH = h - fh - 2!
    If plotH < 4! Then plotH = 4!

    metric = idx
    GetMetricStyle idx, name, col

    mx = DebugMaximum(metric)
    If mx < 0.001 Then mx = 0.001
    mx = mx * 1.1

    ref.r = 1!: ref.g = 1!: ref.b = 1!: ref.a = 0.18
    DrawLine x, plotY, x + w, plotY, ref, 1!

    PlotPolyline x, plotY, w, plotH, metric, mx, col
End Sub

' Polilinha usando o rlgl direto (1 vertice por amostra; bem mais barato que
' DrawLine por segmento).
Private Sub PlotPolyline(ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, ByVal metric As DebugMetric, ByVal maxV As Single, ByRef col As ColorRGBA)
    Dim i As Long
    Dim n As Long
    Dim v As Single, px As Single, py As Single
    Dim stepX As Single

    n = mCount
    If n < 2 Then Exit Sub
    If maxV <= 0! Then Exit Sub

    stepX = w / CSng(n - 1)

    rlSetTexture 0
    rlColor4f col.r, col.g, col.b, col.a
    rlBegin RL_LINE_STRIP
    For i = 0 To n - 1
        v = MetricAt(metric, i)
        If v < 0! Then v = 0!
        If v > maxV Then v = maxV
        px = x + CSng(i) * stepX
        py = (y + h) - (v / maxV) * h
        rlVertex2f px, py
    Next i
    rlEnd
End Sub

' Linhas verticais a cada 1 segundo de historico.
Private Sub DebugGrid(ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    Dim i As Long, n As Long
    Dim acc As Single, lastSec As Long
    Dim stepX As Single, px As Single
    Dim gcol As ColorRGBA

    n = mCount
    If n < 2 Then Exit Sub
    stepX = w / CSng(n - 1)

    gcol.r = 1!: gcol.g = 1!: gcol.b = 1!: gcol.a = 0.12

    lastSec = 0
    acc = 0!
    rlSetTexture 0
    rlColor4f gcol.r, gcol.g, gcol.b, gcol.a
    rlBegin RL_LINES
    For i = 0 To n - 1
        acc = acc + MetricAt(DEBUGM_FRAME, i) / 1000!
        If Int(acc) > lastSec Then
            lastSec = Int(acc)
            px = x + CSng(i) * stepX
            rlVertex2f px, y
            rlVertex2f px, y + h
        End If
    Next i
    rlEnd
End Sub

Private Sub GetMetricStyle(ByVal idx As Long, ByRef name As String, ByRef col As ColorRGBA)
    Select Case idx
        Case 0
            name = "FRAME"
            col.r = 0.45: col.g = 1!: col.b = 0.45: col.a = 1!
        Case 1
            name = "UPDATE"
            col.r = 0.4: col.g = 0.8: col.b = 1!: col.a = 1!
        Case 2
            name = "DRAW"
            col.r = 1!: col.g = 0.85: col.b = 0.3: col.a = 1!
        Case Else
            name = "GPU"
            col.r = 1!: col.g = 0.4: col.b = 0.85: col.a = 1!
    End Select
End Sub
