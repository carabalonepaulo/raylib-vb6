# modRLGL - cheatsheet (rlgl-style batching layer for VB6 / OpenGL 1.1)

`modRLGL` is a small clone of raylib's rlgl. In raylib, rlgl exists mainly to
normalize several OpenGL versions; here only OpenGL 1.1 exists, so the focus is the
other big benefit: BATCHING.

Instead of issuing one GL call per primitive, you record vertices into an internal
buffer. When the buffer is full, or when the mode/texture changes, or when you ask
for it, the whole batch is flushed with `glDrawArrays` (client-side vertex arrays).

The API uses the `rl` prefix (instead of `gl`) so it never clashes with the raw
OpenGL bindings in `modOpenGL`.

- Coordinates in pixels, Y grows downward (same as raylib), as long as the projection
  is set with `rlOrtho 0, w, h, 0, -1, 1`.
- Colors passed to `rlColor4ub` are Bytes (0..255); `rlColor4f`/`rlColor3f` take
  Singles (0.0 .. 1.0) and are converted internally.
- `rlMatrix` stores 16 Singles in COLUMN-MAJOR order: `m(col * 4 + row)`, `m(0..15)`.
- Everything is recorded between `rlBegin` and `rlEnd`.
- `rlglInit()` must be called once before drawing. The raylib layer calls it from
  `InitWindow`, right after the OpenGL context is created.

Typical setup (once per frame):

```vb
rlMatrixMode RL_PROJECTION
rlLoadIdentity
rlOrtho 0, screenWidth, screenHeight, 0, -1, 1

rlMatrixMode RL_MODELVIEW
rlLoadIdentity

' ... record geometry ...
rlDrawRenderBatchActive        ' flush before EndDrawing
```

# Lifecycle

```vb
Sub rlglInit()
Sub rlglClose()
Sub rlLoadRenderBatch(ByVal numBuffers As Long, ByVal maxVertices As Long)
Sub rlUnloadRenderBatch()
```

- `rlglInit` allocates one buffer of 8192 vertices and resets all state. Call it exactly once, together with the raylib init (`InitWindow`). It also creates the internal 1x1 white texture used by shaders for untextured draws.
- `rlglClose` deletes that texture and resets the shader state. The raylib layer calls it from `CloseWindow`, before the context is destroyed.
- `rlLoadRenderBatch` lets you pick the buffer count and max vertices per buffer. Only the first buffer is used for drawing; it is drawn and reused when full.
- `rlUnloadRenderBatch` just clears the pending batch.

# Primitives (recorded into the batch)

```vb
Sub rlBegin(ByVal mode As Long)   ' RL_POINTS / RL_LINES / RL_LINE_LOOP / RL_LINE_STRIP
                                  ' RL_TRIANGLES / RL_TRIANGLE_STRIP / RL_TRIANGLE_FAN / RL_QUADS
Sub rlEnd()
Sub rlVertex2f(ByVal x As Single, ByVal y As Single)
Sub rlVertex2i(ByVal x As Long, ByVal y As Long)
Sub rlVertex3f(ByVal x As Single, ByVal y As Single, ByVal z As Single)
Sub rlVertex3i(ByVal x As Long, ByVal y As Long, ByVal z As Long)
```

Example: one textured quad.

```vb
rlSetTexture tex.Id
rlBegin RL_QUADS
    rlColor4ub 255, 255, 255, 255
    rlTexCoord2f 0, 0: rlVertex2f x, y
    rlTexCoord2f 1, 0: rlVertex2f x + w, y
    rlTexCoord2f 1, 1: rlVertex2f x + w, y + h
    rlTexCoord2f 0, 1: rlVertex2f x, y + h
rlEnd
```

# Per-vertex state

```vb
Sub rlColor4f(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
Sub rlColor3f(ByVal r As Single, ByVal g As Single, ByVal b As Single)
Sub rlColor4ub(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte, ByVal a As Byte)
Sub rlTexCoord2f(ByVal u As Single, ByVal v As Single)
Sub rlTexCoord2i(ByVal u As Long, ByVal v As Long)
Sub rlNormal3f(ByVal x As Single, ByVal y As Single, ByVal z As Single)
```

- Color, texcoord and normal are "current" values: they are copied into every vertex recorded after them. Default color is opaque white, default normal is `(0, 0, 1)`.

# Texture

```vb
Sub rlSetTexture(ByVal id As Long)   ' 0 = no texture
Sub rlEnableTexture(ByVal id As Long)
Sub rlDisableTexture()
Function rlGetTextureId() As Long
```

- The texture is a property of the draw call, not of the whole batch. Changing it starts a new draw call, so one batch can mix textured and untextured geometry (and several textures). Each draw call binds its texture before its `glDrawArrays`.

# Matrices (applied on the CPU)

```vb
Sub rlMatrixMode(ByVal mode As Long)   ' RL_MODELVIEW / RL_PROJECTION
Sub rlPushMatrix()
Sub rlPopMatrix()
Sub rlLoadIdentity()
Sub rlTranslatef(ByVal x As Single, ByVal y As Single, ByVal z As Single)
Sub rlRotatef(ByVal angle As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single)
Sub rlScalef(ByVal x As Single, ByVal y As Single, ByVal z As Single)
Sub rlMultMatrixf(ByRef mat As rlMatrix)
Sub rlOrtho(ByVal left As Double, ByVal right As Double, ByVal bottom As Double, ByVal top As Double, ByVal nearPlane As Double, ByVal farPlane As Double)
Function rlGetMatrixModelview() As rlMatrix
Function rlGetMatrixProjection() As rlMatrix
```

- Vertices are transformed on the CPU by `(modelview * projection)` when they are recorded, and stored already in clip space. At flush time the OpenGL matrices are set to identity, because the vertices are ready.
- The composition follows the OpenGL fixed-function order (post-multiply: `M = M * op`), so `rlTranslatef` / `rlRotatef` / `rlScalef` behave exactly like `glTranslatef` / `glRotatef` / `glScalef`. Example: translate then rotate rotates around the translated origin.

# Flush / batch info

```vb
Sub rlDrawRenderBatchActive()                      ' flush now
Function rlCheckRenderBatchLimit(ByVal vCount As Long) As Boolean
Function rlGetBatchVertexCount() As Long
Function rlGetBatchDrawCallCount() As Long
Function rlGetBatchMaxVertices() As Long
Function rlIsBatchActive() As Boolean              ' True between rlBegin/rlEnd
Sub rlResetFrameStats()
Function rlGetFrameDrawCallCount() As Long
Function rlGetFrameVertexCount() As Long
```

- The batch is flushed automatically when it is full (`rlGetBatchMaxVertices`) and when the draw-call list overflows. Consecutive vertices with the same mode and texture are merged into a single draw call.
- `rlGetBatch*` report the state of the batch that is CURRENTLY being recorded (pending, not yet flushed). `rlGetFrame*` report the ACCUMULATED totals for the whole frame: every flush (partial or forced) plus every raw draw (`rlDrawVertexArray*`) adds to them. `rlResetFrameStats` zeroes those accumulators; call it once at the start of each frame (the debug overlay does this via `DebugBeginFrame`).

# Direct OpenGL emulation

```vb
Sub rlClearColor(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
Sub rlClearScreenBuffers()                         ' glClear(COLOR | DEPTH)
Sub rlViewport(ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)

Sub rlEnable(ByVal state As Long)
Sub rlDisable(ByVal state As Long)
Sub rlSetLineWidth(ByVal width As Single)
Sub rlSetPointSize(ByVal size As Single)

Sub rlDrawVertexArray(ByVal mode As Long, ByVal first As Long, ByVal count As Long)
Sub rlDrawVertexArrayElements(ByVal mode As Long, ByVal count As Long, ByVal indices As Long)
```

- `rlEnable`/`rlDisable` are plain passthroughs to `glEnable`/`glDisable` and do NOT flush (state applies to the batch that will still be drawn).
- `rlSetLineWidth`/`rlSetPointSize` flush the pending batch first IF the value changed, because raster state applies to the whole batch at flush time.
- Raw draws: `rlDrawVertexArray` calls `glDrawArrays` with the arrays you configured yourself (`glVertexPointer` etc); `rlDrawVertexArrayElements` calls `glDrawElements` with `GL_UNSIGNED_INT` indices. They do NOT flush: call `rlDrawRenderBatchActive` YOURSELF before you configure the raw texture and client arrays (otherwise the flush would overwrite that state). They also upload the CPU matrices (modelview/projection) to the GL matrices so the raw array is drawn in object space with the current camera, then restore identity (batched vertices are already in clip space).

# Constants

```vb
RL_POINTS         &H0    ' glDrawArrays mode (values equal the GL enums)
RL_LINES          &H1
RL_LINE_LOOP      &H2
RL_LINE_STRIP     &H3
RL_TRIANGLES      &H4
RL_TRIANGLE_STRIP &H5
RL_TRIANGLE_FAN   &H6
RL_QUADS          &H7

RL_MODELVIEW      &H1700   ' rlMatrixMode
RL_PROJECTION     &H1701
```

# Types

```vb
Type rlMatrix
    m(0 To 15) As Single       ' column-major: m(col * 4 + row)
End Type
```

# Shaders

```vb
Const RL_SHADER_VERTEX   = &H8B31&
Const RL_SHADER_FRAGMENT = &H8B30&
Const RL_MAX_SHADER_LOCATIONS = 32

' program build
Function rlLoadShaderCode(ByVal vsCode As String, ByVal fsCode As String) As Long
Function rlCompileShader(ByVal code As String, ByVal shaderType As Long) As Long
Sub rlUnloadShader(ByVal id As Long)
Function rlGetShaderLocation(ByVal program As Long, ByVal name As String) As Long
Function rlGetShaderLocationAttrib(ByVal program As Long, ByVal name As String) As Long

' selection / queries
Sub rlSetShader(ByVal id As Long, ByRef locs() As Long)
Sub rlEnableShader(ByVal id As Long)
Sub rlDisableShader()
Sub rlUseShaderProgram(ByVal id As Long)
Function rlIsShaderReady(ByVal id As Long) As Boolean
Function rlGetShaderIdDefault() As Long
Sub rlGetShaderLocsDefault(ByRef locs() As Long)
Function rlGetTextureIdDefault() As Long

' uniforms (they bind the program themselves)
Sub rlSetUniform1f(ByVal programId As Long, ByVal locIndex As Long, ByVal value As Single)
Sub rlSetUniform2f(ByVal programId As Long, ByVal locIndex As Long, ByVal v0 As Single, ByVal v1 As Single)
Sub rlSetUniform3f(ByVal programId As Long, ByVal locIndex As Long, ByVal v0 As Single, ByVal v1 As Single, ByVal v2 As Single)
Sub rlSetUniform4f(ByVal programId As Long, ByVal locIndex As Long, ByVal v0 As Single, ByVal v1 As Single, ByVal v2 As Single, ByVal v3 As Single)
Sub rlSetUniform1i(ByVal programId As Long, ByVal locIndex As Long, ByVal value As Long)
Sub rlSetUniform1fv / 2fv / 3fv / 4fv (ByVal programId As Long, ByVal locIndex As Long, ByRef values() As Single, ByVal count As Long)
Sub rlSetUniform1iv / 2iv / 3iv / 4iv (ByVal programId As Long, ByVal locIndex As Long, ByRef values() As Long, ByVal count As Long)
Sub rlSetUniformMatrix4fv(ByVal programId As Long, ByVal locIndex As Long, ByRef mat As rlMatrix, ByVal count As Long)
Sub rlSetUniformSampler(ByVal programId As Long, ByVal locIndex As Long, ByVal textureId As Long, ByVal unit As Long)
Sub rlSetUniformPtr(ByVal programId As Long, ByVal locIndex As Long, ByVal valuePtr As Long, ByVal uniformType As Long, ByVal count As Long)
```

- Shader selection is batch-global: `rlSetShader` flushes the pending batch when the program changes, so a batch never mixes object-space (shader) and clip-space (fixed) vertices.
- `id = 0` is the fixed pipeline; `rlGetShaderIdDefault()` returns 0. `rlDisableShader` is `rlSetShader 0`.
- `rlSetShader` copies `locs` (indexed by raylib's `SHADER_LOC_*` order, length <= `RL_MAX_SHADER_LOCATIONS`); `rlEnableShader` selects a program without locations (the caller must upload the uniforms).
- `rlSetUniform*` bind the program, set the value and unbind, so they are safe to call between draws. `uniformType` uses the `ShaderUniformDataType` numbering. `location < 0` is a no-op.
- The batch uploads `mvp` (per draw call) itself; `rlSetUniformMatrix4fv` is for user uniforms.

# Shader batching rules

- With a shader active, `rlVertex2f/3f` store OBJECT-space coordinates (what you passed), and every draw call carries an MVP snapshot taken when it was created. `FlushBatchShader` binds the program, uploads each draw call's `mvp`, binds the draw texture (or the internal 1x1 white texture when none is set) and calls `glDrawArrays`.
- The fixed path is untouched: vertices are pre-transformed on the CPU, GL matrices are loaded at flush, identity afterwards.
- Matrix calls (`rlTranslatef`, `rlRotatef`, `rlScalef`, `rlPushMatrix`/`rlPopMatrix`, ...) mark the batch dirty; the next `rlBegin` starts a NEW draw call while a shader is active. That keeps per-object transforms correct without splitting a primitive. The fixed path keeps merging across matrix changes (vertices are already baked).
- Consequence: in shader mode, `DrawTexturePro` (push/translate/draw/pop per sprite) costs one draw call per sprite. Batch your static geometry to avoid it.
- `rlDrawVertexArray*` honour the active shader: the program is bound, `mvp` is uploaded from the current CPU matrices, and the GL matrices are left untouched.

# Notes / gotchas

- `rlglInit` resets matrices, colors, current texture and the matrix stack. It is called once from `InitWindow`, after the OpenGL context exists. There is no lazy/defensive re-init on the entry points: the state exists from then on.
- Raw draws (`rlDrawVertexArray*`) do NOT flush. Flush the batch BEFORE binding your own texture and setting the client arrays, or the later flush would clobber them. See `DrawMap` in `modGame` for the pattern.
- `rlPushMatrix`/`rlPopMatrix` keep a SEPARATE stack per matrix mode, exactly like `glPushMatrix`/`glPopMatrix`. `BeginMode2D` pushes both projection and modelview and pops both in `EndMode2D`.
- Only the first internal buffer is drawn; `rlLoadRenderBatch`'s `numBuffers` is accepted but effectively one buffer is used (drawn and reused when full).
- There is no perspective divide: vertices are transformed with `(modelview * projection)` and used as-is, exactly like raylib's rlgl 2D path. This is correct for orthographic projections (the intended use here).
- Normals are recorded per vertex and uploaded via `glNormalPointer`; they only matter if you enable `GL_LIGHTING` yourself.
- VB6 gotcha: do NOT name a local variable `len` - the compiler rejects it with a misleading "Syntax error" reported at the first line of the procedure.
