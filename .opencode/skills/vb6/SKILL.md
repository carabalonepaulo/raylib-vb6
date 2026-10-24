---
name: vb6
description: Use ALWAYS when a task involves the VB6 (Visual Basic 6) language or the vb6.exe IDE/compiler - writing, reviewing, or debugging .bas/.frm/.cls/.vbp code, calling Win32/OpenGL APIs via Declare, compiling with /make, or investigating compile/runtime errors. Collects the known VB6 pitfalls (high-bit hex literals, declaration order, operator precedence, ANSI encoding, DispCallFunc, misleading line errors).
---

# VB6 (Visual Basic 6) - practical guide

VB6 pitfalls that waste the most time. Read this before writing or debugging
anything in VB6.

## Golden rules

1. **Compile without `/nologo`.** When running `vb6.exe` use
   `VB6.EXE /make <project.vbp> /out <build.log>`. **Do NOT pass `/nologo`** - it
   breaks log generation/reading in this setup.
2. **A hex literal with the high bit set becomes a negative `Integer`.** `&H8000`
   to `&HFFFF` are read as `Integer` (16-bit, signed) and end up negative.
   `&H88BF` = -30529, `&H8D40` = -29376. **Always add the `&` suffix** to force
   `Long`: `&H88BF&`, `&H8892&`, `&H8D40&`. This holds even for
   `Const X As Long = &H88BF` - the `As Long` does NOT save the literal; write
   `Const X As Long = &H88BF&`.
3. **Module declarations come before the first procedure.** In a module,
   everything declarative at module scope - variables (`Dim`/`Private`/
   `Public`), `Const`, `Type`, `Enum`, `Declare`, `Option` - must appear BEFORE
   the first `Sub`/`Function`/`Property`. Putting a declaration after a
   procedure is a compile error.
4. **Watch operator precedence** (see table below). The most common mistakes are
   comparisons/concatenation binding tighter than expected.

## IDE / compilation

- Compile from the command line: `VB6.EXE /make <vbp> /out <log>` and read the
  log as **ANSI** (`Get-Content -Encoding Default` in PowerShell). VB6 writes the
  log in ANSI.
- VB6 **frequently reports the wrong file and line**. E.g. it points at
  `modRaylib.bas` line 810 when the real problem is a `Sub` that does not exist in
  another module. Do not blindly trust the file/line: look up the symbol mentioned.
- Do not compile inside the repository when the build overwrites a versioned
  executable (in this project: `Raylib.exe`). Copy to a temporary directory.
- `Attribute VB_Name = "module"` at the top of each file must match the name.
- The startup object lives in the `.vbp` (`Startup="Sub Main"`). An EXE with
  `Sub Main` terminates when `Main` returns.
- VB6 is case-insensitive in identifiers and rewrites the casing in the IDE.

## File encoding

- VB6 reads `.bas`/`.frm`/`.cls` as **ANSI** (system code page). Keep sources in
  pure ASCII; accented comments can turn into mojibake.
- Project convention: `src/*.bas` without accents, purely ASCII.
- When generating test/harness files, save them as ANSI with CRLF if the build
  expects it.

## Operator precedence (strongest to weakest)

```
^                          exponentiation
-  (unary)                 negation
*  /                       multiplication, division
\                          integer division
Mod                        remainder
+  -                       addition, subtraction
&                          string concatenation
=  <>  <  >  <=  >=        comparison (also Like, Is)
Not                        logical/bitwise negation
And                        logical/bitwise conjunction
Or                         disjunction
Xor  Eqv  Imp              remaining logical operators
```

Consequences that catch you off guard:

- `Not a = b`  =>  `Not (a = b)`   (comparison binds tighter than `Not`).
- `a & b = c`  =>  `(a & b) = c`   (concatenation binds tighter than comparison).
- `a And b Or c`  =>  `(a And b) Or c`.
- `1 + 2 & 3`  =>  `(1 + 2) & 3`  =  `"33"`.
- `-2 ^ 2`  =>  `-(2 ^ 2)`  =  `-4`.
- `And`/`Or`/`Not`/`Xor` are also **bitwise** on integers.
- In a `Variant`, `+` is ambiguous (addition or concatenation). Use `&` for strings.

## Language

- `Integer` is 16-bit signed (-32768..32767); `Long` is 32-bit; `Currency` is
  64-bit fixed-point (used for `QueryPerformanceCounter`). Literal suffixes:
  `%` Integer, `&` Long, `!` Single, `#` Double, `@` Currency, `$` String.
- Procedure parameters are **`ByRef` by default**. You cannot pass a literal or
  expression to a `ByRef` parameter.
- Arrays are 0-based by default (`Option Base`). `ReDim Preserve` only resizes
  the last dimension.
- `Set` is required to assign objects; `Let` (implicit) for values.
- `Option Explicit` must be the first statement (after `Attribute VB_Name`).
- Reserved names cannot be used as variables (e.g. `Len`, `Left`, `Right`,
  `Name`, `Date`, `Error`, `Print`). Using `len` as a local variable produces a
  misleading "Syntax error" pointed at the first line of the procedure.
- `Type` (UDT) only accepts fixed-size members for use with APIs/`Put`/`Get`;
  variable-length strings and objects are not allowed.
- Error handling: `On Error GoTo label` / `On Error Resume Next`, `Err.Number`,
  `Err.Description`, `Resume`/`Resume Next`/`Resume label`. An unhandled error
  takes down the application.
- A `Function` called as a statement (without `Call`) discards the return value -
  that is valid.

## API calls (Declare)

- `Declare Function Foo Lib "dll" Alias "Foo" (...) As Long`. Pass `As Any` for
  pointers; `StrPtr`/`VarPtr` to get addresses.
- `Declare` is also `ByRef` by default. To pass by value, use `ByVal`.
- `StrPtr(s)` = pointer to the string data (UTF-16/BSTR); `VarPtr(s)` = pointer to
  the descriptor. For ANSI C-strings, use a `Byte` array + `StrConv`.
- For `As Any` with a UDT, pass `ByRef` (not `ByVal`).

### OpenGL / WGL bindings (modOpenGL conventions)

When declaring GL/WGL functions, map the C types like this:

```
GLenum / GLbitfield / GLint / GLsizei / GLuint -> Long
GLboolean                                       -> Byte
GLfloat / GLclampf                              -> Single
GLdouble / GLclampd                             -> Double
GLvoid*                                         -> ByVal Long   (call: ByVal VarPtr(data(0)))
vectors (GLfloat*, GLint*, GLdouble*, ...)      -> ByRef <type> (call: data(0))
```

- `wglSwapIntervalEXT` (vsync) only takes effect when the current context was
  created with a double buffer, and it must be called with the context current
  (`wglMakeCurrent` already done).

## Function pointers (DispCallFunc)

VB6 has no function pointers. To call an address obtained from
`GetProcAddress`/`wglGetProcAddress`, use `DispCallFunc` (oleaut32):

- `cc = 4` (`CC_STDCALL`), `pvInstance = 0`, `oVft = address`.
- `prgvt` = array of `VARTYPE` (one per argument). For 32-bit use `VT_I4 = 3`.
- `prgpvarg` = array of **pointers to `VARIANT`** (not raw values!). Store each
  argument in a `Variant` with `vt = VT_I4` and pass `VarPtr(variant)`.
- `pvargResult` **must point to a valid `VARIANT`, even if the function returns
  `void`**. Passing `0` crashes.
- `CallWindowProc` does **not** work for calling arbitrary stdcall functions: it
  hangs. Use `DispCallFunc`.
- `wglGetProcAddress` returns `1`, `2`, `3`, or `-1` (not just `0`) when the
  function does not exist. Handle those values and fall back to
  `GetProcAddress(GetModuleHandle("opengl32.dll"), name)`.

Example of an stdcall call with up to 3 32-bit arguments:

```vb
Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As Long, ByVal prgpvarg As Long, ByVal pvargResult As Long) As Long

Private Function CallStd3(ByVal addr As Long, ByVal a1 As Long, ByVal a2 As Long, ByVal a3 As Long, ByVal nargs As Long) As Long
    Dim vt(0 To 2) As Integer
    Dim av(0 To 2) As Long
    Dim vv(0 To 2) As Variant
    Dim res As Variant
    Dim i As Long
    If nargs > 0 Then vv(0) = a1
    If nargs > 1 Then vv(1) = a2
    If nargs > 2 Then vv(2) = a3
    For i = 0 To nargs - 1
        vt(i) = 3            ' VT_I4
        av(i) = VarPtr(vv(i))
    Next i
    res = 0&
    DispCallFunc 0&, addr, 4&, 0, nargs, VarPtr(vt(0)), VarPtr(av(0)), VarPtr(res)
    CallStd3 = 0
End Function
```

## Testing without the IDE (harness)

- The useful pattern: copy the project to a temporary directory, replace the game
  module with a harness, compile with `/make`, and run the EXE.
- There is no game loop in the library anymore: the user owns `Main`. A harness
  `Main` calls `SetConfigFlags` (optional), `InitWindow frmMain, w, h, "title"`,
  then runs its own `Do While WindowShouldClose() = False ... BeginDrawing /
  EndDrawing ... Loop`, then `CloseWindow`. Nothing calls `Game_*` any more, so a
  harness is free to define whatever it needs (and it must NOT leave the form's
  `Form_QueryUnload` calling `CloseWindow`, which is now the teardown).
- Missing procedure errors still point at the WRONG file: VB6 reports
  "Sub or Function not defined" in the module that *calls* it, not where it is
  missing.
- Log to `App.Path & "\out.log"`. Write with `Open ... For Append` and **close
  each line** (or flush) - a crash discards the `Print #` buffer and you lose the
  clue about where it died.
- When investigating a crash, add incremental markers and narrow it down until you
  find the exact statement.
- `QueryPerformanceCounter`/`QueryPerformanceFrequency` use `Currency` (8 bytes)
  as the output parameter.
