# AGENTS.md

## VB6 language / IDE

Whenever a task involves the VB6 (Visual Basic 6) language or the `vb6.exe`
IDE/compiler - writing, reviewing, or debugging `.bas`/`.frm`/`.cls`/`.vbp`
code, calling Win32/OpenGL APIs via `Declare`, compiling with `/make`, or
investigating compile/runtime errors - load the local `vb6` skill
([.opencode/skills/vb6/SKILL.md](.opencode/skills/vb6/SKILL.md)) before starting.

## raylib / rlgl API (VB6)

Whenever a task involves this project's API - `src/modRaylib.bas`,
`src/modRLGL.bas`, or any use, extension, or fix of it - consult the two
Markdown cheatsheets before starting work:

- [cheatsheet.md](cheatsheet.md) - high-level API (raylib): Core, 2D Camera,
  Render Texture, Mouse, Cursor, Keyboard, Shapes, Textures, Text (ImageFont),
  Images, Colors, Audio, Types, Enums.
- [cheatsheet_rlgl.md](cheatsheet_rlgl.md) - batching / OpenGL 1.1 emulation
  layer (rlgl-style).

### Keep the cheatsheets in sync

When adding, removing, renaming, or changing any function, type, enum, or
constant of the API, update the corresponding section in the appropriate
cheatsheet in the same task. The cheatsheets are the API reference and must not
fall out of date with the code.

### Cheatsheet format

- Each section is a `# name` heading.
- All functions of a section are grouped inside a single code block delimited by
  ```.
- Notes and "gotchas" are plain text (outside the block), right after the
  section.

### Project conventions

- `src/*.bas` are read by VB6 as ANSI/28591: keep them purely ASCII, with
  comments without accents.
- On-screen text uses only `ImageFont` (bitmap font); there is no TTF or font
  loading.
- The low-level layer uses the `rl` prefix and only OpenGL 1.1.
- Do not compile the project inside the repository (it would overwrite
  `Raylib.exe`); use a temporary directory.
