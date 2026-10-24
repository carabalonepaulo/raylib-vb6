# AGENTS.md

## API raylib / rlgl (VB6)

Sempre que a tarefa envolver a API deste projeto — `src/modRaylib.bas`,
`src/modRLGL.bas`, ou qualquer uso, extensao ou correcao dela — consulte as duas
cheatsheets em Markdown antes de comecar a trabalhar:

- [cheatsheet.md](cheatsheet.md) — API de alto nivel (raylib): Core, 2D Camera,
  Render Texture, Mouse, Cursor, Keyboard, Shapes, Textures, Text (ImageFont),
  Images, Colors, Audio, Types, Enums.
- [cheatsheet_rlgl.md](cheatsheet_rlgl.md) — camada de batching / emulacao de
  OpenGL 1.1 (estilo rlgl).

### Manter as cheatsheets em sincronia

Ao adicionar, remover, renomear ou alterar qualquer funcao, tipo, enum ou
constante da API, atualize a secao correspondente na cheatsheet apropriada, na
mesma tarefa. As cheatsheets sao a referencia da API e nao devem ficar
desatualizadas em relacao ao codigo.

### Formato das cheatsheets

- Cada secao e um heading `# nome`.
- Todas as funcoes de uma secao ficam agrupadas dentro de um unico bloco de
  codigo delimitado por ```.
- Notas e "gotchas" ficam como texto normal (fora do bloco), logo apos a secao.

### Convencoes do projeto

- `src/*.bas` sao lidos pela VB6 como ANSI/28591: mantenha-os puramente ASCII,
  com comentarios sem acentos.
- Texto na tela usa somente `ImageFont` (bitmap font); nao existe TTF nem
  carregamento de fonte.
- A camada de baixo nivel usa o prefixo `rl` e apenas OpenGL 1.1.
- Nao compile o projeto dentro do repositorio (evita sobrescrever `Raylib.exe`);
  use um diretorio temporario.
