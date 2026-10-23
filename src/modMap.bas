Attribute VB_Name = "modMap"
Option Explicit

Public Type Tilemap
    Id As Integer
    Tileset As Texture2D
    TilesetColumns As Byte
    TileWidth As Byte
    TileHeight As Byte
    Layers As Byte
    Width As Integer
    Height As Integer
    Data() As Integer

    TotalVertices As Long
    Vertices() As Vector2
    TexCoords() As TexCoord2D
End Type

Public Sub DebugTilemap(ByRef map As Tilemap)
    Debug.Print "Map: " & map.Id
    Debug.Print "Tileset: id = " & map.Tileset.Id & ", size = " & map.Tileset.Width & "x" & map.Tileset.Height
    Debug.Print "Tileset Columns: " & map.TilesetColumns
    Debug.Print "Tile Size: " & map.TileWidth & "x" & map.TileHeight
    Debug.Print "Size: " & map.Width & "x" & map.Height
    Debug.Print "Layers: " & map.Layers
End Sub

Public Function LoadTilemap(ByVal Id As Integer) As Tilemap
    Dim basePath As String
    basePath = app.path & "\" & "assets\maps\" & Id

    Dim path As String
    path = basePath & ".ini"

    Dim tilesetPath As String
    tilesetPath = INI_ReadString("Meta", "Tileset", path)

    Dim map As Tilemap
    map.Id = Id
    map.Tileset = LoadTexture(tilesetPath)
    map.TilesetColumns = INI_ReadByte("Meta", "TilesetColumns", path)
    map.TileWidth = INI_ReadByte("Meta", "TileWidth", path)
    map.TileHeight = INI_ReadByte("Meta", "TileHeight", path)
    map.Layers = INI_ReadByte("Meta", "Layers", path)
    map.Width = INI_ReadInt("Meta", "Width", path)
    map.Height = INI_ReadInt("Meta", "Height", path)

    ReDim map.Data(map.Layers - 1, map.Width - 1, map.Height - 1)

    Dim l As Integer, y As Integer, x As Integer
    Dim fileNum As Integer
    Dim fileLine As String
    Dim values() As String

    For l = 0 To map.Layers - 1
        Dim csvPath As String
        csvPath = basePath & "_" & l & ".csv"

        fileNum = FreeFile
        Open csvPath For Input As #fileNum

        y = 0
        While Not EOF(fileNum) And y < map.Height
            Line Input #fileNum, fileLine
            fileLine = Trim$(fileLine)

            If Len(fileLine) > 0 Then
                values = Split(fileLine, ",")

                For x = LBound(values) To UBound(values)
                    If x < map.Width Then
                        map.Data(l, x, y) = val(values(x))
                    End If
                Next x
                y = y + 1
            End If
        Wend

        Close #fileNum
    Next l

    ' =========================================================================
    ' PRÉ-PROCESSAMENTO: Alinhado 100% com o DrawTexturePro original
    ' =========================================================================
    
    Dim texW As Single, texH As Single
    If map.Tileset.Id <> 0 Then
        texW = map.Tileset.Width
        texH = map.Tileset.Height
    Else
        texW = CLng(map.TilesetColumns) * CLng(map.TileWidth)
        texH = 256
    End If

    Dim totalTilesValidos As Long
    totalTilesValidos = 0
    
    Dim ll As Long, yy As Long, xx As Long
    For ll = 0 To CLng(map.Layers) - 1
        For yy = 0 To CLng(map.Height) - 1
            For xx = 0 To CLng(map.Width) - 1
                If map.Data(ll, xx, yy) <> -1 Then
                    totalTilesValidos = totalTilesValidos + 1
                End If
            Next xx
        Next yy
    Next ll
    
    map.TotalVertices = totalTilesValidos * 4
    
    If map.TotalVertices > 0 Then
        ReDim map.Vertices(map.TotalVertices - 1)
        ReDim map.TexCoords(map.TotalVertices - 1)
        
        Dim vIndex As Long
        vIndex = 0
        
        Dim tileId As Long, tileX As Long, tileY As Long
        Dim srcX As Single, srcY As Single
        Dim dstX As Single, dstY As Single
        Dim tw As Single, th As Single
        tw = map.TileWidth
        th = map.TileHeight
        
        Dim cols As Long
        cols = map.TilesetColumns
        
        Dim u1 As Single, v1 As Single, u2 As Single, v2 As Single
        
        For ll = 0 To CLng(map.Layers) - 1
            For yy = 0 To CLng(map.Height) - 1
                For xx = 0 To CLng(map.Width) - 1
                    tileId = map.Data(ll, xx, yy)
                    
                    If tileId <> -1 Then
                        tileX = tileId Mod cols
                        tileY = tileId \ cols
                        
                        ' Cálculo exato de src e dst do seu DrawTexturePro
                        srcX = tileX * tw
                        srcY = tileY * th
                        
                        dstX = xx * tw
                        dstY = yy * th
                        
                        ' UVs calculadas idênticas ao DrawTexturePro
                        u1 = srcX / texW
                        v1 = srcY / texH
                        u2 = (srcX + tw) / texW
                        v2 = (srcY + th) / texH
                        
                        ' Vértices mapeados nas posições absolutas de tela (equivalente ao glTranslate + vertices do glBegin)
                        ' Vértice 1: Top-Left
                        map.Vertices(vIndex).x = dstX:          map.Vertices(vIndex).y = dstY
                        map.TexCoords(vIndex).u = u1:         map.TexCoords(vIndex).v = v1
                        vIndex = vIndex + 1
                        
                        ' Vértice 2: Top-Right
                        map.Vertices(vIndex).x = dstX + tw:     map.Vertices(vIndex).y = dstY
                        map.TexCoords(vIndex).u = u2:         map.TexCoords(vIndex).v = v1
                        vIndex = vIndex + 1
                        
                        ' Vértice 3: Bottom-Right
                        map.Vertices(vIndex).x = dstX + tw:     map.Vertices(vIndex).y = dstY + th
                        map.TexCoords(vIndex).u = u2:         map.TexCoords(vIndex).v = v2
                        vIndex = vIndex + 1
                        
                        ' Vértice 4: Bottom-Left
                        map.Vertices(vIndex).x = dstX:          map.Vertices(vIndex).y = dstY + th
                        map.TexCoords(vIndex).u = u1:         map.TexCoords(vIndex).v = v2
                        vIndex = vIndex + 1
                    End If
                Next xx
            Next yy
        Next ll
    End If

    LoadTilemap = map
End Function






























