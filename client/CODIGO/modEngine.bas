Attribute VB_Name = "modEngine"
Option Explicit

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' This program is free software; you can redistribute it and/or modify
' it under the terms of the Affero General Public License;
' either version 1 of the License, or any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' Affero General Public License for more details.
'
' You should have received a copy of the Affero General Public License
' along with this program; if not, you can find it at http://www.affero.org/oagpl.html
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' / Codeado en un 90% por Emanuel Matías 'Dunkan' D'Urso.
' / Motor gráfico de Boskorcha AO

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Public Type TEXTURE_STATISTICS
    Texture As Direct3DTexture8
    TextureWidth As Integer
    TextureHeight As Integer
End Type

Public polygonCount(1) As Single

Public Wireframe As Boolean

'Quad Draw
Public indexList(0 To 5)    As Integer
Public ibQuad               As DxVBLibA.Direct3DIndexBuffer8
Public vbQuadIdx            As DxVBLibA.Direct3DVertexBuffer8
Public temp_verts(3)        As TLVERTEX

Public ScreenDelayX     As Single
Public ScreenDelayY     As Single
Public Movement_Speed   As Single
Public xLightPos        As Long
Public yLightPos        As Long

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Private Const GrhFogata As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X   As Integer
    Y   As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData

    sX          As Integer
    sY          As Integer
    
    FileNum     As Long
    
    pixelWidth  As Integer
    pixelHeight As Integer
    
    TileWidth   As Single
    TileHeight  As Single
    
    numFrames   As Integer
    Frames()    As Long
    Speed       As Single
    
End Type

Public Type Grh

    GrhIndex        As Integer
    FrameCounter    As Single
    Speed           As Single
    Started         As Byte
    Loops           As Integer
    
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Apariencia del personaje
Public Type Char
    AnimTime    As Byte
    InMoviment  As Boolean
    EsNPC       As Boolean
    Active      As Byte
    Heading     As E_Heading
    Pos         As Position
    timeInvi    As Integer
    iHead       As Integer
    iBody       As Integer
    Body        As BodyData
    Head        As HeadData
    Casco       As HeadData
    Arma        As WeaponAnimData
    Escudo      As ShieldAnimData
    UsandoArma  As Boolean
    
    Aura As Integer
    
    fX      As Grh
    FxIndex As Integer
    
    Criminal As Byte
    Atacable As Boolean
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving      As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    Pie         As Boolean
    Muerto      As Boolean
    Invisible   As Boolean
    Priv        As Byte
End Type

'Info de un objeto
Public Type Obj
    objIndex As Integer
    amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    ' / Models
    Graphic(1 To 4) As Grh
    
    ' / Index's
    CharIndex       As Integer
    NPCIndex        As Integer
    
    ' / Objetos
    OBJGrh          As Grh
    OBJInfo         As Obj
    
    ' / Terreno
    Offset(3)       As Long
    light_value(3)  As Long
    
    ' / Tile Sets
    WaterEffect     As Byte
    
    ' / Triggers
    TileExit        As WorldPos
    Blocked         As Byte
    Trigger         As Integer
    isEscalera      As WorldPos
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

Public Type TLVERTEX
    X           As Single
    Y           As Single
    Z           As Single
    RHW         As Single
    Color       As Long
    Specular    As Long
    TU          As Single
    TV          As Single
End Type

Public IniPath As String
Public MapPath As String

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap       As Integer 'Mapa actual
Public UserIndex    As Integer
Public UserMoving   As Byte
Public UserBody     As Integer
Public UserHead     As Integer
Public UserPos      As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve

Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Single
Public FramesPerSecCounter As Single
Public fpsLastCheck As Single

'Tamaño del la vista en Tiles
Private WindowTileWidth     As Integer
Private WindowTileHeight    As Integer

Private HalfWindowTileWidth     As Integer
Private HalfWindowTileHeight    As Integer

'Offset del desde 0,0 del main view
Private MainViewTop     As Integer
Private MainViewLeft    As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight  As Integer
Public TilePixelWidth   As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime As Single
Dim timerTicksPerFrame As Single
Dim engineBaseSpeed As Single

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth   As Integer
Private MainViewHeight  As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData()            As GrhData 'Guarda todos los grh
Public BodyData()           As BodyData
Public HeadData()           As HeadData
Public FxData()             As tIndiceFx
Public WeaponAnimData()     As WeaponAnimData
Public ShieldAnimData()     As ShieldAnimData
Public CascoAnimData()      As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData()    As MapBlock ' Mapa
Public MapInfo      As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public bTechoAB     As Byte
Public brstTick     As Long

Public charlist(1 To 10000) As Char

' Used by GetTextExtentPoint32
Private Type Size
    cX As Long
    cY As Long
End Type

Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\Init\Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\Init\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open App.path & "\Init\Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
End Sub

Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    N = FreeFile()
    Open App.path & "\Init\Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub

Sub CargarTips()
    Dim N As Integer
    Dim i As Long
    Dim NumTips As Integer
    
    N = FreeFile
    Open App.path & "\Init\Tips.ayu" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumTips
    
    'Resize array
    ReDim Tips(1 To NumTips) As String * 255
    
    For i = 1 To NumTips
        Get #N, , Tips(i)
    Next i
    
    Close #N
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************

On Error GoTo errHandle

    If MainScreenRect.Right - ScreenWidth <> 0 Then
        tX = UserPos.X + viewPortX \ TilePixelWidth - Round(MainScreenRect.Right / 32, 0) \ 2 + IIf((MainScreenRect.Right - ScreenWidth) > 0, 1, 0)
        tY = UserPos.Y + viewPortY \ TilePixelHeight - Round(MainScreenRect.bottom / 32, 0) \ 2
    Else
        tX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
        tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
    End If
    
Exit Sub

errHandle:
    
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        .EsNPC = (Arma = -1)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .Active = 1
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        .Active = 0
        .Criminal = 0
        .Atacable = False
        .FxIndex = 0
        .Invisible = False
        
#If SeguridadAlkon Then
        Call MI(CualMI).ResetInvisible(CharIndex)
#End If
        
        .Moving = 0
        .Muerto = False
        .Nombre = ""
        .Pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    charlist(CharIndex).Active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).numFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).numFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addX As Integer
    Dim addY As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    Dim CharHeadPos As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addY = -1
                CharHeadPos = NORTH
                
            Case E_Heading.EAST
                addX = 1
                CharHeadPos = EAST
        
            Case E_Heading.SOUTH
                addY = 1
                CharHeadPos = SOUTH
            
            Case E_Heading.WEST
                addX = -1
                CharHeadPos = WEST
                
                
        End Select
        
        nX = X + addX
        nY = Y + addY
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)

        .Moving = 1
        .Heading = CharHeadPos
        
        .scrollDirectionX = addX
        .scrollDirectionY = addY
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        If CharIndex <> UserCharIndex Then
            Call EraseChar(CharIndex)
        End If
    End If
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
    
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With charlist(CharIndex)
            If Not .Muerto And EstaPCarea(CharIndex) And (.Priv = 0 Or .Priv > 5) Then
                .Pie = Not .Pie
                
                If .Pie Then
                    Call Audio.PlayWave(SND_PASOS1, .Pos.X, .Pos.Y)
                Else
                    Call Audio.PlayWave(SND_PASOS2, .Pos.X, .Pos.Y)
                End If
            End If
        End With
    Else
' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Audio.PlayWave(SND_NAVEGANDO, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)
    End If
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addX = nX - X
        addY = nY - Y
        
        If Sgn(addX) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addX) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addY) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addY) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.Grande Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        
        Case E_Heading.NORTH
            Y = -1
            
        Case E_Heading.EAST
            X = 1
            
        Case E_Heading.SOUTH
            Y = 1
            
        Case E_Heading.WEST
            X = -1
            
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
    End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).OBJGrh.GrhIndex = GrhFogata Then
                    location.X = j
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim LoopC As Long
    Dim Dale As Boolean
    
    LoopC = 1
    Do While charlist(LoopC).Active And Dale
        LoopC = LoopC + 1
        Dale = (LoopC <= UBound(charlist))
    Loop
    
    NextOpenChar = LoopC
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Open IniPath & "Graficos.ind" For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(0 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        
        With GrhData(Grh)
            'Get number of frames
            Get handle, , .numFrames
            If .numFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(Grh).numFrames)
            
            If .numFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .numFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get handle, , .Speed
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
    Wend
    
    Close handle
    
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(X, Y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).Priv > 0 And charlist(UserCharIndex).Priv < 6 Then
                    If charlist(UserCharIndex).Invisible = True Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Private Sub DDrawGrhtoSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByRef light_value() As Long, Optional ByVal PosX As Byte, Optional ByVal PosY As Byte)
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
On Error GoTo error
        
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).numFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).numFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).numFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.bottom = SourceRect.Top + .pixelHeight
            
        'Draw
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, light_value, , , , , PosX, PosY)
    End With
Exit Sub

error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.number & " ] Error"
        End
    End If
End Sub

Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, Optional ByVal byteAlpha As Byte = 255, _
                                Optional ByVal R As Byte = 255, _
                                Optional ByVal G As Byte = 255, _
                                Optional ByVal B As Byte = 255)
                                
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Texture_PNG_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, , byteAlpha, R, G, B)
    End With
End Sub

Sub DDrawTransGrhtoSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByRef light_value() As Long, Optional ByVal AlphaByte As Byte = 255, Optional ByVal Angle As Single, Optional ByVal Shadow As Boolean, Optional ByVal PosX As Byte, Optional ByVal PosY As Byte)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
'On Error GoTo error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).numFrames / Grh.Speed) * Movement_Speed
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).numFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).numFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If

    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, light_value(), AlphaByte, Angle, Shadow, CurrentGrhIndex, PosX, PosY)
    End With
    
Exit Sub

error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.number & " ] Error"
        End
    End If
End Sub

Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
    Dim BMHeader As BITMAPFILEHEADER
    Dim BINFOHeader As BITMAPINFOHEADER
    
    Open BmpFile For Binary Access Read As #1
    
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    
    Close #1
    
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight
End Function


Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, ByVal srchdc As Long, ByRef SourceRect As RECT, ByRef destRect As RECT, ByVal TransparentColor)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 12/22/2009
'This method is SLOW... Don't use in a loop if you care about
'speed!
'*************************************************************
    Dim Color As Long
    Dim X As Long
    Dim Y As Long
    
    For X = SourceRect.Left To SourceRect.Right
        For Y = SourceRect.Top To SourceRect.bottom
            Color = GetPixel(srchdc, X, Y)
            
            If Color <> TransparentColor Then
                Call SetPixel(dsthdc, destRect.Left + (X - SourceRect.Left), destRect.Top + (Y - SourceRect.Top), Color)
            End If
        Next Y
    Next X
End Sub

Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, ByRef Picture As StdPicture, ByVal X1 As Single, ByVal Y1 As Single, Optional Width1, Optional Height1, Optional X2, Optional Y2, Optional Width2, Optional Height2)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 12/28/2009
'Draw Picture in the PictureBox
'*************************************************************

Call PictureBox.PaintPicture(Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2)
End Sub


Sub RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y           As Long     'Keeps track of where on map we are
    Dim X           As Long     'Keeps track of where on map we are
    Dim screenX     As Integer  'Keeps track of where to place tile on screen
    Dim screenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    Dim j As Long
    Dim Angle As Single
        
    'Figure out Ends and Starts of screen
    ScreenMinY = TileY - HalfWindowTileHeight
    ScreenMaxY = TileY + HalfWindowTileHeight
    ScreenMinX = TileX - HalfWindowTileWidth
    ScreenMaxX = TileX + HalfWindowTileWidth
    
    minY = ScreenMinY - TileBufferSize
    maxY = ScreenMaxY + TileBufferSize
    minX = ScreenMinX - TileBufferSize
    maxX = ScreenMaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If ScreenMinY > YMinMapSize Then
        ScreenMinY = ScreenMinY - 1
    Else
        ScreenMinY = 1
        screenY = 1
    End If
    
    If ScreenMaxY < YMaxMapSize Then ScreenMaxY = ScreenMaxY + 1
    
    If ScreenMinX > XMinMapSize Then
        ScreenMinX = ScreenMinX - 1
    Else
        ScreenMinX = 1
        screenX = 1
    End If
    
    If ScreenMaxX < XMaxMapSize Then ScreenMaxX = ScreenMaxX + 1
    
    'Draw floor layer
    For Y = ScreenMinY To ScreenMaxY
        For X = ScreenMinX To ScreenMaxX
            
                'Layer 1 **********************************
                    Call Engine_Render_Layer1(X, Y, screenX, screenY, PixelOffsetX, PixelOffsetY)
                    'Call DDrawGrhtoSurface(MapData(X, Y).Graphic(1), _
                (screenX - 1) * TilePixelWidth + PixelOffsetX, _
                (screenY - 1) * TilePixelHeight + PixelOffsetY, _
                0, 1, MapData(X, Y).light_value(), X, Y)
                '******************************************
                
            screenX = screenX + 1
        Next X
        
        'Reset ScreenX to original value and increment ScreenY
        screenX = screenX - X + ScreenMinX
        screenY = screenY + 1
    Next Y
    
    'Draw Transparent Layers
    screenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        screenX = minXOffset - TileBufferSize
        For X = minX To maxX
        
            PixelOffsetXTemp = screenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = screenY * TilePixelHeight + PixelOffsetY
            
            With MapData(X, Y)
                                
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call Engine_Char_Water(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                End If
                '*************************************************

            End With
            
            screenX = screenX + 1
        Next X
        screenY = screenY + 1
    Next Y
    
    'Draw floor layer 2
    screenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        screenX = minXOffset - TileBufferSize
        For X = minX To maxX
        
            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(2), _
                        screenX * TilePixelWidth + PixelOffsetX, _
                        screenY * TilePixelHeight + PixelOffsetY, _
                        1, 1, MapData(X, Y).light_value())
            End If
            '******************************************
            
            screenX = screenX + 1
        Next X
        screenY = screenY + 1
    Next Y
        
    'Draw Transparent Layers
    screenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        screenX = minXOffset - TileBufferSize
        For X = minX To maxX
        
            PixelOffsetXTemp = screenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = screenY * TilePixelHeight + PixelOffsetY
            
            With MapData(X, Y)
            
                'Object Layer **********************************
                If .OBJGrh.GrhIndex <> 0 Then
                            Call DDrawTransGrhtoSurface(.OBJGrh, _
                                PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value())
                End If
                '***********************************************
                                
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call Engine_Char_Render(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp - MapData(X, Y).Offset(0), MapData(X, Y).light_value())
                End If
                '*************************************************
                
                
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value())
                End If
                '************************************************
                
            End With
            
            screenX = screenX + 1
        Next X
        screenY = screenY + 1
    Next Y
        
    '************** Projectiles **************
    'Loop to do drawing
    If LastProjectile > 0 Then
        For j = 1 To LastProjectile
            If ProjectileList(j).Grh.GrhIndex Then
            
                'Update the position
                Angle = DegreeToRadian * Engine_GetAngle(ProjectileList(j).X, ProjectileList(j).Y, ProjectileList(j).tX, ProjectileList(j).tY)
                'Angle = Grados2Radianes(GetAngle2Points(ProjectileList(j).X, ProjectileList(j).Y, ProjectileList(j).tX, ProjectileList(j).tY))
                
                ProjectileList(j).X = ProjectileList(j).X + (Sin(Angle) * timerElapsedTime * 0.8)
                ProjectileList(j).Y = ProjectileList(j).Y - (Cos(Angle) * timerElapsedTime * 0.8)
                         
                'Update the rotation
                If ProjectileList(j).RotateSpeed > 0 Then
                    ProjectileList(j).Rotate = ProjectileList(j).Rotate + (ProjectileList(j).RotateSpeed * timerElapsedTime * 0.01)
                    Do While ProjectileList(j).Rotate > 360
                        ProjectileList(j).Rotate = ProjectileList(j).Rotate - 360
                    Loop
                End If

                'Draw if within range
                X = ((-minX - 1) * 32) + ProjectileList(j).X + PixelOffsetX - 180
                Y = ((-minY - 1) * 32) + ProjectileList(j).Y + PixelOffsetY - 180
                
                If Y >= -32 Then
                    If Y <= (ScreenHeight + 32) Then
                        If X >= -32 Then
                            If X <= (ScreenWidth + 32) Then
                                If ProjectileList(j).Rotate = 0 Then
                                    Call DDrawTransGrhtoSurface(ProjectileList(j).Grh, X, Y, 0, 1, DefaultColor(), , ProjectileList(j).Rotate + 128)
                                Else
                                    Call DDrawTransGrhtoSurface(ProjectileList(j).Grh, X, Y, 0, 1, DefaultColor(), , ProjectileList(j).Rotate + 128)
                                End If
                            End If
                        End If
                    End If
                End If
                
            End If
        Next j
        
        'Check if it is close enough to the target to remove
        For j = 1 To LastProjectile
            If ProjectileList(j).Grh.GrhIndex Then
                If Abs(ProjectileList(j).X - ProjectileList(j).tX) <= 20 Then
                    If Abs(ProjectileList(j).Y - ProjectileList(j).tY) <= 20 Then
                        Engine_Projectile_Erase j
                    End If
                End If
            End If
        Next j
        
    End If
    
    '************** Damage text **************
    'Loop to do drawing
    For j = 1 To LastDamage
        If DamageList(j).Counter > 0 Then
            DamageList(j).Counter = DamageList(j).Counter - timerElapsedTime
            X = (((DamageList(j).Pos.X - minX) - 1) * TilePixelWidth) + PixelOffsetX - 154
            Y = (DamageList(j).Counter * 0.02) + (((DamageList(j).Pos.Y - minY) - 1) * TilePixelHeight) + PixelOffsetY - 184
            If Y >= -32 Then
                If Y <= (ScreenHeight + 32) Then
                    If X >= -32 Then
                        If X <= (ScreenWidth + 32) Then
                            Engine_RenderText X, Y, DamageList(j).value, D3DColorXRGB(DamageList(j).R, DamageList(j).G, DamageList(j).B), 10 + (DamageList(j).Counter * 0.09)
                        End If
                    End If
                End If
            End If
        End If
    Next j
            
    'Seperate loop to remove the unused - I dont like removing while drawing
    For j = 1 To LastDamage
        If DamageList(j).Width Then
            If DamageList(j).Counter <= 0 Then Engine_Damage_Erase j
        End If
    Next j

 If Not bTecho Then
    If bTechoAB > 0 Then
        'Draw blocked tiles and grid
        screenY = minYOffset - TileBufferSize
        For Y = minY To maxY
            screenX = minXOffset - TileBufferSize
            For X = minX To maxX
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).GrhIndex Then
                    'Draw
                    Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                        screenX * TilePixelWidth + PixelOffsetX, _
                        screenY * TilePixelHeight + PixelOffsetY, _
                        1, 1, MapData(X, Y).light_value(), bTechoAB)
                End If
                '**********************************
                screenX = screenX + 1
            Next X
            screenY = screenY + 1
        Next Y
    End If

    
End If
End Sub

Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
'    If bLluvia(UserMap) = 1 Then
'        If bRain Then
'            If bTecho Then
'                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
'                    If RainBufferIndex Then _
'                        Call Audio.StopWave(RainBufferIndex)
'                    RainBufferIndex = Audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
'                    frmMain.IsPlaying = PlayLoop.plLluviain
'                End If
'            Else
'                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
'                    If RainBufferIndex Then _
'                        Call Audio.StopWave(RainBufferIndex)
'                    RainBufferIndex = Audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
'                    frmMain.IsPlaying = PlayLoop.plLluviaout
'                End If
'            End If
'        End If
'    End If
    
    DoFogataFx
    
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.Y <= Y
    End If
End Function


Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
        
    IniPath = App.path & "\Init\"
    
    Movement_Speed = 1
    
    bTechoAB = 255
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    
    'Set FPS value to 60 for startup
    FPS = 101
    FramesPerSecCounter = 101
    
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    'Set the dest rect
    With MainScreenRect
        .bottom = frmMain.MainViewPic.ScaleHeight
        .Right = frmMain.MainViewPic.ScaleWidth
    End With
    
On Error GoTo 0
    
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    
    '// Index Buffer
    indexList(0) = 0: indexList(1) = 1: indexList(2) = 2
    indexList(3) = 3: indexList(4) = 4: indexList(5) = 5
    
    Set ibQuad = DirectDevice.CreateIndexBuffer(Len(indexList(0)) * 4, 0, D3DFMT_INDEX32, D3DPOOL_SYSTEMMEM)
    D3DIndexBuffer8SetData ibQuad, 0, Len(indexList(0)) * 4, 0, indexList(0)
    
    Set vbQuadIdx = DirectDevice.CreateVertexBuffer(Len(temp_verts(0)) * 4, 0, FVF, D3DPOOL_SYSTEMMEM)

    ' Initialize Font's
    Call Engine_Init_FontTextures
    Call Engine_Init_FontSettings
    
    
    ' Initialize Color's
    Call Engine_InitColor

    Call SurfaceDB.Initialize(DirectD3D8, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
    
    InitTileEngine = True

End Function


Sub Engine_Show_NextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
    
    If EngineRun Then
        
        Engine_BeginScene ' Iniciamos
             
            'Update mouse position within view area
            Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
            
            '****** Update screen ******
            If UserCiego Then
                DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            Else
                Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX - Engine_Zoom_Offset(1), OffsetCounterY - Engine_Zoom_Offset(0))
            End If
            
            Call Dialogos.Render
            
            Call DialogosClanes.Draw
            
            ' // - FPS
            Engine_FPS_Update
        
        Engine_EndScene MainScreenRect, 0
        
        ' // - Water Effect
        Dim waterHeight As Integer: waterHeight = 4

        Static polygon_one_down As Single
        
            If polygon_one_down = 0 Then
                polygonCount(0) = polygonCount(0) + (4 * 0.042)
                If polygonCount(0) >= waterHeight Then
                    polygonCount(0) = waterHeight
                    polygon_one_down = 1
                End If
            Else
                polygonCount(0) = polygonCount(0) - (4 * 0.042)
                If polygonCount(0) <= -waterHeight Then
                    polygonCount(0) = -waterHeight
                    polygon_one_down = 0
                End If
            End If
                  
            polygonCount(1) = polygonCount(0)
           
            If polygon_one_down = 0 Then
                polygonCount(1) = polygonCount(1) + (waterHeight * 0.5)
                If polygonCount(1) >= waterHeight Then polygonCount(1) = waterHeight - (polygonCount(1) - waterHeight)
            Else
                polygonCount(1) = polygonCount(1) - (waterHeight * 0.5)
                If polygonCount(1) <= -waterHeight Then polygonCount(1) = -waterHeight + Abs(polygonCount(1) + waterHeight)
            End If
           
            polygonCount(1) = -polygonCount(1)
        ' // - Water Effect
        
        If GetTickCount() - AmbientLastCheck >= 1000 Then
            Ambient_Check
            AmbientLastCheck = GetTickCount()
        End If
        
        If modDx8_ambient.Fade Then Ambient_Fade
        
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
        
    End If
    
End Sub


Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Private Sub Engine_Char_Render(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByRef light_value() As Long)

' / Author: Juan Martín Sotuyo Dodero (Maraxus)
' / Organizado y optimizado por Dunkansdk

    Dim Moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim Color As Long

    With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + 8 * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                Moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + 8 * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                Moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not Moved Then
            'Stop animations
            If .AnimTime = 0 Then
            
                .Body.Walk(.Heading).Started = 0
                .Body.Walk(.Heading).FrameCounter = 1
                
            If Not .InMoviment Then
            
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
            
                .Escudo.ShieldWalk(.Heading).Started = 0
                .Escudo.ShieldWalk(.Heading).FrameCounter = 1
                
            End If
            
                .Moving = False
                
            Else
                .AnimTime = .AnimTime - 1
            End If
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        'NPC's names.
        'Diferenciamos npcs de usuarios.
        If .EsNPC Then
           'Render name.
           Dim M_SIZE_TAG   As Byte
           M_SIZE_TAG = Engine_GetTextWidth(cfonts(1), .Nombre) * 0.5
           Call Engine_RenderText(PixelOffsetX - M_SIZE_TAG + 15, PixelOffsetY + 30, .Nombre, -6908266, 170)
        End If
        
        If .Head.Head(.Heading).GrhIndex Then
            If Not .Invisible Then
            
                Movement_Speed = 0.5
                
                'Draw Mini-Shadow
                'Call DDrawSurfacePNG(HDGraphic.SHADOW_HD, PixelOffsetX + 16, PixelOffsetY + 22, 1, 150)
               
                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, light_value())
                 
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, light_value())
    
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, 0, light_value())
                     
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, light_value())
                     
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, light_value())

                'Draw name over head
                If LenB(.Nombre) > 0 Then
                
                    If Nombres Then
                        Pos = getTagPosition(.Nombre)
                        
                        If .Priv = 0 Then
                            If .Atacable Then
                                Color = D3DColorXRGB(ColoresPJ(48).R, ColoresPJ(48).G, ColoresPJ(48).B)
                            Else
                                If .Criminal Then
                                    Color = D3DColorXRGB(ColoresPJ(50).R, ColoresPJ(50).G, ColoresPJ(50).B)
                                Else
                                    Color = D3DColorXRGB(ColoresPJ(49).R, ColoresPJ(49).G, ColoresPJ(49).B)
                                End If
                            End If
                        Else
                            Color = D3DColorXRGB(ColoresPJ(.Priv).R, ColoresPJ(.Priv).G, ColoresPJ(.Priv).B)
                        End If
                            
                        Dim MEDIUM_SIZE_TAG As Byte
                        
                            
                        'Nick
                        line = Left$(.Nombre, Pos - 2)
                        MEDIUM_SIZE_TAG = Engine_GetTextWidth(cfonts(1), line) * 0.5
                        
                        Call Engine_RenderText(PixelOffsetX - MEDIUM_SIZE_TAG + 15, PixelOffsetY + 30, line, Color, 200)
                                
                        'Clan
                        line = mid$(.Nombre, Pos)
                        MEDIUM_SIZE_TAG = Engine_GetTextWidth(cfonts(1), line) * 0.5
                        
                        Call Engine_RenderText(PixelOffsetX - MEDIUM_SIZE_TAG + 15, PixelOffsetY + 45, line, Color, 220)

                    End If
                End If
            End If
        Else
            
            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then _
                Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, light_value())
        
        End If
        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex)   '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        
        Movement_Speed = 1
        
        'Draw FX
        If .FxIndex <> 0 Then

        Call DDrawTransGrhtoSurface(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1, AlphaColor(), 150)
            
            'Check if animation is over
            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
    End With
End Sub

Private Sub Engine_Char_Water(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

On Error Resume Next

    Dim Moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim Color As Long

    With charlist(CharIndex)

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
            
        If .Head.Head(.Heading).GrhIndex Then
            If Not .Invisible Then
            
                Movement_Speed = 0.5
                                 
                If MapData(.Pos.X, .Pos.Y + 2).WaterEffect = 1 Then
                
                     If .Body.Walk(.Heading).GrhIndex Then _
                         Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 0, 0, AlphaColor(), 100, 359, , .Pos.X, .Pos.Y + 2)

                     If .Head.Head(.Heading).GrhIndex Then _
                         Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + 7 + .Body.HeadOffset.Y, 0, 0, AlphaColor(), 100, 359, , .Pos.X, .Pos.Y + 2)

                     If .Casco.Head(.Heading).GrhIndex Then _
                         Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 0, 0, AlphaColor(), 100, 359, , .Pos.X, .Pos.Y + 2)

                     If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                         Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, AlphaColor(), 100, 359, , .Pos.X, .Pos.Y + 2)

                     If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                         Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, AlphaColor(), 100, 1, , .Pos.X, .Pos.Y + 2)
                End If
            End If
        Else
            If MapData(.Pos.X, .Pos.Y + 2).WaterEffect = 1 Then
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 0, 0, AlphaColor(), 100, 1, , .Pos.X, .Pos.Y + 2)
            End If
        End If

    End With
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If
    End With
End Sub

Public Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long)

' / Author: Dunkansdk

    ' * v0      * v1
    ' |        /|
    ' |      /  |
    ' |    /    |
    ' |  /      |
    ' |/        |
    ' * v2      * v3

    Dim x_Cor       As Single
    Dim y_Cor       As Single
    
    ' * - - - - - - - Vertice 0 -
    x_Cor = dest.Left
    y_Cor = dest.bottom
    
    '0 - Bottom left vertex
    If Textures_Width Or Textures_Height Then
        verts(0) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width + 0.001, (src.bottom) / Textures_Height)
    Else
        verts(0) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    ' * - - - - - - - Vertice 0 -
    
    
    ' * - - - - - - - Vertice 1 -
    x_Cor = dest.Left
    y_Cor = dest.Top
       
    '1 - Top left vertex
    If Textures_Width Or Textures_Height Then
        verts(1) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width + 0.001, src.Top / Textures_Height + 0.001)
    Else
        verts(1) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If
    ' * - - - - - - - Vertice 1 -
    

    ' * - - - - - - - Vertice 2 -
    x_Cor = dest.Right
    y_Cor = dest.bottom
    
    '2 - Bottom right vertex
    If Textures_Width Or Textures_Height Then
        verts(2) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right) / Textures_Width, (src.bottom) / Textures_Height)
    Else
        verts(2) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If
    ' * - - - - - - - Vertice 2 -
    
    
    ' * - - - - - - - Vertice 3 -
    x_Cor = dest.Right
    y_Cor = dest.Top
    
    '3 - Top right vertex
    If Textures_Width Or Textures_Height Then
        verts(3) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right) / Textures_Width, src.Top / Textures_Height + 0.001)
    Else
        verts(3) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If
    ' * - - - - - - - Vertice 3 -

End Sub

Public Function CreateVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal RHW As Single, ByVal Color As Long, ByVal Specular As Long, TU As Single, _
                                            ByVal TV As Single) As TLVERTEX

' / Author: Aaron Perkins
' / Last Modify Date: 10/07/2002

    CreateVertex.X = X
    CreateVertex.Y = Y
    CreateVertex.Z = Z
    CreateVertex.RHW = RHW
    CreateVertex.Color = Color
    CreateVertex.Specular = Specular
    CreateVertex.TU = TU
    CreateVertex.TV = TV
    
End Function

Public Sub Device_Textured_Render(ByVal X As Integer, ByVal Y As Integer, _
                                ByVal Texture As Direct3DTexture8, _
                                ByRef src_rect As RECT, _
                                ByRef light_value() As Long, _
                                Optional ByVal AlphaByte As Byte = 255, _
                                Optional ByVal Grados2D As Single, _
                                Optional ByVal Shadow As Boolean, _
                                Optional ByVal GrhIndex As Integer, _
                                Optional ByVal PosX As Byte, _
                                Optional ByVal PosY As Byte)
                                
    ' / - - - - - - - - - - - - - - - - - -
    ' / Author: Dunkansdk
    ' / Note: Dibuja las texturas
    ' / - - - - - - - - - - - - - - - - - -
    
    Dim dest_rect       As RECT
    Dim SRDesc          As D3DSURFACE_DESC
    Dim RadAngle        As Single
    Dim CenterX         As Single
    Dim CenterY         As Single
    Dim Index           As Integer
    Dim NewX            As Single
    Dim NewY            As Single
    Dim SinRad          As Single
    Dim CosRad          As Single
    Dim Width           As Single
    Dim Height          As Single
    Dim SrcBitmapWidth  As Long
    Dim SrcBitmapHeight As Long
    Dim rgb_list(3)     As Long
    
    rgb_list(0) = light_value(0)
    rgb_list(1) = light_value(1)
    rgb_list(2) = light_value(2)
    rgb_list(3) = light_value(3)
    
    If (rgb_list(0) = 0) Then rgb_list(0) = Base_Light
    If (rgb_list(1) = 0) Then rgb_list(1) = Base_Light
    If (rgb_list(2) = 0) Then rgb_list(2) = Base_Light
    If (rgb_list(3) = 0) Then rgb_list(3) = Base_Light
    
    Width = Int(GrhData(GrhIndex).TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
    Height = Int(GrhData(GrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
    
    With dest_rect
        .bottom = Y + (src_rect.bottom - src_rect.Top)
        .Left = X
        .Right = X + (src_rect.Right - src_rect.Left)
        .Top = Y
    End With

    ' Texture Settings - Dunkan
    Texture.GetLevelDesc 0, SRDesc
    SrcBitmapWidth = SRDesc.Width
    SrcBitmapHeight = SRDesc.Height
    
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, rgb_list(), SrcBitmapWidth, SrcBitmapHeight
    
    With DirectDevice
    
        .SetTexture 0, Texture
        
        'If Alpha Then
        '    .SetRenderState D3DRS_SRCBLEND, 3
        '    .SetRenderState D3DRS_DESTBLEND, 2
        'End If
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        
        ' ***************** Angulo *****************
        If Grados2D <> 0 And Grados2D <> 360 Then
            
            RadAngle = Grados2D * DegreeToRadian

            CenterX = X + (Width * 0.5)
            CenterY = Y + (Height * 0.5)
    
            SinRad = Sin(RadAngle)
            CosRad = Cos(RadAngle)
    
            For Index = 0 To 3
    
                NewX = CenterX + (temp_verts(Index).X - CenterX) * -CosRad - (temp_verts(Index).Y - CenterY) * -SinRad
                NewY = CenterY + (temp_verts(Index).Y - CenterY) * -CosRad + (temp_verts(Index).X - CenterX) * -SinRad
    
                temp_verts(Index).X = NewX
                temp_verts(Index).Y = NewY
    
            Next Index
    
        End If
        
        ' ***************** Sombras *****************
        If Shadow Then
        
            temp_verts(0).X = X + (Width * 0.5)
            temp_verts(0).Y = Y - (Height * 0.5)
            temp_verts(0).TU = (GrhData(GrhIndex).sX / SrcBitmapWidth)
            temp_verts(0).TV = (GrhData(GrhIndex).sY / SrcBitmapHeight)
            
            temp_verts(1).X = temp_verts(0).X + Width
            temp_verts(1).TU = ((GrhData(GrhIndex).sX + Width) / SrcBitmapWidth)
    
            temp_verts(2).X = X
            temp_verts(2).TU = (GrhData(GrhIndex).sX / SrcBitmapWidth)
    
            temp_verts(3).X = X + Width
            temp_verts(3).TU = (GrhData(GrhIndex).sX + 32 + 0) / SrcBitmapWidth
        
        End If
        
        ' ***************** Agua *****************
        If PosX > 0 Then
        If MapData(PosX, PosY).WaterEffect = 1 Then
        
        Dim POLYGON_IGNORE_TOP      As Byte
        Dim POLYGON_IGNORE_LOWER    As Byte
                
        POLYGON_IGNORE_LOWER = 0
        POLYGON_IGNORE_TOP = 0
                
        If MapData(PosX, PosY - 1).WaterEffect = 0 Then POLYGON_IGNORE_TOP = 1
        If MapData(PosX, PosY + 1).WaterEffect = 0 Then POLYGON_IGNORE_LOWER = 1
            
            If PosX Mod 2 = 0 Then
                
                If PosY Mod 2 = 0 Then
                    If POLYGON_IGNORE_TOP <> 1 Then
                        temp_verts(0).Y = temp_verts(0).Y - Val(polygonCount(1))
                        temp_verts(1).Y = temp_verts(1).Y + Val(polygonCount(1))
                    End If
                    If POLYGON_IGNORE_LOWER <> 1 Then
                        temp_verts(2).Y = temp_verts(2).Y + Val(polygonCount(0))
                        temp_verts(3).Y = temp_verts(3).Y - Val(polygonCount(0))
                    End If
                Else
                    If POLYGON_IGNORE_TOP <> 1 Then
                        temp_verts(0).Y = temp_verts(0).Y + Val(polygonCount(1))
                        temp_verts(1).Y = temp_verts(1).Y - Val(polygonCount(1))
                    End If
                    If POLYGON_IGNORE_LOWER <> 1 Then
                        temp_verts(2).Y = temp_verts(2).Y - Val(polygonCount(0))
                        temp_verts(3).Y = temp_verts(3).Y + Val(polygonCount(0))
                    End If
                           
                End If
                       
            ElseIf PosX Mod 2 = 1 Then
                   
                If PosY Mod 2 = 0 Then
                    If POLYGON_IGNORE_TOP <> 1 Then
                        temp_verts(0).Y = temp_verts(0).Y + Val(polygonCount(0))
                        temp_verts(1).Y = temp_verts(1).Y - Val(polygonCount(0))
                    End If
                    If POLYGON_IGNORE_LOWER <> 1 Then
                        temp_verts(2).Y = temp_verts(2).Y - Val(polygonCount(1))
                        temp_verts(3).Y = temp_verts(3).Y + Val(polygonCount(1))
                    End If
                Else
                    If POLYGON_IGNORE_TOP <> 1 Then
                        temp_verts(0).Y = temp_verts(0).Y - Val(polygonCount(0))
                        temp_verts(1).Y = temp_verts(1).Y + Val(polygonCount(0))
                    End If
                           
                    If POLYGON_IGNORE_LOWER <> 1 Then
                        temp_verts(2).Y = temp_verts(2).Y + Val(polygonCount(1))
                        temp_verts(3).Y = temp_verts(3).Y - Val(polygonCount(1))
                    End If
                End If
                
            End If
            
        End If
        End If
        
        ' MEDIUM LOAD
        '.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
        
        ' FASTER LOAD - DUNKAN
        .DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, _
                    indexList(0), D3DFMT_INDEX16, _
                    temp_verts(0), Len(temp_verts(0))
                    
        'If Alpha Then
        '    .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        '    .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        'End If
    
    End With
    
End Sub

Public Sub Engine_Render_Connect()
Dim X As Long, Y As Long

    Engine_BeginScene

    For Y = 1 To 20
        For X = 1 To 26
            With MapData(X + 50, Y + 50)
                Call DDrawGrhtoSurface(.Graphic(1), _
                    (X - 1) * TilePixelWidth, _
                    (Y - 1) * TilePixelHeight, _
                    1, 1, DefaultColor(), X + 50, Y + 50)
                
                If .Graphic(2).GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(.Graphic(2), _
                            (X - 1) * TilePixelWidth, _
                            (Y - 1) * TilePixelHeight, _
                            1, 1, DefaultColor(), , , , X, Y)
                End If
            End With
        Next X
    Next Y
    
    For Y = 1 To 20
        For X = 1 To 26
            With MapData(X + 50, Y + 50)
                If .Graphic(3).GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(.Graphic(3), (X - 2) * TilePixelWidth, (Y - 2) * TilePixelHeight, 1, 1, DefaultColor(), , , , X, Y)
                End If
            End With
        Next X
    Next Y
    
    Call Engine_FPS_Update
    
    Call Engine_RenderText(10, 10, FPS & " FPS", D3DColorARGB(150, 255, 255, 255))
    
    timerElapsedTime = GetElapsedTime()
End Sub


Public Sub Engine_DX_End()

On Error Resume Next

    Set DirectD3D = Nothing
    
    Set DirectX = Nothing
    
    Set vbQuadIdx = Nothing
    Set ibQuad = Nothing
        
End Sub



