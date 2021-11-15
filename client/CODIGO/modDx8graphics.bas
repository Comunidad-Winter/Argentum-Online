Attribute VB_Name = "modDx8_graphics"
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

' / Módulo creado por Emanuel Matías D'Urso 'Dunkan'

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'Dunkansdk@ = DX8 Objects
Public DirectX      As New DirectX8
Public DirectD3D8   As D3DX8
Public DirectD3D    As Direct3D8
Public DirectDevice As Direct3DDevice8

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'Dunkansdk@ = Extras
Public Movement_Speed   As Single
Public DefaultColor(3)  As Long
Public ShadowColor(3)   As Long
Public AlphaColor(3)    As Long

'Dunkansdk@ = Rect's
Public MainScreenRect As RECT
Public ConnectScreenRect As RECT


Public Function Engine_DirectX8_Init()

' Author: Emanuel Matías 'Dunkan'
' Note: Inicia el DirectX

    Dim DispMode    As D3DDISPLAYMODE
    Dim D3DWindow   As D3DPRESENT_PARAMETERS
    Dim Caps8       As D3DCAPS8
    
    '// Initialize
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
    
On Local Error GoTo ErrHandler

    DispMode.format = D3DFMT_X8R8G8B8

    '// Mostrar
    DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    '// Window
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY  'Settings
        .BackBufferFormat = DispMode.format
        .BackBufferWidth = 800
        .BackBufferHeight = 600
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.MainViewPic.hWnd
    End With
    
    '// Crea el device
    'D3DCREATE_HARDWARE_VERTEXPROCESSING
    Set DirectDevice = DirectD3D.CreateDevice( _
                        D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                        frmMain.MainViewPic.hWnd, _
                        D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                        D3DWindow)
                        
    Engine_Render_States
    
ErrHandler:

    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If
    
    If Err Then
        MsgBox "No se puede iniciar DirectD3D. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If
    
    If DirectDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectDevice. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function
    End If
    
End Function

Private Sub Engine_Render_States()

' / Author: Dunkan
' / Note: Carga los estados del render.

    With DirectDevice
        
        ' // Shader
        .SetVertexShader FVF
        
        ' // States
        .SetRenderState D3DRS_LIGHTING, True
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

        ' // WireFrame (Desactivado)
        '.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
        
        '.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD

        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        ' // Filtros
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_NONE
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_NONE
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
        .SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_CURRENT

        .SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_MIRROR
        .SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_MIRROR
                
    End With

End Sub


Public Sub Engine_Render_Layer1(ByVal X As Long, ByVal Y As Long, _
                                ByVal screenX As Integer, ByVal screenY As Integer, _
                                ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

' / - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' / Author: Emanuel Matías 'Dunkan'
' / Note: Efectos de la capa 1, movimiento de polígonos y dibujado
' / - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Dim VertexArray(0 To 3) As TLVERTEX
Dim D3DTextures         As TEXTURE_STATISTICS
Dim SRDesc              As D3DSURFACE_DESC
Dim SrcWidth            As Integer
Dim Width               As Integer
Dim SrcHeight           As Integer
Dim Height              As Integer
Dim new_x               As Integer
Dim new_y               As Integer
Dim SrcBitmapWidth      As Long
Dim SrcBitmapHeight     As Long
            
If MapData(X, Y).Graphic(1).GrhIndex Then

    new_x = (screenX - 1) * 32 + PixelOffsetX
    new_y = (screenY - 1) * 32 + PixelOffsetY
           
    If MapData(X, Y).Graphic(1).Started = 1 Then
    
        MapData(X, Y).Graphic(1).FrameCounter = MapData(X, Y).Graphic(1).FrameCounter + ((timerElapsedTime * 0.1) * GrhData(MapData(X, Y).Graphic(1).GrhIndex).numFrames / MapData(X, Y).Graphic(1).Speed)
            
            If MapData(X, Y).Graphic(1).FrameCounter > GrhData(MapData(X, Y).Graphic(1).GrhIndex).numFrames Then
                MapData(X, Y).Graphic(1).FrameCounter = (MapData(X, Y).Graphic(1).FrameCounter Mod GrhData(MapData(X, Y).Graphic(1).GrhIndex).numFrames) + 1
            End If
            
    End If
                        
    Dim iGrhIndex   As Integer
    
    iGrhIndex = GrhData(MapData(X, Y).Graphic(1).GrhIndex).Frames(MapData(X, Y).Graphic(1).FrameCounter)
            
    With GrhData(iGrhIndex)
    
        Set D3DTextures.Texture = SurfaceDB.Surface(.FileNum) 'Cargamos la textura
                
        D3DTextures.Texture.GetLevelDesc 0, SRDesc ' Medimos las texturas
        
        D3DTextures.TextureWidth = SRDesc.Width
        D3DTextures.TextureHeight = SRDesc.Height
                
        SrcWidth = 32
        Width = 32
                   
        Height = 32
        SrcHeight = 32
                
        SrcBitmapWidth = D3DTextures.TextureWidth
        SrcBitmapHeight = D3DTextures.TextureHeight
               
        'Seteamos los RHW a 1
        VertexArray(0).RHW = 1
        VertexArray(1).RHW = 1
        VertexArray(2).RHW = 1
        VertexArray(3).RHW = 1
             
        'Find the left side of the rectangle
        VertexArray(0).X = new_x
        VertexArray(0).TU = (.sX / SrcBitmapWidth)
             
        'Find the top side of the rectangle
        VertexArray(0).Y = new_y
        VertexArray(0).TV = (.sY / SrcBitmapHeight)
               
        'Find the right side of the rectangle
        VertexArray(1).X = new_x + Width
        VertexArray(1).TU = (.sX + SrcWidth) / SrcBitmapWidth
             
        'These values will only equal each other when not a shadow
        VertexArray(2).X = VertexArray(0).X
        VertexArray(3).X = VertexArray(1).X
               
        'Find the bottom of the rectangle
        VertexArray(2).Y = new_y + Height
        VertexArray(2).TV = (.sY + SrcHeight) / SrcBitmapHeight
             
        'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
        VertexArray(1).Y = VertexArray(0).Y
        VertexArray(1).TV = VertexArray(0).TV
        VertexArray(2).TU = VertexArray(0).TU
        VertexArray(3).Y = VertexArray(2).Y
        VertexArray(3).TU = VertexArray(1).TU
        VertexArray(3).TV = VertexArray(2).TV
                            
        VertexArray(0).Y = VertexArray(0).Y - MapData(X, Y).Offset(0)
        VertexArray(1).Y = VertexArray(1).Y - MapData(X, Y).Offset(1)
        VertexArray(2).Y = VertexArray(2).Y - MapData(X, Y).Offset(2)
        VertexArray(3).Y = VertexArray(3).Y - MapData(X, Y).Offset(3)
        
        Static Polygon As Long
        
        For Polygon = 0 To 3
            VertexArray(Polygon).Color = D3DColorARGB(255, _
                        ColorActual.R - MapData(X, Y).Offset(Polygon), _
                        ColorActual.G - MapData(X, Y).Offset(Polygon), _
                        ColorActual.B - MapData(X, Y).Offset(Polygon))
        Next Polygon
        
        If HayAgua(X, Y) Then
        
        Dim POLYGON_IGNORE_TOP      As Byte
        Dim POLYGON_IGNORE_LOWER    As Byte
                
        POLYGON_IGNORE_LOWER = 0
        POLYGON_IGNORE_TOP = 0
                
        If HayAgua(X, Y - 1) = False Then POLYGON_IGNORE_TOP = 1
        If HayAgua(X, Y + 1) = False Then POLYGON_IGNORE_LOWER = 1
            
            If X Mod 2 = 0 Then
                
                If Y Mod 2 = 0 Then
                    If POLYGON_IGNORE_TOP <> 1 Then
                        VertexArray(0).Y = VertexArray(0).Y - Val(polygonCount(0))
                        VertexArray(1).Y = VertexArray(1).Y + Val(polygonCount(0))
                    End If
                    If POLYGON_IGNORE_LOWER <> 1 Then
                        VertexArray(2).Y = VertexArray(2).Y + Val(polygonCount(1))
                        VertexArray(3).Y = VertexArray(3).Y - Val(polygonCount(1))
                    End If
                Else
                    If POLYGON_IGNORE_TOP <> 1 Then
                        VertexArray(0).Y = VertexArray(0).Y + Val(polygonCount(1))
                        VertexArray(1).Y = VertexArray(1).Y - Val(polygonCount(1))
                    End If
                    If POLYGON_IGNORE_LOWER <> 1 Then
                        VertexArray(2).Y = VertexArray(2).Y - Val(polygonCount(0))
                        VertexArray(3).Y = VertexArray(3).Y + Val(polygonCount(0))
                    End If
                           
                End If
                       
            ElseIf X Mod 2 = 1 Then
                   
                If Y Mod 2 = 0 Then
                    If POLYGON_IGNORE_TOP <> 1 Then
                        VertexArray(0).Y = VertexArray(0).Y + Val(polygonCount(0))
                        VertexArray(1).Y = VertexArray(1).Y - Val(polygonCount(0))
                    End If
                    If POLYGON_IGNORE_LOWER <> 1 Then
                        VertexArray(2).Y = VertexArray(2).Y - Val(polygonCount(1))
                        VertexArray(3).Y = VertexArray(3).Y + Val(polygonCount(1))
                    End If
                Else
                    If POLYGON_IGNORE_TOP <> 1 Then
                        VertexArray(0).Y = VertexArray(0).Y - Val(polygonCount(1))
                        VertexArray(1).Y = VertexArray(1).Y + Val(polygonCount(1))
                    End If
                           
                    If POLYGON_IGNORE_LOWER <> 1 Then
                        VertexArray(2).Y = VertexArray(2).Y + Val(polygonCount(0))
                        VertexArray(3).Y = VertexArray(3).Y - Val(polygonCount(0))
                    End If
                End If
                
            End If
            
        End If

        DirectDevice.SetTexture 0, D3DTextures.Texture
        
        If Wireframe = True Then
            DirectDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
        Else
            DirectDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        End If
        
        DirectDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, _
                indexList(0), D3DFMT_INDEX16, _
                VertexArray(0), Len(VertexArray(0))
                
        'Draw the triangles that make up our square Textures
        'DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), Len(VertexArray(0))
    End With
    
End If

End Sub

Public Sub Engine_FPS_Update()

' / Author: Dunkansdk
' / Note: Limit FPS & Calculate later

        ' Limitar las fps
        ' While (GetTickCount - fpsLastCheck) \ 14 < FramesPerSecCounter: Wend

        If fpsLastCheck + 1000.1 < GetTickCount Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            fpsLastCheck = GetTickCount
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If

        frmMain.lblFPS.Caption = FPS

End Sub

' - Screen - - - - - - - - - - - - - -
Public Sub Engine_BeginScene(Optional ByVal Color As Long = 0)

    ' / Author: Dunkansdk
    ' / Note: DD Clear & BeginScene
    
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, Color, 1#, 0
    'DirectDevice.BeginScene

End Sub

Public Sub Engine_EndScene(ByRef destRect As RECT, Optional ByVal hWndDest As Long = 0)
    
    ' / Author: Dunkansdk
    ' / Note: DD EndScene & Present
    
    If hWndDest = 0 Then
        'DirectDevice.EndScene
        DirectDevice.Present destRect, ByVal 0, ByVal 0, ByVal 0
    Else
        'DirectDevice.EndScene
        DirectDevice.Present destRect, ByVal 0, hWndDest, ByVal 0
    End If
    
End Sub

Public Sub Engine_Zoom_In()

    ' / Author: Dunkansdk
    ' / Note: + Zoom

    With MainScreenRect
        .bottom = IIf(.bottom - 1 <= 367, .bottom, .bottom - 1)
        .Right = IIf(.Right - 1 <= 491, .Right, .Right - 1)
    End With
    
End Sub

Public Sub Engine_Zoom_Out()

    ' / Author: Dunkansdk
    ' / Note: - Zoom
    
    With MainScreenRect
        .bottom = IIf(.bottom + 1 >= 459, .bottom, .bottom + 1)
        .Right = IIf(.Right + 1 >= 583, .Right, .Right + 1)
    End With
    
End Sub

Public Sub Engine_Zoom_Normal()
    
    ' / Author: Dunkansdk
    ' / Note: Sin Zoom
    
    With MainScreenRect
        .bottom = frmMain.MainViewPic.ScaleHeight
        .Right = frmMain.MainViewPic.ScaleWidth
    End With
    
End Sub

Public Function Engine_Zoom_Offset(ByVal Offset As Byte) As Single

    ' / Author: Dunkansdk
    ' / Note: Offset

    Engine_Zoom_Offset = IIf((Offset = 1), (ScreenHeight - MainScreenRect.bottom) / 2, (ScreenWidth - MainScreenRect.Right) / 2)
    
End Function

' - Screen - - - - - - - - - - - - - -

Public Sub Engine_Convert_List(rgb_list() As Long, Long_Color As Long)

    ' / Author: Dunkansdk
    ' / Note: Convierte en array's los D3DColorArgb

    rgb_list(0) = Long_Color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
    
End Sub

Public Sub Engine_Draw_Box(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As Long)

    ' / Author: Ezequiel Juárez (Standelf)
    ' / Note: Extract to Blisse AO, modified by Dunkansdk

    Dim b_Rect As RECT
    Dim b_Color(0 To 3) As Long
    Dim b_Vertex(0 To 3) As TLVERTEX
    
    With b_Rect
        .bottom = Y + Height
        .Left = X
        .Right = X + Width
        .Top = Y
    End With

    Engine_Convert_List b_Color(), Color

    Geometry_Create_Box b_Vertex(), b_Rect, b_Rect, b_Color(), 0, 0
    
    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, b_Vertex(0), Len(b_Vertex(0))

End Sub

Public Sub Engine_InitColor()
    
    ' / Author: Dunkansd
    ' / Note: Inicializa los array de los colores
    
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    ShadowColor(0) = ShadowColor(0)
    ShadowColor(0) = ShadowColor(0)
    ShadowColor(0) = ShadowColor(0)
    ShadowColor(0) = ShadowColor(0)
    
    AlphaColor(0) = D3DColorARGB(70, 255, 255, 255)
    AlphaColor(1) = AlphaColor(0)
    AlphaColor(2) = AlphaColor(0)
    AlphaColor(3) = AlphaColor(0)
    
End Sub

Public Sub Engine_LoadMap_Connect()

    UserMap = 2
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    If FileExist(App.path & "\Mapas\Mapa" & CStr(UserMap) & ".map", vbNormal) Then
        Call SwitchMap(UserMap)
    Else
        MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        Call CloseClient
    End If
    
End Sub

