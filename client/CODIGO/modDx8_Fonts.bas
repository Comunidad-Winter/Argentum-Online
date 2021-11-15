Attribute VB_Name = "modDx8_Fonts"
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

' / Sistema de Font de Boskorcha AO - Dunkansdk

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
    Width       As Long
    Height      As Long
    Depth       As Long
    MipLevels   As Long
    format      As CONST_D3DFORMAT
    
    ResourceType    As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Private Type VFH
    BitmapWidth         As Long     'Size of the bitmap itself
    BitmapHeight        As Long
    CellWidth           As Long     'Size of the cells (area for each character)
    CellHeight          As Long
    BaseCharOffset      As Byte     'The character we start from
    CharWidth(0 To 255) As Byte     'The actual factual width of each character
    CharVA(0 To 255)    As CharVA
End Type

Public Type CustomFont
    HeaderInfo  As VFH              'Holds the header information
    Texture     As Direct3DTexture8 'Holds the texture of the text
    RowPitch    As Integer          'Number of characters per row
    RowFactor   As Single           'Percentage of the texture width each character takes
    ColFactor   As Single           'Percentage of the texture height each character takes
    CharHeight  As Byte             'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI         'Size of the texture
End Type

'Private Const Font_Default_TextureNum As Long = -1   'The texture number used to represent this font - only used for AlternateRendering - keep negative to prevent interfering with game textures
Public cfonts(1 To 3) As CustomFont ' _Default2 As CustomFont
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)

Public Sub Engine_RenderText(ByVal Left As Long, ByVal Top As Long, ByVal Text As String, ByVal Color As Long, Optional ByVal Alpha As Byte = 255, Optional ByVal Center As Boolean = False)
    
On Error Resume Next ' < maTih JAJAJA
' Author: Dunkan
' Note: Render text

    If Alpha <> 255 Then
        Dim newRGB As D3DCOLORVALUE
        ARGBtoD3DCOLORVALUE Color, newRGB
        Color = D3DColorARGB(Alpha, newRGB.R, newRGB.G, newRGB.B)
    End If
    
    Engine_Render_Text cfonts(1), Text, Left - 1, Top - 1, D3DColorARGB(Alpha - 40, 0, 0, 0), Center, Alpha - 40
        
    Engine_Render_Text cfonts(1), Text, Left, Top, Color, Center, Alpha
    
End Sub

Public Function ARGBtoD3DCOLORVALUE(ByVal ARGB As Long, ByRef Color As D3DCOLORVALUE)
Dim dest(3) As Byte

CopyMemory dest(0), ARGB, 4
    Color.a = dest(3)
    Color.R = dest(2)
    Color.G = dest(1)
    Color.B = dest(0)
    
End Function

Private Sub Engine_Render_Text(ByRef UseFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal Color As Long, Optional ByVal Center As Boolean = False, Optional ByVal Alpha As Byte = 255)

' / Author: Spodi
' / Optimizado por: Dunkan

Dim TempVA(0 To 3)  As TLVERTEX
Dim tempstr()       As String
Dim Count           As Integer
Dim ascii()         As Byte
Dim Row             As Integer
Dim U               As Single
Dim V               As Single
Dim i               As Long
Dim j               As Long
Dim KeyPhrase       As Byte
Dim TempColor       As Long
Dim ResetColor      As Byte
Dim SrcRect         As RECT
Dim v2              As D3DVECTOR2
Dim v3              As D3DVECTOR2
Dim YOffset         As Single

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
    
    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    
    'Get the text
    tempstr = Split(Text, vbCrLf)
    
    'Set the temp color (or else the first character has no color)
    TempColor = Color

    'Set the texture
    DirectDevice.SetTexture 0, UseFont.Texture
    
    If Center Then
        X = X - Engine_GetTextWidth(cfonts(1), Text) * 0.5
    End If
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)
        If Len(tempstr(i)) > 0 Then
            YOffset = i * UseFont.CharHeight
            Count = 0
        
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
    
            'Loop through the characters
            For j = 1 To Len(tempstr(i))

                'Check for a key phrase
                'If ascii(j - 1) = 124 Then 'If Ascii = "|"
                '    KeyPhrase = (Not KeyPhrase)  'TempColor = ARGB 255/255/0/0
                '    If KeyPhrase Then TempColor = ARGB(255, 0, 0, alpha) Else ResetColor = 1
                'Else

                    'Render with triangles
                    'If AlternateRender = 0 Then

                        'Copy from the cached vertex array to the temp vertex array
                        CopyMemory TempVA(0), UseFont.HeaderInfo.CharVA(ascii(j - 1)).Vertex(0), 32 * 4

                        'Set up the verticies
                        TempVA(0).X = X + Count
                        TempVA(0).Y = Y + YOffset
                        
                        TempVA(1).X = TempVA(1).X + X + Count
                        TempVA(1).Y = TempVA(0).Y

                        TempVA(2).X = TempVA(0).X
                        TempVA(2).Y = TempVA(2).Y + TempVA(0).Y

                        TempVA(3).X = TempVA(1).X
                        TempVA(3).Y = TempVA(2).Y
                        
                        'Set the colors
                        TempVA(0).Color = TempColor
                        TempVA(1).Color = TempColor
                        TempVA(2).Color = TempColor
                        TempVA(3).Color = TempColor
                        
                        'Draw the verticies
                        'DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0))
                        
                        ' / Faster LOAD - Dunkan
                        DirectDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, _
                                                            indexList(0), D3DFMT_INDEX16, _
                                                            TempVA(0), Len(TempVA(0))
                      
                    'Shift over the the position to render the next character
                    Count = Count + UseFont.HeaderInfo.CharWidth(ascii(j - 1))
                
                'End If
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = Color
                End If
                
            Next j
            
        End If
    Next i
    
End Sub

Public Function Engine_GetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
Dim i As Integer

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    'Loop through the text
    For i = 1 To Len(Text)
        
        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
        
    Next i

End Function

Sub Engine_Init_FontTextures()
On Error GoTo eDebug:
Dim TexInfo As D3DXIMAGE_INFO_A

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '*** Default font ***
    
    'Set the texture
    Set cfonts(1).Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, App.path & "\Data\texdefault.bmp", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)

    'Store the size of the texture
    cfonts(1).TextureSize.X = TexInfo.Width
    cfonts(1).TextureSize.Y = TexInfo.Height
    
    Exit Sub
eDebug:
    If Err.number = "-2005529767" Then
        MsgBox "Error en la texdefault.png (Font)", vbCritical
        End
    End If
    
End

End Sub

Sub Engine_Init_FontSettings()

' / Author: Unknow

Dim FileNum     As Byte
Dim LoopChar    As Long
Dim Row         As Single
Dim U           As Single
Dim V           As Single

    '*** Default font ***

    'Load the header information
    FileNum = FreeFile
    Open App.path & "\Data\FontData.dat" For Binary As #FileNum
        Get #FileNum, , cfonts(1).HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    cfonts(1).CharHeight = cfonts(1).HeaderInfo.CellHeight - 4
    cfonts(1).RowPitch = cfonts(1).HeaderInfo.BitmapWidth \ cfonts(1).HeaderInfo.CellWidth
    cfonts(1).ColFactor = cfonts(1).HeaderInfo.CellWidth / cfonts(1).HeaderInfo.BitmapWidth
    cfonts(1).RowFactor = cfonts(1).HeaderInfo.CellHeight / cfonts(1).HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) \ cfonts(1).RowPitch
        U = ((LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) - (Row * cfonts(1).RowPitch)) * cfonts(1).ColFactor
        V = Row * cfonts(1).RowFactor

        'Set the verticies
        With cfonts(1).HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).TU = U
            .Vertex(0).TV = V
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).Z = 0
            
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).TU = U + cfonts(1).ColFactor
            .Vertex(1).TV = V
            .Vertex(1).X = cfonts(1).HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).Z = 0
            
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).TU = U
            .Vertex(2).TV = V + cfonts(1).RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = cfonts(1).HeaderInfo.CellHeight
            .Vertex(2).Z = 0
            
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).TU = U + cfonts(1).ColFactor
            .Vertex(3).TV = V + cfonts(1).RowFactor
            .Vertex(3).X = cfonts(1).HeaderInfo.CellWidth
            .Vertex(3).Y = cfonts(1).HeaderInfo.CellHeight
            .Vertex(3).Z = 0
        End With
        
    Next LoopChar

End Sub


