Attribute VB_Name = "modDx8_HDC"
Option Explicit

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' This program is free software; you can redistribute it and/or modify
' it under the terms of the Affero General Public License;
' either version 1 of the License, or any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the ifmplied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' Affero General Public License for more details.
'
' You should have received a copy of the Affero General Public License
' along with this program; if not, you can find it at http://www.affero.org/oagpl.html
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' / Módulo creado por Emanuel Matías D'Urso 'Dunkan'
' / Note: Manejo de .PNG's y efectos con ellos.

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public Enum HDGraphic

    SHADOW_HD = 23651
    LIGHT1_HD = 23652
    LIGHT2_HD = 23656
    RADIUS_HD = 23653
    
End Enum

Public Sub Device_Texture_PNG_Render(ByVal X As Integer, ByVal Y As Integer, _
                                ByVal Texture As Direct3DTexture8, _
                                ByRef src_rect As RECT, _
                                Optional ByVal Alpha As Byte, _
                                Optional ByVal AlphaByte As Byte = 255, _
                                Optional ByVal pixelR As Byte = 0, _
                                Optional ByVal pixelG As Byte = 0, _
                                Optional ByVal pixelB As Byte = 0, _
                                Optional ByVal setFactorAlpha As Boolean = False)
                                
    ' / - - - - - - - - - - - - - - - - - -
    ' / Author: Dunkansdk
    ' / Note: Dibuja las texturas HD
    ' / - - - - - - - - - - - - - - - - - -
    
    Dim dest_rect       As RECT
    Dim SRDesc          As D3DSURFACE_DESC
    Dim light_value(3)  As Long
    Dim texture_width   As Long
    Dim texture_height  As Long
    
    'On Error Resume Next

    light_value(0) = D3DColorARGB(AlphaByte, pixelR, pixelG, pixelB)
    light_value(1) = D3DColorARGB(AlphaByte, pixelR, pixelG, pixelB)
    light_value(2) = D3DColorARGB(AlphaByte, pixelR, pixelG, pixelB)
    light_value(3) = D3DColorARGB(AlphaByte, pixelR, pixelG, pixelB)

    With dest_rect
        .bottom = Y + (src_rect.bottom - src_rect.Top)
        .Left = X
        .Right = X + (src_rect.Right - src_rect.Left)
        .Top = Y
    End With

    ' Texture Settings - Dunkan
    Texture.GetLevelDesc 0, SRDesc
    texture_width = SRDesc.Width
    texture_height = SRDesc.Height
    
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), texture_width, texture_height
    
    With DirectDevice
    
        .SetTexture 0, Texture
        
        If Alpha = 1 Then
            .SetRenderState D3DRS_SRCBLEND, 1
            .SetRenderState D3DRS_DESTBLEND, 3
        ElseIf Alpha = 2 Then
            .SetRenderState D3DRS_SRCBLEND, 4
            .SetRenderState D3DRS_DESTBLEND, 3
        ElseIf Alpha = 3 Then
            .SetRenderState D3DRS_SRCBLEND, 3
            .SetRenderState D3DRS_DESTBLEND, 2
        ElseIf Alpha = 4 Then
            .SetRenderState D3DRS_SRCBLEND, 2
            .SetRenderState D3DRS_DESTBLEND, 2
        End If
        
        If setFactorAlpha = True Then
            .SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(AlphaByte, pixelR, pixelG, pixelB)
            .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTA_TFACTOR
        End If
    
        ' MEDIUM LOAD
        '.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
        
        ' FASTER LOAD - DUNKAN
        .DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, _
                    indexList(0), D3DFMT_INDEX16, _
                    temp_verts(0), Len(temp_verts(0))
        
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
    End With
    
End Sub

Public Sub DDrawSurfacePNG(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, _
                    Optional ByVal byteAlpha As Byte = 255, _
                    Optional ByVal alphaType As Byte, _
                    Optional ByVal Rotate As Single = 0, _
                    Optional ByVal pixelR As Byte = 255, _
                    Optional ByVal pixelG As Byte = 255, _
                    Optional ByVal pixelB As Byte = 255, _
                    Optional ByVal setFactorAlpha As Boolean = False)
                    
    ' / Author: Dunkansdk
    
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.bottom = SourceRect.Top + .pixelHeight
        
        If Center Then
            X = X - (.pixelHeight / 2)
            Y = Y - (.pixelWidth / 2)
        End If
        
        'DirectDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

        Call Device_Texture_PNG_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, alphaType, byteAlpha, pixelR, pixelG, pixelB, setFactorAlpha)
    
    End With
    
End Sub

Public Function Render_Radio(ByVal Alpha As Byte)

' / Author: Dunkansdk

    Call DDrawSurfacePNG(HDGraphic.RADIUS_HD, -242, 200, 1, Alpha)
    Call DDrawSurfacePNG(HDGraphic.RADIUS_HD, 270, 200, 1, Alpha)
    Call DDrawSurfacePNG(HDGraphic.RADIUS_HD, 782, 200, 1, Alpha)
    
End Function

