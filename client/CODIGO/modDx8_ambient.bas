Attribute VB_Name = "modDx8_ambient"
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

' / Author: Emanuel Matías (Dunkan)
' / Note: Maneja los estados climáticos

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public ColorActual  As D3DCOLORVALUE
Public ColorFinal   As D3DCOLORVALUE
Public Fade         As Boolean
Public Base_Light   As Long

Public AmbientLastCheck As Long

Public WeatherEffectIndex   As Integer
Public LastWeather          As Byte

Public WeatherFogX1     As Single
Public WeatherFogY1     As Single
Public WeatherFogX2     As Single
Public WeatherFogY2     As Single
Public WeatherDoFog     As Byte
Public WeatherFogCount  As Byte
Public LightningTimer   As Single
Public FlashTimer       As Single

Sub Engine_Weather_UpdateFog()

' / Author: Emanuel Matías (Dunkan)
' / Note: Adaptado de vbGore, niebla.

Dim TempGrh As Grh
Dim i As Long
Dim X As Long
Dim Y As Long
Dim c As Long

    'Make sure we have the fog value
    If WeatherFogCount = 0 Then WeatherFogCount = 13
    
    'Update the fog's position
    WeatherFogX1 = WeatherFogX1 + (timerElapsedTime * (0.018 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY1 = WeatherFogY1 + (timerElapsedTime * (0.013 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
    Do While WeatherFogX1 < -512
        WeatherFogX1 = WeatherFogX1 + 512
    Loop
    Do While WeatherFogY1 < -512
        WeatherFogY1 = WeatherFogY1 + 512
    Loop
    Do While WeatherFogX1 > 0
        WeatherFogX1 = WeatherFogX1 - 512
    Loop
    Do While WeatherFogY1 > 0
        WeatherFogY1 = WeatherFogY1 - 512
    Loop
    
    WeatherFogX2 = WeatherFogX2 - (timerElapsedTime * (0.037 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY2 = WeatherFogY2 - (timerElapsedTime * (0.021 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
    Do While WeatherFogX2 < -512
        WeatherFogX2 = WeatherFogX2 + 512
    Loop
    Do While WeatherFogY2 < -512
        WeatherFogY2 = WeatherFogY2 + 512
    Loop
    Do While WeatherFogX2 > 0
        WeatherFogX2 = WeatherFogX2 - 512
    Loop
    Do While WeatherFogY2 > 0
        WeatherFogY2 = WeatherFogY2 - 512
    Loop

    TempGrh.FrameCounter = 1
    
    'Render fog 2
    TempGrh.GrhIndex = 23654
    
    X = 2
    Y = -1
    c = D3DColorARGB(100, 255, 255, 255)
    
    For i = 1 To WeatherFogCount
        DDrawTransGrhIndextoSurface TempGrh.GrhIndex, (X * 512) + WeatherFogX1, (Y * 512) + WeatherFogY1, 1
        X = X + 1
        If X > (1 + (ScreenWidth \ 512)) Then
            X = 0
            Y = Y + 1
        End If
    Next i
            
    'Render fog 1
    TempGrh.GrhIndex = 23655
    X = 0
    Y = 0
    c = D3DColorARGB(100, 255, 255, 255)
    For i = 1 To WeatherFogCount
        DDrawTransGrhIndextoSurface TempGrh.GrhIndex, (X * 512) + WeatherFogX1, (Y * 512) + WeatherFogY1, 1
        X = X + 1
        If X > (2 + (ScreenWidth \ 512)) Then
            X = 0
            Y = Y + 1
        End If
    Next i

End Sub

Public Sub Ambient_Set(LightR As Byte, LightG As Byte, LightB As Byte, Optional ByVal Fade As Boolean = True)

' Author: Emanuel Matías (Dunkan)
' Note: Setea el RGB

Dim Light As D3DCOLORVALUE

Light.R = LightR
Light.G = LightG
Light.B = LightB

ColorFinal = Light

If Not Fade Then
    ColorActual = Light
End If

End Sub

Public Sub Ambient_Check()

' Author: Emanuel Matías (Dunkan)
' Note: Verifica la hora para cambiar el clima

Dim Hora As Byte, Minutos As Byte
Hora = Hour(time)
Minutos = Minute(time)
    
    If LastWeather = 1 Then
    
        Ambient_SetFinal 169, 169, 197
        
        
    Else
        
        If Hora >= 6 And Hora < 8 Then
            'Amanecer
            Ambient_SetFinal 193, 176, 121
        ElseIf Hora >= 8 And Hora < 16 Then
            'Dia
            Ambient_SetFinal 255, 255, 255
        ElseIf Hora >= 16 And (Hora < 19 Or (Hora = 19 And Minutos < 30)) Then
            'Tarde
            Ambient_SetFinal 215, 213, 215
        ElseIf (Hora = 19 And Minutos >= 30 And Minutos < 35) Then
            'Anochecer
            Ambient_SetFinal 187, 189, 192
        ElseIf (((Hora > 19) Or (Hora = 19 And Minutos >= 35)) And Hora < 23) Or _
                  Hora >= 3 And Hora < 6 Then
            'Noche
            Ambient_SetFinal 169, 169, 197
        ElseIf (Hora = 23 Or Hora = 0) Or (Hora > 0 And Hora < 3) Then
            'Noche mas oscura
            Ambient_SetFinal 159, 155, 197
        End If
        
    End If
   
With ColorActual
    'Red
    If .R < ColorFinal.R Then
        .R = .R + 1
    ElseIf .R > ColorFinal.R Then
        .R = .R - 1
    End If
    'Green
    If .G < ColorFinal.G Then
        .G = .G + 1
    ElseIf .G > ColorFinal.G Then
        .G = .G - 1
    End If
    'Blue
    If .B < ColorFinal.B Then
        .B = .B + 1
    ElseIf .B > ColorFinal.B Then
        .B = .B - 1
    End If

End With

End Sub

Public Sub Ambient_Fade()

' Author: Emanuel Matías (Dunkan)
' Note: Cambio progresivo del clima

With ColorActual

    'Red
    If .R < ColorFinal.R Then
        .R = .R + 1
    ElseIf .R > ColorFinal.R Then
        .R = .R - 1
    End If
    
    'Green
    If .G < ColorFinal.G Then
        .G = .G + 1
    ElseIf .G > ColorFinal.G Then
        .G = .G - 1
    End If
    
    'Blue
    If .B < ColorFinal.B Then
        .B = .B + 1
    ElseIf .B > ColorFinal.B Then
        .B = .B - 1
    End If
    
End With

Fade = Not (ColorFinal.R = ColorActual.R And ColorFinal.G = ColorActual.G And ColorFinal.B = ColorActual.B)

Base_Light = D3DColorARGB(255, ColorActual.R, ColorActual.G, ColorActual.B)

End Sub

Public Sub Ambient_Start()

' / Author: Emanuel Matías (Dunkan)
' / Note: Inicia el clima (En desuso)

Dim Hora As Byte, Minutos As Byte
Hora = Hour(time)
Minutos = Minute(time)

If Hora >= 6 And Hora < 8 Then
    'Amanecer
    Ambient_SetFinal 193, 176, 121
    Ambient_SetActual 193, 176, 121
    
ElseIf Hora >= 8 And Hora < 16 Then
    'Dia
    Ambient_SetFinal 255, 255, 255
    Ambient_SetActual 255, 255, 255
    
ElseIf Hora >= 16 And (Hora < 20 Or (Hora = 20 And Minutos < 30)) Then
    'Tarde
    Ambient_SetFinal 215, 213, 215
    Ambient_SetActual 215, 213, 215
    
ElseIf (Hora = 20 And Minutos >= 30 And Minutos < 35) Then
    'Anochecer
    Ambient_SetFinal 187, 189, 192
    Ambient_SetActual 187, 189, 192
    
ElseIf (((Hora > 20) Or (Hora = 20 And Minutos >= 35)) And Hora < 23) Or _
          Hora >= 3 And Hora < 6 Then
    'Noche
    Ambient_SetFinal 159, 159, 167
    Ambient_SetActual 159, 159, 167
    
ElseIf (Hora = 23 Or Hora = 0) Or (Hora > 0 And Hora < 3) Then
    'Noche mas oscura
    Ambient_SetFinal 149, 150, 197
    Ambient_SetActual 149, 150, 197
    
End If

End Sub

Public Sub Ambient_SetFinal(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)

' / Author: Emanuel Matías (Dunkan)

    ColorFinal.R = R: ColorFinal.G = G: ColorFinal.B = B
    Fade = Not (ColorFinal.R = ColorActual.R And ColorFinal.G = ColorActual.G And ColorFinal.B = ColorActual.B)
    
End Sub

Public Sub Ambient_SetActual(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)

' / Author: Emanuel Matías (Dunkan)

    ColorActual.R = R: ColorActual.G = G: ColorActual.B = B
    Fade = Not (ColorFinal.R = ColorActual.R And ColorFinal.G = ColorActual.G And ColorFinal.B = ColorActual.B)
    
End Sub


