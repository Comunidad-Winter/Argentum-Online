Attribute VB_Name = "modMath"
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

' / Author: Dunkansdk
' / Note: Modulo donde se almacenan todas las funciones matemáticas de Boskorcha AO

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Sub vASM_PUTMEMx(ByVal ptr As Long, ByVal NewVal As Long, ByVal nB As Long)
    Dim Acode(25)   As Byte: Acode(25) = &HC3
 
    Acode(17) = &H8A: Acode(18) = &H10: Acode(19) = &H88: Acode(20) = &H17
    Acode(21) = &H40: Acode(22) = &H47: Acode(23) = &HE2: Acode(24) = &HF8
    
    Dim i           As Long
    
      '       MOV EAX,OFFSET newval
      '       MOV EDI,OFFSET ptr
      '       XOR ECX,ECX
      '       MOV ECX,nB
      'INI:   MOV DL,[EAX]
      '       MOV [EDI], DL
      '       INC EAX
      '       INC EDI
      '       LOOP INI
    
      Acode(0) = &HB8
      i = LongToByte(NewVal, Acode(), i + 1)
    
      Acode(5) = &HBF
    
      i = LongToByte(ptr, Acode(), i + 1)
      Acode(10) = &H33: Acode(11) = &HC9
    
      Acode(12) = &HB9
      i = LongToByte(nB, Acode(), i + 3)
    
      Call CallWindowProc(ByVal VarPtr(Acode(0)), 0&, 0&, 0&, 0&)
      
End Sub

Private Function LongToByte(ByVal lLong As Long, ByRef bReturn() As Byte, Optional i As Integer = 0) As Long
   
   ' / Author: BlackZeroX
   
   bReturn(i) = lLong And &HFF
   bReturn(i + 1) = (lLong And &HFF00&) \ &H100
   bReturn(i + 2) = (lLong And &HFF0000) \ &H10000
   bReturn(i + 3) = (lLong And &HFF000000) \ &H1000000
   LongToByte = i + 4
   
End Function

Public Function GetAngle2Points(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Double
     
'Author: Dunkan
'Note: Calcula el ángulo entre dos puntos
    
    GetAngle2Points = GetAngleXY((X2 - X1), (Y2 - Y1))
    
End Function
 
Public Function GetAngleXY(ByVal X As Double, ByVal Y As Double) As Double

'Author: Dunkan
'Note: Calcula el ángulo entre dos puntos
    
Dim dblres              As Double

    dblres = 0
    
    If (Y <> 0) Then
        dblres = Radianes2Grados(Atn(X / Y))
        If (X <= 0 And Y < 0) Then
            dblres = dblres + 180
        ElseIf (X > 0 And Y < 0) Then
            dblres = dblres + 180
        ElseIf (X < 0 And Y > 0) Then
            dblres = dblres + 360
        End If
    Else
        If (X > 0) Then
            dblres = 90
        ElseIf (X < 0) Then
            dblres = 270
        End If
    End If
    
    GetAngleXY = dblres
    
End Function
 
Public Function Grados2Radianes(ByVal Grados As Double) As Double

'Author: Dunkan
'Note: Convierte grados en radianes
    
    Grados2Radianes = Grados * (3.14159265358979 / 180) ' PI / 180
    
End Function
 
Public Function Radianes2Grados(ByVal Radianes As Double) As Double

'Author: Dunkan
'Note: Convierte radianes en grados

    Radianes2Grados = Radianes * 180 / 3.14159265358979 ' 180 / PI
    
End Function


Function Get_Distance(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Single

' Author: Emanuel Matías 'Dunkan'
' Note: Distancia entre dos puntos.

    Get_Distance = (Abs(X2 - X1) + Abs(Y2 - Y1)) * 0.5
    
End Function

Public Sub Generate_Mountain(ByVal Altura As Integer, ByVal Radio_X As Integer, ByVal Radio_Y As Integer, _
                            ByVal Pos_X As Integer, ByVal Pos_Y As Integer)
                            
    ' / Author: Dunkansdk
    ' / Note: Montañas, Colinas y toda esa parafernalia.
    
    Dim X           As Byte
    Dim Y           As Byte
    Dim minX        As Integer
    Dim maxX        As Integer
   
    For Y = Pos_Y To Pos_Y + Radio_Y / 2
    
        minX = Pos_X - Radio_X / 2 + Y - Pos_Y - 1
        maxX = Pos_X + Radio_X / 2 - (Y - Pos_Y - 1)
        
        For X = minX To maxX
        
                                
                MapData(X, Y).Offset(0) = Altura * Sin(DegreeToRadian * 180 * (X - (Pos_X - Radio_X / 2 + Y - Pos_Y)) / (maxX - minX))
                MapData(X, Y).Offset(1) = Altura * Sin(DegreeToRadian * 180 * (X + 1 - (Pos_X - Radio_X / 2 + Y - Pos_Y)) / (maxX - minX))
                
                If MapData(X, Y).Offset(0) < 0 Then MapData(X, Y).Offset(0) = 0
                
                MapData(X, Y - 1).Offset(2) = MapData(X, Y).Offset(0)
                MapData(X, Y - 1).Offset(3) = MapData(X, Y).Offset(1)
                                               
                'If Pos_X < X Then
                '    MapData(X, Y).Altura_Pie = Max_(CInt(MapData(X, Y).Offset(0)), CInt(MapData(X, Y).Offset(1))) / 32
                'Else
                '    MapData(X, Y).Altura_Pie = -(Max_(CInt(MapData(X, Y).Offset(0)), CInt(MapData(X, Y).Offset(1))) / 32)
                'End If
                
        Next X
    
    Next Y

    
                            End Sub
 
Function Max_(N_1 As Integer, N_2 As Integer) As Integer

    ' Author: Dunkansdk
    ' Note: Nro Máximo -
    
    If N_1 < N_2 Then
        Max_ = N_2
    Else
        Max_ = N_1
    End If
    
End Function

Function Min_(N_1 As Integer, N_2 As Integer) As Integer

    ' Author: Dunkansdk
    ' Note: Nro Mínimo -
    
    If N_1 > N_2 Then
        Min_ = N_2
    Else
        Min_ = N_1
    End If
    
End Function


