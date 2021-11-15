Attribute VB_Name = "Carteles"
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

' / Este módulo es horrible, sacarlo a la mierda cuando se valla la paja. Atte Dunkan.

Const XPosCartel = 360
Const YPosCartel = 335
Const MAXLONG = 40

'Carteles
Public Cartel       As Boolean
Public Leyenda      As String
Public Textura      As Integer

Public LeyendaFormateada() As String


Sub InitCartel(Ley As String, Grh As Integer)

If Not Cartel Then

    Leyenda = Ley
    Textura = Grh
    Cartel = True
    
    ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
                
    Dim i As Integer, k As Integer, anti As Integer
    anti = 1
    k = 0
    i = 0
    
    Call DarFormato(Leyenda, i, k, anti)
    
    i = 0
    
    Do While LeyendaFormateada(i) <> "" And i < UBound(LeyendaFormateada)
       i = i + 1
    Loop
    
    ReDim Preserve LeyendaFormateada(0 To i)

Else

    Exit Sub
    
End If

End Sub


Private Function DarFormato(s As String, i As Integer, k As Integer, anti As Integer)
If anti + i <= Len(s) + 1 Then
    If ((i >= MAXLONG) And mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
        LeyendaFormateada(k) = mid(s, anti, i + 1)
        k = k + 1
        anti = anti + i + 1
        i = 0
    Else
        i = i + 1
    End If
    Call DarFormato(s, i, k, anti)
End If
End Function

Sub DibujarCartel()

If Not Cartel Then Exit Sub

Dim X As Integer, Y As Integer
X = XPosCartel + 20
Y = YPosCartel + 60

Dim j As Integer, desp As Integer

For j = 0 To UBound(LeyendaFormateada)
    Engine_RenderText X, Y + desp, LeyendaFormateada(j), D3DColorARGB(255, 255, 255, 255)
    desp = desp + (frmMain.Font.Size) + 5
Next

Call Engine_Draw_Box(X, Y, LeyendaFormateada(j), desp, D3DColorARGB(100, 100, 100, 100))

End Sub

