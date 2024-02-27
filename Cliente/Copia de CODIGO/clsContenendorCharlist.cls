VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContenendorCharlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clsContenedorCharlist
'Esta clase es una especie de anticheating para evitar aquellos
'programas que modifican una posicion de memoria fija una vez que
'el programa se ha cargado.

'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Private p_dummyInts() As Integer
Private p_charlist() As Char


Public Sub Inicializar(ByVal relleno As Integer, ByVal MaxChars As Integer)

    Dim I As Integer
    
        ReDim p_dummyInts(1 To relleno) As Integer
        For I = 1 To relleno
            p_dummyInts(relleno) = RandomNumber(1, 30000)
        Next I
    
        ReDim p_charlist(1 To MaxChars) As Char
    
End Sub

Friend Function CharList(ByVal I As Integer) As Char
    CharList = p_charlist(I)
End Function


Public Sub CopyFrom(ByRef C As clsContenendorCharlist)
Dim I As Integer
    p_charlist(I) = C.CharList(I)
Next I
End Sub
