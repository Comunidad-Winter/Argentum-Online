VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBmpManNoDyn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Argentum Online 0.11.2
'
'Copyright (C) 2002 Márquez Pablo Ignacio
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


'BMP Manager No Dinámico
'Por Alejandro Santos
'Basado en el BMPMan Dinamico por Federico Torres

Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer 'borrar de aca para abajo?
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type BMP_INFO
    Height As Long
    Width As Long
    Size As Long
End Type

Private GrhPath As String
Private mDDraw As DirectDraw7
Private mTotalGraficos As Long
Private mSurfaceDB() As DirectDrawSurface7
Private mUsarMemoriaVideo As Boolean

'Para compatibilidad con el BMPMan Dinamico
Public Property Get MaxEntries() As Long

End Property

Public Property Let MaxEntries(ByVal vNewValue As Long)

End Property

Public Property Let lpDirectDraw7(ByRef objDirectDraw7 As DirectDraw7)
Set mDDraw = objDirectDraw7
End Property

Public Property Let Path(ByVal inString As String)
GrhPath = inString
End Property

Public Property Get TotalGraficos() As Long
TotalGraficos = mTotalGraficos
End Property

Public Property Let TotalGraficos(ByVal lNewValue As Long)
mTotalGraficos = lNewValue
ReDim mSurfaceDB(1 To mTotalGraficos)

End Property

Private Sub Class_Initialize()
mTotalGraficos = 0
GrhPath = App.Path
Set mDDraw = Nothing
mUsarMemoriaVideo = True
End Sub

Public Property Get EsDinamico() As Variant
EsDinamico = False
End Property

Private Sub Class_Terminate()
On Local Error Resume Next
Dim I As Long

For I = LBound(mSurfaceDB) To UBound(mSurfaceDB)
    Set mSurfaceDB(I) = Nothing
Next I
End Sub

Private Function GetBMPHeader(Archivo As String, ByRef lpBMPInfo As BMP_INFO) As Long
Dim nfile As Integer, H_BMPH As BITMAPFILEHEADER, H_BMPI As BITMAPINFOHEADER

If FileExist(Archivo, vbNormal) Then
    nfile = FreeFile()
    Open Archivo For Binary Access Read As #nfile
    Get #nfile, , H_BMPH
    Get #nfile, , H_BMPI
    Close #nfile
    lpBMPInfo.Height = H_BMPI.biHeight
    lpBMPInfo.Width = H_BMPI.biWidth
    lpBMPInfo.Size = H_BMPH.bfSize
    GetBMPHeader = 1
Else
    GetBMPHeader = 0
End If

End Function

Public Function CargarGrafico(Archivo As Integer) As Integer
'12/1/3, colorkey
On Local Error GoTo HayError
Dim ddsd As DDSURFACEDESC2, BMPInfo As BMP_INFO, RetVal As Long, ddck As DDCOLORKEY
ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH

'para alto y ancho de la surface
If GetBMPHeader(GrhPath & Archivo & ".bmp", BMPInfo) Then

    If mUsarMemoriaVideo Then
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    
    ddsd.lHeight = BMPInfo.Height
    ddsd.lWidth = BMPInfo.Width
    
    ddck.high = 0: ddck.low = 0
    
    Set mSurfaceDB(Archivo) = mDDraw.CreateSurfaceFromFile(GrhPath & Archivo & ".bmp", ddsd) 'creo la surface
    mSurfaceDB(Archivo).SetColorKey DDCKEY_SRCBLT, ddck
    
    CargarGrafico = 1
Else
HayError:

    CargarGrafico = 0
End If

End Function

Public Function GetBMP(ByVal FileName As Integer, Optional ByVal NoUsado As Integer) As DirectDrawSurface7
Set GetBMP = mSurfaceDB(FileName)
End Function

Public Function Init(UsarMemoriaVideo As Boolean) As Long
Init = 1
mUsarMemoriaVideo = UsarMemoriaVideo
End Function

Public Function BorrarTodo() As Long
On Local Error Resume Next
Dim I As Long

For I = LBound(mSurfaceDB) To UBound(mSurfaceDB)
    Set mSurfaceDB(I) = Nothing
Next I

End Function

Public Property Get EfectoPred() As Integer
EfectoPred = 0
End Property

Public Property Let EfectoPred(ByVal iNewValue As Integer)

End Property

