VERSION 5.00
Begin VB.Form frmLibrerias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Librerias"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   Icon            =   "frmLibrerias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LibName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "MSCOMCTL.OCX"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox LibName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "MSWINSCK.OCX"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox LibName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "CSWSK32.OCX"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox LibName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "RICHTX32.OCX"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox LibName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "AAM532.DLL"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox LibName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "MSINET.OCX"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuraci�n de Proxy para Descargas"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   3255
      Begin VB.TextBox txtProxy 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox ChkProxy 
         Caption         =   "Usar servidor"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin AOSetup.chameleonButton bCancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Aceptar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibrerias.frx":0442
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton cSolucion 
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibrerias.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton cVerificar 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Verificar nuevamente"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibrerias.frx":047A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton cSolucion 
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibrerias.frx":0496
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton cSolucion 
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibrerias.frx":04B2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton cSolucion 
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibrerias.frx":04CE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton cSolucion 
      Height          =   375
      Index           =   4
      Left            =   2280
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibrerias.frx":04EA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton cSolucion 
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibrerias.frx":0506
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   2280
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   2280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1560
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   2280
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   2280
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   2280
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmLibrerias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
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
'**************************************************************

Option Explicit

' MD5 de Referencia:
' RICHTX32.OCX 722435ba4d18f1704b43e823a12e489a
' CSWSK32.OCX 5181704b2772e050e4a8331e15ee4bb4
' MSINET.OCX 40d81470a19269d88bf44e766be7f84a
' MSWINSCK.OCX 3d8fd62d17a44221e07d5c535950449b

Public descargando As Boolean


Private Sub bCancelar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
Unload Me
End Sub

Sub LibError(ByVal Index As Byte, ByVal Solucion As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
    lblOK(Index).Caption = "ERROR"
    lblOK(Index).ForeColor = RGB(255, 0, 0)
    lblOK(Index).Visible = True
    cSolucion(Index).Caption = Solucion
    cSolucion(Index).Visible = True
    LibName(Index).BackColor = lblOK(Index).ForeColor
End Sub

Sub LibOK(ByVal Index As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
    lblOK(Index).Caption = "OK"
    lblOK(Index).ForeColor = &H8000&
    lblOK(Index).Visible = True
    cSolucion(Index).Visible = False
    LibName(Index).BackColor = lblOK(Index).ForeColor
End Sub

Private Sub cSolucion_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 10/01/07
'Last Modified by: Lucas Tavolaro Ortiz (Tavo)
'De ahora en mas se utiliza la funcion LibraryExist()
'*************************************************
    If cSolucion(Index).Caption = "&Registrar" Then
        ' registrar
        Dim fsoObject As FileSystemObject
        
        Select Case Index
            Case 0  ' inet
                If Not LibraryExist("MSINET.OCX", vbNormal) Then
                    MsgBox "ERROR, el archivo MSINET.OCX descargado tiene que ser copiado a este directorio.", vbCritical, "Argentum Online Setup"
                Else
                    Set fsoObject = New FileSystemObject
                    fsoObject.CopyFile "MSINET.OCX", fsoObject.GetSpecialFolder(SystemFolder) & "\", True
                    If Err Then MsgBox Err.Description
                    Shell "regsvr32 /s " & fsoObject.GetSpecialFolder(SystemFolder) & "\MSINET.OCX"
                    MsgBox "Copia y registro realizados con exito.", vbOKOnly, "Argentum Online Setup"
                End If
            
            Case 2  ' Rich
                If Not LibraryExist("richtx32.ocx", vbNormal) Then
                    MsgBox "ERROR, el archivo RichTx32.ocx descargado tiene que ser copiado a este directorio.", vbCritical, "Argentum Online Setup"
                Else
                    Set fsoObject = New FileSystemObject
                    fsoObject.CopyFile "RichTx32.ocx", fsoObject.GetSpecialFolder(SystemFolder) & "\", True
                    If Err Then MsgBox Err.Description
                    Shell "regsvr32 /s " & fsoObject.GetSpecialFolder(SystemFolder) & "\RichTx32.ocx"
                    MsgBox "Copia y registro realizados con exito.", vbOKOnly, "Argentum Online Setup"
                End If
            
            Case 3  ' CS
                If Not LibraryExist("CSWSK32.OCX", vbNormal) Then
                    MsgBox "ERROR, el archivo CSWSK32.OCX descargado tiene que ser copiado a este directorio.", vbCritical, "Argentum Online Setup"
                Else
                    Set fsoObject = New FileSystemObject
                    fsoObject.CopyFile "CSWSK32.OCX", fsoObject.GetSpecialFolder(SystemFolder) & "\", True
                    If Err Then MsgBox Err.Description
                    Shell "regsvr32 /s " & fsoObject.GetSpecialFolder(SystemFolder) & "\CSWSK32.OCX"
                    MsgBox "Copia y registro realizados con exito.", vbOKOnly, "Argentum Online Setup"
                End If
            
            Case 4  ' WS
                If Not LibraryExist("MSWINSCK.OCX", vbNormal) Then
                    MsgBox "ERROR, el archivo MSWINSCK.OCX descargado tiene que ser copiado a este directorio.", vbCritical, "Argentum Online Setup"
                Else
                    Set fsoObject = New FileSystemObject
                    fsoObject.CopyFile "MSWINSCK.OCX", fsoObject.GetSpecialFolder(SystemFolder) & "\", True
                    If Err Then MsgBox Err.Description
                    Shell "regsvr32 /s " & fsoObject.GetSpecialFolder(SystemFolder) & "\MSWINSCK.OCX"
                    MsgBox "Copia y registro realizados con exito.", vbOKOnly, "Argentum Online Setup"
                End If
            
            Case 5 'MSCOMCTL
                If Not LibraryExist("MSCOMCTL.OCX", vbNormal) Then
                    MsgBox "ERROR, el archivo MSCOMCTL.OCX descargado tiene que ser copiado a este directorio.", vbCritical, "Argentum Online Setup"
                Else
                    Set fsoObject = New FileSystemObject
                    fsoObject.CopyFile "MSCOMCTL.OCX", fsoObject.GetSpecialFolder(SystemFolder) & "\", True
                    If Err Then MsgBox Err.Description
                    Shell "regsvr32 /s " & fsoObject.GetSpecialFolder(SystemFolder) & "\MSCOMCTL.OCX"
                    MsgBox "Copia y registro realizados con exito.", vbOKOnly, "Argentum Online Setup"
                End If
        End Select
        DoEvents
        Call cVerificar_Click
    Else
        ' descargar
        If descargando = True Then
            MsgBox "Debes esperar a que se termine la descarga actual", vbCritical
            Exit Sub
        End If
        
        Dim rta As VbMsgBoxResult
        
        Select Case Index
            Case 0  ' inet
                ' revisar :O
                rta = MsgBox("Necesita descargar el archivo MSINET.OCX." & vbCrLf & _
                    "Es necesario que este archivo sea descargando manualmente y colocado en el directorio del juego, si esta de acuerdo precione Si", vbInformation + vbYesNo, "Soluci�n al problema")
                
                If rta = vbYes Then
                    Call ShellExecute(hwnd, "open", "http://ao.alkon.com.ar/descargas/MSINET.OCX", vbNullString, vbNullString, SW_SHOWNORMAL)
                End If
            
            Case 1  ' AA
                rta = MsgBox("Necesita descargar el archivo AAMD532.DLL." & vbCrLf & _
                    "Si desea descargarlo y registrarlo automaticamente precione Si.", vbYesNo, "Soluci�n al problema")
                
                    If rta = vbYes Then
                    'Bajarlo
                    descargando = True
                    
                    If ChkProxy.Value = 1 Then
                        Call DownloadForm.DownloadFile("http://ao.alkon.com.ar/descargas/AAMD532.DLL", "AAMD532.DLL", , , 2, txtProxy.Text)
                    Else
                        Call DownloadForm.DownloadFile("http://ao.alkon.com.ar/descargas/AAMD532.DLL", "AAMD532.DLL")
                    End If
                    
                    If (Not DownloadForm.DownloadSuccess) Or (DownloadForm.BotonCancel = True) Then
                       Unload DownloadForm
                       MsgBox "Descarga cancelada", vbInformation, "Error no solucionado"
                       Exit Sub
                    Else
                       Unload DownloadForm
                    End If
                    
                    descargando = False
                    
                    If FileExist("aamd532.dll", vbNormal) Then
                        If mod_MD5.MD5File("aamd532.dll") <> "cefd956a1ef122cda4d53007bab6c694" Then
                            MsgBox "No se puede comprobar la originalidad del archivo descargado, no se instalara.", vbCritical, "Error en MD5"
                            Exit Sub
                        Else
                            DoEvents
                            Call cVerificar_Click
                        End If
                    Else
                        MsgBox "No se pudo descargar el archivo", vbInformation, "Falta archivo"
                    End If
                End If
                
            Case 2  ' Rich
                rta = MsgBox("Necesita descargar el archivo RICHTX32.OCX." & vbCrLf & _
                    "Si desea descargarlo y registrarlo automaticamente precione Si.", vbYesNo, "Soluci�n al problema")
                
                If rta = vbYes Then
                    'Bajarlo
                    descargando = True
                    If ChkProxy.Value = 1 Then
                        Call DownloadForm.DownloadFile("http://ao.alkon.com.ar/descargas/RICHTX32.OCX", "RICHTX32.OCX", , , 2, txtProxy.Text)
                    Else
                        Call DownloadForm.DownloadFile("http://ao.alkon.com.ar/descargas/RICHTX32.OCX", "RICHTX32.OCX")
                    End If
                    
                    If (Not DownloadForm.DownloadSuccess) Or (DownloadForm.BotonCancel = True) Then
                       Unload DownloadForm
                       MsgBox "Descarga cancelada", vbInformation, "Error no solucionado"
                       Exit Sub
                    Else
                       Unload DownloadForm
                    End If
                    
                    descargando = False
                    
                    If FileExist("richtx32.ocx", vbNormal) Then
                        If mod_MD5.MD5File("richtx32.ocx") <> "722435ba4d18f1704b43e823a12e489a" Then
                            MsgBox "No se puede comprobar la originalidad del archivo descargado, no se instalara.", vbCritical, "Error en MD5"
                            Exit Sub
                        Else
                            DoEvents
                            Call cVerificar_Click
                        End If
                    Else
                        MsgBox "No se pudo descargar el archivo", vbInformation, "Falta archivo"
                    End If
                End If
            
            Case 3  ' CS
                rta = MsgBox("Necesita descargar el archivo CSWSK32.OCX." & Chr(10) & _
                    "Si desea descargarlo y registrarlo automaticamente precione Si.", vbYesNo, "Soluci�n al problema")
                
                If rta = vbYes Then
                    'Bajarlo
                    descargando = True
                    If ChkProxy.Value = 1 Then
                        Call DownloadForm.DownloadFile("http://ao.alkon.com.ar/descargas/CSWSK32.OCX", "CSWSK32.OCX", , , 2, txtProxy.Text)
                    Else
                        Call DownloadForm.DownloadFile("http://ao.alkon.com.ar/descargas/CSWSK32.OCX", "CSWSK32.OCX")
                    End If
                    
                    If (Not DownloadForm.DownloadSuccess) Or (DownloadForm.BotonCancel = True) Then
                       Unload DownloadForm
                       MsgBox "Descarga cancelada", vbInformation, "Error no solucionado"
                       Exit Sub
                    Else
                       Unload DownloadForm
                    End If
                    
                    descargando = False
                    
                    If FileExist("cswsk32.ocx", vbNormal) Then
                        If mod_MD5.MD5File("cswsk32.ocx") <> "5181704b2772e050e4a8331e15ee4bb4" Then
                            MsgBox "No se puede comprobar la originalidad del archivo descargado, no se instalara.", vbCritical, "Error en MD5"
                            Exit Sub
                        Else
                            DoEvents
                            Call cVerificar_Click
                        End If
                    Else
                        MsgBox "No se pudo descargar el archivo", vbInformation, "Falta archivo"
                    End If
                End If
            
            Case 4  ' WS
                rta = MsgBox("Necesita descargar el archivo MSWINSCK.OCX." & Chr(10) & _
                    "Si desea descargarlo y registrarlo automaticamente precione Si.", vbYesNo, "Soluci�n al problema")
                
                If rta = vbYes Then
                    'Bajarlo
                    descargando = True
                    
                    If ChkProxy.Value = 1 Then
                        Call DownloadForm.DownloadFile("http://ao.alkon.com.ar/descargas/MSWINSCK.OCX", "MSWINSCK.OCX", , , 2, txtProxy.Text)
                    Else
                        Call DownloadForm.DownloadFile("http://ao.alkon.com.ar/descargas/MSWINSCK.OCX", "MSWINSCK.OCX")
                    End If
                    
                    If (Not DownloadForm.DownloadSuccess) Or (DownloadForm.BotonCancel = True) Then
                       Unload DownloadForm
                       MsgBox "Descarga cancelada", vbInformation, "Error no solucionado"
                       Exit Sub
                    Else
                       Unload DownloadForm
                    End If
                    
                    descargando = False
                    
                    If FileExist("MSWINSCK.OCX", vbNormal) Then
                        If mod_MD5.MD5File("MSWINSCK.OCX") <> "3d8fd62d17a44221e07d5c535950449b" Then
                            MsgBox "No se puede comprobar la originalidad del archivo descargado, no se instalara.", vbCritical, "Error en MD5"
                            Exit Sub
                        Else
                            DoEvents
                            Call cVerificar_Click
                        End If
                    Else
                        MsgBox "No se pudo descargar el archivo", vbInformation, "Falta archivo"
                    End If
                End If
            
            Case 5 'MSCOMCTL
                rta = MsgBox("Necesita descargar el archivo MSCOMCTL.OCX." & Chr(10) & _
                    "Si desea descargarlo y registrarlo automaticamente precione Si.", vbYesNo, "Soluci�n al problema")
                
                If rta = vbYes Then
                    'Bajarlo
                    descargando = True
                    If ChkProxy.Value = 1 Then
                        Call DownloadFormSimple.DownloadFile("http://ao.alkon.com.ar/descargas/MSCOMCTL.OCX", "MSCOMCTL.OCX", , , 2, txtProxy.Text)
                    Else
                        Call DownloadFormSimple.DownloadFile("http://ao.alkon.com.ar/descargas/MSCOMCTL.OCX", "MSCOMCTL.OCX")
                    End If
                    
                    If (Not DownloadFormSimple.DownloadSuccess) Or (DownloadFormSimple.BotonCancel = True) Then
                       Unload DownloadFormSimple
                       MsgBox "Descarga cancelada", vbInformation, "Error no solucionado"
                       Exit Sub
                    Else
                       Unload DownloadFormSimple
                    End If
                    
                    descargando = False
                    
                    If FileExist("MSCOMCTL.OCX", vbNormal) Then
                        If mod_MD5.MD5File("MSCOMCTL.OCX") <> "ECC7D7F0D3446DE36045D1D9E964FAFE" Then
                            MsgBox "No se puede comprobar la originalidad del archivo descargado, no se instalara.", vbCritical, "Error en MD5"
                            Exit Sub
                        Else
                            DoEvents
                            Call cVerificar_Click
                        End If
                    Else
                        MsgBox "No se pudo descargar el archivo", vbInformation, "Falta archivo"
                    End If
                End If
        End Select
    End If
End Sub

Private Sub cVerificar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 04/11/08
'Last Modified by: NicoNZ
'Busca la existencia de la libreria "mscomctl.ocx"
'*************************************************
On Error Resume Next
    Err.Clear
    
    Load DownloadForm
    If Err Then
        If Not LibraryExist("mscomctl.ocx", vbNormal) Then
            Call LibError(5, "&Explorar")
        Else
            Call LibError(5, "&Registrar")
        End If
    Else
        Call LibOK(5)
    End If
        
    
    Load frmTestINET
    If Err Then
        If Not LibraryExist("msinet.ocx", vbNormal) Then
            Call LibError(0, "&Explorar")
        Else
            Call LibError(0, "&Registrar")
        End If
    Else
        Call LibOK(0)
    End If
    
    If Not LibraryExist("aamd532.dll", vbNormal) Then
        Call LibError(1, "&Descargar")
    Else
        Call LibOK(1)
    End If
    
    Err.Clear
    Load frmTestRICH
    If Err Then
        If Not LibraryExist("richtx32.ocx", vbNormal) Then
            Call LibError(2, "&Descargar")
        Else
            Call LibError(2, "&Registrar")
        End If
    Else
        Call LibOK(2)
    End If

    Err.Clear
    Load frmTestCS
    If Err Then
        If Not LibraryExist("cswsk32.ocx", vbNormal) Then
            Call LibError(3, "&Descargar")
        Else
            Call LibError(3, "&Registrar")
        End If
    Else
        Call LibOK(3)
    End If
    
    Err.Clear
    Load frmTestWS
    If Err Then
        If Not LibraryExist("mswinsck.ocx", vbNormal) Then
            Call LibError(4, "&Descargar")
        Else
            Call LibError(4, "&Registrar")
        End If
    Else
        Call LibOK(4)
    End If
    
End Sub

Private Sub Form_Load()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
Me.Show
DoEvents
Call cVerificar_Click
End Sub

