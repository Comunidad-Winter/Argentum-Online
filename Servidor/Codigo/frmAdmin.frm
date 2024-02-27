VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Administraci�n del servidor"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Personajes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   495
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Echar todos los PJS no privilegiados"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   2280
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "R"
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.ComboBox cboPjs 
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Echar"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   1800
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.20
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez


Option Explicit


Private Sub cboPjs_Change()
Call ActualizaPjInfo
End Sub

Private Sub cboPjs_Click()
Call ActualizaPjInfo
End Sub

Private Sub Command1_Click()
Dim tIndex As Long

tIndex = NameIndex(cboPjs.Text)
If tIndex > 0 Then
    Call SendData(ToAll, 0, 0, "||Servidor> " & UserList(tIndex).Name & " ha sido hechado. " & FONTTYPE_SERVER)
    Call CloseSocket(tIndex)
End If

End Sub

Public Sub ActualizaListaPjs()
Dim LoopC As Long

With cboPjs
    .Clear
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
            If UserList(LoopC).flags.Privilegios < 1 Then
                .AddItem UserList(LoopC).Name
                .ItemData(.NewIndex) = LoopC
            End If
        End If
    Next LoopC
End With

End Sub

Private Sub Command3_Click()
Call EcharPjsNoPrivilegiados

End Sub

Private Sub Label1_Click()
Call ActualizaPjInfo

End Sub

Private Sub ActualizaPjInfo()
Dim tIndex As Long

tIndex = NameIndex(cboPjs.Text)
If tIndex > 0 Then
    With UserList(tIndex)
        Text1.Text = .ColaSalida.Count & " elementos en cola." & vbCrLf
    End With
End If

End Sub
