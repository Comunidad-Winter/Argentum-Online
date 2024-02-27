VERSION 5.00
Begin VB.Form frmUserList 
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Echar todos los no Logged"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualiza"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmUserList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.20
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


Option Explicit

Private Sub Command1_Click()
Dim LoopC As Integer

Text2.Text = "MaxUsers: " & MaxUsers & vbCrLf
Text2.Text = Text2.Text & "LastUser: " & LastUser & vbCrLf
Text2.Text = Text2.Text & "NumUsers: " & NumUsers & vbCrLf
'Text2.Text = Text2.Text & "" & vbCrLf

List1.Clear

For LoopC = 1 To MaxUsers
    List1.AddItem Format(LoopC, "000") & " " & IIf(UserList(LoopC).flags.UserLogged, UserList(LoopC).Name, "")
    List1.ItemData(List1.NewIndex) = LoopC
Next LoopC


End Sub

Private Sub Command2_Click()
Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 And Not UserList(LoopC).flags.UserLogged Then
        Call CloseSocket(LoopC)
    End If
Next LoopC

End Sub

Private Sub List1_Click()
Dim UserIndex As Integer
If List1.ListIndex <> -1 Then
    UserIndex = List1.ItemData(List1.ListIndex)
    If UserIndex > 0 And UserIndex <= MaxUsers Then
        With UserList(UserIndex)
            Text1.Text = "UserLogged: " & .flags.UserLogged & vbCrLf
            Text1.Text = Text1.Text & "IdleCount: " & .Counters.IdleCount & vbCrLf
            Text1.Text = Text1.Text & "ConnId: " & .ConnID & vbCrLf
            Text1.Text = Text1.Text & "ConnIDValida: " & .ConnIDValida & vbCrLf
        End With
    End If
End If

End Sub
