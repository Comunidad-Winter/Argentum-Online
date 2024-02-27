Attribute VB_Name = "TCP_HandleData2"
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

Public Sub HandleData_2(ByVal UserIndex As Integer, rdata As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim Name As String
Dim ind
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim t() As String
Dim i As Integer

Procesado = True 'ver al final del sub


    Select Case UCase$(rdata)
    
        Case "/ONLINE"
            N = 0
            tStr = ""
            For LoopC = 1 To LastUser
                If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios <= 1 Then
                    N = N + 1
                    tStr = tStr & UserList(LoopC).Name & ", "
                End If
            Next LoopC
            If Len(tStr) > 2 Then
                tStr = Left(tStr, Len(tStr) - 2)
            End If
            Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
            Call SendData(ToIndex, UserIndex, 0, "||N�mero de usuarios: " & N & FONTTYPE_INFO)
            Exit Sub
        Case "/SALIR"
            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes salir estando paralizado." & FONTTYPE_WARNING)
                Exit Sub
            End If
            ''mato los comercios seguros
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                        Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                    End If
                End If
                Call SendData(ToIndex, UserIndex, 0, "||Comercio cancelado. " & FONTTYPE_TALK)
                Call FinComerciarUsu(UserIndex)
            End If
            Cerrar_Usuario (UserIndex)
            Exit Sub
    ''    Case "/SALIRCLAN"
    ''        If UserList(UserIndex).GuildInfo.GuildName <> "" Then
    ''            Call EacharMember(UserIndex, UserList(UserIndex).Name)
    ''            UserList(UserIndex).GuildInfo.GuildName = ""
    ''            UserList(UserIndex).GuildInfo.EsGuildLeader = 0
    ''        End If
    ''        Exit Sub
        Case "/FUNDARCLAN"
            If UserList(UserIndex).GuildInfo.FundoClan = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||Ya has fundado un clan, solo se puede fundar uno por personaje." & FONTTYPE_INFO)
                Exit Sub
            End If
            If CanCreateGuild(UserIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "SHOWFUN" & FONTTYPE_INFO)
            End If
            Exit Sub
            
        '[Barrin 1-12-03]
        Case "/SALIRCLAN"
            If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Eres l�der de un clan, no puedes salir del mismo." & FONTTYPE_INFO)
                      Exit Sub
            ElseIf UserList(UserIndex).GuildInfo.GuildName = "" Then
                      Call SendData(ToIndex, UserIndex, 0, "||No perteneces a ning�n clan." & FONTTYPE_INFO)
                      Exit Sub
            Else
                Call SendData(ToGuildMembers, UserIndex, 0, "||" & UserList(UserIndex).Name & " decidi� dejar al clan." & FONTTYPE_GUILD)
                Dim oGuild As cGuild
                Set oGuild = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)
                Call oGuild.RemoveMember(UserList(UserIndex).Name)
                Call AddtoVar(UserList(UserIndex).GuildInfo.Echadas, 1, 1000)
                UserList(UserIndex).GuildInfo.GuildPoints = 0
                UserList(UserIndex).GuildInfo.GuildName = ""
            '''''''''''''''''
            End If
            Exit Sub
        '[/Barrin 1-12-03]
            
        Case "/BALANCE"
            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||��Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                      Exit Sub
            End If
            Select Case Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype
            Case NPCTYPE_BANQUERO
                If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
                      Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                      CloseSocket (UserIndex)
                      Exit Sub
                End If
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "�" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & FONTTYPE_INFO)
            Case NPCTYPE_TIMBERO
                If UserList(UserIndex).flags.Privilegios > 0 Then
                    tLong = Apuestas.Ganancias - Apuestas.Perdidas
                    N = 0
                    If tLong >= 0 And Apuestas.Ganancias <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Ganancias)
                    End If
                    If tLong < 0 And Apuestas.Perdidas <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Perdidas)
                    End If
                    Call SendData(ToIndex, UserIndex, 0, "||Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & tLong & " (" & N & "%) Jugadas: " & Apuestas.Jugadas & FONTTYPE_INFO)
                End If
            End Select
            Exit Sub
        Case "/QUIETO" ' << Comando a mascotas
             '�Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                          Call SendData(ToIndex, UserIndex, 0, "||��Estas muerto!!" & FONTTYPE_INFO)
                          Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                      Exit Sub
             End If
             If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                          Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                          Exit Sub
             End If
             If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
                UserIndex Then Exit Sub
             Npclist(UserList(UserIndex).flags.TargetNPC).Movement = ESTATICO
             Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
             Exit Sub
        Case "/ACOMPA�AR"
            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||��Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
              UserIndex Then Exit Sub
            Call FollowAmo(UserList(UserIndex).flags.TargetNPC)
            Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
            Exit Sub
        Case "/ENTRENAR"
            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||��Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Sub
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Exit Sub
        Case "/DESCANSAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||��Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If HayOBJarea(UserList(UserIndex).Pos, FOGATA) Then
                    Call SendData(ToIndex, UserIndex, 0, "DOK")
                    If Not UserList(UserIndex).flags.Descansar Then
                        Call SendData(ToIndex, UserIndex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                    End If
                    UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
            Else
                    If UserList(UserIndex).flags.Descansar Then
                        Call SendData(ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                        
                        UserList(UserIndex).flags.Descansar = False
                        Call SendData(ToIndex, UserIndex, 0, "DOK")
                        Exit Sub
                    End If
                    Call SendData(ToIndex, UserIndex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/MEDITAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||��Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).Stats.MaxMAN = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Solo las clases m�gicas conocen el arte de la meditaci�n" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).flags.Privilegios > 0 Then
                UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                Call SendData(ToIndex, UserIndex, 0, "||Mana restaurado" & FONTTYPE_VENENO)
                Call SendUserStatsBox(val(UserIndex))
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, 0, "MEDOK")
            If Not UserList(UserIndex).flags.Meditando Then
               Call SendData(ToIndex, UserIndex, 0, "||Comenzas a meditar." & FONTTYPE_INFO)
            Else
               Call SendData(ToIndex, UserIndex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
            End If
           UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando
            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(UserIndex).flags.Meditando Then
                UserList(UserIndex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                Call SendData(ToIndex, UserIndex, 0, "||Te est�s concentrando. En " & TIEMPO_INICIOMEDITAR & " segundos comenzar�s a meditar." & FONTTYPE_INFO)
                
                UserList(UserIndex).Char.loops = LoopAdEternum
                If UserList(UserIndex).Stats.ELV < 15 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & FXMEDITARCHICO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARCHICO
                ElseIf UserList(UserIndex).Stats.ELV < 30 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & FXMEDITARMEDIANO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARMEDIANO
                ElseIf UserList(UserIndex).Stats.ELV < 45 Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & FXMEDITARGRANDE & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARGRANDE
                Else
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & FXMEDITARXGRANDE & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXMEDITARXGRANDE
                End If
            Else
                UserList(UserIndex).Counters.bPuedeMeditar = False
                
                UserList(UserIndex).Char.FX = 0
                UserList(UserIndex).Char.loops = 0
                Call SendData(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.charindex & "," & 0 & "," & 0)
            End If
            Exit Sub
        Case "/RESUCITAR"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 1 _
           Or UserList(UserIndex).flags.Muerto <> 1 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede resucitarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
               Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
               CloseSocket (UserIndex)
               Exit Sub
           End If
           Call RevivirUsuario(UserIndex)
           Call SendData(ToIndex, UserIndex, 0, "||��H�s sido resucitado!!" & FONTTYPE_INFO)
           Exit Sub
        Case "/CURAR"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 1 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
           Call SendUserStatsBox(val(UserIndex))
           Call SendData(ToIndex, UserIndex, 0, "||��H�s sido curado!!" & FONTTYPE_INFO)
           Exit Sub
        Case "/AYUDA"
           Call SendHelp(UserIndex)
           Exit Sub
                  
        Case "/EST"
            Call SendUserStatsTxt(UserIndex, UserIndex)
            Exit Sub
        
        Case "/SEG"
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(ToIndex, UserIndex, 0, "SEGOFF")
            Else
                Call SendData(ToIndex, UserIndex, 0, "SEGON")
            End If
            UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            Exit Sub
    
    
        Case "/COMERCIAR"
            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||��Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            
            If UserList(UserIndex).flags.Comerciando Then
                    Call SendData(ToIndex, UserIndex, 0, "||Ya est�s comerciando" & FONTTYPE_INFO)
                    Exit Sub
            End If
            
            If UserList(UserIndex).flags.Privilegios = 1 Then
                Exit Sub
            End If
            '�El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                  '�El NPC puede comerciar?
                  If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                     If Len(Npclist(UserList(UserIndex).flags.TargetNPC).Desc) > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "�" & "No tengo ningun interes en comerciar." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                     Exit Sub
                  End If
                  If Npclist(UserList(UserIndex).flags.TargetNPC).Name = "SR" Then
                     If UserList(UserIndex).Faccion.ArmadaReal <> 1 Then
                        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "�" & "Muestra tu bandera antes de comprar ropa del ej�rcito" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                        Exit Sub
                     End If
                  End If
                  If Npclist(UserList(UserIndex).flags.TargetNPC).Name = "SC" Then
                     If UserList(UserIndex).Faccion.FuerzasCaos <> 1 Then
                        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbRed & "�" & "�Vete de aqu�!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                        Exit Sub
                     End If
                  End If
                  If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                      Exit Sub
                  End If
                  'Iniciamos la rutina pa' comerciar.
                  Call IniciarCOmercioNPC(UserIndex)
             '[Alejo]
            ElseIf UserList(UserIndex).flags.TargetUser > 0 Then
            
                'Call SendData(ToIndex, UserIndex, 0, "||COMERCIO SEGURO ENTRE USUARIOS TEMPORALMENTE DESHABILITADO" & FONTTYPE_INFO)
                'Exit Sub
                'Comercio con otro usuario
                'Puede comerciar ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||��No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                'soy yo ?
                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(ToIndex, UserIndex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'ta muy lejos ?
                If Distancia(UserList(UserList(UserIndex).flags.TargetUser).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del usuario." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'Ya ta comerciando ? es con migo o con otro ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                    Call SendData(ToIndex, UserIndex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'inicializa unas variables...
                UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).Name
                UserList(UserIndex).ComUsu.Cant = 0
                UserList(UserIndex).ComUsu.Objeto = 0
                UserList(UserIndex).ComUsu.Acepto = False
                
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
        '[/Alejo]
        '[KEVIN]------------------------------------------
        Case "/BOVEDA"
            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||��Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            '�El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                  If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                      Exit Sub
                  End If
                  If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = 4 Then
                    Call IniciarDeposito(UserIndex)
                  Else
                    Exit Sub
                  End If
            Else
              Call SendData(ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
        '[/KEVIN]------------------------------------
    
        Case "/ENLISTAR"
            'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(ToIndex, UserIndex, 0, "||Debes acercarte m�s." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(UserIndex)
           Else
                  Call EnlistarCaos(UserIndex)
           End If
           
           Exit Sub
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "No perteneces a las tropas reales!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    Exit Sub
                End If
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "No perteneces a la legi�n oscura!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    Exit Sub
                End If
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
           End If
           Exit Sub
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "No perteneces a las tropas reales!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    Exit Sub
                End If
                Call RecompensaArmadaReal(UserIndex)
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "No perteneces a la legi�n oscura!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    Exit Sub
                End If
                Call RecompensaCaos(UserIndex)
           End If
           Exit Sub
           
        Case "/MOTD"
            Call SendMOTD(UserIndex)
            Exit Sub
            
        Case "/UPTIME"
            tLong = Int(((GetTickCount() And &H7FFFFFFF) - tInicioServer) / 1000)
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(ToIndex, UserIndex, 0, "||Uptime: " & tStr & FONTTYPE_INFO)
            
            tLong = IntervaloAutoReiniciar
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(ToIndex, UserIndex, 0, "||Pr�ximo mantenimiento autom�tico: " & tStr & FONTTYPE_INFO)
            
            Exit Sub
        
        Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(UserIndex)
            Exit Sub
        
'        Case "/PARTY"
'            Dim iRet As Integer, NewMember As Integer

'            UserList(UserIndex).PartyData.TargetUser = UserList(UserIndex).flags.TargetUser
'            NewMember = UserList(UserIndex).PartyData.TargetUser

'            If UserList(NewMember).flags.TargetUser = UserIndex Then _
'                iRet = NewUserInParty(UserIndex, NewMember)

'            Select Case iRet
                
'                Case 1
                    '�xito!
'                    Call SendData(ToIndex, UserIndex, 0, "||Has a�adido un nuevo miembro a tu grupo" & FONTTYPE_INFO)
'                    Call SendData(ToIndex, NewMember, 0, "||Has ingresado al grupo" & FONTTYPE_INFO)

'                Case PartyERR.NoEsLider
'                    Call SendData(ToIndex, UserIndex, 0, "||Solo el l�der puede invitar gente a unirse al grupo" & FONTTYPE_INFO)

'                Case PartyERR.ArmadaProhibe
'                    Call SendData(ToIndex, UserIndex, 0, "||Un soldado de las filas reales no puede aliarse con criminales" & FONTTYPE_INFO)
'                    Call SendData(ToIndex, NewMember, 0, "||Un soldado de las filas reales no puede aliarse con criminales" & FONTTYPE_INFO)

'                Case PartyERR.LegionProhibe
'                    Call SendData(ToIndex, UserIndex, 0, "||Un legionario oscuro no puede aliarse con ciudadanos" & FONTTYPE_INFO)
'                    Call SendData(ToIndex, NewMember, 0, "||Un legionario oscuro no puede aliarse con ciudadanos" & FONTTYPE_INFO)

'                Case PartyERR.NivelProhibe
'                    Call SendData(ToIndex, UserIndex, 0, "||Hay demasiada diferencia de experiencia" & FONTTYPE_INFO)
'                    Call SendData(ToIndex, NewMember, 0, "||Hay demasiada diferencia de experiencia" & FONTTYPE_INFO)
                
'                Case PartyERR.PrivilegiosProhiben
'                    Call SendData(ToIndex, UserIndex, 0, "||No puedes aliarte con administradores del juego" & FONTTYPE_INFO)
'                    Call SendData(ToIndex, NewMember, 0, "||No puedes aliarte con administradores del juego" & FONTTYPE_INFO)
            
'                Case PartyERR.DemasiadoLejos
'                    Call SendData(ToIndex, UserIndex, 0, "||Debes acercarte a esa persona para concretar la alianza" & FONTTYPE_INFO)
'                    Call SendData(ToIndex, NewMember, 0, "||Debes acercarte a esa persona para concretar la alianza" & FONTTYPE_INFO)

'                Case PartyERR.ESTAMUERTO
'                    Call SendData(ToIndex, UserIndex, 0, "||No puedes aliarte con los muertos!" & FONTTYPE_INFO)
'                    Call SendData(ToIndex, NewMember, 0, "||No puedes aliarte con los muertos!" & FONTTYPE_INFO)
'            End Select
'            Exit Sub
        Case "/CREARPARTY"
            If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
            Call mdParty.CrearParty(UserIndex)
            Exit Sub
        Case "/PARTY"
            Call mdParty.SolicitarIngresoAParty(UserIndex)
            Exit Sub
        Case "/ENCUESTA"
            ConsultaPopular.SendInfoEncuesta (UserIndex)
        
'        Case "/ONLINEPARTY"
'            Dim sRet As String

'            sRet = GetPartyList(UserList(UserIndex).PartyData.PIndex)
'            Call SendData(ToIndex, UserIndex, 0, "||Miembros de tu grupo: " & sRet & FONTTYPE_INFO)
    End Select
    
    
    
    If UCase$(Left$(rdata, 6)) = "/CMSG " Then
    
        If Len(UserList(UserIndex).GuildInfo.GuildName) = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||No perteneces a ningun clan " & FONTTYPE_INFO)
                Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 6)
        If rdata <> "" And UserList(UserIndex).GuildInfo.GuildName <> "" Then
            'Call SendData(ToGuildMembers, UserIndex, 0, "||" & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_GUILDMSG)
            Call SendData(ToGuildMembers, UserIndex, 0, "||" & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_GUILDMSG)
            Call SendData(ToClanArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "�< " & rdata & " >�" & str(UserList(UserIndex).Char.charindex))
        End If
        Exit Sub
    End If
    
    If UCase$(Left$(rdata, 6)) = "/PMSG " Then
        Call mdParty.BroadCastParty(UserIndex, Mid$(rdata, 7))
        Exit Sub
    End If
    
    If UCase$(rdata) = "/ONLINECLAN" Then
    
        If UserList(UserIndex).GuildInfo.GuildName = "" Then Exit Sub
    
        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") And UserList(LoopC).GuildInfo.GuildName = UserList(UserIndex).GuildInfo.GuildName Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||Usuarios de tu clan conectados: " & tStr & FONTTYPE_GUILDMSG)
        
        Exit Sub
    
    End If
    
    If UCase$(rdata) = "/ONLINEPARTY" Then
        Call mdParty.OnlineParty(UserIndex)
        Exit Sub
    End If
    
    '[yb]
    If UCase$(Left$(rdata, 6)) = "/BMSG " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        If UserList(UserIndex).flags.PertAlCons = 1 Then
            Call SendData(ToConsejo, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_CONSEJO)
        End If
        If UserList(UserIndex).flags.PertAlConsCaos = 1 Then
            Call SendData(ToConsejoCaos, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).Name & "> " & rdata & FONTTYPE_CONSEJOCAOS)
        End If
        Exit Sub
    End If
    '[/yb]
    
    If UCase$(Left$(rdata, 5)) = "/ROL " Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Call SendData(ToRolesMasters, 0, 0, "|| " & LCase$(UserList(UserIndex).Name) & " PREGUNTA ROL: " & rdata & FONTTYPE_GUILDMSG)
        Exit Sub
    End If
    
    
    'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rdata, 6)) = "/GMSG " And UserList(UserIndex).flags.Privilegios > 0 Then
        rdata = Right$(rdata, Len(rdata) - 6)
        Call LogGM(UserList(UserIndex).Name, "Mensaje a Gms:" & rdata, (UserList(UserIndex).flags.Privilegios = 1))
        If rdata <> "" Then
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & "> " & rdata & "~255~255~255~0~1")
        End If
        Exit Sub
    End If
    
    Select Case UCase$(Left$(rdata, 3))
        Case "/GM"
            If Not Ayuda.Existe(UserList(UserIndex).Name) Then
                Call SendData(ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                Call Ayuda.Push(rdata, UserList(UserIndex).Name)
            Else
                Call Ayuda.Quitar(UserList(UserIndex).Name)
                Call Ayuda.Push(rdata, UserList(UserIndex).Name)
                Call SendData(ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
            End If
            Exit Sub
    End Select
    
    
    
    Select Case UCase(Left(rdata, 5))
        Case "/_BUG "
            N = FreeFile
            Open App.Path & "\LOGS\BUGs.log" For Append Shared As N
            Print #N,
            Print #N,
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N, "Usuario:" & UserList(UserIndex).Name & "  Fecha:" & Date & "    Hora:" & Time
            Print #N, "########################################################################"
            Print #N, "BUG:"
            Print #N, Right$(rdata, Len(rdata) - 5)
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N,
            Print #N,
            Close #N
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rdata, 6))
        Case "/DESC "
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes cambiar la descripci�n estando muerto." & FONTTYPE_INFO)
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 6)
            If Not AsciiValidos(rdata) Then
                Call SendData(ToIndex, UserIndex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Sub
            End If
            UserList(UserIndex).Desc = rdata
            Call SendData(ToIndex, UserIndex, 0, "||La descripcion a cambiado." & FONTTYPE_INFO)
            Exit Sub
        Case "/VOTO "
                rdata = Right$(rdata, Len(rdata) - 6)
                Call ComputeVote(UserIndex, rdata)
                Exit Sub
    End Select
    
    If UCase$(Left$(rdata, 7)) = "/PENAS " Then
        Name = Right$(rdata, Len(rdata) - 7)
        If Name = "" Then Exit Sub
        
        Name = Replace(Name, "\", "")
        Name = Replace(Name, "/", "")
        
        If FileExist(CharPath & Name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
            If tInt = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Sin prontuario.." & FONTTYPE_INFO)
            Else
                While tInt > 0
                    Call SendData(ToIndex, UserIndex, 0, "||" & tInt & "- " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & tInt) & FONTTYPE_INFO)
                    tInt = tInt - 1
                Wend
            End If
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Personaje """ & Name & """ inexistente." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    
    
    Select Case UCase$(Left$(rdata, 8))
        Case "/PASSWD "
            rdata = Right$(rdata, Len(rdata) - 8)
            If Len(rdata) < 6 Then
                 Call SendData(ToIndex, UserIndex, 0, "||El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
            Else
                 Call SendData(ToIndex, UserIndex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
                 UserList(UserIndex).Password = rdata
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 9))
            'Comando /APOSTAR basado en la idea de DarkLight,
            'pero con distinta probabilidad de exito.
        Case "/APOSTAR "
            rdata = Right(rdata, Len(rdata) - 9)
            tLong = CLng(val(rdata))
            If tLong > 32000 Then tLong = 32000
            N = tLong
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||��Estas muerto!!" & FONTTYPE_INFO)
            ElseIf UserList(UserIndex).flags.TargetNPC = 0 Then
                'Se asegura que el target es un npc
                Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
            ElseIf Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
            ElseIf Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_TIMBERO Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "No tengo ningun interes en apostar." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
            ElseIf N < 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "El minimo de apuesta es 1 moneda." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
            ElseIf N > 5000 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "El maximo de apuesta es 5000 monedas." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
            ElseIf UserList(UserIndex).Stats.GLD < N Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "No tienes esa cantidad." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
            Else
                If RandomNumber(1, 100) <= 47 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + N
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Felicidades! Has ganado " & CStr(N) & " monedas de oro!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    
                    Apuestas.Perdidas = Apuestas.Perdidas + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - N
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Lo siento, has perdido " & CStr(N) & " monedas de oro." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                
                    Apuestas.Ganancias = Apuestas.Ganancias + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                End If
                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
                
                Call SendUserStatsBox(UserIndex)
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 10))
            'consultas populares muchacho'
        Case "/ENCUESTA "
            rdata = Right(rdata, Len(rdata) - 10)
            If Len(rdata) = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "|| Aca va la info de la encuesta" & FONTTYPE_GUILD)
                Exit Sub
            End If
            DummyInt = CLng(val(rdata))
            Call SendData(ToIndex, UserIndex, 0, "|| " & ConsultaPopular.doVotar(UserIndex, DummyInt) & FONTTYPE_GUILD)
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rdata, 8))
        Case "/RETIRAR" 'RETIRA ORO EN EL BANCO o te saca de la armada
             '�Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||��Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
             End If
             
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = 5 Then
                
                'Se quiere retirar de la armada
                If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                        Call ExpulsarFaccionReal(UserIndex)
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ser�s bienvenido a las fuerzas imperiales si deseas regresar." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                        Debug.Print "||" & vbWhite & "�" & "Ser�s bienvenido a las fuerzas imperiales si deseas regresar." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "���Sal de aqu� buf�n!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    End If
                ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 1 Then
                        Call ExpulsarFaccionCaos(UserIndex)
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ya volver�s arrastrandote." & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Sal de aqu� maldito criminal" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "�No perteneces a ninguna fuerza!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex))
                End If
                Exit Sub
             
             End If
             
             If Len(rdata) = 8 Then
                Call SendData(ToIndex, UserIndex, 0, "||Debes indicar el monto de cuanto quieres retirar" & FONTTYPE_INFO)
                Exit Sub
             End If
             
             rdata = Right$(rdata, Len(rdata) - 9)
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_BANQUERO _
             Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
             If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
             End If
             If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
                  Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                  CloseSocket (UserIndex)
                  Exit Sub
             End If
             If val(rdata) > 0 And val(rdata) <= UserList(UserIndex).Stats.Banco Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(rdata)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rdata)
                  Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "�" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & FONTTYPE_INFO)
             Else
                  Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & " No tenes esa cantidad." & "�" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & FONTTYPE_INFO)
             End If
             Call SendUserStatsBox(val(UserIndex))
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rdata, 11))
        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
            '�Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||��Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 11)
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> NPCTYPE_BANQUERO _
            Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
            If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If CLng(val(rdata)) > 0 And CLng(val(rdata)) <= UserList(UserIndex).Stats.GLD Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rdata)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rdata)
                  Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "�" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & FONTTYPE_INFO)
            Else
                  Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "�" & " No tenes esa cantidad." & "�" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.charindex & FONTTYPE_INFO)
            End If
            Call SendUserStatsBox(val(UserIndex))
            Exit Sub
        Case "/DENUNCIAR "
            If UserList(UserIndex).flags.Silenciado = 1 Then
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 11)
            Call SendData(ToAdmins, 0, 0, "|| " & LCase$(UserList(UserIndex).Name) & " DENUNCIA: " & rdata & FONTTYPE_GUILDMSG)
            Call SendData(ToIndex, UserIndex, 0, "|| Denuncia enviada, espere.." & FONTTYPE_INFO)
            Exit Sub
    End Select

    Select Case UCase$(Left$(rdata, 12))
        Case "/ECHARPARTY "
            rdata = Right$(rdata, Len(rdata) - 12)
            tInt = NameIndex(rdata)
            If tInt > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tInt)
            Else
                Call SendData(ToIndex, UserIndex, 0, "|| El personaje no est� online." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/PARTYLIDER "
            rdata = Right$(rdata, Len(rdata) - 12)
            tInt = NameIndex(rdata)
            If tInt > 0 Then
                Call mdParty.TransformarEnLider(UserIndex, tInt)
            Else
                Call SendData(ToIndex, UserIndex, 0, "|| El personaje no est� online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rdata, 13))
        Case "/ACCEPTPARTY "
            rdata = Right$(rdata, Len(rdata) - 13)
            tInt = NameIndex(rdata)
            If tInt > 0 Then
                Call mdParty.AprobarIngresoAParty(UserIndex, tInt)
            Else
                Call SendData(ToIndex, UserIndex, 0, "|| El personaje no est� online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select
    

    Select Case UCase$(Left$(rdata, 14))
        Case "/MIEMBROSCLAN "
            rdata = Trim(Right(rdata, Len(rdata) - 14))
            Name = Replace(rdata, "\", "")
            Name = Replace(rdata, "/", "")
    
            If Not FileExist(App.Path & "\guilds\" & rdata & "-members.mem") Then
                Call SendData(ToIndex, UserIndex, 0, "|| No existe el clan: " & rdata & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tInt = val(GetVar(App.Path & "\Guilds\" & rdata & "-Members" & ".mem", "INIT", "NroMembers"))
            
            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rdata & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call SendData(ToIndex, UserIndex, 0, "||" & tStr & "<" & rdata & ">." & FONTTYPE_INFO)
            Next i
        
            Exit Sub
        End Select



Procesado = False

End Sub


