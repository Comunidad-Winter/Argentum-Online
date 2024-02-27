Attribute VB_Name = "NPCs"
'Argentum Online 0.9.0.2
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


'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Option Explicit

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

Dim I As Integer
UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
For I = 1 To MAXMASCOTAS
  If UserList(UserIndex).MascotasIndex(I) = NpcIndex Then
     UserList(UserIndex).MascotasIndex(I) = 0
     UserList(UserIndex).MascotasType(I) = 0
     Exit For
  End If
Next I

End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer, ByVal Mascota As Integer)

Dim I As Integer

Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1

'For i = 1 To UBound(Npclist(Maestro).Criaturas)
'  If Npclist(Maestro).Criaturas(i).NpcIndex = Mascota Then
'     Npclist(Maestro).Criaturas(i).NpcIndex = 0
'     Npclist(Maestro).Criaturas(i).NpcName = ""
'     Exit For
'  End If
'Next i

End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
On Error GoTo errhandler

'   Call LogTarea("Sub MuereNpc")
   
   Dim MiNPC As npc
   MiNPC = Npclist(NpcIndex)
   
   'Quitamos el npc
   Call QuitarNPC(NpcIndex)
   
   
    
   If UserIndex > 0 Then ' Lo mato un usuario?
        If MiNPC.flags.Snd3 > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & MiNPC.flags.Snd3)
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        
        'El user que lo mato tiene mascotas?
        If UserList(UserIndex).NroMacotas > 0 Then
                Dim T As Integer
                For T = 1 To MAXMASCOTAS
                      If UserList(UserIndex).MascotasIndex(T) > 0 Then
                          If Npclist(UserList(UserIndex).MascotasIndex(T)).TargetNpc = NpcIndex Then
                                  Call FollowAmo(UserList(UserIndex).MascotasIndex(T))
                          End If
                      End If
                Next T
        End If
        
        Call AddtoVar(UserList(UserIndex).Stats.Exp, MiNPC.GiveEXP, MAXEXP)
        Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & MiNPC.GiveEXP & " puntos de experiencia." & FONTTYPE_FIGHT)
        Call SendData(ToIndex, UserIndex, 0, "||Has matado la criatura!" & FONTTYPE_FIGHT)
        Call AddtoVar(UserList(UserIndex).Stats.NPCsMuertos, 1, 32000)
        
        If MiNPC.Stats.Alineacion = 0 Then
              If MiNPC.Numero = Guardias Then
                    Call VolverCriminal(UserIndex)
              End If
              If Not EsDios(UserList(UserIndex).Name) Then Call AddtoVar(UserList(UserIndex).Reputacion.AsesinoRep, vlASESINO, MAXREP)
        ElseIf MiNPC.Stats.Alineacion = 1 Then
          Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR, MAXREP)
        ElseIf MiNPC.Stats.Alineacion = 2 Then
          Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, vlASESINO / 2, MAXREP)
        ElseIf MiNPC.Stats.Alineacion = 4 Then
          Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR, MAXREP)
        End If
        If Not Criminal(UserIndex) And UserList(UserIndex).Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(UserIndex)
        
        'Controla el nivel del usuario
        Call CheckUserLevel(UserIndex)
       
   End If ' Userindex > 0

   
   If MiNPC.MaestroUser = 0 Then
        'Tiramos el oro
        Call NPCTirarOro(MiNPC)
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(MiNPC)
   End If
   
   'ReSpawn o no
   Call ReSpawnNpc(MiNPC)
   
Exit Sub

errhandler:
    Call LogError("Error en MuereNpc")
    
End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)
'Clear the npc's flags

Npclist(NpcIndex).flags.AfectaParalisis = 0
Npclist(NpcIndex).flags.AguaValida = 0
Npclist(NpcIndex).flags.AttackedBy = ""
Npclist(NpcIndex).flags.Attacking = 0
Npclist(NpcIndex).flags.BackUp = 0
Npclist(NpcIndex).flags.Bendicion = 0
Npclist(NpcIndex).flags.Domable = 0
Npclist(NpcIndex).flags.Envenenado = 0
Npclist(NpcIndex).flags.Faccion = 0
Npclist(NpcIndex).flags.Follow = False
Npclist(NpcIndex).flags.LanzaSpells = 0
Npclist(NpcIndex).flags.GolpeExacto = 0
Npclist(NpcIndex).flags.Invisible = 0
Npclist(NpcIndex).flags.Maldicion = 0
Npclist(NpcIndex).flags.OldHostil = 0
Npclist(NpcIndex).flags.OldMovement = 0
Npclist(NpcIndex).flags.Paralizado = 0
Npclist(NpcIndex).flags.Respawn = 0
Npclist(NpcIndex).flags.RespawnOrigPos = 0
Npclist(NpcIndex).flags.Snd1 = 0
Npclist(NpcIndex).flags.Snd2 = 0
Npclist(NpcIndex).flags.Snd3 = 0
Npclist(NpcIndex).flags.Snd4 = 0
Npclist(NpcIndex).flags.TierraInvalida = 0
Npclist(NpcIndex).flags.UseAINow = False

End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Contadores.Paralisis = 0
Npclist(NpcIndex).Contadores.TiempoExistencia = 0

End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Char.Body = 0
Npclist(NpcIndex).Char.CascoAnim = 0
Npclist(NpcIndex).Char.CharIndex = 0
Npclist(NpcIndex).Char.FX = 0
Npclist(NpcIndex).Char.Head = 0
Npclist(NpcIndex).Char.Heading = 0
Npclist(NpcIndex).Char.loops = 0
Npclist(NpcIndex).Char.ShieldAnim = 0
Npclist(NpcIndex).Char.WeaponAnim = 0


End Sub


Sub ResetNpcCriatures(ByVal NpcIndex As Integer)


Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroCriaturas
    Npclist(NpcIndex).Criaturas(j).NpcIndex = 0
    Npclist(NpcIndex).Criaturas(j).NpcName = ""
Next j

Npclist(NpcIndex).NroCriaturas = 0

End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)

Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroExpresiones: Npclist(NpcIndex).Expresiones(j) = "": Next j

Npclist(NpcIndex).NroExpresiones = 0

End Sub


Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Attackable = 0
Npclist(NpcIndex).CanAttack = 0
Npclist(NpcIndex).Comercia = 0
Npclist(NpcIndex).GiveEXP = 0
Npclist(NpcIndex).GiveGLD = 0
Npclist(NpcIndex).Hostile = 0
Npclist(NpcIndex).Inflacion = 0
Npclist(NpcIndex).InvReSpawn = 0
Npclist(NpcIndex).level = 0

If Npclist(NpcIndex).MaestroUser > 0 Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)

If Npclist(NpcIndex).MaestroNpc > 0 Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNpc, NpcIndex)

Npclist(NpcIndex).MaestroUser = 0
Npclist(NpcIndex).MaestroNpc = 0

Npclist(NpcIndex).Mascotas = 0
Npclist(NpcIndex).Movement = 0
Npclist(NpcIndex).Name = "NPC SIN INICIAR"
Npclist(NpcIndex).NPCtype = 0
Npclist(NpcIndex).Numero = 0
Npclist(NpcIndex).Orig.Map = 0
Npclist(NpcIndex).Orig.X = 0
Npclist(NpcIndex).Orig.Y = 0
Npclist(NpcIndex).PoderAtaque = 0
Npclist(NpcIndex).PoderEvasion = 0
Npclist(NpcIndex).Pos.Map = 0
Npclist(NpcIndex).Pos.X = 0
Npclist(NpcIndex).Pos.Y = 0
Npclist(NpcIndex).SkillDomar = 0
Npclist(NpcIndex).Target = 0
Npclist(NpcIndex).TargetNpc = 0
Npclist(NpcIndex).TipoItems = 0
Npclist(NpcIndex).Veneno = 0
Npclist(NpcIndex).Desc = ""

Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroSpells
    Npclist(NpcIndex).Spells(j) = 0
Next j

Call ResetNpcCharInfo(NpcIndex)
Call ResetNpcCriatures(NpcIndex)
Call ResetExpresiones(NpcIndex)

End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer)

On Error GoTo errhandler

Npclist(NpcIndex).flags.NPCActive = False


If InMapBounds(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then
    Call EraseNPCChar(ToMap, 0, Npclist(NpcIndex).Pos.Map, NpcIndex)
End If

'Nos aseguramos de que el inventario sea removido...
'asi los lobos no volveran a tirar armaduras ;))
Call ResetNpcInv(NpcIndex)
Call ResetNpcFlags(NpcIndex)
Call ResetNpcCounters(NpcIndex)

Call ResetNpcMainInfo(NpcIndex)

If NpcIndex = LastNPC Then
    Do Until Npclist(LastNPC).flags.NPCActive
        LastNPC = LastNPC - 1
        If LastNPC < 1 Then Exit Do
    Loop
End If
    
  
If NumNPCs <> 0 Then
    NumNPCs = NumNPCs - 1
End If

Exit Sub

errhandler:
    Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(Pos As WorldPos) As Boolean


If LegalPos(Pos.Map, Pos.X, Pos.Y) Then
    TestSpawnTrigger = _
    MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 3 And _
    MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 2 And _
    MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 1
End If

End Function

Sub CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)
'Call LogTarea("Sub CrearNPC")
'Crea un NPC del tipo NRONPC

Dim Pos As WorldPos
Dim newpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long


Dim Map As Integer
Dim X As Integer
Dim Y As Integer

nIndex = OpenNPC(NroNPC) 'Conseguimos un indice

If nIndex > MAXNPCS Then Exit Sub

'Necesita ser respawned en un lugar especifico
If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
    
    Map = OrigPos.Map
    X = OrigPos.X
    Y = OrigPos.Y
    Npclist(nIndex).Orig = OrigPos
    Npclist(nIndex).Pos = OrigPos
   
Else
    
    Pos.Map = mapa 'mapa
    
    Do While Not PosicionValida
    
        Randomize (Timer)
        Pos.X = CInt(Rnd * 100 + 1) 'Obtenemos posicion al azar en x
        Pos.Y = CInt(Rnd * 100 + 1) 'Obtenemos posicion al azar en y
        
        Call ClosestLegalPos(Pos, newpos)  'Nos devuelve la posicion valida mas cercana
        
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, Npclist(nIndex).flags.AguaValida) And _
           Not HayPCarea(newpos) And TestSpawnTrigger(newpos) Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(nIndex).Pos.Map = newpos.Map
            Npclist(nIndex).Pos.X = newpos.X
            Npclist(nIndex).Pos.Y = newpos.Y
            PosicionValida = True
        Else
            newpos.X = 0
            newpos.Y = 0
        
        End If
            
        'for debug
        Iteraciones = Iteraciones + 1
        If Iteraciones > MAXSPAWNATTEMPS Then
                Call QuitarNPC(nIndex)
                Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNPC)
                Exit Sub
        End If
    Loop
    
    'asignamos las nuevas coordenas
    Map = newpos.Map
    X = Npclist(nIndex).Pos.X
    Y = Npclist(nIndex).Pos.Y
End If

'Crea el NPC
Call MakeNPCChar(ToMap, 0, Map, nIndex, Map, X, Y)

End Sub

Sub MakeNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)


Dim CharIndex As Integer

If Npclist(NpcIndex).Char.CharIndex = 0 Then
    CharIndex = NextOpenCharIndex
    Npclist(NpcIndex).Char.CharIndex = CharIndex
    CharList(CharIndex) = NpcIndex
End If


MapData(Map, X, Y).NpcIndex = NpcIndex

Call SendData(sndRoute, sndIndex, sndMap, "CC" & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading & "," & Npclist(NpcIndex).Char.CharIndex & "," & X & "," & Y)

End Sub

Sub ChangeNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, NpcIndex As Integer, Body As Integer, Head As Integer, Heading As Byte)

If NpcIndex > 0 Then
    Npclist(NpcIndex).Char.Body = Body
    Npclist(NpcIndex).Char.Head = Head
    Npclist(NpcIndex).Char.Heading = Heading
    Call SendData(sndRoute, sndIndex, sndMap, "CP" & Npclist(NpcIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading)
End If

End Sub

Sub EraseNPCChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

If Npclist(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar < 1 Then Exit Do
    Loop
End If

'Quitamos del mapa
MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

'Actualizamos los cliente
Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "BP" & Npclist(NpcIndex).Char.CharIndex)

'Update la lista npc
Npclist(NpcIndex).Char.CharIndex = 0


'update NumChars
NumChars = NumChars - 1


End Sub

Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)

On Error GoTo errh
    Dim nPos As WorldPos
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(nHeading, nPos)
    
    'Es mascota ????
    If Npclist(NpcIndex).MaestroUser > 0 Then
            ' es una posicion legal
            If LegalPos(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then
            
                If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
                If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
                
                '[Alejo-18-5]
                Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
                'Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).Pos.Map, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
            
                'Update map and user pos
                MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
                Npclist(NpcIndex).Pos = nPos
                Npclist(NpcIndex).Char.Heading = nHeading
                MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex
            End If
    Else ' No es mascota
            ' Controlamos que la posicion sea legal, los npc que
            ' no son mascotas tienen mas restricciones de movimiento.
            If LegalPosNPC(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y, Npclist(NpcIndex).flags.AguaValida) Then
                
                If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
                If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.Map, nPos.X, nPos.Y) Then Exit Sub
                
                '[Alejo-18-5]
                Call SendData(ToMap, 0, Npclist(NpcIndex).Pos.Map, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
                'Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).Pos.Map, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
                
                'Update map and user pos
                MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
                Npclist(NpcIndex).Pos = nPos
                Npclist(NpcIndex).Char.Heading = nHeading
                MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex
            Else
                If Npclist(NpcIndex).Movement = NPC_PATHFINDING Then
                    'Someone has blocked the npc's way, we must to seek a new path!
                    Npclist(NpcIndex).PFINFO.PathLenght = 0
                End If
            
            End If
    End If

Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)


End Sub

Function NextOpenNPC() As Integer
'Call LogTarea("Sub NextOpenNPC")

On Error GoTo errhandler

Dim LoopC As Integer
  
For LoopC = 1 To MAXNPCS + 1
    If LoopC > MAXNPCS Then Exit For
    If Not Npclist(LoopC).flags.NPCActive Then Exit For
Next LoopC
  
NextOpenNPC = LoopC


Exit Function
errhandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)

Dim N As Integer
N = RandomNumber(1, 100)
If N < 30 Then
    UserList(UserIndex).flags.Envenenado = 1
    Call SendData(ToIndex, UserIndex, 0, "||¡¡La criatura te ha envenenado!!" & FONTTYPE_FIGHT)
End If

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
'Crea un NPC del tipo Npcindex
'Call LogTarea("Sub SpawnNpc")

Dim newpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean


Dim Map As Integer
Dim X As Integer
Dim Y As Integer
Dim it As Integer

nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

it = 0

If nIndex > MAXNPCS Then
    SpawnNpc = nIndex
    Exit Function
End If

Do While Not PosicionValida
        
        Call ClosestLegalPos(Pos, newpos)  'Nos devuelve la posicion valida mas cercana
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
        If LegalPos(newpos.Map, newpos.X, newpos.Y) Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(nIndex).Pos.Map = newpos.Map
            Npclist(nIndex).Pos.X = newpos.X
            Npclist(nIndex).Pos.Y = newpos.Y
            PosicionValida = True
        Else
            newpos.X = 0
            newpos.Y = 0
        End If
        
        it = it + 1
        
        If it > MAXSPAWNATTEMPS Then
            Call QuitarNPC(nIndex)
            SpawnNpc = MAXNPCS
            Call LogError("Mas de " & MAXSPAWNATTEMPS & " iteraciones en SpawnNpc Mapa:" & Pos.Map & " Index:" & NpcIndex)
            Exit Function
        End If
Loop
    
'asignamos las nuevas coordenas
Map = newpos.Map
X = Npclist(nIndex).Pos.X
Y = Npclist(nIndex).Pos.Y

'Crea el NPC
Call MakeNPCChar(ToMap, 0, Map, nIndex, Map, X, Y)

If FX Then
    Call SendData(ToMap, 0, Map, "TW" & SND_WARP)
    Call SendData(ToMap, 0, Map, "CFX" & Npclist(nIndex).Char.CharIndex & "," & FXWARP & "," & 0)
End If

SpawnNpc = nIndex

End Function

Sub ReSpawnNpc(MiNPC As npc)

If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)

End Sub

'Devuelve el nro de enemigos que hay en el Mapa Map
Function NPCHostiles(ByVal Map As Integer) As Integer

Dim NpcIndex As Integer
Dim cont As Integer

'Contador
cont = 0
For NpcIndex = 1 To LastNPC

    '¿esta vivo?
    If Npclist(NpcIndex).flags.NPCActive _
       And Npclist(NpcIndex).Pos.Map = Map _
       And Npclist(NpcIndex).Hostile = 1 And _
       Npclist(NpcIndex).Stats.Alineacion = 2 Then
            cont = cont + 1
           
    End If
    
Next NpcIndex

NPCHostiles = cont

End Function

Sub NPCTirarOro(MiNPC As npc)

'SI EL NPC TIENE ORO LO TIRAMOS
If MiNPC.GiveGLD > 0 Then
    Dim MiObj As Obj
    MiObj.Amount = MiNPC.GiveGLD
    MiObj.ObjIndex = iORO
    Call TirarItemAlPiso(MiNPC.Pos, MiObj)
End If

End Sub



Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer

Dim NpcIndex As Integer
'Dim npcfile As String
Dim A As Long, S As Long

If NpcNumber > 499 Then
'        npcfile = DatPath & "NPCs-HOSTILES.dat"
    A = Anpc_host
Else
'        npcfile = DatPath & "NPCs.dat"
    A = ANpc
End If

S = INIBuscarSeccion(A, "NPC" & NpcNumber)

NpcIndex = NextOpenNPC

If NpcIndex > MAXNPCS Then 'Limite de npcs
    OpenNPC = NpcIndex
    Exit Function
End If

Npclist(NpcIndex).Numero = NpcNumber
'Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
'Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
'
'Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
'Npclist(NpcIndex).Flags.OldMovement = Npclist(NpcIndex).Movement

Npclist(NpcIndex).Name = INIDarClaveStr(A, S, "Name")
Npclist(NpcIndex).Desc = INIDarClaveStr(A, S, "Desc")

Npclist(NpcIndex).Movement = INIDarClaveInt(A, S, "Movement")
Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

Npclist(NpcIndex).flags.AguaValida = INIDarClaveInt(A, S, "AguaValida")
Npclist(NpcIndex).flags.TierraInvalida = INIDarClaveInt(A, S, "TierraInValida")
Npclist(NpcIndex).flags.Faccion = INIDarClaveInt(A, S, "Faccion")

Npclist(NpcIndex).NPCtype = INIDarClaveInt(A, S, "NpcType")

Npclist(NpcIndex).Char.Body = INIDarClaveInt(A, S, "Body")
Npclist(NpcIndex).Char.Head = INIDarClaveInt(A, S, "Head")
Npclist(NpcIndex).Char.Heading = INIDarClaveInt(A, S, "Heading")

Npclist(NpcIndex).Attackable = INIDarClaveInt(A, S, "Attackable")
Npclist(NpcIndex).Comercia = INIDarClaveInt(A, S, "Comercia")
Npclist(NpcIndex).Hostile = INIDarClaveInt(A, S, "Hostile")
Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile

Npclist(NpcIndex).GiveEXP = INIDarClaveInt(A, S, "GiveEXP")

Npclist(NpcIndex).Veneno = INIDarClaveInt(A, S, "Veneno")

Npclist(NpcIndex).flags.Domable = INIDarClaveInt(A, S, "Domable")


Npclist(NpcIndex).GiveGLD = INIDarClaveInt(A, S, "GiveGLD")

Npclist(NpcIndex).PoderAtaque = INIDarClaveInt(A, S, "PoderAtaque")
Npclist(NpcIndex).PoderEvasion = INIDarClaveInt(A, S, "PoderEvasion")

Npclist(NpcIndex).InvReSpawn = INIDarClaveInt(A, S, "InvReSpawn")


Npclist(NpcIndex).Stats.MaxHP = INIDarClaveInt(A, S, "MaxHP")
Npclist(NpcIndex).Stats.MinHP = INIDarClaveInt(A, S, "MinHP")
Npclist(NpcIndex).Stats.MaxHIT = INIDarClaveInt(A, S, "MaxHIT")
Npclist(NpcIndex).Stats.MinHIT = INIDarClaveInt(A, S, "MinHIT")
Npclist(NpcIndex).Stats.Def = INIDarClaveInt(A, S, "DEF")
Npclist(NpcIndex).Stats.Alineacion = INIDarClaveInt(A, S, "Alineacion")
Npclist(NpcIndex).Stats.ImpactRate = INIDarClaveInt(A, S, "ImpactRate")


Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = INIDarClaveInt(A, S, "NROITEMS")
For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = INIDarClaveStr(A, S, "Obj" & LoopC)
    Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
Next LoopC

Npclist(NpcIndex).flags.LanzaSpells = INIDarClaveInt(A, S, "LanzaSpells")
If Npclist(NpcIndex).flags.LanzaSpells > 0 Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
For LoopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
    Npclist(NpcIndex).Spells(LoopC) = INIDarClaveInt(A, S, "Sp" & LoopC)
Next LoopC


If Npclist(NpcIndex).NPCtype = NPCTYPE_ENTRENADOR Then
    Npclist(NpcIndex).NroCriaturas = INIDarClaveInt(A, S, "NroCriaturas")
    ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
    For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
        Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = INIDarClaveInt(A, S, "CI" & LoopC)
        Npclist(NpcIndex).Criaturas(LoopC).NpcName = INIDarClaveStr(A, S, "CN" & LoopC)
    Next LoopC
End If


Npclist(NpcIndex).Inflacion = INIDarClaveInt(A, S, "Inflacion")

Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False

If Respawn Then
    Npclist(NpcIndex).flags.Respawn = INIDarClaveInt(A, S, "ReSpawn")
Else
    Npclist(NpcIndex).flags.Respawn = 1
End If

Npclist(NpcIndex).flags.BackUp = INIDarClaveInt(A, S, "BackUp")
Npclist(NpcIndex).flags.RespawnOrigPos = INIDarClaveInt(A, S, "OrigPos")
Npclist(NpcIndex).flags.AfectaParalisis = INIDarClaveInt(A, S, "AfectaParalisis")
Npclist(NpcIndex).flags.GolpeExacto = INIDarClaveInt(A, S, "GolpeExacto")


Npclist(NpcIndex).flags.Snd1 = INIDarClaveInt(A, S, "Snd1")
Npclist(NpcIndex).flags.Snd2 = INIDarClaveInt(A, S, "Snd2")
Npclist(NpcIndex).flags.Snd3 = INIDarClaveInt(A, S, "Snd3")
Npclist(NpcIndex).flags.Snd4 = INIDarClaveInt(A, S, "Snd4")

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

Dim aux As Long
aux = INIDarClaveInt(A, S, "NROEXP")
Npclist(NpcIndex).NroExpresiones = (aux)
    
If aux > 0 Then
    ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
    For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
        Npclist(NpcIndex).Expresiones(LoopC) = INIDarClaveStr(A, S, "Exp" & LoopC)
    Next LoopC
End If

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = INIDarClaveInt(A, S, "TipoItems")

'Update contadores de NPCs
If NpcIndex > LastNPC Then LastNPC = NpcIndex
NumNPCs = NumNPCs + 1


'Devuelve el nuevo Indice
OpenNPC = NpcIndex

End Function


Function OpenNPC_Viejo(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer

Dim NpcIndex As Integer
Dim npcfile As String

If NpcNumber > 499 Then
        npcfile = DatPath & "NPCs-HOSTILES.dat"
Else
        npcfile = DatPath & "NPCs.dat"
End If


NpcIndex = NextOpenNPC

If NpcIndex > MAXNPCS Then 'Limite de npcs
    OpenNPC_Viejo = NpcIndex
    Exit Function
End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")

Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

Npclist(NpcIndex).flags.AguaValida = val(GetVar(npcfile, "NPC" & NpcNumber, "AguaValida"))
Npclist(NpcIndex).flags.TierraInvalida = val(GetVar(npcfile, "NPC" & NpcNumber, "TierraInValida"))
Npclist(NpcIndex).flags.Faccion = val(GetVar(npcfile, "NPC" & NpcNumber, "Faccion"))

Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile

Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))

Npclist(NpcIndex).Veneno = val(GetVar(npcfile, "NPC" & NpcNumber, "Veneno"))

Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))


Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))

Npclist(NpcIndex).PoderAtaque = val(GetVar(npcfile, "NPC" & NpcNumber, "PoderAtaque"))
Npclist(NpcIndex).PoderEvasion = val(GetVar(npcfile, "NPC" & NpcNumber, "PoderEvasion"))

Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))


Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.Def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
Npclist(NpcIndex).Stats.ImpactRate = val(GetVar(npcfile, "NPC" & NpcNumber, "ImpactRate"))


Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
    Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
Next LoopC

Npclist(NpcIndex).flags.LanzaSpells = val(GetVar(npcfile, "NPC" & NpcNumber, "LanzaSpells"))
If Npclist(NpcIndex).flags.LanzaSpells > 0 Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
For LoopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
    Npclist(NpcIndex).Spells(LoopC) = val(GetVar(npcfile, "NPC" & NpcNumber, "Sp" & LoopC))
Next LoopC


If Npclist(NpcIndex).NPCtype = NPCTYPE_ENTRENADOR Then
    Npclist(NpcIndex).NroCriaturas = val(GetVar(npcfile, "NPC" & NpcNumber, "NroCriaturas"))
    ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
    For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
        Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = GetVar(npcfile, "NPC" & NpcNumber, "CI" & LoopC)
        Npclist(NpcIndex).Criaturas(LoopC).NpcName = GetVar(npcfile, "NPC" & NpcNumber, "CN" & LoopC)
    Next LoopC
End If


Npclist(NpcIndex).Inflacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Inflacion"))

Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False

If Respawn Then
    Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Else
    Npclist(NpcIndex).flags.Respawn = 1
End If

Npclist(NpcIndex).flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))
Npclist(NpcIndex).flags.AfectaParalisis = val(GetVar(npcfile, "NPC" & NpcNumber, "AfectaParalisis"))
Npclist(NpcIndex).flags.GolpeExacto = val(GetVar(npcfile, "NPC" & NpcNumber, "GolpeExacto"))


Npclist(NpcIndex).flags.Snd1 = val(GetVar(npcfile, "NPC" & NpcNumber, "Snd1"))
Npclist(NpcIndex).flags.Snd2 = val(GetVar(npcfile, "NPC" & NpcNumber, "Snd2"))
Npclist(NpcIndex).flags.Snd3 = val(GetVar(npcfile, "NPC" & NpcNumber, "Snd3"))
Npclist(NpcIndex).flags.Snd4 = val(GetVar(npcfile, "NPC" & NpcNumber, "Snd4"))

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

Dim aux As String
aux = GetVar(npcfile, "NPC" & NpcNumber, "NROEXP")
If aux = "" Then
    Npclist(NpcIndex).NroExpresiones = 0
Else
    Npclist(NpcIndex).NroExpresiones = val(aux)
    ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
    For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
        Npclist(NpcIndex).Expresiones(LoopC) = GetVar(npcfile, "NPC" & NpcNumber, "Exp" & LoopC)
    Next LoopC
End If

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

'Update contadores de NPCs
If NpcIndex > LastNPC Then LastNPC = NpcIndex
NumNPCs = NumNPCs + 1


'Devuelve el nuevo Indice
OpenNPC_Viejo = NpcIndex

End Function

Sub EnviarListaCriaturas(ByVal UserIndex As Integer, ByVal NpcIndex)
  Dim SD As String
  Dim k As Integer
  SD = SD & Npclist(NpcIndex).NroCriaturas & ","
  For k = 1 To Npclist(NpcIndex).NroCriaturas
        SD = SD & Npclist(NpcIndex).Criaturas(k).NpcName & ","
  Next k
  SD = "LSTCRI" & SD
  Call SendData(ToIndex, UserIndex, 0, SD)
End Sub


Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)

If Npclist(NpcIndex).flags.Follow Then
  Npclist(NpcIndex).flags.AttackedBy = ""
  Npclist(NpcIndex).flags.Follow = False
  Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
  Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
Else
  Npclist(NpcIndex).flags.AttackedBy = UserName
  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = 4 'follow
  Npclist(NpcIndex).Hostile = 0
End If

End Sub

Sub FollowAmo(ByVal NpcIndex As Integer)

  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = SIGUE_AMO 'follow
  Npclist(NpcIndex).Hostile = 0
  Npclist(NpcIndex).Target = 0
  Npclist(NpcIndex).TargetNpc = 0

End Sub

