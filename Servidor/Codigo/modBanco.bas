Attribute VB_Name = "modBanco"
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

'MODULO PROGRAMADO POR NEB
'Kevin Birmingham
'kbneb@hotmail.com

Sub IniciarDeposito(ByVal UserIndex As Integer)
On Error GoTo errhandler

'Hacemos un Update del inventario del usuario
Call UpdateBanUserInv(True, UserIndex, 0)
'Atcualizamos el dinero
Call SendUserStatsBox(UserIndex)
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
SendData ToIndex, UserIndex, 0, "INITBANCO"
UserList(UserIndex).flags.Comerciando = True

errhandler:

End Sub

Sub SendBanObj(UserIndex As Integer, Slot As Byte, Object As UserOBJ)


UserList(UserIndex).BancoInvent.Object(Slot) = Object

If Object.ObjIndex > 0 Then

    Call SendData(ToIndex, UserIndex, 0, "SBO" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
    & ObjData(Object.ObjIndex).ObjType & "," _
    & ObjData(Object.ObjIndex).MaxHIT & "," _
    & ObjData(Object.ObjIndex).MinHIT & "," _
    & ObjData(Object.ObjIndex).MaxDef)

Else

    Call SendData(ToIndex, UserIndex, 0, "SBO" & Slot & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")

End If


End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex > 0 Then
        Call SendBanObj(UserIndex, Slot, UserList(UserIndex).BancoInvent.Object(Slot))
    Else
        Call SendBanObj(UserIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex > 0 Then
            Call SendBanObj(UserIndex, LoopC, UserList(UserIndex).BancoInvent.Object(LoopC))
        Else
            
            Call SendBanObj(UserIndex, LoopC, NullObj)
            
        End If

    Next LoopC

End If

End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)
On Error GoTo errhandler


If Cantidad < 1 Then Exit Sub


Call SendUserStatsBox(UserIndex)

   
       If UserList(UserIndex).BancoInvent.Object(i).Amount > 0 Then
            If Cantidad > UserList(UserIndex).BancoInvent.Object(i).Amount Then Cantidad = UserList(UserIndex).BancoInvent.Object(i).Amount
            'Agregamos el obj que compro al inventario
            Call UserReciveObj(UserIndex, CInt(i), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el banco
            Call UpdateBanUserInv(True, UserIndex, 0)
            'Actualizamos la ventana de comercio
            Call UpdateVentanaBanco(i, 0, UserIndex)
       End If



errhandler:

End Sub

Sub UserReciveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim Slot As Integer
Dim obji As Integer


If UserList(UserIndex).BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub

obji = UserList(UserIndex).BancoInvent.Object(ObjIndex).ObjIndex


'�Ya tiene un objeto de este tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji And _
   UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
    
    Slot = Slot + 1
    If Slot > MAX_INVENTORY_SLOTS Then
        Exit Do
    End If
Loop

'Sino se fija por un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(ToIndex, UserIndex, 0, "||No pod�s tener mas objetos." & FONTTYPE_INFO)
                Exit Sub
            End If
        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If



'Mete el obj en el slot
If UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
    'Menor que MAX_INV_OBJS
    UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad
    
    Call QuitarBancoInvItem(UserIndex, CByte(ObjIndex), Cantidad)
Else
    Call SendData(ToIndex, UserIndex, 0, "||No pod�s tener mas objetos." & FONTTYPE_INFO)
End If


End Sub

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)



Dim ObjIndex As Integer
ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex

    'Quita un Obj

       UserList(UserIndex).BancoInvent.Object(Slot).Amount = UserList(UserIndex).BancoInvent.Object(Slot).Amount - Cantidad
        
        If UserList(UserIndex).BancoInvent.Object(Slot).Amount <= 0 Then
            UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems - 1
            UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = 0
            UserList(UserIndex).BancoInvent.Object(Slot).Amount = 0
        End If

    
    
End Sub

Sub UpdateVentanaBanco(ByVal Slot As Integer, ByVal NpcInv As Byte, ByVal UserIndex As Integer)
 
 
 Call SendData(ToIndex, UserIndex, 0, "BANCOOK" & Slot & "," & NpcInv)
 
End Sub

Sub UserDepositaItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

On Error GoTo errhandler

'El usuario deposita un item
Call SendUserStatsBox(UserIndex)
   
If UserList(UserIndex).Invent.Object(Item).Amount > 0 And UserList(UserIndex).Invent.Object(Item).Equipped = 0 Then
            
            If Cantidad > 0 And Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
            'Agregamos el obj que compro al inventario
            Call UserDejaObj(UserIndex, CInt(Item), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el inventario del banco
            Call UpdateBanUserInv(True, UserIndex, 0)
            'Actualizamos la ventana del banco
            
            Call UpdateVentanaBanco(Item, 1, UserIndex)
            
End If

errhandler:

End Sub

Sub UserDejaObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim Slot As Integer
Dim obji As Integer

If Cantidad < 1 Then Exit Sub

obji = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex

'�Ya tiene un objeto de este tipo?
Slot = 1
Do Until UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = obji And _
         UserList(UserIndex).BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
        
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
Loop

'Sino se fija por un slot vacio antes del slot devuelto
If Slot > MAX_BANCOINVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Call SendData(ToIndex, UserIndex, 0, "||No tienes mas espacio en el banco!!" & FONTTYPE_INFO)
                Exit Sub
                Exit Do
            End If
        Loop
        If Slot <= MAX_BANCOINVENTORY_SLOTS Then UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1
        
        
End If

If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
    'Mete el obj en el slot
    If UserList(UserIndex).BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        
        'Menor que MAX_INV_OBJS
        UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = obji
        UserList(UserIndex).BancoInvent.Object(Slot).Amount = UserList(UserIndex).BancoInvent.Object(Slot).Amount + Cantidad
        
        Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)

    Else
        Call SendData(ToIndex, UserIndex, 0, "||El banco no puede cargar tantos objetos." & FONTTYPE_INFO)
    End If

Else
    Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
End If

End Sub

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos." & FONTTYPE_INFO)
For j = 1 To MAX_BANCOINVENTORY_SLOTS
    If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
        Call SendData(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(j).Amount & FONTTYPE_INFO)
    End If
Next

End Sub

Sub SendUserBovedaTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & CharName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call SendData(ToIndex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
    Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "BancoInventory", "CantidadItems") & " objetos." & FONTTYPE_INFO)
    For j = 1 To MAX_BANCOINVENTORY_SLOTS
        Tmp = GetVar(CharFile, "BancoInventory", "Obj" & j)
        ObjInd = ReadField(1, Tmp, Asc("-"))
        ObjCant = ReadField(2, Tmp, Asc("-"))
        If ObjInd > 0 Then
            Call SendData(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant & FONTTYPE_INFO)
        End If
    Next
Else
    Call SendData(ToIndex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)
End If

End Sub

