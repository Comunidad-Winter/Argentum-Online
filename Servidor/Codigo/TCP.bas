Attribute VB_Name = "TCP"
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
Option Explicit

'Buffer en bytes de cada socket
Public Const SOCKET_BUFFER_SIZE = 2048

'Cuantos comandos de cada cliente guarda el server
Public Const COMMAND_BUFFER_SIZE = 1000

Public Const NingunArma = 2

'RUTAS DE ENVIO DE DATOS
Public Const ToIndex = 0 'Envia a un solo User
Public Const ToAll = 1 'A todos los Users
Public Const ToMap = 2 'Todos los Usuarios en el mapa
Public Const ToPCArea = 3 'Todos los Users en el area de un user determinado
Public Const ToNone = 4 'Ninguno
Public Const ToAllButIndex = 5 'Todos menos el index
Public Const ToMapButIndex = 6 'Todos en el mapa menos el indice
Public Const ToGM = 7
Public Const ToNPCArea = 8 'Todos los Users en el area de un user determinado
Public Const ToGuildMembers = 9
Public Const ToAdmins = 10
Public Const ToPCAreaButIndex = 11

#If Not (UsarAPI = 1) Then
' General constants used with most of the controls
Public Const INVALID_HANDLE = -1
Public Const CONTROL_ERRIGNORE = 0
Public Const CONTROL_ERRDISPLAY = 1


' SocietWrench Control Actions
Public Const SOCKET_OPEN = 1
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_LISTEN = 3
Public Const SOCKET_ACCEPT = 4
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_CLOSE = 7
Public Const SOCKET_DISCONNECT = 7
Public Const SOCKET_ABORT = 8

' SocketWrench Control States
Public Const SOCKET_NONE = 0
Public Const SOCKET_IDLE = 1
Public Const SOCKET_LISTENING = 2
Public Const SOCKET_CONNECTING = 3
Public Const SOCKET_ACCEPTING = 4
Public Const SOCKET_RECEIVING = 5
Public Const SOCKET_SENDING = 6
Public Const SOCKET_CLOSING = 7

' Societ Address Families
Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2

' Societ Types
Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5

' Protocol Types
Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_GGP = 2
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77
Public Const IPPROTO_RAW = 255
Public Const IPPROTO_MAX = 256


' Network Addpesses
Public Const INADDR_ANY = "0.0.0.0"
Public Const INADDR_LOOPBACK = "127.0.0.1"
Public Const INADDR_NONE = "255.055.255.255"

' Shutdown Values
Public Const SOCKET_READ = 0
Public Const SOCKET_WRITE = 1
Public Const SOCKET_READWRITE = 2

' SocketWrench Error Pesponse
Public Const SOCKET_ERRIGNORE = 0
Public Const SOCKET_ERRDISPLAY = 1

' SocketWrench Error Aodes
Public Const WSABASEERR = 24000
Public Const WSAEINTR = 24004
Public Const WSAEBADF = 24009
Public Const WSAEACCES = 24013
Public Const WSAEFAULT = 24014
Public Const WSAEINVAL = 24022
Public Const WSAEMFILE = 24024
Public Const WSAEWOULDBLOCK = 24035
Public Const WSAEINPROGRESS = 24036
Public Const WSAEALREADY = 24037
Public Const WSAENOTSOCK = 24038
Public Const WSAEDESTADDRREQ = 24039
Public Const WSAEMSGSIZE = 24040
Public Const WSAEPROTOTYPE = 24041
Public Const WSAENOPROTOOPT = 24042
Public Const WSAEPROTONOSUPPORT = 24043
Public Const WSAESOCKTNOSUPPORT = 24044
Public Const WSAEOPNOTSUPP = 24045
Public Const WSAEPFNOSUPPORT = 24046
Public Const WSAEAFNOSUPPORT = 24047
Public Const WSAEADDRINUSE = 24048
Public Const WSAEADDRNOTAVAIL = 24049
Public Const WSAENETDOWN = 24050
Public Const WSAENETUNREACH = 24051
Public Const WSAENETRESET = 24052
Public Const WSAECONNABORTED = 24053
Public Const WSAECONNRESET = 24054
Public Const WSAENOBUFS = 24055
Public Const WSAEISCONN = 24056
Public Const WSAENOTCONN = 24057
Public Const WSAESHUTDOWN = 24058
Public Const WSAETOOMANYREFS = 24059
Public Const WSAETIMEDOUT = 24060
Public Const WSAECONNREFUSED = 24061
Public Const WSAELOOP = 24062
Public Const WSAENAMETOOLONG = 24063
Public Const WSAEHOSTDOWN = 24064
Public Const WSAEHOSTUNREACH = 24065
Public Const WSAENOTEMPTY = 24066
Public Const WSAEPROCLIM = 24067
Public Const WSAEUSERS = 24068
Public Const WSAEDQUOT = 24069
Public Const WSAESTALE = 24070
Public Const WSAEREMOTE = 24071
Public Const WSASYSNOTREADY = 24091
Public Const WSAVERNOTSUPPORTED = 24092
Public Const WSANOTINITIALISED = 24093
Public Const WSAHOST_NOT_FOUND = 25001
Public Const WSATRY_AGAIN = 25002
Public Const WSANO_RECOVERY = 25003
Public Const WSANO_DATA = 25004
Public Const WSANO_ADDRESS = 2500
#End If

'Esta funcion calcula el CRC de cada paquete que se
'envía al servidor.

Public Function GenCrC(ByVal Key As Long, ByVal sdData As String) As Long

End Function



Sub DarCuerpoYCabeza(UserBody As Integer, UserHead As Integer, Raza As String, Gen As String)

Select Case Gen
   Case "Hombre"
        Select Case Raza
        
                Case "Humano"
                    UserHead = CInt(RandomNumber(1, 11))
                    If UserHead > 11 Then UserHead = 11
                    UserBody = 1
                Case "Elfo"
                    UserHead = CInt(RandomNumber(1, 4)) + 100
                    If UserHead > 104 Then UserHead = 104
                    UserBody = 2
                Case "Elfo Oscuro"
                    UserHead = CInt(RandomNumber(1, 3)) + 200
                    If UserHead > 203 Then UserHead = 203
                    UserBody = 3
                Case "Enano"
                    UserHead = RandomNumber(1, 1) + 300
                    If UserHead > 301 Then UserHead = 301
                    UserBody = 52
                Case "Gnomo"
                    UserHead = RandomNumber(1, 1) + 400
                    If UserHead > 401 Then UserHead = 401
                    UserBody = 52
                Case Else
                    UserHead = 1
                    UserBody = 1
            
        End Select
   Case "Mujer"
        Select Case Raza
                Case "Humano"
                    UserHead = CInt(RandomNumber(1, 3)) + 69
                    If UserHead > 72 Then UserHead = 72
                    UserBody = 1
                Case "Elfo"
                    UserHead = CInt(RandomNumber(1, 3)) + 169
                    If UserHead > 172 Then UserHead = 172
                    UserBody = 2
                Case "Elfo Oscuro"
                    UserHead = CInt(RandomNumber(1, 3)) + 269
                    If UserHead > 272 Then UserHead = 272
                    UserBody = 3
                Case "Gnomo"
                    UserHead = RandomNumber(1, 2) + 469
                    If UserHead > 471 Then UserHead = 471
                    UserBody = 52
                Case "Enano"
                    UserHead = 370
                    UserBody = 52
                Case Else
                    UserHead = 70
                    UserBody = 1
        End Select
End Select

   
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function


Function NombrePermitido(ByVal Nombre As String) As Boolean
Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
    End If
Next i

NombrePermitido = True

End Function

Function ValidateAtrib(ByVal UserIndex As Integer) As Boolean
Dim LoopC As Integer

For LoopC = 1 To NUMATRIBUTOS
    If UserList(UserIndex).Stats.UserAtributos(LoopC) > 18 Or UserList(UserIndex).Stats.UserAtributos(LoopC) < 1 Then Exit Function
Next LoopC

ValidateAtrib = True

End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    

End Function

Sub ConnectNewUser(UserIndex As Integer, Name As String, Password As String, Body As Integer, Head As Integer, UserRaza As String, UserSexo As String, UserClase As String, _
UA1 As String, UA2 As String, UA3 As String, UA4 As String, UA5 As String, _
US1 As String, US2 As String, US3 As String, US4 As String, US5 As String, _
US6 As String, US7 As String, US8 As String, US9 As String, US10 As String, _
US11 As String, US12 As String, US13 As String, US14 As String, US15 As String, _
US16 As String, US17 As String, US18 As String, US19 As String, US20 As String, _
US21 As String, UserEmail As String, Hogar As String)

If Not NombrePermitido(Name) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRLos nombres de los personajes deben pertencer a la fantasia, el nombre indicado es invalido.")
    Exit Sub
End If

If Not AsciiValidos(Name) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRNombre invalido.")
    Exit Sub
End If

Dim LoopC As Integer
Dim totalskpts As Long
  
'¿Existe el personaje?
If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
    Call SendData(ToIndex, UserIndex, 0, "ERRYa existe el personaje.")
    Exit Sub
End If

UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).flags.Escondido = 0



UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.BurguesRep = 0
UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.NobleRep = 1000
UserList(UserIndex).Reputacion.PlebeRep = 30

UserList(UserIndex).Reputacion.Promedio = 30 / 6


UserList(UserIndex).Name = Name
UserList(UserIndex).Clase = UserClase
UserList(UserIndex).Raza = UserRaza
UserList(UserIndex).Genero = UserSexo
UserList(UserIndex).Email = UserEmail
UserList(UserIndex).Hogar = Hogar

'UserList(UserIndex).Stats.UserAtributos(Fuerza) = Abs(CInt(UA1))
'UserList(UserIndex).Stats.UserAtributos(Inteligencia) = Abs(CInt(UA2))
'UserList(UserIndex).Stats.UserAtributos(Agilidad) = Abs(CInt(UA3))
'UserList(UserIndex).Stats.UserAtributos(Carisma) = Abs(CInt(UA4))
'UserList(UserIndex).Stats.UserAtributos(Constitucion) = Abs(CInt(UA5))


'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS ATRIBUTOS %%%%%%%%%%%%%
If Not ValidateAtrib(UserIndex) Then
        Call SendData(ToIndex, UserIndex, 0, "ERRAtributos invalidos.")
        Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS ATRIBUTOS %%%%%%%%%%%%%

Select Case UCase$(UserRaza)
    Case "HUMANO"
        UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) + 2
        UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) + 1
        UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 2
        UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 1
    Case "ELFO"
        UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) + 2
        UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 2
        UserList(UserIndex).Stats.UserAtributos(Carisma) = UserList(UserIndex).Stats.UserAtributos(Carisma) + 2
    Case "ELFO OSCURO"
        UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) + 1
        UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) + 2
        UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 2
        UserList(UserIndex).Stats.UserAtributos(Carisma) = UserList(UserIndex).Stats.UserAtributos(Carisma) + 2
    Case "ENANO"
        UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) + 3
        UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 3
        UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) - 6
    Case "GNOMO"
        UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) - 5
        UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 3
        UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) + 3
End Select



UserList(UserIndex).Stats.UserSkills(1) = val(US1)
UserList(UserIndex).Stats.UserSkills(2) = val(US2)
UserList(UserIndex).Stats.UserSkills(3) = val(US3)
UserList(UserIndex).Stats.UserSkills(4) = val(US4)
UserList(UserIndex).Stats.UserSkills(5) = val(US5)
UserList(UserIndex).Stats.UserSkills(6) = val(US6)
UserList(UserIndex).Stats.UserSkills(7) = val(US7)
UserList(UserIndex).Stats.UserSkills(8) = val(US8)
UserList(UserIndex).Stats.UserSkills(9) = val(US9)
UserList(UserIndex).Stats.UserSkills(10) = val(US10)
UserList(UserIndex).Stats.UserSkills(11) = val(US11)
UserList(UserIndex).Stats.UserSkills(12) = val(US12)
UserList(UserIndex).Stats.UserSkills(13) = val(US13)
UserList(UserIndex).Stats.UserSkills(14) = val(US14)
UserList(UserIndex).Stats.UserSkills(15) = val(US15)
UserList(UserIndex).Stats.UserSkills(16) = val(US16)
UserList(UserIndex).Stats.UserSkills(17) = val(US17)
UserList(UserIndex).Stats.UserSkills(18) = val(US18)
UserList(UserIndex).Stats.UserSkills(19) = val(US19)
UserList(UserIndex).Stats.UserSkills(20) = val(US20)
UserList(UserIndex).Stats.UserSkills(21) = val(US21)

totalskpts = 0

'Abs PREVINENE EL HACKEO DE LOS SKILLS %%%%%%%%%%%%%
For LoopC = 1 To NUMSKILLS
    totalskpts = totalskpts + Abs(UserList(UserIndex).Stats.UserSkills(LoopC))
Next LoopC


If totalskpts > 10 Then
    Call LogHackAttemp(UserList(UserIndex).Name & " intento hackear los skills.")
    Call BorrarUsuario(UserList(UserIndex).Name)
    Call CloseSocket(UserIndex)
    Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

UserList(UserIndex).Password = Password
UserList(UserIndex).Char.Heading = SOUTH

Call Randomize(Timer)
Call DarCuerpoYCabeza(UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Raza, UserList(UserIndex).Genero)
UserList(UserIndex).OrigChar = UserList(UserIndex).Char
   
 
UserList(UserIndex).Char.WeaponAnim = NingunArma
UserList(UserIndex).Char.ShieldAnim = NingunEscudo
UserList(UserIndex).Char.CascoAnim = NingunCasco

UserList(UserIndex).Stats.MET = 1
Dim MiInt
MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 3)

UserList(UserIndex).Stats.MaxHP = 15 + MiInt
UserList(UserIndex).Stats.MinHP = 15 + MiInt

UserList(UserIndex).Stats.FIT = 1


MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(UserIndex).Stats.MaxSta = 20 * MiInt
UserList(UserIndex).Stats.MinSta = 20 * MiInt


UserList(UserIndex).Stats.MaxAGU = 100
UserList(UserIndex).Stats.MinAGU = 100

UserList(UserIndex).Stats.MaxHam = 100
UserList(UserIndex).Stats.MinHam = 100


'<-----------------MANA----------------------->
If UserClase = "Mago" Then
    MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Inteligencia)) / 3
    UserList(UserIndex).Stats.MaxMAN = 100 + MiInt
    UserList(UserIndex).Stats.MinMAN = 100 + MiInt
ElseIf UserClase = "Clerigo" Or UserClase = "Druida" _
    Or UserClase = "Bardo" Or UserClase = "Asesino" Then
        MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Inteligencia)) / 4
        UserList(UserIndex).Stats.MaxMAN = 50
        UserList(UserIndex).Stats.MinMAN = 50
Else
    UserList(UserIndex).Stats.MaxMAN = 0
    UserList(UserIndex).Stats.MinMAN = 0
End If

If UserClase = "Mago" Or UserClase = "Clerigo" Or _
   UserClase = "Druida" Or UserClase = "Bardo" Or _
   UserClase = "Asesino" Then
        UserList(UserIndex).Stats.UserHechizos(1) = 2
End If

UserList(UserIndex).Stats.MaxHIT = 2
UserList(UserIndex).Stats.MinHIT = 1

UserList(UserIndex).Stats.GLD = 0




UserList(UserIndex).Stats.Exp = 0
UserList(UserIndex).Stats.ELU = 300
UserList(UserIndex).Stats.ELV = 1


'???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
UserList(UserIndex).Invent.NroItems = 4

UserList(UserIndex).Invent.Object(1).ObjIndex = 467
UserList(UserIndex).Invent.Object(1).Amount = 100

UserList(UserIndex).Invent.Object(2).ObjIndex = 468
UserList(UserIndex).Invent.Object(2).Amount = 100

UserList(UserIndex).Invent.Object(3).ObjIndex = 460
UserList(UserIndex).Invent.Object(3).Amount = 1
UserList(UserIndex).Invent.Object(3).Equipped = 1

Select Case UserRaza
    Case "Humano"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 463
    Case "Elfo"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 464
    Case "Elfo Oscuro"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 465
    Case "Enano"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 466
    Case "Gnomo"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 466
End Select

UserList(UserIndex).Invent.Object(4).Amount = 1
UserList(UserIndex).Invent.Object(4).Equipped = 1

'UserList(UserIndex).Invent.Object(5).ObjIndex = 461
'UserList(UserIndex).Invent.Object(5).Amount = 50
'
'UserList(UserIndex).Invent.Object(6).ObjIndex = 462
'UserList(UserIndex).Invent.Object(6).Amount = 50

UserList(UserIndex).Invent.ArmourEqpSlot = 4
UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(4).ObjIndex

UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(3).ObjIndex
UserList(UserIndex).Invent.WeaponEqpSlot = 3



Call SaveUser(UserIndex, CharPath & UCase$(Name) & ".chr")
  
'Open User
Call ConnectUser(UserIndex, Name, Password)
  
End Sub

Sub CloseSocket(ByVal UserIndex As Integer)
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

'Call LogTarea("Close Socket")

#If Not (UsarAPI = 1) Then
On Error GoTo errhandler
#End If

    
    Call aDos.RestarConexion(UserList(UserIndex).ip)
    
    If UserList(UserIndex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            Call CloseUser(UserIndex)
            
            Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    End If
    
    #If UsarAPI Then
    
    If UserList(UserIndex).ConnID <> -1 Then
        Call apiclosesocket(UserList(UserIndex).ConnID)
    End If
    
    #Else
    
    'frmMain.Socket2(UserIndex).Disconnect
    frmMain.Socket2(UserIndex).Cleanup
    Unload frmMain.Socket2(UserIndex)
    
    #End If
    
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
            
    Call ResetUserSlot(UserIndex)

Exit Sub

errhandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
'    Unload frmMain.Socket2(UserIndex) OJOOOOOOOOOOOOOOOOO
'    If NumUsers > 0 Then NumUsers = NumUsers - 1
    Call ResetUserSlot(UserIndex)
    
    #If UsarAPI Then
    If UserList(UserIndex).ConnID <> -1 Then
        Call apiclosesocket(UserList(UserIndex).ConnID)
    End If
    #End If
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
    
End Sub


Sub CloseSocket_NUEVA(ByVal UserIndex As Integer)
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

'Call LogTarea("Close Socket")

On Error GoTo errhandler
    

    
    Call aDos.RestarConexion(frmMain.Socket2(UserIndex).PeerAddress)
    
    'UserList(UserIndex).ConnID = -1
    'UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
            
    If UserList(UserIndex).flags.UserLogged Then
        If NumUsers <> 0 Then NumUsers = NumUsers - 1
        Call CloseUser(UserIndex)
        UserList(UserIndex).ConnID = -1: UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
        frmMain.Socket2(UserIndex).Disconnect
        frmMain.Socket2(UserIndex).Cleanup
        'Unload frmMain.Socket2(UserIndex)
        Call ResetUserSlot(UserIndex)
        'Call Cerrar_Usuario(UserIndex)
    Else
        UserList(UserIndex).ConnID = -1
        UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
        
        frmMain.Socket2(UserIndex).Disconnect
        frmMain.Socket2(UserIndex).Cleanup
        Call ResetUserSlot(UserIndex)
        'Unload frmMain.Socket2(UserIndex)
    End If

Exit Sub

errhandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
'    Unload frmMain.Socket2(UserIndex) OJOOOOOOOOOOOOOOOOO
'    If NumUsers > 0 Then NumUsers = NumUsers - 1
    Call ResetUserSlot(UserIndex)
    
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
    
End Sub


Sub SendData(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, sndData As String)


On Error Resume Next

Dim LoopC As Integer
Dim x As Integer
Dim y As Integer
Dim aux$
Dim dec$
Dim nfile As Integer
Dim Ret As Long

sndData = sndData & ENDC



Select Case sndRoute


    Case ToNone
        Exit Sub
        
    Case ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 Then
               If EsDios(UserList(LoopC).Name) Or EsSemiDios(UserList(LoopC).Name) Then
                        'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                        #If UsarAPI Then
                        Call WsApiEnviar(LoopC, sndData)
                        #Else
                        frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        #End If
               End If
            End If
        Next LoopC
        Exit Sub
        
    Case ToAll
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                    #If UsarAPI Then
                    Call WsApiEnviar(LoopC, sndData)
                    #Else
                    frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                    #End If
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToAllButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) And (LoopC <> sndIndex) Then
                If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                    'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                    #If UsarAPI Then
                    Call WsApiEnviar(LoopC, sndData)
                    #Else
                    frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                    #End If
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToMap
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) Then
                If UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).Pos.Map = sndMap Then
                        'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                        'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        #If UsarAPI Then
                        Call WsApiEnviar(LoopC, sndData)
                        #Else
                        frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        #End If
                    End If
                End If
            End If
        Next LoopC
        Exit Sub
      
    Case ToMapButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) And LoopC <> sndIndex Then
                If UserList(LoopC).Pos.Map = sndMap Then
                    'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        #If UsarAPI Then
                        Call WsApiEnviar(LoopC, sndData)
                        #Else
                        frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        #End If
                End If
            End If
        Next LoopC
        Exit Sub
            
    Case ToGuildMembers
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) Then
                If UserList(sndIndex).GuildInfo.GuildName = UserList(LoopC).GuildInfo.GuildName Then
                        'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        #If UsarAPI Then
                        Call WsApiEnviar(LoopC, sndData)
                        #Else
                        frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        #End If
                End If
            End If
        Next LoopC
        Exit Sub
    
    Case ToPCArea
        For y = UserList(sndIndex).Pos.y - MinYBorder + 1 To UserList(sndIndex).Pos.y + MinYBorder - 1
            For x = UserList(sndIndex).Pos.x - MinXBorder + 1 To UserList(sndIndex).Pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, y) Then
                    If MapData(sndMap, x, y).UserIndex > 0 Then
                       If UserList(MapData(sndMap, x, y).UserIndex).ConnID > -1 Then
                            'Call AddtoVar(UserList(MapData(sndMap, X, Y).UserIndex).BytesTransmitidosSvr, LenB(sndData), 100000)
                            'frmMain.Socket2(MapData(sndMap, X, Y).UserIndex).Write sndData, Len(sndData)
                            #If UsarAPI Then
                            Call WsApiEnviar(MapData(sndMap, x, y).UserIndex, sndData)
                            #Else
                            frmMain.Socket2(MapData(sndMap, x, y).UserIndex).Write sndData, Len(sndData)
                            #End If
                       End If
                    End If
               End If
            Next x
        Next y
        Exit Sub

    '[Alejo-18-5]
    Case ToPCAreaButIndex
        For y = UserList(sndIndex).Pos.y - MinYBorder + 1 To UserList(sndIndex).Pos.y + MinYBorder - 1
            For x = UserList(sndIndex).Pos.x - MinXBorder + 1 To UserList(sndIndex).Pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, y) Then
                    If (MapData(sndMap, x, y).UserIndex > 0) And (MapData(sndMap, x, y).UserIndex <> sndIndex) Then
                       If UserList(MapData(sndMap, x, y).UserIndex).ConnID > -1 Then
                            'Call AddtoVar(UserList(MapData(sndMap, X, Y).UserIndex).BytesTransmitidosSvr, LenB(sndData), 100000)
                            'frmMain.Socket2(MapData(sndMap, X, Y).UserIndex).Write sndData, Len(sndData)
                            #If UsarAPI Then
                            Call WsApiEnviar(MapData(sndMap, x, y).UserIndex, sndData)
                            #Else
                            frmMain.Socket2(MapData(sndMap, x, y).UserIndex).Write sndData, Len(sndData)
                            #End If
                       End If
                    End If
               End If
            Next x
        Next y
        Exit Sub

    Case ToNPCArea
        For y = Npclist(sndIndex).Pos.y - MinYBorder + 1 To Npclist(sndIndex).Pos.y + MinYBorder - 1
            For x = Npclist(sndIndex).Pos.x - MinXBorder + 1 To Npclist(sndIndex).Pos.x + MinXBorder - 1
               If InMapBounds(sndMap, x, y) Then
                    If MapData(sndMap, x, y).UserIndex > 0 Then
                       If UserList(MapData(sndMap, x, y).UserIndex).ConnID > -1 Then
                            'Call AddtoVar(UserList(MapData(sndMap, X, Y).UserIndex).BytesTransmitidosSvr, LenB(sndData), 100000)
                            'frmMain.Socket2(MapData(sndMap, X, Y).UserIndex).Write sndData, Len(sndData)
                            #If UsarAPI Then
                            Call WsApiEnviar(MapData(sndMap, x, y).UserIndex, sndData)
                            #Else
                            frmMain.Socket2(MapData(sndMap, x, y).UserIndex).Write sndData, Len(sndData)
                            #End If
                       End If
                    End If
               End If
            Next x
        Next y
        Exit Sub

    Case ToIndex
        If UserList(sndIndex).ConnID > -1 Then
             'Call AddtoVar(UserList(sndIndex).BytesTransmitidosSvr, LenB(sndData), 100000)
             'frmMain.Socket2(sndIndex).Write sndData, Len(sndData)
             #If UsarAPI Then
             Call WsApiEnviar(sndIndex, sndData)
             #Else
             frmMain.Socket2(sndIndex).Write sndData, Len(sndData)
             #End If
             Exit Sub
        End If

End Select

End Sub
Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean


Dim x As Integer, y As Integer
For y = UserList(Index).Pos.y - MinYBorder + 1 To UserList(Index).Pos.y + MinYBorder - 1
        For x = UserList(Index).Pos.x - MinXBorder + 1 To UserList(Index).Pos.x + MinXBorder - 1

            If MapData(UserList(Index).Pos.Map, x, y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next x
Next y
EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean


Dim x As Integer, y As Integer
For y = Pos.y - MinYBorder + 1 To Pos.y + MinYBorder - 1
        For x = Pos.x - MinXBorder + 1 To Pos.x + MinXBorder - 1
            If x > 0 And y > 0 And x < 101 And y < 101 Then
                If MapData(Pos.Map, x, y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next x
Next y
HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean


Dim x As Integer, y As Integer
For y = Pos.y - MinYBorder + 1 To Pos.y + MinYBorder - 1
        For x = Pos.x - MinXBorder + 1 To Pos.x + MinXBorder - 1
            If MapData(Pos.Map, x, y).OBJInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next x
Next y
HayOBJarea = False
End Function

Sub CorregirSkills(ByVal UserIndex As Integer)
Dim k As Integer

For k = 1 To NUMSKILLS
  If UserList(UserIndex).Stats.UserSkills(k) > MAXSKILLPOINTS Then UserList(UserIndex).Stats.UserSkills(k) = MAXSKILLPOINTS
Next

For k = 1 To NUMATRIBUTOS
 If UserList(UserIndex).Stats.UserAtributos(k) > MAXATRIBUTOS Then
    Call SendData(ToIndex, UserIndex, 0, "ERREl personaje tiene atributos invalidos.")
    Exit Sub
 End If
Next k
 
End Sub


Function ValidateChr(ByVal UserIndex As Integer) As Boolean

ValidateChr = UserList(UserIndex).Char.Head <> 0 And _
UserList(UserIndex).Char.Body <> 0 And ValidateSkills(UserIndex)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, Name As String, Password As String)
Dim N As Integer

'Reseteamos los FLAGS
UserList(UserIndex).flags.Escondido = 0
UserList(UserIndex).flags.TargetNpc = 0
UserList(UserIndex).flags.TargetNpcTipo = 0
UserList(UserIndex).flags.TargetObj = 0
UserList(UserIndex).flags.TargetUser = 0
UserList(UserIndex).Char.FX = 0



'Controlamos no pasar el maximo de usuarios
If NumUsers >= MaxUsers Then
    Call SendData(ToIndex, UserIndex, 0, "ERREl servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If
  
'¿Este IP ya esta conectado?
If AllowMultiLogins = 0 Then
    If CheckForSameIP(UserIndex, UserList(UserIndex).ip) = True Then
        Call SendData(ToIndex, UserIndex, 0, "ERRNo es posible usar mas de un personaje al mismo tiempo.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
End If

'¿Ya esta conectado el personaje?
If CheckForSameName(UserIndex, Name) = True Then
    Call SendData(ToIndex, UserIndex, 0, "ERRPerdon, un usuario con el mismo nombre se há logoeado.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Existe el personaje?
If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = False Then
    Call SendData(ToIndex, UserIndex, 0, "ERREl personaje no existe.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Es el passwd valido?
If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password")) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRPassword incorrecto.")
    'Call frmMain.Socket2(UserIndex).Disconnect
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'Cargamos los datos del personaje
Call LoadUserInit(UserIndex, CharPath & UCase$(Name) & ".chr")
Call LoadUserStats(UserIndex, CharPath & UCase$(Name) & ".chr")
'Call CorregirSkills(UserIndex)

If Not ValidateChr(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRError en el personaje.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

Call LoadUserReputacion(UserIndex, CharPath & UCase$(Name) & ".chr")


If UserList(UserIndex).Invent.EscudoEqpSlot = 0 Then UserList(UserIndex).Char.ShieldAnim = NingunEscudo
If UserList(UserIndex).Invent.CascoEqpSlot = 0 Then UserList(UserIndex).Char.CascoAnim = NingunCasco
If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then UserList(UserIndex).Char.WeaponAnim = NingunArma


Call UpdateUserInv(True, UserIndex, 0)
Call UpdateUserHechizos(True, UserIndex, 0)

If UserList(UserIndex).flags.Navegando = 1 Then
     UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje
     UserList(UserIndex).Char.Head = 0
     UserList(UserIndex).Char.WeaponAnim = NingunArma
     UserList(UserIndex).Char.ShieldAnim = NingunEscudo
     UserList(UserIndex).Char.CascoAnim = NingunCasco
End If


If UserList(UserIndex).flags.Paralizado Then Call SendData(ToIndex, UserIndex, 0, "PARADOK")

'Posicion de comienzo
If UserList(UserIndex).Pos.Map = 0 Then
    If UCase$(UserList(UserIndex).Hogar) = "NIX" Then
             UserList(UserIndex).Pos = Nix
    ElseIf UCase$(UserList(UserIndex).Hogar) = "ULLATHORPE" Then
             UserList(UserIndex).Pos = Ullathorpe
    ElseIf UCase$(UserList(UserIndex).Hogar) = "BANDERBILL" Then
             UserList(UserIndex).Pos = Banderbill
    ElseIf UCase$(UserList(UserIndex).Hogar) = "LINDOS" Then
             UserList(UserIndex).Pos = Lindos
    Else
        UserList(UserIndex).Hogar = "ULLATHORPE"
        UserList(UserIndex).Pos = Ullathorpe
    End If
Else
   
   If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).UserIndex <> 0 Then Call CloseSocket(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).UserIndex)
   
End If

'Nombre de sistema
UserList(UserIndex).Name = Name

UserList(UserIndex).Password = Password
'UserList(UserIndex).ip = frmMain.Socket2(UserIndex).PeerAddress
  
'Info
Call SendData(ToIndex, UserIndex, 0, "IU" & UserIndex) 'Enviamos el User index
Call SendData(ToIndex, UserIndex, 0, "CM" & UserList(UserIndex).Pos.Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion) 'Carga el mapa
Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(UserList(UserIndex).Pos.Map).Music)

If Lloviendo Then Call SendData(ToIndex, UserIndex, 0, "LLU")

Call UpdateUserMap(UserIndex)
Call SendUserStatsBox(UserIndex)
Call EnviarHambreYsed(UserIndex)

Call SendMOTD(UserIndex)

If haciendoBK Then
    Call SendData(ToIndex, UserIndex, 0, "BKW")
    Call SendData(ToIndex, UserIndex, 0, "||Por favor espera algunos segundo, WorldSave esta ejecutandose." & FONTTYPE_INFO)
End If

'Actualiza el Num de usuarios
If UserIndex > LastUser Then LastUser = UserIndex

NumUsers = NumUsers + 1
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)

UserList(UserIndex).flags.UserLogged = True

MapInfo(UserList(UserIndex).Pos.Map).NumUsers = MapInfo(UserList(UserIndex).Pos.Map).NumUsers + 1

If UserList(UserIndex).Stats.SkillPts > 0 Then
    Call EnviarSkills(UserIndex)
    Call EnviarSubirNivel(UserIndex, UserList(UserIndex).Stats.SkillPts)
End If

If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers

If NumUsers > recordusuarios Then
    Call SendData(ToAll, 0, 0, "||Record de usuarios conectados simultaniamente." & "Hay " & NumUsers & " usuarios." & FONTTYPE_INFO)
    recordusuarios = NumUsers
    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
    
    Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
End If

If EsDios(Name) Then
    UserList(UserIndex).flags.Privilegios = 3
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip, False)
ElseIf EsSemiDios(Name) Then
    UserList(UserIndex).flags.Privilegios = 2
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip, False)
ElseIf EsConsejero(Name) Then
    UserList(UserIndex).flags.Privilegios = 1
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip, True)
Else
    UserList(UserIndex).flags.Privilegios = 0
End If

Set UserList(UserIndex).GuildRef = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

UserList(UserIndex).Counters.IdleCount = 0

If UserList(UserIndex).NroMacotas > 0 Then
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(i) > 0 Then
            UserList(UserIndex).MascotasIndex(i) = SpawnNpc(UserList(UserIndex).MascotasType(i), UserList(UserIndex).Pos, True, True)
            
            If UserList(UserIndex).MascotasIndex(i) <= MAXNPCS Then
                  Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
                  Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
            Else
                  UserList(UserIndex).MascotasIndex(i) = 0
            End If
        End If
    Next i
End If


If UserList(UserIndex).flags.Navegando = 1 Then Call SendData(ToIndex, UserIndex, 0, "NAVEG")

UserList(UserIndex).flags.Seguro = True

'Crea  el personaje del usuario
Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y)
Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 0)
Call SendData(ToIndex, UserIndex, 0, "LOGGED")

Call SendGuildNews(UserIndex)


Call MostrarNumUsers

N = FreeFile
Open App.Path & "\logs\numusers.log" For Output As N
Print #N, NumUsers
Close #N

N = FreeFile
'Log
Open App.Path & "\logs\Connect.log" For Append Shared As #N
Print #N, UserList(UserIndex).Name & " ha entrado al juego. UserIndex:" & UserIndex & " " & Time & " " & Date
Close #N

End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
Dim j As Integer
Call SendData(ToIndex, UserIndex, 0, "||Message of the day:" & FONTTYPE_INFO)
For j = 1 To MaxLines
    Call SendData(ToIndex, UserIndex, 0, "||" & MOTD(j).texto)
Next j
End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)

UserList(UserIndex).Faccion.ArmadaReal = 0
UserList(UserIndex).Faccion.FuerzasCaos = 0
UserList(UserIndex).Faccion.CiudadanosMatados = 0
UserList(UserIndex).Faccion.CriminalesMatados = 0
UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0
UserList(UserIndex).Faccion.RecibioArmaduraReal = 0
UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0
UserList(UserIndex).Faccion.RecibioExpInicialReal = 0
UserList(UserIndex).Faccion.RecompensasCaos = 0
UserList(UserIndex).Faccion.RecompensasReal = 0

End Sub

Sub ResetContadores(ByVal UserIndex As Integer)

UserList(UserIndex).Counters.AGUACounter = 0
UserList(UserIndex).Counters.AttackCounter = 0
UserList(UserIndex).Counters.Ceguera = 0
UserList(UserIndex).Counters.COMCounter = 0
UserList(UserIndex).Counters.Estupidez = 0
UserList(UserIndex).Counters.Frio = 0
UserList(UserIndex).Counters.HPCounter = 0
UserList(UserIndex).Counters.IdleCount = 0
UserList(UserIndex).Counters.Invisibilidad = 0
UserList(UserIndex).Counters.Paralisis = 0
UserList(UserIndex).Counters.Pasos = 0
UserList(UserIndex).Counters.Pena = 0
UserList(UserIndex).Counters.PiqueteC = 0
UserList(UserIndex).Counters.STACounter = 0
UserList(UserIndex).Counters.Veneno = 0

End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)

UserList(UserIndex).Char.Body = 0
UserList(UserIndex).Char.CascoAnim = 0
UserList(UserIndex).Char.CharIndex = 0
UserList(UserIndex).Char.FX = 0
UserList(UserIndex).Char.Head = 0
UserList(UserIndex).Char.loops = 0
UserList(UserIndex).Char.Heading = 0
UserList(UserIndex).Char.loops = 0
UserList(UserIndex).Char.ShieldAnim = 0
UserList(UserIndex).Char.WeaponAnim = 0

End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)

UserList(UserIndex).Name = ""
UserList(UserIndex).modName = ""
UserList(UserIndex).Password = ""
UserList(UserIndex).Desc = ""
UserList(UserIndex).Pos.Map = 0
UserList(UserIndex).Pos.x = 0
UserList(UserIndex).Pos.y = 0
UserList(UserIndex).ip = ""
UserList(UserIndex).RDBuffer = ""
UserList(UserIndex).Clase = ""
UserList(UserIndex).Email = ""
UserList(UserIndex).Genero = ""
UserList(UserIndex).Hogar = ""
UserList(UserIndex).Raza = ""

UserList(UserIndex).RandKey = 0
UserList(UserIndex).PrevCRC = 0
UserList(UserIndex).PacketNumber = 0

UserList(UserIndex).Stats.Banco = 0
UserList(UserIndex).Stats.ELV = 0
UserList(UserIndex).Stats.ELU = 0
UserList(UserIndex).Stats.Exp = 0
UserList(UserIndex).Stats.Def = 0
UserList(UserIndex).Stats.CriminalesMatados = 0
UserList(UserIndex).Stats.NPCsMuertos = 0
UserList(UserIndex).Stats.UsuariosMatados = 0

End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)

UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.BurguesRep = 0
UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.NobleRep = 0
UserList(UserIndex).Reputacion.PlebeRep = 0
UserList(UserIndex).Reputacion.NobleRep = 0
UserList(UserIndex).Reputacion.Promedio = 0

End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)

UserList(UserIndex).GuildInfo.ClanFundado = ""
UserList(UserIndex).GuildInfo.Echadas = 0
UserList(UserIndex).GuildInfo.EsGuildLeader = 0
UserList(UserIndex).GuildInfo.FundoClan = 0
UserList(UserIndex).GuildInfo.GuildName = ""
UserList(UserIndex).GuildInfo.Solicitudes = 0
UserList(UserIndex).GuildInfo.SolicitudesRechazadas = 0
UserList(UserIndex).GuildInfo.VecesFueGuildLeader = 0
UserList(UserIndex).GuildInfo.YaVoto = 0
UserList(UserIndex).GuildInfo.ClanesParticipo = 0
UserList(UserIndex).GuildInfo.GuildPoints = 0

End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)

UserList(UserIndex).flags.Comerciando = False
UserList(UserIndex).flags.Ban = 0
UserList(UserIndex).flags.Escondido = 0
UserList(UserIndex).flags.DuracionEfecto = 0
UserList(UserIndex).flags.NpcInv = 0
UserList(UserIndex).flags.StatsChanged = 0
UserList(UserIndex).flags.TargetNpc = 0
UserList(UserIndex).flags.TargetNpcTipo = 0
UserList(UserIndex).flags.TargetObj = 0
UserList(UserIndex).flags.TargetObjMap = 0
UserList(UserIndex).flags.TargetObjX = 0
UserList(UserIndex).flags.TargetObjY = 0
UserList(UserIndex).flags.TargetUser = 0
UserList(UserIndex).flags.TipoPocion = 0
UserList(UserIndex).flags.TomoPocion = False
UserList(UserIndex).flags.Descuento = ""
UserList(UserIndex).flags.Hambre = 0
UserList(UserIndex).flags.Sed = 0
UserList(UserIndex).flags.Descansar = False
UserList(UserIndex).flags.ModoCombate = False
UserList(UserIndex).flags.Vuela = 0
UserList(UserIndex).flags.Navegando = 0
UserList(UserIndex).flags.Oculto = 0
UserList(UserIndex).flags.Envenenado = 0
UserList(UserIndex).flags.Invisible = 0
UserList(UserIndex).flags.Paralizado = 0
UserList(UserIndex).flags.Maldicion = 0
UserList(UserIndex).flags.Bendicion = 0
UserList(UserIndex).flags.Meditando = 0
UserList(UserIndex).flags.Privilegios = 0
UserList(UserIndex).flags.PuedeMoverse = 0
UserList(UserIndex).flags.PuedeLanzarSpell = 0
UserList(UserIndex).Stats.SkillPts = 0
UserList(UserIndex).flags.OldBody = 0
UserList(UserIndex).flags.OldHead = 0
UserList(UserIndex).flags.AdminInvisible = 0
UserList(UserIndex).flags.ValCoDe = 0
UserList(UserIndex).flags.Hechizo = 0

End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)

Dim LoopC As Integer
For LoopC = 1 To MAXUSERHECHIZOS
    UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
Next

End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)

Dim LoopC As Integer

UserList(UserIndex).NroMacotas = 0
    
For LoopC = 1 To MAXMASCOTAS
    UserList(UserIndex).MascotasIndex(LoopC) = 0
    UserList(UserIndex).MascotasType(LoopC) = 0
Next LoopC

End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
Dim LoopC As Integer
For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
      UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
      UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
      UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
Next
UserList(UserIndex).BancoInvent.NroItems = 0
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)

Set UserList(UserIndex).CommandsBuffer = Nothing
Set UserList(UserIndex).GuildRef = Nothing

UserList(UserIndex).AntiCuelgue = 0

Call ResetFacciones(UserIndex)
Call ResetContadores(UserIndex)
Call ResetCharInfo(UserIndex)
Call ResetBasicUserInfo(UserIndex)
Call ResetReputacion(UserIndex)
Call ResetGuildInfo(UserIndex)
Call ResetUserFlags(UserIndex)
Call LimpiarInventario(UserIndex)
Call ResetUserSpells(UserIndex)
Call ResetUserPets(UserIndex)
Call ResetUserBanco(UserIndex)

'UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
'UserList(UserIndex).BytesTransmitidosUser = 0
'UserList(UserIndex).BytesTransmitidosSvr = 0





End Sub


Sub CloseUser(ByVal UserIndex As Integer)
'Call LogTarea("CloseUser " & UserIndex)

On Error GoTo errhandler

Dim N As Integer
Dim x As Integer
Dim y As Integer
Dim LoopC As Integer
Dim Map As Integer
Dim Name As String
Dim Raza As String
Dim Clase As String
Dim i As Integer

Dim aN As Integer

aN = UserList(UserIndex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = ""
End If

Map = UserList(UserIndex).Pos.Map
x = UserList(UserIndex).Pos.x
y = UserList(UserIndex).Pos.y
Name = UCase$(UserList(UserIndex).Name)
Raza = UserList(UserIndex).Raza
Clase = UserList(UserIndex).Clase

UserList(UserIndex).Char.FX = 0
UserList(UserIndex).Char.loops = 0
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
   

UserList(UserIndex).flags.UserLogged = False
UserList(UserIndex).Counters.Saliendo = False

'Le devolvemos el body y head originales
If UserList(UserIndex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)

' Grabamos el personaje del usuario
Call SaveUser(UserIndex, CharPath & Name & ".chr")

'Quitar el dialogo
If MapInfo(Map).NumUsers > 0 Then
    Call SendData(ToMapButIndex, UserIndex, Map, "QDL" & UserList(UserIndex).Char.CharIndex)
End If

'Borrar el personaje
If UserList(UserIndex).Char.CharIndex > 0 Then
    Call EraseUserChar(ToMapButIndex, UserIndex, Map, UserIndex)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).flags.NPCActive Then _
                Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

If UserIndex = LastUser Then
    Do Until UserList(LastUser).flags.UserLogged
        LastUser = LastUser - 1
        If LastUser < 1 Then Exit Do
    Loop
End If
  
'If NumUsers <> 0 Then
'    NumUsers = NumUsers - 1
'End If

'Update Map Users
MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1

If MapInfo(Map).NumUsers < 0 Then
    MapInfo(Map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(UserIndex).Name) Then Call Ayuda.Quitar(UserList(UserIndex).Name)

Call ResetUserSlot(UserIndex)

Call MostrarNumUsers

N = FreeFile(1)
Open App.Path & "\logs\Connect.log" For Append Shared As #N
Print #N, Name & " há dejado el juego. " & "User Index:" & UserIndex & " " & Time & " " & Date
Close #N

Exit Sub

errhandler:
Call LogError("Error en CloseUser")


End Sub


Sub HandleData(ByVal UserIndex As Integer, ByVal rdata As String)

Call LogTarea("Sub HandleData :" & rdata & " " & UserList(UserIndex).Name)

On Error GoTo ErrorHandler:




Dim sndData As String
Dim CadenaOriginal As String

Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim auxind As Integer
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
Dim x As Integer
Dim y As Integer
Dim cliMD5 As String

Dim ClientCRC As String
Dim ServerSideCRC As Long

CadenaOriginal = rdata

'¿Tiene un indece valido?
If UserIndex <= 0 Then
    Call CloseSocket(UserIndex)
    Exit Sub
End If

If Left$(rdata, 13) = "gIvEmEvAlcOde" Then
   '<<<<<<<<<<< MODULO PRIVADO DE CADA IMPLEMENTACION >>>>>>
   UserList(UserIndex).flags.ValCoDe = CInt(RandomNumber(20000, 32000))
   UserList(UserIndex).RandKey = CLng(RandomNumber(0, 99999))
   UserList(UserIndex).PrevCRC = UserList(UserIndex).RandKey
   UserList(UserIndex).PacketNumber = 100
   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   Call SendData(ToIndex, UserIndex, 0, "VAL" & UserList(UserIndex).RandKey & "," & UserList(UserIndex).flags.ValCoDe)
   Exit Sub
ElseIf UserList(UserIndex).flags.UserLogged = False And Left(rdata, 12) = "CLIENTEVIEJO" Then
    Dim ElMsg As String, LaLong As String
    ElMsg = "ERRLa version del cliente que usas es obsoleta. Si deseas conectarte a este servidor, entra a www.argentum-online.com.ar y alli podrás enterarte como hacer."
    If Len(ElMsg) > 255 Then ElMsg = Left(ElMsg, 255)
    LaLong = Chr(0) & Chr(Len(ElMsg))
    Call SendData(ToIndex, UserIndex, 0, LaLong & ElMsg)
    Call CloseSocket(UserIndex)
    Exit Sub
Else
   '<<<<<<<<<<< MODULO PRIVADO DE CADA IMPLEMENTACION >>>>>>
   'ClientCRC = ReadField(2, rdata, 126)
   ClientCRC = Right(rdata, Len(rdata) - InStrRev(rdata, Chr(126)))
   tStr = Left$(rdata, Len(rdata) - Len(ClientCRC) - 1)
   'ServerSideCRC = GenCrC(UserList(UserIndex).PrevCRC, tStr)
   'If CLng(ClientCRC) <> ServerSideCRC Then Call CloseSocket(UserIndex): Debug.Print "ERR CRC"
   UserList(UserIndex).PrevCRC = ServerSideCRC
   rdata = tStr
   tStr = ""
   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>
End If

UserList(UserIndex).Counters.IdleCount = 0
   
   If Not UserList(UserIndex).flags.UserLogged Then

        Select Case Left$(rdata, 6)
            Case "OLOGIN"
                rdata = Right$(rdata, Len(rdata) - 6)
                cliMD5 = Right$(rdata, 16)
                rdata = Left$(rdata, Len(rdata) - 16)
                If Not MD5ok(cliMD5) Then
                    Call SendData(ToIndex, UserIndex, 0, "ERREl cliente está dañado, por favor descarguelo nuevamente desde el sitio.")
                    Exit Sub
                End If
                Ver = ReadField(3, rdata, 44)
                If VersionOK(Ver) Then
                    tName = ReadField(1, rdata, 44)
                    
                    If Not AsciiValidos(tName) Then
                        Call SendData(ToIndex, UserIndex, 0, "ERRNombre invalido.")
                        Exit Sub
                    End If
                    
                    If Not PersonajeExiste(tName) Then
                        Call SendData(ToIndex, UserIndex, 0, "ERREl personaje no existe.")
                        Exit Sub
                    End If
                    
                    If Not BANCheck(tName) Then
                        
                        If (UserList(UserIndex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(UserIndex).flags.ValCoDe) <> CInt(val(ReadField(4, rdata, 44)))) Then
                              Call LogHackAttemp("IP:" & UserList(UserIndex).ip & " intento crear un bot.")
                              Call CloseSocket(UserIndex)
                              Exit Sub
                        End If
                        Dim Pass11 As String
                        Pass11 = ReadField(2, rdata, 44)
                        Call ConnectUser(UserIndex, tName, Pass11)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a Argentum debido a tu mal comportamiento.")
                    End If
                    
                Else
                     Call SendData(ToIndex, UserIndex, 0, "ERREsta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en nuestra pagina.")
                     'Call CloseSocket(UserIndex)
                     Exit Sub
                End If
                Exit Sub
            Case "TIRDAD"
                UserList(UserIndex).Stats.UserAtributos(1) = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
                UserList(UserIndex).Stats.UserAtributos(2) = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
                UserList(UserIndex).Stats.UserAtributos(3) = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
                UserList(UserIndex).Stats.UserAtributos(4) = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
                UserList(UserIndex).Stats.UserAtributos(5) = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))

                Call SendData(ToIndex, UserIndex, 0, "DADOS" & UserList(UserIndex).Stats.UserAtributos(1) & "," & UserList(UserIndex).Stats.UserAtributos(2) & "," & UserList(UserIndex).Stats.UserAtributos(3) & "," & UserList(UserIndex).Stats.UserAtributos(4) & "," & UserList(UserIndex).Stats.UserAtributos(5))
                
                Exit Sub

            Case "NLOGIN"
            
                If PuedeCrearPersonajes = 0 Then
                        Call SendData(ToIndex, UserIndex, 0, "ERRNo se pueden crear mas personajes en este servidor.")
                        Call CloseSocket(UserIndex)
                        Exit Sub
                End If
                
                If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
                        Call SendData(ToIndex, UserIndex, 0, "ERRHas creado demasiados personajes.")
                        Call CloseSocket(UserIndex)
                        Exit Sub
                End If
                
                rdata = Right$(rdata, Len(rdata) - 6)
                cliMD5 = Right$(rdata, 16)
                rdata = Left$(rdata, Len(rdata) - 16)
                If Not MD5ok(cliMD5) Then
                    Call SendData(ToIndex, UserIndex, 0, "ERREl cliente está dañado, por favor descarguelo nuevamente desde el sitio.")
                    Exit Sub
                End If
'                If Not ValidInputNP(rdata) Then Exit Sub
                
                Ver = ReadField(5, rdata, 44)
                If VersionOK(Ver) Then
                     Dim miinteger As Integer
                     miinteger = CInt(val(ReadField(37, rdata, 44)))
                     
                     If (UserList(UserIndex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(UserIndex).flags.ValCoDe) <> CInt(val(ReadField(37, rdata, 44)))) Then
                         Call LogHackAttemp("IP:" & UserList(UserIndex).ip & " intento crear un bot.")
                         Call CloseSocket(UserIndex)
                         Exit Sub
                     End If
                     
                     Call ConnectNewUser(UserIndex, ReadField(1, rdata, 44), ReadField(2, rdata, 44), val(ReadField(3, rdata, 44)), ReadField(4, rdata, 44), ReadField(6, rdata, 44), ReadField(7, rdata, 44), _
                     ReadField(8, rdata, 44), ReadField(9, rdata, 44), ReadField(10, rdata, 44), ReadField(11, rdata, 44), ReadField(12, rdata, 44), ReadField(13, rdata, 44), _
                     ReadField(14, rdata, 44), ReadField(15, rdata, 44), ReadField(16, rdata, 44), ReadField(17, rdata, 44), ReadField(18, rdata, 44), ReadField(19, rdata, 44), _
                     ReadField(20, rdata, 44), ReadField(21, rdata, 44), ReadField(22, rdata, 44), ReadField(23, rdata, 44), ReadField(24, rdata, 44), ReadField(25, rdata, 44), _
                     ReadField(26, rdata, 44), ReadField(27, rdata, 44), ReadField(28, rdata, 44), ReadField(29, rdata, 44), ReadField(30, rdata, 44), ReadField(31, rdata, 44), _
                     ReadField(32, rdata, 44), ReadField(33, rdata, 44), ReadField(34, rdata, 44), ReadField(35, rdata, 44), ReadField(36, rdata, 44))
                Else
                     Call SendData(ToIndex, UserIndex, 0, "!!Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en nuestra pagina.")
                     Exit Sub
                End If
                
                Exit Sub
        End Select
    End If
    
Select Case Left$(rdata, 4)
    Case "BORR" ' <<< borra personajes
       On Error GoTo ExitErr1
        rdata = Right$(rdata, Len(rdata) - 4)
        If (UserList(UserIndex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(UserIndex).flags.ValCoDe) <> CInt(val(ReadField(3, rdata, 44)))) Then
                      Call LogHackAttemp("IP:" & UserList(UserIndex).ip & " intento borrar un personaje.")
                      Call CloseSocket(UserIndex)
                      Exit Sub
        End If
        Arg1 = ReadField(1, rdata, 44)
        
        If Not AsciiValidos(Arg1) Then Exit Sub
        
        '¿Existe el personaje?
        If Not FileExist(CharPath & UCase$(Arg1) & ".chr", vbNormal) Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If

        '¿Es el passwd valido?
        If UCase$(ReadField(2, rdata, 44)) <> UCase$(GetVar(CharPath & UCase$(Arg1) & ".chr", "INIT", "Password")) Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If

        'If FileExist(CharPath & ucase$(Arg1) & ".chr", vbNormal) Then
            Dim rt$
            rt$ = App.Path & "\ChrBackUp\" & UCase$(Arg1) & ".bak"
            If FileExist(rt$, vbNormal) Then Kill rt$
            Name CharPath & UCase$(Arg1) & ".chr" As rt$
            Call SendData(ToIndex, UserIndex, 0, "BORROK")
            Exit Sub
ExitErr1:
    Call LogError(Err.Description & " " & rdata)
    Exit Sub
        'End If
End Select

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Si no esta logeado y envia un comando diferente a los
'de arriba cerramos la conexion.
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
If Not UserList(UserIndex).flags.UserLogged Then
    Call LogHackAttemp("Mesaje enviado sin logearse:" & rdata)
'    Call frmMain.Socket2(UserIndex).Disconnect
    Call CloseSocket(UserIndex)
    Exit Sub
End If
  


Select Case UCase$(Left$(rdata, 1))
    Case ";" 'Hablar
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 1)
        If InStr(rdata, "°") Then
            Exit Sub
        End If
    
        '[Consejeros]
        If UserList(UserIndex).flags.Privilegios = 1 Then
            Call LogGM(UserList(UserIndex).Name, "Dijo: " & rdata, True)
        End If
        
        ind = UserList(UserIndex).Char.CharIndex
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & rdata & "°" & str(ind))
        Exit Sub
    Case "-" 'Gritar
        If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
                Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 1)
        If InStr(rdata, "°") Then
            Exit Sub
        End If
        '[Consejeros]
        If UserList(UserIndex).flags.Privilegios = 1 Then
            Call LogGM(UserList(UserIndex).Name, "Grito: " & rdata, True)
        End If

        ind = UserList(UserIndex).Char.CharIndex
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbRed & "°" & rdata & "°" & str(ind))
        Exit Sub
    Case "\" 'Susurrar al oido
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 1)
        tName = ReadField(1, rdata, 32)
        tIndex = NameIndex(tName)
        If tIndex <> 0 Then
            If Len(rdata) <> Len(tName) Then
                tMessage = Right$(rdata, Len(rdata) - (1 + Len(tName)))
            Else
                tMessage = " "
            End If
            If Not EstaPCarea(UserIndex, tIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "||Estas muy lejos del usuario." & FONTTYPE_INFO)
                Exit Sub
            End If
            ind = UserList(UserIndex).Char.CharIndex
            If InStr(tMessage, "°") Then
                Exit Sub
            End If
            
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = 1 Then
                Call LogGM(UserList(UserIndex).Name, "Le dijo a '" & UserList(tIndex).Name & "' " & tMessage, True)
            End If

            Call SendData(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbBlue & "°" & tMessage & "°" & str(ind))
            Call SendData(ToIndex, tIndex, UserList(UserIndex).Pos.Map, "||" & vbBlue & "°" & tMessage & "°" & str(ind))
            Exit Sub
        End If
        Call SendData(ToIndex, UserIndex, 0, "||Usuario inexistente. " & FONTTYPE_INFO)
        Exit Sub
    Case "M" 'Moverse
        
        rdata = Right$(rdata, Len(rdata) - 1)
        
        If Not UserList(UserIndex).flags.Descansar And Not UserList(UserIndex).flags.Meditando _
           And UserList(UserIndex).flags.Paralizado = 0 Then
              Call MoveUserChar(UserIndex, val(rdata))
        ElseIf UserList(UserIndex).flags.Descansar Then
          UserList(UserIndex).flags.Descansar = False
          Call SendData(ToIndex, UserIndex, 0, "DOK")
          Call SendData(ToIndex, UserIndex, 0, "||Has dejado de descansar." & FONTTYPE_INFO)
          Call MoveUserChar(UserIndex, val(rdata))
        ElseIf UserList(UserIndex).flags.Meditando Then
          UserList(UserIndex).flags.Meditando = False
          Call SendData(ToIndex, UserIndex, 0, "MEDOK")
          Call SendData(ToIndex, UserIndex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
          UserList(UserIndex).Char.FX = 0
          UserList(UserIndex).Char.loops = 0
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
          Call MoveUserChar(UserIndex, val(rdata))
        Else
          Call SendData(ToIndex, UserIndex, 0, "||No podes moverte porque estas paralizado." & FONTTYPE_INFO)
        End If
        
        If UserList(UserIndex).flags.Oculto = 1 Then
            
            If UCase$(UserList(UserIndex).Clase) <> "LADRON" Then
                Call SendData(ToIndex, UserIndex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
                UserList(UserIndex).flags.Oculto = 0
                UserList(UserIndex).flags.Invisible = 0
                Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",0")
            End If
            
        End If

        Exit Sub
End Select

Select Case UCase$(rdata)
    Case "RPU" 'Pedido de actualizacion de la posicion
        Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.x & "," & UserList(UserIndex).Pos.y)
        Exit Sub
    Case "AT"
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡No podes atacar a nadie porque estas muerto!!. " & FONTTYPE_INFO)
            Exit Sub
        End If
        '[Consejeros]
        If UserList(UserIndex).flags.Privilegios = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar a nadie. " & FONTTYPE_INFO)
            Exit Sub
        End If
        If Not UserList(UserIndex).flags.ModoCombate Then
            Call SendData(ToIndex, UserIndex, 0, "||No estas en modo de combate, presiona la tecla ""C"" para pasar al modo combate. " & FONTTYPE_INFO)
        Else
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                            Call SendData(ToIndex, UserIndex, 0, "||No podés usar asi esta arma." & FONTTYPE_INFO)
                            Exit Sub
                End If
            End If
            Call UsuarioAtaca(UserIndex)
        End If
        Exit Sub
    Case "AG"
        If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden tomar objetos. " & FONTTYPE_INFO)
                Exit Sub
        End If
        '[Consejeros]
        If UserList(UserIndex).flags.Privilegios = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes tomar ningun objeto. " & FONTTYPE_INFO)
                Exit Sub
        End If
        Call GetObj(UserIndex)
        Exit Sub
    Case "TAB" 'Entrar o salir modo combate
        If UserList(UserIndex).flags.ModoCombate Then
            Call SendData(ToIndex, UserIndex, 0, "||Has salido del modo de combate. " & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Has pasado al modo de combate. " & FONTTYPE_INFO)
        End If
        UserList(UserIndex).flags.ModoCombate = Not UserList(UserIndex).flags.ModoCombate
        Exit Sub
    Case "SEG" 'Activa / desactiva el seguro
        If UserList(UserIndex).flags.Seguro Then
              Call SendData(ToIndex, UserIndex, 0, "||Has desactivado el seguro. " & FONTTYPE_INFO)
        Else
              Call SendData(ToIndex, UserIndex, 0, "||Has activado el seguro. " & FONTTYPE_INFO)
        End If
        UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
        Exit Sub
    Case "ACTUALIZAR"
        Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.x & "," & UserList(UserIndex).Pos.y)
        Exit Sub
    Case "/ONLINE"
        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||Número de usuarios: " & NumUsers & FONTTYPE_INFO)
        Exit Sub
    Case "/SALIR"
        'Call SendData(ToIndex, UserIndex, 0, "FINOK")
        Cerrar_Usuario (UserIndex)
        Exit Sub
    Case "/FUNDARCLAN"
        If UserList(UserIndex).GuildInfo.FundoClan = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya has fundado un clan, solo se puede fundar uno por personaje." & FONTTYPE_INFO)
            Exit Sub
        End If
        If CanCreateGuild(UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "SHOWFUN" & FONTTYPE_INFO)
        End If
        Exit Sub
    Case "GLINFO"
        If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
                    Call SendGuildLeaderInfo(UserIndex)
        Else
                    Call SendGuildsList(UserIndex)
        End If
        Exit Sub
    Case "/BALANCE"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                  Exit Sub
        End If
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
        Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
              Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
              CloseSocket (UserIndex)
              Exit Sub
        End If
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
        Exit Sub
    Case "/QUIETO" ' << Comando a mascotas
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
         End If
         'Se asegura que el target es un npc
         If UserList(UserIndex).flags.TargetNpc = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
         End If
         If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
         End If
         If Npclist(UserList(UserIndex).flags.TargetNpc).MaestroUser <> _
            UserIndex Then Exit Sub
         Npclist(UserList(UserIndex).flags.TargetNpc).Movement = ESTATICO
         Call Expresar(UserList(UserIndex).flags.TargetNpc, UserIndex)
         Exit Sub
    Case "/ACOMPAÑAR"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
        End If
        If Npclist(UserList(UserIndex).flags.TargetNpc).MaestroUser <> _
          UserIndex Then Exit Sub
        Call FollowAmo(UserList(UserIndex).flags.TargetNpc)
        Call Expresar(UserList(UserIndex).flags.TargetNpc, UserIndex)
        Exit Sub
    Case "/ENTRENAR"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
        End If
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Sub
        Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNpc)
        Exit Sub
    Case "/DESCANSAR"
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
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
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        Call SendData(ToIndex, UserIndex, 0, "MEDOK")
        If Not UserList(UserIndex).flags.Meditando Then
           Call SendData(ToIndex, UserIndex, 0, "||Comenzas a meditar." & FONTTYPE_INFO)
        Else
           Call SendData(ToIndex, UserIndex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
        End If
        UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando
        If UserList(UserIndex).flags.Meditando Then
            UserList(UserIndex).Char.loops = LoopAdEternum
            If UserList(UserIndex).Stats.ELV < 15 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARCHICO & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARCHICO
            ElseIf UserList(UserIndex).Stats.ELV < 30 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARMEDIANO & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARMEDIANO
            Else
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARGRANDE & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARGRANDE
            End If
        Else
            UserList(UserIndex).Char.FX = 0
            UserList(UserIndex).Char.loops = 0
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
        End If
        Exit Sub
    Case "/RESUCITAR"
       'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 1 _
       Or UserList(UserIndex).flags.Muerto <> 1 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
           Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede resucitarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
           Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
           CloseSocket (UserIndex)
           Exit Sub
       End If
       Call RevivirUsuario(UserIndex)
       Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sido resucitado!!" & FONTTYPE_INFO)
       Exit Sub
    Case "/CURAR"
       'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 1 _
       Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
           Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
       Call SendUserStatsBox(val(UserIndex))
       Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sido curado!!" & FONTTYPE_INFO)
       Exit Sub
    Case "/AYUDA"
       Call SendHelp(UserIndex)
       Exit Sub
     Case "/EST"
        Call SendUserStatsTxt(UserIndex, UserIndex)
        Exit Sub
    Case "ATRI"
        Call EnviarAtrib(UserIndex)
        Exit Sub
    Case "FAMA"
        Call EnviarFama(UserIndex)
        Exit Sub
    Case "ESKI"
        Call EnviarSkills(UserIndex)
        Exit Sub
    Case "/COMERCIAR"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        If UserList(UserIndex).flags.Privilegios = 1 Then
            Exit Sub
        End If
        '¿El target es un NPC valido?
        If UserList(UserIndex).flags.TargetNpc > 0 Then
              '¿El NPC puede comerciar?
              If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                 If Len(Npclist(UserList(UserIndex).flags.TargetNpc).Desc) > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                 Exit Sub
              End If
              If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                  Exit Sub
              End If
              'Iniciamos la rutina pa' comerciar.
              Call IniciarCOmercioNPC(UserIndex)
         '[Alejo]
#If False = True Then
        ElseIf UserList(UserIndex).flags.TargetUser > 0 Then
            'Comercio con otro usuario
            'Puede comerciar ?
            If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
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
            UserList(UserIndex).ComUsu.Cant = 0
            UserList(UserIndex).ComUsu.Objeto = 0
            UserList(UserIndex).ComUsu.Acepto = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)
#End If
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
        End If
        Exit Sub
    '[/Alejo]
    '[KEVIN]------------------------------------------
    Case "/BOVEDA"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        '¿El target es un NPC valido?
        If UserList(UserIndex).flags.TargetNpc > 0 Then
              If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                  Exit Sub
              End If
              If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype = 4 Then
                Call IniciarDeposito(UserIndex)
              Else
                Exit Sub
              End If
        Else
          Call SendData(ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
        End If
        Exit Sub
    '[/KEVIN]------------------------------------
    '[Alejo]
    Case "FINCOM"
        'User sale del modo COMERCIO
        UserList(UserIndex).flags.Comerciando = False
        Call SendData(ToIndex, UserIndex, 0, "FINCOMOK")
        Exit Sub
        Case "FINCOMUSU"
        'Sale modo comercio Usuario
        If UserList(UserIndex).ComUsu.DestUsu > 0 And _
            UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
            Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha dejado de comerciar con vos." & FONTTYPE_TALK)
            Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
        End If
        
        Call FinComerciarUsu(UserIndex)
        Exit Sub
    '[KEVIN]---------------------------------------
    '******************************************************
    Case "FINBAN"
        'User sale del modo BANCO
        UserList(UserIndex).flags.Comerciando = False
        Call SendData(ToIndex, UserIndex, 0, "FINBANOK")
        Exit Sub
    '-------------------------------------------------------
    '[/KEVIN]**************************************
    Case "COMUSUOK"
        'Aceptar el cambio
        Call AceptarComercioUsu(UserIndex)
        Exit Sub
    Case "COMUSUNO"
        'Rechazar el cambio
        If UserList(UserIndex).ComUsu.DestUsu > 0 Then
            Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha rechazado tu oferta." & FONTTYPE_TALK)
            Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
        End If
        Call SendData(ToIndex, UserIndex, 0, "||Has rechazado la oferta del otro usuario." & FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)
        Exit Sub
    '[/Alejo]

    Case "/ENLISTAR"
        'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 5 _
       Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
       
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 0 Then
              Call EnlistarArmadaReal(UserIndex)
       Else
              Call EnlistarCaos(UserIndex)
       End If
       
       Exit Sub
    Case "/INFORMACION"
       'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 5 _
       Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
       
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 0 Then
            If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
       Else
            If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las fuerzas del caos!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
       End If
       Exit Sub
    Case "/RECOMPENSA"
       'Se asegura que el target es un npc
       If UserList(UserIndex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 5 _
       Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).flags.TargetNpc).flags.Faccion = 0 Then
            If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call RecompensaArmadaReal(UserIndex)
       Else
            If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las fuerzas del caos!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call RecompensaCaos(UserIndex)
       End If
       Exit Sub
End Select


Select Case UCase$(Left$(rdata, 2))
'    Case "/Z"
'        Dim Pos As WorldPos, Pos2 As WorldPos
'        Dim O As Obj
'
'        For LoopC = 1 To 100
'            Pos = UserList(UserIndex).Pos
'            O.Amount = 1
'            O.ObjIndex = iORO
'            'Exit For
'            Call TirarOro(100000, UserIndex)
'            'Call Tilelibre(Pos, Pos2)
'            'If Pos2.x = 0 Or Pos2.y = 0 Then Exit For
'
'            'Call MakeObj(ToMap, 0, UserList(UserIndex).Pos.Map, O, Pos2.Map, Pos2.x, Pos2.y)
'        Next LoopC
'
'        Exit Sub
    Case "TI" 'Tirar item
            If UserList(UserIndex).flags.Navegando = 1 Or _
               UserList(UserIndex).flags.Muerto = 1 Or _
               UserList(UserIndex).flags.Privilegios = 1 Then Exit Sub
               '[Consejeros]
            
            rdata = Right$(rdata, Len(rdata) - 2)
            Arg1 = ReadField(1, rdata, 44)
            Arg2 = ReadField(2, rdata, 44)
            If val(Arg1) = FLAGORO Then
                Call TirarOro(val(Arg2), UserIndex)
                Call SendUserStatsBox(UserIndex)
                Exit Sub
            Else
                If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
                    If UserList(UserIndex).Invent.Object(val(Arg1)).ObjIndex = 0 Then
                            Exit Sub
                    End If
                    Call DropObj(UserIndex, val(Arg1), val(Arg2), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y)
                Else
                    Exit Sub
                End If
            End If
            Exit Sub
    Case "LH" ' Lanzar hechizo
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!." & FONTTYPE_INFO)
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 2)
        UserList(UserIndex).flags.Hechizo = val(rdata)
        Exit Sub
    Case "LC" 'Click izquierdo
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        x = CInt(Arg1)
        y = CInt(Arg2)
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, y)
        Exit Sub
    Case "RC" 'Click derecho
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        x = CInt(Arg1)
        y = CInt(Arg2)
        Call Accion(UserIndex, UserList(UserIndex).Pos.Map, x, y)
        Exit Sub
    Case "UK"
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!." & FONTTYPE_INFO)
            Exit Sub
        End If

        rdata = Right$(rdata, Len(rdata) - 2)
        Select Case val(rdata)
            Case Robar
                Call SendData(ToIndex, UserIndex, 0, "T01" & Robar)
            Case Magia
                Call SendData(ToIndex, UserIndex, 0, "T01" & Magia)
            Case Domar
                Call SendData(ToIndex, UserIndex, 0, "T01" & Domar)
            Case Ocultarse
                
                If UserList(UserIndex).flags.Navegando = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||No podes ocultarte si estas navegando." & FONTTYPE_INFO)
                      Exit Sub
                End If
                
                If UserList(UserIndex).flags.Oculto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Ya estas oculto." & FONTTYPE_INFO)
                      Exit Sub
                End If
                
                Call DoOcultarse(UserIndex)
        End Select
        Exit Sub
End Select

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
    Case "USA"
        rdata = Right$(rdata, Len(rdata) - 3)
        If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) > 0 Then
            If UserList(UserIndex).Invent.Object(val(rdata)).ObjIndex = 0 Then Exit Sub
        Else
            Exit Sub
        End If
        Call UseInvItem(UserIndex, val(rdata))
        Exit Sub
    Case "CNS" ' Construye herreria
        rdata = Right$(rdata, Len(rdata) - 3)
        x = CInt(rdata)
        If x < 1 Then Exit Sub
        If ObjData(x).SkHerreria = 0 Then Exit Sub
        Call HerreroConstruirItem(UserIndex, x)
        Exit Sub
    Case "CNC" ' Construye carpinteria
        rdata = Right$(rdata, Len(rdata) - 3)
        x = CInt(rdata)
        If x < 1 Or ObjData(x).SkCarpinteria = 0 Then Exit Sub
        Call CarpinteroConstruirItem(UserIndex, x)
        Exit Sub
    Case "WLC" 'Click izquierdo en modo trabajo
        rdata = Right$(rdata, Len(rdata) - 3)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        Arg3 = ReadField(3, rdata, 44)
        If Arg3 = "" Or Arg2 = "" Or Arg1 = "" Then Exit Sub
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(Arg3) Then Exit Sub
        
        x = CInt(Arg1)
        y = CInt(Arg2)
        tLong = CInt(Arg3)
        
        If UserList(UserIndex).flags.Muerto = 1 Or _
           UserList(UserIndex).flags.Descansar Or _
           UserList(UserIndex).flags.Meditando Or _
           Not InMapBounds(UserList(UserIndex).Pos.Map, x, y) Then Exit Sub
                          
        If Not InRangoVision(UserIndex, x, y) Then
            Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.x & "," & UserList(UserIndex).Pos.y)
            Exit Sub
        End If
        
        Select Case tLong
        
        Case Proyectiles
            Dim TU As Integer, tN As Integer
            'Nos aseguramos que este usando un arma de proyectiles
            If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then Exit Sub
            
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then Exit Sub
             
            If UserList(UserIndex).Invent.MunicionEqpObjIndex = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes municiones." & FONTTYPE_INFO)
                    Exit Sub
            End If
             
            'Quitamos stamina
            If UserList(UserIndex).Stats.MinSta >= 10 Then
                 Call QuitarSta(UserIndex, RandomNumber(1, 10))
            Else
                 Call SendData(ToIndex, UserIndex, 0, "||Estas muy cansado para luchar." & FONTTYPE_INFO)
                 Exit Sub
            End If
             
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, Arg1, Arg2)
            
            TU = UserList(UserIndex).flags.TargetUser
            tN = UserList(UserIndex).flags.TargetNpc
            
            
            If tN > 0 Then
                If Npclist(tN).Attackable = 0 Then Exit Sub
            Else
                If TU = 0 Then Exit Sub
            End If
            
            If tN > 0 Then Call UsuarioAtacaNpc(UserIndex, tN)
                
            If TU > 0 Then
                If UserList(UserIndex).flags.Seguro Then
                        If Not Criminal(TU) Then
                                Call SendData(ToIndex, UserIndex, 0, "||No podes atacar ciudadanos, para hacerlo debes desactivar el seguro apretando la tecla S" & FONTTYPE_FIGHT)
                                Exit Sub
                        End If
                End If

                Call UsuarioAtacaUsuario(UserIndex, TU)
            End If
            
            Dim DummyInt As Integer
            DummyInt = UserList(UserIndex).Invent.MunicionEqpSlot
            Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 1)
            If DummyInt < 1 Or DummyInt > MAX_INVENTORY_SLOTS Then Exit Sub
            If UserList(UserIndex).Invent.Object(DummyInt).Amount > 0 Then
                UserList(UserIndex).Invent.Object(DummyInt).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpSlot = DummyInt
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(DummyInt).ObjIndex
                Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
            Else
                Call UpdateUserInv(False, UserIndex, DummyInt)
                UserList(UserIndex).Invent.MunicionEqpSlot = 0
                UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
            End If
            
        Case Magia
            If UserList(UserIndex).flags.PuedeLanzarSpell = 0 Then Exit Sub
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = 1 Then Exit Sub
            
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, y)
            
            If UserList(UserIndex).flags.Hechizo > 0 Then
                Call LanzarHechizo(UserList(UserIndex).flags.Hechizo, UserIndex)
                UserList(UserIndex).flags.PuedeLanzarSpell = 0
                UserList(UserIndex).flags.Hechizo = 0
            Else
                Call SendData(ToIndex, UserIndex, 0, "||¡Primero selecciona el hechizo que quieres lanzar!" & FONTTYPE_INFO)
            End If
        Case Pesca
                  
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> OBJTYPE_CAÑA Then
                    Call CloseSocket(UserIndex)
                    Exit Sub
            End If
            
            If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
            
            If HayAgua(UserList(UserIndex).Pos.Map, x, y) Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_PESCAR)
                Call DoPescar(UserIndex)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No hay agua donde pescar busca un lago, rio o mar." & FONTTYPE_INFO)
            End If
            
        Case Robar
           If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
                If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, y)
                
                If UserList(UserIndex).flags.TargetUser > 0 And UserList(UserIndex).flags.TargetUser <> UserIndex Then
                   If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 0 Then
                        wpaux.Map = UserList(UserIndex).Pos.Map
                        wpaux.x = val(ReadField(1, rdata, 44))
                        wpaux.y = val(ReadField(2, rdata, 44))
                        If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                            Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        '17/09/02
                        'No aseguramos que el trigger le permite robar
                        If MapData(UserList(UserList(UserIndex).flags.TargetUser).Pos.Map, UserList(UserList(UserIndex).flags.TargetUser).Pos.x, UserList(UserList(UserIndex).flags.TargetUser).Pos.y).trigger = 4 Then
                            Call SendData(ToIndex, UserIndex, 0, "||No podes robar aquí." & FONTTYPE_WARNING)
                            Exit Sub
                        End If

                        Call DoRobar(UserIndex, UserList(UserIndex).flags.TargetUser)
                   End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||No a quien robarle!." & FONTTYPE_INFO)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||¡No podes robarle en zonas seguras!." & FONTTYPE_INFO)
            End If
        Case Talar
            
            If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Deberías equiparte el hacha." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                    Call CloseSocket(UserIndex)
                    Exit Sub
            End If
            
            auxind = MapData(UserList(UserIndex).Pos.Map, x, y).OBJInfo.ObjIndex
            If auxind > 0 Then
                wpaux.Map = UserList(UserIndex).Pos.Map
                wpaux.x = x
                wpaux.y = y
                If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                    Exit Sub
                End If
                '¿Hay un arbol donde clickeo?
                If ObjData(auxind).ObjType = OBJTYPE_ARBOLES Then
                    Call SendData(ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SOUND_TALAR)
                    Call DoTalar(UserIndex)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No hay ningun arbol ahi." & FONTTYPE_INFO)
            End If
        Case Mineria
            
            If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> PIQUETE_MINERO Then
                    Call CloseSocket(UserIndex)
                    Exit Sub
            End If
            
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, y)
            
            auxind = MapData(UserList(UserIndex).Pos.Map, x, y).OBJInfo.ObjIndex
            If auxind > 0 Then
                wpaux.Map = UserList(UserIndex).Pos.Map
                wpaux.x = x
                wpaux.y = y
                If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                    Exit Sub
                End If
                '¿Hay un yacimiento donde clickeo?
                If ObjData(auxind).ObjType = OBJTYPE_YACIMIENTO Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_MINERO)
                    Call DoMineria(UserIndex)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
            End If
        Case Domar
          'Modificado 25/11/02
          'Optimizado y solucionado el bug de la doma de
          'criaturas hostiles.
          Dim CI As Integer
          
          Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, y)
          CI = UserList(UserIndex).flags.TargetNpc
          
          If CI > 0 Then
                   If Npclist(CI).flags.Domable > 0 Then
                        wpaux.Map = UserList(UserIndex).Pos.Map
                        wpaux.x = x
                        wpaux.y = y
                        If Distancia(wpaux, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 2 Then
                              Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                              Exit Sub
                        End If
                        If Npclist(CI).flags.AttackedBy <> "" Then
                              Call SendData(ToIndex, UserIndex, 0, "||No podés domar una criatura que está luchando con un jugador." & FONTTYPE_INFO)
                              Exit Sub
                        End If
                        Call DoDomar(UserIndex, CI)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||No podes domar a esa criatura." & FONTTYPE_INFO)
                    End If
          Else
                 Call SendData(ToIndex, UserIndex, 0, "||No hay ninguna criatura alli!." & FONTTYPE_INFO)
          End If
          
        Case FundirMetal
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, y)
            
            If UserList(UserIndex).flags.TargetObj > 0 Then
                If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = OBJTYPE_FRAGUA Then
                    Call FundirMineral(UserIndex)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)
            End If
            
        Case Herreria
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, y)
            
            If UserList(UserIndex).flags.TargetObj > 0 Then
                If ObjData(UserList(UserIndex).flags.TargetObj).ObjType = OBJTYPE_YUNQUE Then
                    Call EnivarArmasConstruibles(UserIndex)
                    Call EnivarArmadurasConstruibles(UserIndex)
                    Call SendData(ToIndex, UserIndex, 0, "SFH")
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)
            End If
            
        End Select
        
        UserList(UserIndex).flags.PuedeTrabajar = 0
        Exit Sub
    Case "CIG"
        rdata = Right$(rdata, Len(rdata) - 3)
        x = Guilds.Count
        
        If CreateGuild(UserList(UserIndex).Name, UserList(UserIndex).Reputacion.Promedio, UserIndex, rdata) Then
            If x = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Felicidades has creado el primer clan de Argentum!!!." & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Felicidades has creado el clan numero " & x + 1 & " de Argentum!!!." & FONTTYPE_INFO)
            End If
            Call SaveGuildsDB
        End If
        
        Exit Sub
End Select

Select Case UCase$(Left$(rdata, 4))
    Case "INFS" 'Informacion del hechizo
            rdata = Right$(rdata, Len(rdata) - 4)
            If val(rdata) > 0 And val(rdata) < MAXUSERHECHIZOS + 1 Then
                Dim H As Integer
                H = UserList(UserIndex).Stats.UserHechizos(val(rdata))
                If H > 0 And H < NumeroHechizos + 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "||Nombre:" & Hechizos(H).Nombre & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "||Descripcion:" & Hechizos(H).Desc & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "||Skill requerido: " & Hechizos(H).MinSkill & " de magia." & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "||Mana necesario: " & Hechizos(H).ManaRequerido & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "||%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%" & FONTTYPE_INFO)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||¡Primero selecciona el hechizo.!" & FONTTYPE_INFO)
            End If
            Exit Sub
   Case "EQUI"
            If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 4)
            If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) > 0 Then
                 If UserList(UserIndex).Invent.Object(val(rdata)).ObjIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
            Call EquiparInvItem(UserIndex, val(rdata))
            Exit Sub
    Case "CHEA" 'Cambiar Heading ;-)
        rdata = Right$(rdata, Len(rdata) - 4)
        If val(rdata) > 0 And val(rdata) < 5 Then
            UserList(UserIndex).Char.Heading = rdata
            Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
        Exit Sub
    Case "SKSE" 'Modificar skills
        Dim i As Integer
        Dim sumatoria As Integer
        Dim incremento As Integer
        rdata = Right$(rdata, Len(rdata) - 4)
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rdata, 44))
            
            If incremento < 0 Then
                'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                UserList(UserIndex).Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            sumatoria = sumatoria + incremento
        Next i
        
        If sumatoria > UserList(UserIndex).Stats.SkillPts Then
            'UserList(UserIndex).Flags.AdministrativeBan = 1
            'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
            Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rdata, 44))
            UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts - incremento
            UserList(UserIndex).Stats.UserSkills(i) = UserList(UserIndex).Stats.UserSkills(i) + incremento
            If UserList(UserIndex).Stats.UserSkills(i) > 100 Then UserList(UserIndex).Stats.UserSkills(i) = 100
        Next i
        Exit Sub
    Case "ENTR" 'Entrena hombre!
        
        If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
        
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 3 Then Exit Sub
        
        rdata = Right$(rdata, Len(rdata) - 4)
        
        If Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas < MAXMASCOTASENTRENADOR Then
            If val(rdata) > 0 And val(rdata) < Npclist(UserList(UserIndex).flags.TargetNpc).NroCriaturas + 1 Then
                    Dim SpawnedNpc As Integer
                    SpawnedNpc = SpawnNpc(Npclist(UserList(UserIndex).flags.TargetNpc).Criaturas(val(rdata)).NpcIndex, Npclist(UserList(UserIndex).flags.TargetNpc).Pos, True, False)
                    If SpawnedNpc <= MAXNPCS Then
                        Npclist(SpawnedNpc).MaestroNpc = UserList(UserIndex).flags.TargetNpc
                        Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas = Npclist(UserList(UserIndex).flags.TargetNpc).Mascotas + 1
                    End If
            End If
        Else
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No puedo traer mas criaturas, mata las existentes!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
        End If
        
        Exit Sub
    Case "COMP"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                   Exit Sub
         End If
         '¿El target es un NPC valido?
         If UserList(UserIndex).flags.TargetNpc > 0 Then
               '¿El NPC puede comerciar?
               If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                   Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 5)
         'User compra el item del slot rdata
         Call NPCVentaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(UserIndex).flags.TargetNpc)
         Exit Sub
    '[KEVIN]*********************************************************************
    '------------------------------------------------------------------------------------
    Case "RETI"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                   Exit Sub
         End If
         '¿El target es un NPC valido?
         If UserList(UserIndex).flags.TargetNpc > 0 Then
               '¿Es el banquero?
               If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 4 Then
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right(rdata, Len(rdata) - 5)
         'User retira el item del slot rdata
         Call UserRetiraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    '-----------------------------------------------------------------------------------
    '[/KEVIN]****************************************************************************
    Case "VEND"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                   Exit Sub
         End If
         '¿El target es un NPC valido?
         If UserList(UserIndex).flags.TargetNpc > 0 Then
               '¿El NPC puede comerciar?
               If Npclist(UserList(UserIndex).flags.TargetNpc).Comercia = 0 Then
                   Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 5)
         'User compra el item del slot rdata
         Call NPCCompraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    '[KEVIN]-------------------------------------------------------------------------
    '****************************************************************************************
    Case "DEPO"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                   Exit Sub
         End If
         '¿El target es un NPC valido?
         If UserList(UserIndex).flags.TargetNpc > 0 Then
               '¿El NPC puede comerciar?
               If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> 4 Then
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right(rdata, Len(rdata) - 5)
         'User deposita el item del slot rdata
         Call UserDepositaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    '****************************************************************************************
    '[/KEVIN]---------------------------------------------------------------------------------
         
End Select

Select Case UCase$(Left$(rdata, 5))
    Case "DEMSG"
        If UserList(UserIndex).flags.TargetObj > 0 Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Dim f As String, Titu As String, msg As String, f2 As String
        f = App.Path & "\foros\"
        f = f & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
        Titu = ReadField(1, rdata, 176)
        msg = ReadField(2, rdata, 176)
        Dim n2 As Integer, loopme As Integer
        If FileExist(f, vbNormal) Then
            Dim num As Integer
            num = val(GetVar(f, "INFO", "CantMSG"))
            If num > MAX_MENSAJES_FORO Then
                For loopme = 1 To num
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & loopme & ".for"
                Next
                Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
                num = 0
            End If
            n2 = FreeFile
            f2 = Left$(f, Len(f) - 4)
            f2 = f2 & num + 1 & ".for"
            Open f2 For Output As n2
            Print #n2, Titu
            Print #n2, msg
            Call WriteVar(f, "INFO", "CantMSG", num + 1)
        Else
            n2 = FreeFile
            f2 = Left$(f, Len(f) - 4)
            f2 = f2 & "1" & ".for"
            Open f2 For Output As n2
            Print #n2, Titu
            Print #n2, msg
            Call WriteVar(f, "INFO", "CantMSG", 1)
        End If
        Close #n2
        End If
        Exit Sub
    Case "/BUG "
        N = FreeFile
        Open App.Path & "\BUGS\BUGs.log" For Append Shared As N
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
        rdata = Right$(rdata, Len(rdata) - 6)
        If Not AsciiValidos(rdata) Then
            Call SendData(ToIndex, UserIndex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(UserIndex).Desc = rdata
        Call SendData(ToIndex, UserIndex, 0, "||La descripcion a cambiado." & FONTTYPE_INFO)
        Exit Sub
    Case "DESCOD" 'Informacion del hechizo
            rdata = Right$(rdata, Len(rdata) - 6)
            Call UpdateCodexAndDesc(rdata, UserIndex)
            Exit Sub
    Case "/VOTO "
            rdata = Right$(rdata, Len(rdata) - 6)
            Call ComputeVote(UserIndex, rdata)
            Exit Sub
            
 End Select

'[Alejo]
Select Case UCase$(Left$(rdata, 7))
Case "OFRECER"
        rdata = Right$(rdata, Len(rdata) - 7)
        Arg1 = ReadField(1, rdata, Asc(","))
        Arg2 = ReadField(2, rdata, Asc(","))

        If val(Arg1) <= 0 Or val(Arg2) <= 0 Then
            Exit Sub
        End If
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged = False Then
            'sigue vivo el usuario ?
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        Else
            'esta vivo ?
            If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.Muerto = 1 Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            End If
            '//Tiene la cantidad que ofrece ??//'
            If val(Arg1) = FLAGORO Then
                'oro
                If val(Arg2) > UserList(UserIndex).Stats.GLD Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                    Exit Sub
                End If
            Else
                'inventario
                If val(Arg2) > UserList(UserIndex).Invent.Object(val(Arg1)).Amount Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            '[Consejeros]
            If UserList(UserIndex).ComUsu.Objeto > 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes cambiar tu oferta." & FONTTYPE_TALK)
                Exit Sub
            End If
            UserList(UserIndex).ComUsu.Objeto = val(Arg1)
            UserList(UserIndex).ComUsu.Cant = val(Arg2)
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            Else
                '[CORREGIDO]
                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                    'NO NO NO vos te estas pasando de listo...
                    UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False
                    Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha cambiado su oferta." & FONTTYPE_TALK)
                End If
                '[/CORREGIDO]
                'Es la ofrenda de respuesta :)
                Call EnviarObjetoTransaccion(UserList(UserIndex).ComUsu.DestUsu)
            End If
        End If
        Exit Sub
End Select
'[/Alejo]

Select Case UCase$(Left$(rdata, 8))
    Case "ACEPPEAT"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call AcceptPeaceOffer(UserIndex, rdata)
        Exit Sub
    Case "PEACEOFF"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call RecievePeaceOffer(UserIndex, rdata)
        Exit Sub
    Case "PEACEDET"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendPeaceRequest(UserIndex, rdata)
        Exit Sub
    Case "ENVCOMEN"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendPeticion(UserIndex, rdata)
        Exit Sub
    Case "ENVPROPP"
        Call SendPeacePropositions(UserIndex)
        Exit Sub
    Case "DECGUERR"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DeclareWar(UserIndex, rdata)
        Exit Sub
    Case "DECALIAD"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DeclareAllie(UserIndex, rdata)
        Exit Sub
    Case "NEWWEBSI"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SetNewURL(UserIndex, rdata)
        Exit Sub
    Case "ACEPTARI"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call AcceptClanMember(UserIndex, rdata)
        Exit Sub
    Case "RECHAZAR"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DenyRequest(UserIndex, rdata)
        Exit Sub
    Case "ECHARCLA"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call EacharMember(UserIndex, rdata)
        Exit Sub
    Case "/PASSWD "
        rdata = Right$(rdata, Len(rdata) - 8)
        If Len(rdata) < 6 Then
             Call SendData(ToIndex, UserIndex, 0, "||El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
        Else
             Call SendData(ToIndex, UserIndex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
             UserList(UserIndex).Password = rdata
        End If
        Exit Sub
    Case "ACTGNEWS"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call UpdateGuildNews(rdata, UserIndex)
        Exit Sub
    Case "1HRINFO<"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendCharInfo(rdata, UserIndex)
        Exit Sub
End Select


Select Case UCase$(Left$(rdata, 9))
    Case "SOLICITUD"
         rdata = Right$(rdata, Len(rdata) - 9)
         Call SolicitudIngresoClan(UserIndex, rdata)
         Exit Sub
    Case "/RETIRAR " 'RETIRA ORO EN EL BANCO
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
         End If
         'Se asegura que el target es un npc
         If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
              Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 9)
         If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
         Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
         If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
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
              Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
         Else
              Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
         End If
         Call SendUserStatsBox(val(UserIndex))
         Exit Sub
End Select


Select Case UCase$(Left$(rdata, 11))
    Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 11)
        If Npclist(UserList(UserIndex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
        Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNpc).Pos) > 10 Then
              Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
              Exit Sub
        End If
        If CLng(val(rdata)) > 0 And CLng(val(rdata)) <= UserList(UserIndex).Stats.GLD Then
              UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rdata)
              UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rdata)
              Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
        Else
              Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
        End If
        Call SendUserStatsBox(val(UserIndex))
        Exit Sub
  Case "CLANDETAILS"
        rdata = Right$(rdata, Len(rdata) - 11)
        Call SendGuildDetails(UserIndex, rdata)
        Exit Sub
End Select



'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
 If UserList(UserIndex).flags.Privilegios = 0 Then Exit Sub
'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<


'<<<<<<<<<<<<<<<<<<<< Consejeros <<<<<<<<<<<<<<<<<<<<

'/rem comentario
If UCase$(Left$(rdata, 4)) = "/REM" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    Call LogGM(UserList(UserIndex).Name, "Comentario: " & rdata, (UserList(UserIndex).flags.Privilegios = 1))
    Call SendData(ToIndex, UserIndex, 0, "||Comentario salvado..." & FONTTYPE_INFO)
    Exit Sub
End If

'HORA
If UCase$(Left$(rdata, 5)) = "/HORA" Then
    Call LogGM(UserList(UserIndex).Name, "Hora.", (UserList(UserIndex).flags.Privilegios = 1))
    rdata = Right$(rdata, Len(rdata) - 5)
    Call SendData(ToAll, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_INFO)
    Exit Sub
End If

'¿Donde esta?
If UCase$(Left$(rdata, 7)) = "/DONDE " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call SendData(ToIndex, UserIndex, 0, "||Ubicacion  " & UserList(tIndex).Name & ": " & UserList(tIndex).Pos.Map & ", " & UserList(tIndex).Pos.x & ", " & UserList(tIndex).Pos.y & "." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "/Donde", (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If

'Nro de enemigos
If UCase$(Left$(rdata, 6)) = "/NENE " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    If MapaValido(val(rdata)) Then
        Call SendData(ToIndex, UserIndex, 0, "NENE" & NPCHostiles(val(rdata)))
        Call LogGM(UserList(UserIndex).Name, "Numero enemigos en mapa " & rdata, (UserList(UserIndex).flags.Privilegios = 1))
    End If
    Exit Sub
End If

'[Consejeros] '[Consejeros] '[Consejeros] '[Consejeros]

If UCase$(rdata) = "/TELEPLOC" Then
    Call WarpUserChar(UserIndex, UserList(UserIndex).flags.TargetMap, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, True)
    Call LogGM(UserList(UserIndex).Name, "/TELEPLOC a x:" & UserList(UserIndex).flags.TargetX & " Y:" & UserList(UserIndex).flags.TargetY & " Map:" & UserList(UserIndex).Pos.Map, (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If

'Teleportar
If UCase$(Left$(rdata, 7)) = "/TELEP " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    mapa = val(ReadField(2, rdata, 32))
    If Not MapaValido(mapa) Then Exit Sub
    Name = ReadField(1, rdata, 32)
    If Name = "" Then Exit Sub
    If UCase$(Name) <> "YO" Then
        If UserList(UserIndex).flags.Privilegios = 1 Then
            Exit Sub
        End If
        tIndex = NameIndex(Name)
    Else
        tIndex = UserIndex
    End If
    x = val(ReadField(3, rdata, 32))
    y = val(ReadField(4, rdata, 32))
    If Not InMapBounds(mapa, x, y) Then Exit Sub
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call WarpUserChar(tIndex, mapa, x, y, True)
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " transportado." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(tIndex).Name & " hacia " & "Mapa" & mapa & " X:" & x & " Y:" & y, (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/SHOW SOS" Then
    Dim M As String
    For N = 1 To Ayuda.Longitud
        M = Ayuda.VerElemento(N)
        Call SendData(ToIndex, UserIndex, 0, "RSOS" & M)
    Next N
    Call SendData(ToIndex, UserIndex, 0, "MSOS")
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "SOSDONE" Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Call Ayuda.Quitar(rdata)
    Exit Sub
End If

'IR A
If UCase$(Left$(rdata, 5)) = "/IRA " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    

    Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.x, UserList(tIndex).Pos.y + 1, True)
    
    If UserList(UserIndex).flags.AdminInvisible = 0 Then Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " se ha trasportado hacia donde te encontras." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "/IRA " & UserList(tIndex).Name & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.x & " Y:" & UserList(tIndex).Pos.y, (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If

'Haceme invisible vieja!
If UCase$(rdata) = "/INVISIBLE" Then
    Call DoAdminInvisible(UserIndex)
    Call LogGM(UserList(UserIndex).Name, "/INVISIBLE", (UserList(UserIndex).flags.Privilegios = 1))
    Exit Sub
End If


'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
If UserList(UserIndex).flags.Privilegios < 2 Then
    Exit Sub
End If

'INFO DE USER
If UCase$(Left$(rdata, 6)) = "/INFO " Then
    Call LogGM(UserList(UserIndex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 6)
    
    tIndex = NameIndex(rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If

    SendUserStatsTxt UserIndex, tIndex
    Exit Sub
End If

'INV DEL USER
If UCase$(Left$(rdata, 5)) = "/INV " Then
    Call LogGM(UserList(UserIndex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If

    SendUserInvTxt UserIndex, tIndex
    Exit Sub
End If

'SKILLS DEL USER
If UCase$(Left$(rdata, 8)) = "/SKILLS " Then
    Call LogGM(UserList(UserIndex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    tIndex = NameIndex(rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If

    SendUserSkillsTxt UserIndex, tIndex
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/REVIVIR " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    Name = rdata
    If UCase$(Name) <> "YO" Then
        tIndex = NameIndex(Name)
    Else
        tIndex = UserIndex
    End If
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    UserList(tIndex).flags.Muerto = 0
    UserList(tIndex).Stats.MinHP = UserList(tIndex).Stats.MaxHP
    Call DarCuerpoDesnudo(tIndex)
    Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, val(tIndex), UserList(tIndex).Char.Body, UserList(tIndex).OrigChar.Head, UserList(tIndex).Char.Heading, UserList(tIndex).Char.WeaponAnim, UserList(tIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Call SendUserStatsBox(val(tIndex))
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " te há resucitado." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "Resucito a " & UserList(tIndex).Name, False)
    Exit Sub
End If

If UCase$(rdata) = "/ONLINEGM" Then
        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios <> 0 Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No hay GMs Online" & FONTTYPE_INFO)
        End If
        Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/CARCEL " Then
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    Name = ReadField(1, rdata, 32)
    i = val(ReadField(1, rdata, 32))
    Name = Right$(rdata, Len(rdata) - (Len(Name) + 1))
    
    tIndex = NameIndex(Name)
    
'    If ucase$(Name) = "MORGOLOCK" Then Exit Sub
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes encarcelar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If i > 30 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes encarcelar por mas de 30 minutos." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call Encarcelar(tIndex, i, UserList(UserIndex).Name)
    
    Exit Sub
End If

'PERDON
If UCase$(Left$(rdata, 7)) = "/PERDON" Then
    rdata = Right$(rdata, Len(rdata) - 8)
    tIndex = NameIndex(rdata)
    If tIndex > 0 Then
        
        If EsNewbie(tIndex) Then
                Call VolverCiudadano(tIndex)
        Else
                Call LogGM(UserList(UserIndex).Name, "Intento perdonar un personaje de nivel avanzado.", False)
                Call SendData(ToIndex, UserIndex, 0, "||Solo se permite perdonar newbies." & FONTTYPE_INFO)
        End If
        
    End If
    Exit Sub
End If

'Echar usuario
If UCase$(Left$(rdata, 7)) = "/ECHAR " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(rdata)
    If UCase$(rdata) = "MORGOLOCK" Then Exit Sub
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes echar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
        Exit Sub
    End If
        
    Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
    Call CloseSocket(tIndex)
    Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(tIndex).Name, False)
    Exit Sub
End If


If UCase$(Left$(rdata, 5)) = "/BAN " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    tIndex = NameIndex(ReadField(2, rdata, Asc("@")))
    Name = ReadField(1, rdata, Asc("@"))
    
    If UCase$(rdata) = "MORGOLOCK" Then Exit Sub
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > UserList(UserIndex).flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes banear a al alguien de mayor jerarquia." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call LogBan(tIndex, UserIndex, Name)
    
    Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
    Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " Banned a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
    
    'Ponemos el flag de ban a 1
    UserList(tIndex).flags.Ban = 1
    
    If UserList(tIndex).flags.Privilegios > 0 Then
            UserList(UserIndex).flags.Ban = 1
            Call CloseSocket(UserIndex)
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " banned by the server por bannear un Administrador." & FONTTYPE_FIGHT)
    End If
    
    Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(tIndex).Name, False)
    Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(tIndex).Name, False)
    Call CloseSocket(tIndex)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/UNBAN " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Call UnBan(rdata)
    Call LogGM(UserList(UserIndex).Name, "/UNBAN a " & rdata, False)
    Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " unbanned." & FONTTYPE_INFO)
    Exit Sub
End If


'SEGUIR
If UCase$(rdata) = "/SEGUIR" Then
    If UserList(UserIndex).flags.TargetNpc > 0 Then
        Call DoFollow(UserList(UserIndex).flags.TargetNpc, UserList(UserIndex).Name)
    End If
    Exit Sub
End If

'Summon
If UCase$(Left$(rdata, 5)) = "/SUM " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " há sido trasportado." & FONTTYPE_INFO)
    Call WarpUserChar(tIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y + 1, True)
    
    Call LogGM(UserList(UserIndex).Name, "/SUM " & UserList(tIndex).Name & " Map:" & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.x & " Y:" & UserList(UserIndex).Pos.y, False)
    Exit Sub
End If

'Crear criatura
If UCase$(Left$(rdata, 3)) = "/CC" Then
   Call EnviarSpawnList(UserIndex)
   Exit Sub
End If

'Spawn!!!!!
If UCase$(Left$(rdata, 3)) = "SPA" Then
    rdata = Right$(rdata, Len(rdata) - 3)
    
    If val(rdata) > 0 And val(rdata) < UBound(SpawnList) + 1 Then _
          Call SpawnNpc(SpawnList(val(rdata)).NpcIndex, UserList(UserIndex).Pos, True, False)
          
          Call LogGM(UserList(UserIndex).Name, "Sumoneo " & SpawnList(val(rdata)).NpcName, False)
          
    Exit Sub
End If

'Resetea el inventario
If UCase$(rdata) = "/RESETINV" Then
    rdata = Right$(rdata, Len(rdata) - 9)
    If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
    Call ResetNpcInv(UserList(UserIndex).flags.TargetNpc)
    Call LogGM(UserList(UserIndex).Name, "/RESETINV " & Npclist(UserList(UserIndex).flags.TargetNpc).Name, False)
    Exit Sub
End If

'/Clean
If UCase$(rdata) = "/LIMPIAR" Then
    Call LimpiarMundo
    Exit Sub
End If

'Mensaje del servidor
If UCase$(Left$(rdata, 6)) = "/RMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(UserIndex).Name, "Mensaje Broadcast:" & rdata, False)
    If rdata <> "" Then
        Call SendData(ToAll, 0, 0, "||" & rdata & FONTTYPE_TALK & ENDC)
    End If
    Exit Sub
End If

'Ip del nick
If UCase$(Left$(rdata, 8)) = "/IPNICK " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    tIndex = NameIndex(UCase$(rdata))
    If tIndex > 0 Then
       Call SendData(ToIndex, UserIndex, 0, "||El ip de " & rdata & " es " & UserList(tIndex).ip & FONTTYPE_INFO)
    End If
    Exit Sub
End If

'Ip del nick
If UCase$(Left$(rdata, 8)) = "/NICKIP " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    tIndex = IP_Index(rdata)
    If tIndex > 0 Then
       Call SendData(ToIndex, UserIndex, 0, "||El nick del ip " & rdata & " es " & UserList(tIndex).Name & FONTTYPE_INFO)
    End If
    Exit Sub
End If



'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
If UserList(UserIndex).flags.Privilegios < 3 Then
    Exit Sub
End If

'Ban x IP
If UCase(Left(rdata, 6)) = "/BANIP" Then
    Dim BanIP As String, XNick As Boolean
    
    rdata = Right(rdata, Len(rdata) - 7)
    'busca primero la ip del nick
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        XNick = False
        Call LogGM(UserList(UserIndex).Name, "/BanIP " & rdata, False)
        BanIP = rdata
    Else
        XNick = True
        Call LogGM(UserList(UserIndex).Name, "/BanIP " & UserList(tIndex).Name & " - " & UserList(tIndex).ip, False)
        BanIP = UserList(tIndex).ip
    End If
    
    'se fija si esta baneada
    For LoopC = 1 To BanIps.Count
        If BanIps.Item(LoopC) = BanIP Then
            Call SendData(ToIndex, UserIndex, 0, "||La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next LoopC
    
    BanIps.Add BanIP
    Call SendData(ToAdmins, UserIndex, 0, "||" & UserList(UserIndex).Name & " Baneo la IP " & BanIP & FONTTYPE_FIGHT)
    
    If XNick = True Then
        Call LogBan(tIndex, UserIndex, "Ban por IP desde Nick")
        
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " Banned a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
        
        'Ponemos el flag de ban a 1
        UserList(tIndex).flags.Ban = 1
        
        Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(tIndex).Name, False)
        Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(tIndex).Name, False)
        Call CloseSocket(tIndex)
    End If
    
    Exit Sub
End If

'Desbanea una IP
If UCase(Left(rdata, 8)) = "/UNBANIP" Then
    
    
    rdata = Right(rdata, Len(rdata) - 9)
    Call LogGM(UserList(UserIndex).Name, "/UNBANIP " & rdata, False)
    
    For LoopC = 1 To BanIps.Count
        If BanIps.Item(LoopC) = rdata Then
            BanIps.Remove LoopC
            Call SendData(ToIndex, UserIndex, 0, "||La IP " & BanIP & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next LoopC
    
    Call SendData(ToIndex, UserIndex, 0, "||La IP " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
    
    Exit Sub
End If

'Crear Teleport
If UCase(Left(rdata, 3)) = "/CT" Then
    '/ct mapa_dest x_dest y_dest
    rdata = Right(rdata, Len(rdata) - 4)
    Call LogGM(UserList(UserIndex).Name, "/CT: " & rdata, False)
    mapa = ReadField(1, rdata, 32)
    x = ReadField(2, rdata, 32)
    y = ReadField(3, rdata, 32)
    
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y - 1).OBJInfo.ObjIndex > 0 Then
        Exit Sub
    End If
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y - 1).TileExit.Map > 0 Then
        Exit Sub
    End If
    If MapaValido(mapa) = False Or InMapBounds(mapa, x, y) = False Then
        Exit Sub
    End If
    
    Dim ET As Obj
    ET.Amount = 1
    ET.ObjIndex = 378
    
    Call MakeObj(ToMap, 0, UserList(UserIndex).Pos.Map, ET, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y - 1)
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y - 1).TileExit.Map = mapa
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y - 1).TileExit.x = x
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y - 1).TileExit.y = y
    
    Exit Sub
End If

'Destruir Teleport
'toma el ultimo click
If UCase(Left(rdata, 3)) = "/DT" Then
    '/dt
    Call LogGM(UserList(UserIndex).Name, "/DT", False)
    
    mapa = UserList(UserIndex).flags.TargetMap
    x = UserList(UserIndex).flags.TargetX
    y = UserList(UserIndex).flags.TargetY
    
    If ObjData(MapData(mapa, x, y).OBJInfo.ObjIndex).ObjType = OBJTYPE_TELEPORT And _
        MapData(mapa, x, y).TileExit.Map > 0 Then
        Call EraseObj(ToMap, 0, mapa, MapData(mapa, x, y).OBJInfo.Amount, mapa, x, y)
        MapData(mapa, x, y).TileExit.Map = 0
        MapData(mapa, x, y).TileExit.x = 0
        MapData(mapa, x, y).TileExit.y = 0
    End If
    
    Exit Sub
End If

'Destruir
If UCase$(Left$(rdata, 5)) = "/DEST" Then
    Call LogGM(UserList(UserIndex).Name, "/DEST", False)
    rdata = Right$(rdata, Len(rdata) - 5)
    Call EraseObj(ToMap, UserIndex, UserList(UserIndex).Pos.Map, 10000, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y)
    Exit Sub
End If

'Bloquear
If UCase$(Left$(rdata, 5)) = "/BLOQ" Then
    Call LogGM(UserList(UserIndex).Name, "/BLOQ", False)
    rdata = Right$(rdata, Len(rdata) - 5)
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).Blocked = 0 Then
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).Blocked = 1
        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y, 1)
    Else
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).Blocked = 0
        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y, 0)
    End If
    Exit Sub
End If

'Quitar NPC
If UCase$(rdata) = "/MATA" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    If UserList(UserIndex).flags.TargetNpc = 0 Then Exit Sub
    Call QuitarNPC(UserList(UserIndex).flags.TargetNpc)
    Call LogGM(UserList(UserIndex).Name, "/MATA " & Npclist(UserList(UserIndex).flags.TargetNpc).Name, False)
    Exit Sub
End If

'Quita todos los NPCs del area
If UCase$(rdata) = "/MASSKILL" Then
    For y = UserList(UserIndex).Pos.y - MinYBorder + 1 To UserList(UserIndex).Pos.y + MinYBorder - 1
            For x = UserList(UserIndex).Pos.x - MinXBorder + 1 To UserList(UserIndex).Pos.x + MinXBorder - 1
                If x > 0 And y > 0 And x < 101 And y < 101 Then _
                    If MapData(UserList(UserIndex).Pos.Map, x, y).NpcIndex > 0 Then Call QuitarNPC(MapData(UserList(UserIndex).Pos.Map, x, y).NpcIndex)
            Next x
    Next y
    Call LogGM(UserList(UserIndex).Name, "/MASSKILL", False)
    Exit Sub
End If

'Quita todos los NPCs del area
If UCase$(rdata) = "/LIMPIAR" Then
        Call LimpiarMundo
        Exit Sub
End If

'Mensaje del sistema
If UCase$(Left$(rdata, 6)) = "/SMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(UserIndex).Name, "Mensaje de sistema:" & rdata, False)
    Call SendData(ToAll, 0, 0, "!!" & rdata & ENDC)
    
    Exit Sub
End If

'Crear criatura, toma directamente el indice
If UCase$(Left$(rdata, 5)) = "/ACC " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   Call SpawnNpc(val(rdata), UserList(UserIndex).Pos, True, False)
   Exit Sub
End If

'Crear criatura con respawn, toma directamente el indice
If UCase$(Left$(rdata, 6)) = "/RACC " Then
   rdata = Right$(rdata, Len(rdata) - 6)
   Call SpawnNpc(val(rdata), UserList(UserIndex).Pos, True, True)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AI1 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraImperial1 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AI2 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraImperial1 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AI3 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraImperial3 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AI4 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   TunicaMagoImperial = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AC1 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraCaos1 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AC2 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraCaos2 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AC3 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraCaos3 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AC4 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   TunicaMagoCaos = val(rdata)
   Exit Sub
End If



'Comando para depurar la navegacion
If UCase$(rdata) = "/NAVE" Then
    If UserList(UserIndex).flags.Navegando = 1 Then
        UserList(UserIndex).flags.Navegando = 0
    Else
        UserList(UserIndex).flags.Navegando = 1
    End If
    Exit Sub
End If

'Apagamos
If UCase$(rdata) = "/APAGAR" Then
    If UCase$(UserList(UserIndex).Name) <> "MORGOLOCK" Then
        Call LogGM(UserList(UserIndex).Name, "¡¡¡Intento apagar el server!!!", False)
        Exit Sub
    End If
    'Log
    mifile = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #mifile
    Print #mifile, Date & " " & Time & " server apagado por " & UserList(UserIndex).Name & ". "
    Close #mifile
    Unload frmMain
    Exit Sub
End If

'CONDENA
If UCase$(Left$(rdata, 7)) = "/CONDEN" Then
    rdata = Right$(rdata, Len(rdata) - 8)
    tIndex = NameIndex(rdata)
    If tIndex > 0 Then Call VolverCriminal(tIndex)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/RAJAR " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(UCase$(rdata))
    If tIndex > 0 Then
        Call ResetFacciones(tIndex)
    End If
    Exit Sub
End If

'MODIFICA CARACTER
If UCase$(Left$(rdata, 5)) = "/MOD " Then
    Call LogGM(UserList(UserIndex).Name, rdata, False)
    rdata = Right$(rdata, Len(rdata) - 5)
    tIndex = NameIndex(ReadField(1, rdata, 32))
    Arg1 = ReadField(2, rdata, 32)
    Arg2 = ReadField(3, rdata, 32)
    Arg3 = ReadField(4, rdata, 32)
    Arg4 = ReadField(5, rdata, 32)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Select Case UCase$(Arg1)
    
        Case "ORO"
            If val(Arg2) < 95001 Then
                UserList(tIndex).Stats.GLD = val(Arg2)
                Call SendUserStatsBox(tIndex)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No esta permitido utilizar valores mayores a 95000. Su comando ha quedado en los logs del juego." & FONTTYPE_INFO)
                Exit Sub
            End If
        Case "EXP"
            If val(Arg2) < 9995001 Then
                If UserList(tIndex).Stats.Exp + val(Arg2) > _
                   UserList(tIndex).Stats.ELU Then
                   Dim resto
                   resto = val(Arg2) - UserList(tIndex).Stats.ELU
                   UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + UserList(tIndex).Stats.ELU
                   Call CheckUserLevel(tIndex)
                   UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + resto
                Else
                   UserList(tIndex).Stats.Exp = val(Arg2)
                End If
                Call SendUserStatsBox(tIndex)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No esta permitido utilizar valores mayores a 5000. Su comando ha quedado en los logs del juego." & FONTTYPE_INFO)
                Exit Sub
            End If
   
       
        Case "BODY"
            Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, val(Arg2), UserList(tIndex).Char.Head, UserList(tIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            Exit Sub
        Case "HEAD"
            Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, UserList(tIndex).Char.Body, val(Arg2), UserList(tIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            Exit Sub
        Case "CRI"
            UserList(tIndex).Faccion.CriminalesMatados = val(Arg2)
            Exit Sub
        Case "CIU"
            UserList(tIndex).Faccion.CiudadanosMatados = val(Arg2)
            Exit Sub
        Case "LEVEL"
            UserList(tIndex).Stats.ELV = val(Arg2)
            Exit Sub
        Case Else
            Call SendData(ToIndex, UserIndex, 0, "||Comando no permitido." & FONTTYPE_INFO)
            Exit Sub
    End Select

    Exit Sub
End If


If UCase$(Left$(rdata, 9)) = "/DOBACKUP" Then
    Call DoBackUp
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/GRABAR" Then
    Call GuardarUsuarios
    Exit Sub
End If

If UCase$(Left$(rdata, 11)) = "/BORRAR SOS" Then
    Call Ayuda.Reset
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/SHOW INT" Then
    Call frmMain.mnuMostrar_Click
    Exit Sub
End If

If UCase$(rdata) = "/LLUVIA" Then
    Lloviendo = Not Lloviendo
    Call SendData(ToAll, 0, 0, "LLU")
    Exit Sub
End If

If UCase$(rdata) = "/PASSDAY" Then
    Call DayElapsed
    Exit Sub
End If


Exit Sub


ErrorHandler:
 Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(UserIndex).Name & "UI:" & UserIndex & " N: " & Err.number & " D: " & Err.Description)
 'Call CloseSocket(UserIndex)
 Call Cerrar_Usuario(UserIndex)
 
 

End Sub

Sub ReloadSokcet()

On Error GoTo errhandler

    frmMain.Socket1.Cleanup
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)


Exit Sub
errhandler:
    Call LogError("Error en CheckSocketState," & Err.Description)

End Sub



Public Sub EventoSockAccept(SockID As Long)
#If UsarAPI Then
'==========================================================
'USO DE LA API DE WINSOCK
'========================

'Call LogApiSock("EventoSockAccept")

If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Pedido de conexion SocketID:" & SockID & vbCrLf

'On Error Resume Next
    
    Dim NewIndex As Integer
    Dim Ret As Long
    Dim Tam As Long, sa As sockaddr
    Dim NuevoSock As Long
    Dim i As Long
    
    If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "NextOpenUser" & vbCrLf
    
    NewIndex = NextOpenUser ' Nuevo indice
    If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "UserIndex asignado " & NewIndex & vbCrLf
    
    If NewIndex <= MaxUsers Then
        If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Cargando Socket " & NewIndex & vbCrLf
        '=============================================
        'SockID es en este caso es el socket de escucha,
        'a diferencia de socketwrench que es el nuevo
        'socket de la nueva conn
        
        Tam = sockaddr_size
        
        Ret = accept(SockID, sa, Tam)
        If Ret = INVALID_SOCKET Then
            Call LogCriticEvent("Error en Accept() API")
            Exit Sub
        End If
        NuevoSock = Ret
        
        UserList(NewIndex).ip = GetAscIP(sa.sin_addr)
        
        'Busca si esta banneada la ip
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = UserList(NewIndex).ip Then
                Call apiclosesocket(NuevoSock)
                Exit Sub
            End If
        Next i
        
        Call LogApiSock("EventoSockAccept NewIndex: " & NewIndex & " NuevoSock: " & NuevoSock & " IP: " & UserList(NewIndex).ip)
        '=============================================
        If aDos.MaxConexiones(UserList(NewIndex).ip) Then
            UserList(NewIndex).ConnID = -1
            If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "User slot reseteado " & NewIndex & vbCrLf
            If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Socket unloaded" & NewIndex & vbCrLf
            'Call LogCriticEvent(UserList(NewIndex).ip & " intento crear mas de 3 conexiones.")
            Call aDos.RestarConexion(UserList(NewIndex).ip)
            Call apiclosesocket(NuevoSock)
            'Exit Sub
        End If
        
        UserList(NewIndex).ConnID = NuevoSock
        Set UserList(NewIndex).CommandsBuffer = New CColaArray

        If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & UserList(NewIndex).ip & " logged." & vbCrLf
    Else
        Call LogCriticEvent("No acepte conexion porque no tenia slots")
    End If
    
#End If
End Sub

Public Sub EventoSockRead(Slot As Integer, ByRef Datos As String)
#If UsarAPI Then

Dim T() As String
Dim LoopC As Long

UserList(Slot).RDBuffer = UserList(Slot).RDBuffer & Datos

If InStr(1, UserList(Slot).RDBuffer, Chr(2)) > 0 Then
    UserList(Slot).RDBuffer = "CLIENTEVIEJO" & ENDC
    Debug.Print "CLIENTEVIEJO"
End If

T = Split(UserList(Slot).RDBuffer, ENDC)
If UBound(T) > 0 Then
    UserList(Slot).RDBuffer = T(UBound(T))
    
    For LoopC = 0 To UBound(T) - 1
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
        '%%% EL PROBLEMA DEL SPEEDHACK          %%%
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        If ClientsCommandsQueue = 1 Then
            If T(LoopC) <> "" Then If Not UserList(Slot).CommandsBuffer.Push(T(LoopC)) Then Call Cerrar_Usuario(Slot)
        
        Else ' SH tiebe efecto
              If UserList(Slot).ConnID <> -1 Then
                Call HandleData(Slot, T(LoopC))
              Else
                Exit Sub
              End If
        End If
    Next LoopC
End If

#End If
End Sub

Public Sub EventoSockClose(Slot As Integer)
#If UsarAPI Then
    If UserList(Slot).flags.UserLogged Then
        Call Cerrar_Usuario(Slot)
    Else
        Call CloseSocket(Slot)
    End If
#End If
End Sub

