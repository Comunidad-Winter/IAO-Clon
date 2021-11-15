Attribute VB_Name = "TCP"
'ImperiumAO 0.11.6
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'ImperiumAO is based on Baronsoft's VB6 Online RPG
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

#If UsarQueSocket = 0 Then
' General constants used with most of the controls
Public Const INVALID_HANDLE As Integer = -1
Public Const CONTROL_ERRIGNORE As Integer = 0
Public Const CONTROL_ERRDISPLAY As Integer = 1


' SocietWrench Control Actions
Public Const SOCKET_OPEN As Integer = 1
Public Const SOCKET_CONNECT As Integer = 2
Public Const SOCKET_LISTEN As Integer = 3
Public Const SOCKET_ACCEPT As Integer = 4
Public Const SOCKET_CANCEL As Integer = 5
Public Const SOCKET_FLUSH As Integer = 6
Public Const SOCKET_CLOSE As Integer = 7
Public Const SOCKET_DISCONNECT As Integer = 7
Public Const SOCKET_ABORT As Integer = 8

' SocketWrench Control States
Public Const SOCKET_NONE As Integer = 0
Public Const SOCKET_IDLE As Integer = 1
Public Const SOCKET_LISTENING As Integer = 2
Public Const SOCKET_CONNECTING As Integer = 3
Public Const SOCKET_ACCEPTING As Integer = 4
Public Const SOCKET_RECEIVING As Integer = 5
Public Const SOCKET_SENDING As Integer = 6
Public Const SOCKET_CLOSING As Integer = 7

' Societ Address Families
Public Const AF_UNSPEC As Integer = 0
Public Const AF_UNIX As Integer = 1
Public Const AF_INET As Integer = 2

' Societ Types
Public Const SOCK_STREAM As Integer = 1
Public Const SOCK_DGRAM As Integer = 2
Public Const SOCK_RAW As Integer = 3
Public Const SOCK_RDM As Integer = 4
Public Const SOCK_SEQPACKET As Integer = 5

' Protocol Types
Public Const IPPROTO_IP As Integer = 0
Public Const IPPROTO_ICMP As Integer = 1
Public Const IPPROTO_GGP As Integer = 2
Public Const IPPROTO_TCP As Integer = 6
Public Const IPPROTO_PUP As Integer = 12
Public Const IPPROTO_UDP As Integer = 17
Public Const IPPROTO_IDP As Integer = 22
Public Const IPPROTO_ND As Integer = 77
Public Const IPPROTO_RAW As Integer = 255
Public Const IPPROTO_MAX As Integer = 256


' Network Addpesses
Public Const INADDR_ANY As String = "0.0.0.0"
Public Const INADDR_LOOPBACK As String = "127.0.0.1"
Public Const INADDR_NONE As String = "255.055.255.255"

' Shutdown Values
Public Const SOCKET_READ As Integer = 0
Public Const SOCKET_WRITE As Integer = 1
Public Const SOCKET_READWRITE As Integer = 2

' SocketWrench Error Pesponse
Public Const SOCKET_ERRIGNORE As Integer = 0
Public Const SOCKET_ERRDISPLAY As Integer = 1

' SocketWrench Error Codes
Public Const WSABASEERR As Integer = 24000
Public Const WSAEINTR As Integer = 24004
Public Const WSAEBADF As Integer = 24009
Public Const WSAEACCES As Integer = 24013
Public Const WSAEFAULT As Integer = 24014
Public Const WSAEINVAL As Integer = 24022
Public Const WSAEMFILE As Integer = 24024
Public Const WSAEWOULDBLOCK As Integer = 24035
Public Const WSAEINPROGRESS As Integer = 24036
Public Const WSAEALREADY As Integer = 24037
Public Const WSAENOTSOCK As Integer = 24038
Public Const WSAEDESTADDRREQ As Integer = 24039
Public Const WSAEMSGSIZE As Integer = 24040
Public Const WSAEPROTOTYPE As Integer = 24041
Public Const WSAENOPROTOOPT As Integer = 24042
Public Const WSAEPROTONOSUPPORT As Integer = 24043
Public Const WSAESOCKTNOSUPPORT As Integer = 24044
Public Const WSAEOPNOTSUPP As Integer = 24045
Public Const WSAEPFNOSUPPORT As Integer = 24046
Public Const WSAEAFNOSUPPORT As Integer = 24047
Public Const WSAEADDRINUSE As Integer = 24048
Public Const WSAEADDRNOTAVAIL As Integer = 24049
Public Const WSAENETDOWN As Integer = 24050
Public Const WSAENETUNREACH As Integer = 24051
Public Const WSAENETRESET As Integer = 24052
Public Const WSAECONNABORTED As Integer = 24053
Public Const WSAECONNRESET As Integer = 24054
Public Const WSAENOBUFS As Integer = 24055
Public Const WSAEISCONN As Integer = 24056
Public Const WSAENOTCONN As Integer = 24057
Public Const WSAESHUTDOWN As Integer = 24058
Public Const WSAETOOMANYREFS As Integer = 24059
Public Const WSAETIMEDOUT As Integer = 24060
Public Const WSAECONNREFUSED As Integer = 24061
Public Const WSAELOOP As Integer = 24062
Public Const WSAENAMETOOLONG As Integer = 24063
Public Const WSAEHOSTDOWN As Integer = 24064
Public Const WSAEHOSTUNREACH As Integer = 24065
Public Const WSAENOTEMPTY As Integer = 24066
Public Const WSAEPROCLIM As Integer = 24067
Public Const WSAEUSERS As Integer = 24068
Public Const WSAEDQUOT As Integer = 24069
Public Const WSAESTALE As Integer = 24070
Public Const WSAEREMOTE As Integer = 24071
Public Const WSASYSNOTREADY As Integer = 24091
Public Const WSAVERNOTSUPPORTED As Integer = 24092
Public Const WSANOTINITIALISED As Integer = 24093
Public Const WSAHOST_NOT_FOUND As Integer = 25001
Public Const WSATRY_AGAIN As Integer = 25002
Public Const WSANO_RECOVERY As Integer = 25003
Public Const WSANO_DATA As Integer = 25004
Public Const WSANO_ADDRESS As Integer = 2500
#End If


Sub DarCuerpoYCabeza(ByVal UserIndex As Integer)
Dim NewBody As Integer
Dim NewHead As Integer
Dim UserRaza As Byte
Dim UserGenero As Byte
UserGenero = UserList(UserIndex).genero
UserRaza = UserList(UserIndex).raza
Select Case UserGenero
   Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                NewHead = RandomNumber(1, 40)
                NewBody = 1
            Case eRaza.Elfo
                NewHead = RandomNumber(101, 112)
                NewBody = 2
            Case eRaza.Drow
                NewHead = RandomNumber(200, 210)
                NewBody = 3
            Case eRaza.Enano
                NewHead = RandomNumber(300, 306)
                NewBody = 52
            Case eRaza.Orco
                NewHead = RandomNumber(501, 514) 'shermie80
                NewBody = 248
            Case eRaza.Gnomo
                NewHead = RandomNumber(401, 406)
                NewBody = 52
        End Select
   Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                NewHead = RandomNumber(70, 79)
                NewBody = 1
            Case eRaza.Elfo
                NewHead = RandomNumber(170, 178)
                NewBody = 2
            Case eRaza.Drow
                NewHead = RandomNumber(270, 278)
                NewBody = 3
            Case eRaza.Gnomo
                NewHead = RandomNumber(370, 372)
                NewBody = 138
            Case eRaza.Enano
                NewHead = RandomNumber(470, 476)
                NewBody = 138
            Case eRaza.Orco
                NewHead = RandomNumber(570, 573) ' shermie
                NewBody = 111 'maycolito (:
        End Select
End Select
UserList(UserIndex).Char.Head = NewHead
UserList(UserIndex).Char.body = NewBody
End Sub


Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
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
    car = Asc(mid$(cad, i, 1))
    
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

Sub ConnectNewUser(ByVal UserIndex As Integer, ByRef name As String, ByRef Account As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, _
                    ByRef skills() As Byte, ByRef UserEmail As String, ByVal Hogar As eCiudad, _
                    ByVal Fuerza As Byte, ByVal Agilidad As Byte, ByVal Inteligencia As Byte, _
                    ByVal Carisma As Byte, ByVal Constitucion As Byte, ByVal Cabeza As Integer)
'*************************************************
'Author: Unknown
'Last modified: 20/4/2007
'Conecta un nuevo Usuario
'23/01/2007 Pablo (ToxicWaste) - Agregu� ResetFaccion al crear usuario
'24/01/2007 Pablo (ToxicWaste) - Agregu� el nuevo mana inicial de los magos.
'12/02/2007 Pablo (ToxicWaste) - Puse + 1 de const al Elfo normal.
'20/04/2007 Pablo (ToxicWaste) - Puse -1 de fuerza al Elfo.
'09/01/2008 Pablo (ToxicWaste) - Ahora los modificadores de Raza se controlan desde Balance.dat
'*************************************************

Dim i As Long

If Not AsciiValidos(name) Or LenB(name) = 0 Then
    Call WriteErrorMsg(UserIndex, "Nombre invalido.")
    Exit Sub
End If

If UserList(UserIndex).flags.UserLogged Then
    Call LogCheating("El usuario " & UserList(UserIndex).name & " ha intentado crear a " & name & " desde la IP " & UserList(UserIndex).ip)
    
    'Kick player ( and leave character inside :D )!
    Call CloseSocketSL(UserIndex)
    Call Cerrar_Usuario(UserIndex)
    
    Exit Sub
End If

Dim LoopC As Long
Dim totalskpts As Long

'�Existe el personaje?
If FileExist(CharPath & UCase$(name) & ".pjs", vbNormal) = True Then
    Call WriteErrorMsg(UserIndex, "Ya existe el personaje.")
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

UserList(UserIndex).name = name
UserList(UserIndex).Clase = UserClase
UserList(UserIndex).raza = UserRaza
UserList(UserIndex).genero = UserSexo
UserList(UserIndex).email = UserEmail
UserList(UserIndex).Hogar = Hogar
UserList(UserIndex).flags.Vip = 0

If Not UserList(UserIndex).flags.Privilegios = PlayerType.Dios Then
If UserList(UserIndex).Hogar = cBanderbill Then
        UserList(UserIndex).Faccion.Ciudadano = 1
        UserList(UserIndex).Faccion.Renegado = 0
        UserList(UserIndex).Faccion.Republicano = 0
        UserList(UserIndex).Faccion.FuerzasCaos = 0
        UserList(UserIndex).Faccion.ArmadaReal = 0
        UserList(UserIndex).Hogar = cBanderbill
    ElseIf UserList(UserIndex).Hogar = cRinkel Then
        UserList(UserIndex).Faccion.Renegado = 1
        UserList(UserIndex).Faccion.Ciudadano = 0
        UserList(UserIndex).Faccion.Republicano = 0
        UserList(UserIndex).Faccion.FuerzasCaos = 0
        UserList(UserIndex).Faccion.ArmadaReal = 0
        UserList(UserIndex).Hogar = cRinkel
    ElseIf UserList(UserIndex).Hogar = cSuramei Then
        UserList(UserIndex).Faccion.Republicano = 1
        UserList(UserIndex).Faccion.Ciudadano = 0
        UserList(UserIndex).Faccion.Renegado = 0
        UserList(UserIndex).Faccion.FuerzasCaos = 0
        UserList(UserIndex).Faccion.ArmadaReal = 0
        UserList(UserIndex).Hogar = cSuramei
    End If
Else
        UserList(UserIndex).Faccion.Rango = 1
        UserList(UserIndex).Faccion.Ciudadano = 0
        UserList(UserIndex).Faccion.Renegado = 0
        UserList(UserIndex).Faccion.FuerzasCaos = 0
        UserList(UserIndex).Faccion.ArmadaReal = 0
        UserList(UserIndex).Faccion.Republicano = 0
End If

UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = Fuerza + ModRaza(UserRaza).Fuerza
UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = Agilidad + ModRaza(UserRaza).Agilidad
UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = Inteligencia + ModRaza(UserRaza).Inteligencia
UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = Carisma + ModRaza(UserRaza).Carisma
UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) = Constitucion + ModRaza(UserRaza).Constitucion

'Tir� los dados antes de llegar ac�??
If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
    Call WriteErrorMsg(UserIndex, "Atributos inv�lidos.")
    Exit Sub
End If

For LoopC = 1 To NUMSKILLS
    UserList(UserIndex).Stats.UserSkills(LoopC) = skills(LoopC - 1)
    totalskpts = totalskpts + Abs(UserList(UserIndex).Stats.UserSkills(LoopC))
Next LoopC

Dim asds As Integer
For asds = 1 To MAXAMIGOS
UserList(UserIndex).Amigos(asds).Nombre = "Nadies"
UserList(UserIndex).Amigos(asds).Ignorado = 0
UserList(UserIndex).Amigos(asds).Index = 0
Next asds


If totalskpts > 10 Then
    Call LogHackAttemp(UserList(UserIndex).name & " intento hackear los skills.")
    Call BorrarUsuario(UserList(UserIndex).name)
    Call CloseSocket(UserIndex)
    Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

UserList(UserIndex).Char.heading = eHeading.SOUTH

Call DarCuerpoYCabeza(UserIndex)
UserList(UserIndex).OrigChar = UserList(UserIndex).Char
UserList(UserIndex).OrigChar.Head = Cabeza
 
UserList(UserIndex).Char.WeaponAnim = NingunArma
UserList(UserIndex).Char.ShieldAnim = NingunEscudo
UserList(UserIndex).Char.CascoAnim = NingunCasco

Dim MiInt As Long
MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 3)

UserList(UserIndex).Stats.MaxHP = 15 + MiInt
UserList(UserIndex).Stats.MinHP = 15 + MiInt

MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(UserIndex).Stats.MaxSta = 20 * MiInt
UserList(UserIndex).Stats.MinSta = 20 * MiInt


UserList(UserIndex).Stats.MaxAGU = 100
UserList(UserIndex).Stats.MinAGU = 100

UserList(UserIndex).Stats.MaxHam = 100
UserList(UserIndex).Stats.MinHam = 100


'<-----------------MANA----------------------->
If UserClase = eClass.Mago Then  'Cambio en mana inicial (ToxicWaste)
    MiInt = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * 3
    UserList(UserIndex).Stats.MaxMAN = MiInt
    UserList(UserIndex).Stats.MinMAN = MiInt
ElseIf UserClase = eClass.Clerigo Or UserClase = eClass.Druida _
    Or UserClase = eClass.Bardo Or UserClase = eClass.Asesino _
    Or UserClase = eClass.Nigromante Then
        UserList(UserIndex).Stats.MaxMAN = 50
        UserList(UserIndex).Stats.MinMAN = 50
        
ElseIf UserClase = eClass.Guerrero Or UserClase = eClass.Cazador Then
      UserList(UserIndex).Stats.MinMAN = 0
      UserList(UserIndex).Stats.MaxMAN = 0
End If
 
If UserClase = eClass.Mago Or UserClase = eClass.Clerigo Or _
   UserClase = eClass.Druida Or UserClase = eClass.Bardo Or _
   UserClase = eClass.Asesino Or UserClase = eClass.Nigromante Then
        UserList(UserIndex).Stats.UserHechizos(1) = 2
End If

UserList(UserIndex).Stats.MaxHIT = 2
UserList(UserIndex).Stats.MinHIT = 1

UserList(UserIndex).Stats.GLD = 0

UserList(UserIndex).Stats.Exp = 0
UserList(UserIndex).Stats.ELU = 15
UserList(UserIndex).Stats.ELV = 1

'???????????????? INVENTARIO ��������������������

'organizacion:
'-cuerpo =1
'-arma = 2
'-pociones = 3
'-comidas = 4
'-runa , etc = 5

UserList(UserIndex).Invent.NroItems = 9

Select Case UserClase

    Case eClass.Asesino
    
    UserList(UserIndex).Invent.Object(1).ObjIndex = 31 ' vestimenta
    UserList(UserIndex).Invent.Object(1).amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1

    UserList(UserIndex).Invent.Object(2).ObjIndex = 574 ' daga
    UserList(UserIndex).Invent.Object(2).amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1

    UserList(UserIndex).Invent.Object(3).ObjIndex = 576 ' daga arrojadisa
    UserList(UserIndex).Invent.Object(3).amount = 100
    UserList(UserIndex).Invent.Object(3).Equipped = 0

    UserList(UserIndex).Invent.Object(4).ObjIndex = 38 ' Pocion Roja
    UserList(UserIndex).Invent.Object(4).amount = 100
    
    UserList(UserIndex).Invent.Object(5).ObjIndex = 573 ' Comida
    UserList(UserIndex).Invent.Object(5).amount = 100

    UserList(UserIndex).Invent.Object(6).ObjIndex = 572 ' Bebida
    UserList(UserIndex).Invent.Object(6).amount = 100
    
    UserList(UserIndex).Invent.Object(7).ObjIndex = 650 ' Runa
    UserList(UserIndex).Invent.Object(7).amount = 1
    
    UserList(UserIndex).Invent.Object(8).ObjIndex = 879 ' Mapa
    UserList(UserIndex).Invent.Object(8).amount = 1
    
'**********************************************************************************
   
    Case eClass.Bardo
    
    UserList(UserIndex).Invent.Object(1).ObjIndex = 31 ' vestimenta
    UserList(UserIndex).Invent.Object(1).amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1

    UserList(UserIndex).Invent.Object(2).ObjIndex = 1354 ' Nudillos de hierro
    UserList(UserIndex).Invent.Object(2).amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1

    UserList(UserIndex).Invent.Object(3).ObjIndex = 38 ' Pocion Roja
    UserList(UserIndex).Invent.Object(3).amount = 100
    
    UserList(UserIndex).Invent.Object(4).ObjIndex = 573 ' Comida
    UserList(UserIndex).Invent.Object(4).amount = 100

    UserList(UserIndex).Invent.Object(5).ObjIndex = 572 ' Bebida
    UserList(UserIndex).Invent.Object(5).amount = 100
    
    UserList(UserIndex).Invent.Object(6).ObjIndex = 650 ' Runa
    UserList(UserIndex).Invent.Object(6).amount = 1
    
    UserList(UserIndex).Invent.Object(7).ObjIndex = 879 ' Mapa
    UserList(UserIndex).Invent.Object(7).amount = 1
    
'***********************************************************************************
    
    Case eClass.Gladiador
    
    UserList(UserIndex).Invent.Object(1).ObjIndex = 31 ' vestimenta
    UserList(UserIndex).Invent.Object(1).amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1

    UserList(UserIndex).Invent.Object(2).ObjIndex = 1354 ' Nudillos de hierro
    UserList(UserIndex).Invent.Object(2).amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1

    UserList(UserIndex).Invent.Object(3).ObjIndex = 38 ' Pocion Roja
    UserList(UserIndex).Invent.Object(3).amount = 100
    
    UserList(UserIndex).Invent.Object(4).ObjIndex = 573 ' Comida
    UserList(UserIndex).Invent.Object(4).amount = 100

    UserList(UserIndex).Invent.Object(5).ObjIndex = 572 ' Bebida
    UserList(UserIndex).Invent.Object(5).amount = 100
    
    UserList(UserIndex).Invent.Object(6).ObjIndex = 650 ' Runa
    UserList(UserIndex).Invent.Object(6).amount = 1
    
    UserList(UserIndex).Invent.Object(7).ObjIndex = 879 ' Mapa
    UserList(UserIndex).Invent.Object(7).amount = 1

'***********************************************************************************

    Case eClass.Mercenario
    
    UserList(UserIndex).Invent.Object(1).ObjIndex = 31 ' vestimenta
    UserList(UserIndex).Invent.Object(1).amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1

    UserList(UserIndex).Invent.Object(2).ObjIndex = 574 ' Martillo
    UserList(UserIndex).Invent.Object(2).amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1

    UserList(UserIndex).Invent.Object(3).ObjIndex = 38 ' Pocion Roja
    UserList(UserIndex).Invent.Object(3).amount = 100
    
    UserList(UserIndex).Invent.Object(4).ObjIndex = 573 ' Comida
    UserList(UserIndex).Invent.Object(4).amount = 100

    UserList(UserIndex).Invent.Object(5).ObjIndex = 572 ' Bebida
    UserList(UserIndex).Invent.Object(5).amount = 100
    
    UserList(UserIndex).Invent.Object(6).ObjIndex = 650 ' Runa
    UserList(UserIndex).Invent.Object(6).amount = 1
    
    UserList(UserIndex).Invent.Object(7).ObjIndex = 879 ' Mapa
    UserList(UserIndex).Invent.Object(7).amount = 1

'***********************************************************************************

    Case eClass.Clerigo
    
    UserList(UserIndex).Invent.Object(1).ObjIndex = 31 ' vestimenta
    UserList(UserIndex).Invent.Object(1).amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1

    UserList(UserIndex).Invent.Object(2).ObjIndex = 574 ' Martillo
    UserList(UserIndex).Invent.Object(2).amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1

    UserList(UserIndex).Invent.Object(3).ObjIndex = 38 ' Pocion Roja
    UserList(UserIndex).Invent.Object(3).amount = 100
    
    UserList(UserIndex).Invent.Object(4).ObjIndex = 573 ' Comida
    UserList(UserIndex).Invent.Object(4).amount = 100

    UserList(UserIndex).Invent.Object(5).ObjIndex = 572 ' Bebida
    UserList(UserIndex).Invent.Object(5).amount = 100
    
    UserList(UserIndex).Invent.Object(6).ObjIndex = 650 ' Runa
    UserList(UserIndex).Invent.Object(6).amount = 1
    
    UserList(UserIndex).Invent.Object(7).ObjIndex = 879 ' Mapa
    UserList(UserIndex).Invent.Object(7).amount = 1
    
'*********************************************************************************
 
    Case eClass.Druida
    UserList(UserIndex).Invent.Object(1).ObjIndex = 31 ' vestimenta
    UserList(UserIndex).Invent.Object(1).amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1

    UserList(UserIndex).Invent.Object(2).ObjIndex = 1355 ' arco
    UserList(UserIndex).Invent.Object(2).amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1

    UserList(UserIndex).Invent.Object(3).ObjIndex = 1357 ' Flecha
    UserList(UserIndex).Invent.Object(3).amount = 100
    UserList(UserIndex).Invent.Object(3).Equipped = 0

    UserList(UserIndex).Invent.Object(4).ObjIndex = 38 ' Pocion Roja
    UserList(UserIndex).Invent.Object(4).amount = 100
    
    UserList(UserIndex).Invent.Object(5).ObjIndex = 573 ' Comida
    UserList(UserIndex).Invent.Object(5).amount = 100

    UserList(UserIndex).Invent.Object(6).ObjIndex = 572 ' Bebida
    UserList(UserIndex).Invent.Object(6).amount = 100
    
    UserList(UserIndex).Invent.Object(7).ObjIndex = 650 ' Runa
    UserList(UserIndex).Invent.Object(7).amount = 1
    
    UserList(UserIndex).Invent.Object(8).ObjIndex = 879 ' Mapa
    UserList(UserIndex).Invent.Object(8).amount = 1
    
'*********************************************************************************
    
    Case eClass.Cazador
    
    UserList(UserIndex).Invent.Object(1).ObjIndex = 31 ' vestimenta
    UserList(UserIndex).Invent.Object(1).amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1

    UserList(UserIndex).Invent.Object(2).ObjIndex = 1355 ' arco
    UserList(UserIndex).Invent.Object(2).amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1

    UserList(UserIndex).Invent.Object(3).ObjIndex = 1357 ' Flecha
    UserList(UserIndex).Invent.Object(3).amount = 100
    UserList(UserIndex).Invent.Object(3).Equipped = 0

    UserList(UserIndex).Invent.Object(4).ObjIndex = 38 ' Pocion Roja
    UserList(UserIndex).Invent.Object(4).amount = 100
    
    UserList(UserIndex).Invent.Object(5).ObjIndex = 573 ' Comida
    UserList(UserIndex).Invent.Object(5).amount = 100

    UserList(UserIndex).Invent.Object(6).ObjIndex = 572 ' Bebida
    UserList(UserIndex).Invent.Object(6).amount = 100
    
    UserList(UserIndex).Invent.Object(7).ObjIndex = 650 ' Runa
    UserList(UserIndex).Invent.Object(7).amount = 1
    
    UserList(UserIndex).Invent.Object(8).ObjIndex = 879 ' Mapa
    UserList(UserIndex).Invent.Object(8).amount = 1
    
'********************************************************************************
   
    Case eClass.Mago
    UserList(UserIndex).Invent.Object(1).ObjIndex = 31 ' vestimenta
    UserList(UserIndex).Invent.Object(1).amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1

    UserList(UserIndex).Invent.Object(2).ObjIndex = 1356 ' baculo
    UserList(UserIndex).Invent.Object(2).amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1
    
    UserList(UserIndex).Invent.Object(3).ObjIndex = 37 ' Pocion Azul
    UserList(UserIndex).Invent.Object(3).amount = 100

    UserList(UserIndex).Invent.Object(4).ObjIndex = 38 ' Pocion Roja
    UserList(UserIndex).Invent.Object(4).amount = 100
    
    UserList(UserIndex).Invent.Object(5).ObjIndex = 573 ' Comida
    UserList(UserIndex).Invent.Object(5).amount = 100

    UserList(UserIndex).Invent.Object(6).ObjIndex = 572 ' Bebida
    UserList(UserIndex).Invent.Object(6).amount = 100
    
    UserList(UserIndex).Invent.Object(7).ObjIndex = 650 ' Runa
    UserList(UserIndex).Invent.Object(7).amount = 1
    
    UserList(UserIndex).Invent.Object(8).ObjIndex = 879 ' Mapa
    UserList(UserIndex).Invent.Object(8).amount = 1
    
'********************************************************************************
    
    Case eClass.Nigromante
    UserList(UserIndex).Invent.Object(1).ObjIndex = 31 ' vestimenta
    UserList(UserIndex).Invent.Object(1).amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1

    UserList(UserIndex).Invent.Object(2).ObjIndex = 1356 ' baculo
    UserList(UserIndex).Invent.Object(2).amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1
    
    UserList(UserIndex).Invent.Object(3).ObjIndex = 37 ' Pocion Azul
    UserList(UserIndex).Invent.Object(3).amount = 100

    UserList(UserIndex).Invent.Object(4).ObjIndex = 38 ' Pocion Roja
    UserList(UserIndex).Invent.Object(4).amount = 100
    
    UserList(UserIndex).Invent.Object(5).ObjIndex = 573 ' Comida
    UserList(UserIndex).Invent.Object(5).amount = 100

    UserList(UserIndex).Invent.Object(6).ObjIndex = 572 ' Bebida
    UserList(UserIndex).Invent.Object(6).amount = 100
    
    UserList(UserIndex).Invent.Object(7).ObjIndex = 650 ' Runa
    UserList(UserIndex).Invent.Object(7).amount = 1
    
    UserList(UserIndex).Invent.Object(8).ObjIndex = 879 ' Mapa
    UserList(UserIndex).Invent.Object(8).amount = 1
    
'********************************************************************************
   
    Case eClass.Paladin
    
    UserList(UserIndex).Invent.Object(1).ObjIndex = 31 ' vestimenta
    UserList(UserIndex).Invent.Object(1).amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1

    UserList(UserIndex).Invent.Object(2).ObjIndex = 574 ' Martillo
    UserList(UserIndex).Invent.Object(2).amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1

    UserList(UserIndex).Invent.Object(3).ObjIndex = 38 ' Pocion Roja
    UserList(UserIndex).Invent.Object(3).amount = 100
    
    UserList(UserIndex).Invent.Object(4).ObjIndex = 573 ' Comida
    UserList(UserIndex).Invent.Object(4).amount = 100

    UserList(UserIndex).Invent.Object(5).ObjIndex = 572 ' Bebida
    UserList(UserIndex).Invent.Object(5).amount = 100
    
    UserList(UserIndex).Invent.Object(6).ObjIndex = 650 ' Runa
    UserList(UserIndex).Invent.Object(6).amount = 1
    
    UserList(UserIndex).Invent.Object(7).ObjIndex = 879 ' Mapa
    UserList(UserIndex).Invent.Object(7).amount = 1
    
'***********************************************************************************
    
    Case eClass.Guerrero
    
    UserList(UserIndex).Invent.Object(1).ObjIndex = 31 ' vestimenta
    UserList(UserIndex).Invent.Object(1).amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1

    UserList(UserIndex).Invent.Object(2).ObjIndex = 574 ' Martillo
    UserList(UserIndex).Invent.Object(2).amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1

    UserList(UserIndex).Invent.Object(3).ObjIndex = 38 ' Pocion Roja
    UserList(UserIndex).Invent.Object(3).amount = 100
    
    UserList(UserIndex).Invent.Object(4).ObjIndex = 573 ' Comida
    UserList(UserIndex).Invent.Object(4).amount = 100

    UserList(UserIndex).Invent.Object(5).ObjIndex = 572 ' Bebida
    UserList(UserIndex).Invent.Object(5).amount = 100
    
    UserList(UserIndex).Invent.Object(6).ObjIndex = 650 ' Runa
    UserList(UserIndex).Invent.Object(6).amount = 1
    
    UserList(UserIndex).Invent.Object(7).ObjIndex = 879 ' Mapa
    UserList(UserIndex).Invent.Object(7).amount = 1
    
'**********************************************************************************
    
    Case eClass.Ladron
    UserList(UserIndex).Invent.Object(1).ObjIndex = 31 ' vestimenta
    UserList(UserIndex).Invent.Object(1).amount = 1
    UserList(UserIndex).Invent.Object(1).Equipped = 1

    UserList(UserIndex).Invent.Object(2).ObjIndex = 574 ' daga
    UserList(UserIndex).Invent.Object(2).amount = 1
    UserList(UserIndex).Invent.Object(2).Equipped = 1

    UserList(UserIndex).Invent.Object(3).ObjIndex = 576 ' daga arrojadisa
    UserList(UserIndex).Invent.Object(3).amount = 100
    UserList(UserIndex).Invent.Object(3).Equipped = 0

    UserList(UserIndex).Invent.Object(4).ObjIndex = 38 ' Pocion Roja
    UserList(UserIndex).Invent.Object(4).amount = 100
    
    UserList(UserIndex).Invent.Object(5).ObjIndex = 573 ' Comida
    UserList(UserIndex).Invent.Object(5).amount = 100

    UserList(UserIndex).Invent.Object(6).ObjIndex = 572 ' Bebida
    UserList(UserIndex).Invent.Object(6).amount = 100
    
    UserList(UserIndex).Invent.Object(7).ObjIndex = 650 ' Runa
    UserList(UserIndex).Invent.Object(7).amount = 1
    
    UserList(UserIndex).Invent.Object(8).ObjIndex = 879 ' Mapa
    UserList(UserIndex).Invent.Object(8).amount = 1
    
'***********************************************************************************
    
  For i = 1 To MAXAMIGOS
  UserList(UserIndex).Amigos(i).Nombre = "Nadies"
  UserList(UserIndex).Amigos(i).Ignorado = 0
  UserList(UserIndex).Amigos(i).Index = 0
  Next i
    
End Select

'Vestimentas de raza
Select Case UserRaza
    Case eRaza.Humano
    
        UserList(UserIndex).Invent.Object(1).ObjIndex = 463
        UserList(UserIndex).Invent.Object(1).amount = 1
        UserList(UserIndex).Invent.Object(1).Equipped = 1
        
    Case eRaza.Elfo
    
        UserList(UserIndex).Invent.Object(1).ObjIndex = 464
        UserList(UserIndex).Invent.Object(1).amount = 1
        UserList(UserIndex).Invent.Object(1).Equipped = 1
        
    Case eRaza.Drow
    
        UserList(UserIndex).Invent.Object(1).ObjIndex = 465
        UserList(UserIndex).Invent.Object(1).amount = 1
        UserList(UserIndex).Invent.Object(1).Equipped = 1
        
    Case eRaza.Enano
    
        UserList(UserIndex).Invent.Object(1).ObjIndex = 240
        UserList(UserIndex).Invent.Object(1).amount = 1
        UserList(UserIndex).Invent.Object(1).Equipped = 1
        
    Case eRaza.Orco
    
        UserList(UserIndex).Invent.Object(1).ObjIndex = 465
        UserList(UserIndex).Invent.Object(1).amount = 1
        UserList(UserIndex).Invent.Object(1).Equipped = 1
        
    Case eRaza.Gnomo
    
        UserList(UserIndex).Invent.Object(1).ObjIndex = 240
        UserList(UserIndex).Invent.Object(1).amount = 1
        UserList(UserIndex).Invent.Object(1).Equipped = 1
        
End Select
 
#If ConUpTime Then
    UserList(UserIndex).LogOnTime = Now
    UserList(UserIndex).UpTime = 0
#End If

'Valores Default de facciones al Activar nuevo usuario
Call ResetFacciones(UserIndex)

'Valores default de correos al activar nuevo usuario
'Call ResetCorreos(UserIndex)

'Empezar en Dungeon Newbie


Call SaveUser(UserIndex, CharPath & UCase$(name) & ".pjs", Account, name)
  
'Open User
Call ConnectUser(UserIndex, name, Account)
  
End Sub

#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

Sub CloseSocket(ByVal UserIndex As Integer)

On Error GoTo errhandler
    
    Call aDos.RestarConexion(UserList(UserIndex).ip)
    
    If UserIndex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
    'Call SecurityIp.IpRestarConexion(GetLongIp(UserList(UserIndex).ip))
    
    If UserList(UserIndex).ConnID <> -1 Then
        Call CloseSocketSL(UserIndex)
    End If
    
    'Es el mismo user al que est� revisando el centinela??
    'IMPORTANTE!!! hacerlo antes de resetear as� todav�a sabemos el nombre del user
    ' y lo podemos loguear
    If Centinela.RevisandoUserIndex = UserIndex Then _
        Call modCentinela.CentinelaUserLogout
    
    'mato los comercios seguros
    If UserList(UserIndex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(UserList(UserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                Call FlushBuffer(UserList(UserIndex).ComUsu.DestUsu)
            End If
        End If
    End If
    
    'Empty buffer for reuse
    Call UserList(UserIndex).incomingData.ReadASCIIStringFixed(UserList(UserIndex).incomingData.length)
    
    If UserList(UserIndex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        Call CloseUser(UserIndex)
        
        Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Else
        Call ResetUserSlot(UserIndex)
    End If
    
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    
Exit Sub

errhandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    Call ResetUserSlot(UserIndex)

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripci�n = " & Err.description & " - UserIndex = " & UserIndex)
End Sub

#ElseIf UsarQueSocket = 0 Then

Sub CloseSocket(ByVal UserIndex As Integer)
On Error GoTo errhandler
    
    
    
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0

    If UserIndex = LastUser And LastUser > 1 Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop
    End If

    If UserList(UserIndex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            Call CloseUser(UserIndex)
    End If

    frmMain.Socket2(UserIndex).Cleanup
    Unload frmMain.Socket2(UserIndex)
    Call ResetUserSlot(UserIndex)

Exit Sub

errhandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    Call ResetUserSlot(UserIndex)
End Sub







#ElseIf UsarQueSocket = 3 Then

Sub CloseSocket(ByVal UserIndex As Integer, Optional ByVal cerrarlo As Boolean = True)

On Error GoTo errhandler

Dim NURestados As Boolean
Dim CoNnEcTiOnId As Long


    NURestados = False
    CoNnEcTiOnId = UserList(UserIndex).ConnID
    
    'call logindex(UserIndex, "******> Sub CloseSocket. ConnId: " & CoNnEcTiOnId & " Cerrarlo: " & Cerrarlo)
    
    
  
    UserList(UserIndex).ConnID = -1 'inabilitamos operaciones en socket
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0

    If UserIndex = LastUser And LastUser > 1 Then
        Do
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop While UserList(LastUser).flags.UserLogged = True
    End If

    If UserList(UserIndex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            NURestados = True
            Call CloseUser(UserIndex)
    End If
    
    Call ResetUserSlot(UserIndex)
    
    'limpiada la userlist... reseteo el socket, si me lo piden
    'Me lo piden desde: cerrada intecional del servidor (casi todas
    'las llamadas a CloseSocket del codigo)
    'No me lo piden desde: disconnect remoto (el on_close del control
    'de alejo realiza la desconexion automaticamente). Esto puede pasar
    'por ejemplo, si el cliente cierra el AO.
    If cerrarlo Then Call frmMain.TCPServ.CerrarSocket(CoNnEcTiOnId)

Exit Sub

errhandler:
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    Call LogError("CLOSESOCKETERR: " & Err.description & " UI:" & UserIndex)
    
    If Not NURestados Then
        If UserList(UserIndex).flags.UserLogged Then
            If NumUsers > 0 Then
                NumUsers = NumUsers - 1
            End If
            Call LogError("Cerre sin grabar a: " & UserList(UserIndex).name)
        End If
    End If
    
    Call LogError("El usuario no guardado tenia connid " & CoNnEcTiOnId & ". Socket no liberado.")
    Call ResetUserSlot(UserIndex)

End Sub


#End If

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)

#If UsarQueSocket = 1 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    Call BorraSlotSock(UserList(UserIndex).ConnID)
    Call WSApiCloseSocket(UserList(UserIndex).ConnID)
    UserList(UserIndex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 0 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    frmMain.Socket2(UserIndex).Cleanup
    Unload frmMain.Socket2(UserIndex)
    UserList(UserIndex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 2 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    Call frmMain.Serv.CerrarSocket(UserList(UserIndex).ConnID)
    UserList(UserIndex).ConnIDValida = False
End If

#End If
End Sub

''
' Send an string to a Slot
'
' @param userIndex The index of the User
' @param Datos The string that will be send
' @remarks If UsarQueSocket is 3 it won`t use the clsByteQueue

Public Function EnviarDatosASlot(ByVal UserIndex As Integer, ByRef Datos As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: 01/10/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Now it uses the clsByteQueue class and don`t make a FIFO Queue of String
'***************************************************

#If UsarQueSocket = 1 Then '**********************************************
    On Error GoTo Err
    
    Dim Ret As Long
    
    Ret = WsApiEnviar(UserIndex, Datos)
    
    If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
        ' Close the socket avoiding any critical error
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    End If
Exit Function
    
Err:
        'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("EnviarDatosASlot:: ERR Handler. userindex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida & " ERR: " & Err.Description)

#ElseIf UsarQueSocket = 0 Then '**********************************************
    
    If frmMain.Socket2(UserIndex).Write(Datos, Len(Datos)) < 0 Then
        If frmMain.Socket2(UserIndex).LastError = WSAEWOULDBLOCK Then
            ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(Datos)
        Else
            'Close the socket avoiding any critical error
            Call Cerrar_Usuario(UserIndex)
        End If
    End If
#ElseIf UsarQueSocket = 2 Then '**********************************************

    'Return value for this Socket:
    '--0) OK
    '--1) WSAEWOULDBLOCK
    '--2) ERROR
    
    Dim Ret As Long

    Ret = frmMain.Serv.Enviar(.ConnID, Datos, Len(Datos))
            
    If Ret = 1 Then
        ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
        Call .outgoingData.WriteASCIIStringFixed(Datos)
    ElseIf Ret = 2 Then
        'Close socket avoiding any critical error
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    End If
    

#ElseIf UsarQueSocket = 3 Then
    'THIS SOCKET DOESN`T USE THE BYTE QUEUE CLASS
    Dim rv As Long
    'al carajo, esto encola solo!!! che, me aprobar� los
    'parciales tambi�n?, este control hace todo solo!!!!
    On Error GoTo ErrorHandler
        
        If UserList(UserIndex).ConnID = -1 Then
            Call LogError("TCP::EnviardatosASlot, se intento enviar datos a un userIndex con ConnId=-1")
            Exit Function
        End If
        
        If frmMain.TCPServ.Enviar(UserList(UserIndex).ConnID, Datos, Len(Datos)) = 2 Then Call CloseSocket(UserIndex)

Exit Function
ErrorHandler:
    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & UserIndex & "/" & UserList(UserIndex).ConnID & "/" & Datos)
#End If '**********************************************

End Function
Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
        For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

            If MapData(UserList(Index).Pos.Map, X, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next X
Next Y
EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.Map, X, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.Map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next X
Next Y
HayOBJarea = False
End Function
Function ValidateChr(ByVal UserIndex As Integer) As Boolean

ValidateChr = UserList(UserIndex).Char.Head <> 0 _
                And UserList(UserIndex).Char.body <> 0 _
                And ValidateSkills(UserIndex)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, ByRef name As String, ByRef Account As String)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa - Agrego por default que el color de dialogo de los dioses, sea como el de su nick.
'***************************************************
Dim N As Integer
Dim tStr As String

If UserList(UserIndex).flags.UserLogged Then
    Call LogCheating("El usuario " & UserList(UserIndex).name & " ha intentado loguear a " & name & " desde la IP " & UserList(UserIndex).ip)
    'Kick player ( and leave character inside :D )!
    Call CloseSocketSL(UserIndex)
    Call Cerrar_Usuario(UserIndex)
    Exit Sub
End If

'Reseteamos los FLAGS
UserList(UserIndex).flags.Escondido = 0
UserList(UserIndex).flags.TargetNPC = 0
UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
UserList(UserIndex).flags.TargetObj = 0
UserList(UserIndex).flags.TargetUser = 0
UserList(UserIndex).Char.FX = 0

'Controlamos no pasar el maximo de usuarios
Call WriteConsoleMsg(UserIndex, "Servidor> �Recuerda visitar nuestro sitio web oficial www.imperiumao.com.ar ", FontTypeNames.FONTTYPE_GUILD)
If NumUsers >= MaxUsers Then
    Call WriteErrorMsg(UserIndex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
    Call FlushBuffer(UserIndex)
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'�Este IP ya esta conectado?
If AllowMultiLogins = 0 Then
    If CheckForSameIP(UserIndex, UserList(UserIndex).ip) = True Then
        Call WriteErrorMsg(UserIndex, "No es posible usar mas de un personaje al mismo tiempo.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
End If

'�Existe el personaje?
If Not FileExist(CharPath & UCase$(name) & ".pjs", vbNormal) Then
    Call WriteErrorMsg(UserIndex, "El personaje no existe.")
    Call FlushBuffer(UserIndex)
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'�Es el passwd valido?
If Not IsPJOfAccount(name, Account) Then
    Call WriteErrorMsg(UserIndex, "El personaje que con el que intenta conectarse no es de su cuenta.")
    Call FlushBuffer(UserIndex)
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'�Ya esta conectado el personaje?
If CheckForSameName(name) Then
    If UserList(NameIndex(name)).Counters.Saliendo Then
        Call WriteErrorMsg(UserIndex, "El usuario est� saliendo.")
    Else
        Call WriteErrorMsg(UserIndex, "Perdon, un usuario con el mismo nombre se ha logoeado.")
    End If
    Call FlushBuffer(UserIndex)
    'Call CloseSocket(UserIndex)
    Exit Sub
End If

'Reseteamos los privilegios
UserList(UserIndex).flags.Privilegios = 0

'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
If EsAdmin(name) Then
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.Admin
    Call LogGM(name, "Se conecto con ip:" & UserList(UserIndex).ip)
ElseIf EsDios(name) Then
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.Dios
    Call LogGM(name, "Se conecto con ip:" & UserList(UserIndex).ip)
ElseIf EsSemiDios(name) Then
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.SemiDios
    Call LogGM(name, "Se conecto con ip:" & UserList(UserIndex).ip)
ElseIf EsConsejero(name) Then
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.Consejero
    Call LogGM(name, "Se conecto con ip:" & UserList(UserIndex).ip)
Else
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.User
    UserList(UserIndex).flags.AdminPerseguible = True
End If

'Add RM flag if needed
If EsRolesMaster(name) Then
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.RoleMaster
End If

If ServerSoloGMs > 0 Then
    If (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
        Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
End If

'Cargamos el personaje
Dim Leer As New clsIniReader

'//Leemos los correos
'Call LeerCorreos(UserIndex, Leer)

Call Leer.Initialize(CharPath & UCase$(name) & ".pjs")

'Cargamos los datos del personaje
Call LoadUserInit(UserIndex, Leer)

Call LoadUserStats(UserIndex, Leer)

If Not ValidateChr(UserIndex) Then
    Call WriteErrorMsg(UserIndex, "Error en el personaje.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If


Set Leer = Nothing

If UserList(UserIndex).Invent.EscudoEqpSlot = 0 Then UserList(UserIndex).Char.ShieldAnim = NingunEscudo
If UserList(UserIndex).Invent.CascoEqpSlot = 0 Then UserList(UserIndex).Char.CascoAnim = NingunCasco
If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then UserList(UserIndex).Char.WeaponAnim = NingunArma



Call UpdateUserInv(True, UserIndex, 0)
Call UpdateUserHechizos(True, UserIndex, 0)

Call ActualizarSlotAmigo(UserIndex, 0, True)
Call ObtenerIndexAmigos(UserIndex, False)

If UserList(UserIndex).flags.Paralizado Then
    Call WriteParalizeOK(UserIndex)
End If

''
'TODO : Feo, esto tiene que ser parche cliente
If UserList(UserIndex).flags.Estupidez = 0 Then
    Call WriteDumbNoMore(UserIndex)
End If

'Posicion de comienzo
If UserList(UserIndex).Pos.Map = 0 Then
    Select Case UserList(UserIndex).Hogar
        Case eCiudad.cBanderbill
            UserList(UserIndex).Pos = Banderbill
        Case eCiudad.cRinkel
            UserList(UserIndex).Pos = Rinkel
        Case eCiudad.cSuramei
            UserList(UserIndex).Pos = Suramei
    End Select
Else
    If Not MapaValido(UserList(UserIndex).Pos.Map) Then
        Call WriteErrorMsg(UserIndex, "EL PJ se encuenta en un mapa invalido.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
End If

'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Mart�n Sotuyo Dodero (Maraxus)
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex <> 0 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).NpcIndex <> 0 Then
    Dim FoundPlace As Boolean
    Dim esAgua As Boolean
    Dim tX As Long
    Dim tY As Long
    
    FoundPlace = False
    esAgua = HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    
    For tY = UserList(UserIndex).Pos.Y - 1 To UserList(UserIndex).Pos.Y + 1
        For tX = UserList(UserIndex).Pos.X - 1 To UserList(UserIndex).Pos.X + 1
            If esAgua Then
                'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                If LegalPos(UserList(UserIndex).Pos.Map, tX, tY, True, False) Then
                    FoundPlace = True
                    Exit For
                End If
            Else
                'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                If LegalPos(UserList(UserIndex).Pos.Map, tX, tY, False, True) Then
                    FoundPlace = True
                    Exit For
                End If
            End If
        Next tX
        
        If FoundPlace Then _
            Exit For
    Next tY
    
    If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
        UserList(UserIndex).Pos.X = tX
        UserList(UserIndex).Pos.Y = tY
    Else
        'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
        If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex <> 0 Then
            'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
            If UserList(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex).ComUsu.DestUsu > 0 Then
                'Le avisamos al que estaba comerciando que se tuvo que ir.
                If UserList(UserList(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    Call FinComerciarUsu(UserList(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex).ComUsu.DestUsu)
                    Call WriteConsoleMsg(UserList(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                    Call FlushBuffer(UserList(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex).ComUsu.DestUsu)
                End If
                'Lo sacamos.
                If UserList(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex).flags.UserLogged Then
                    Call FinComerciarUsu(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex)
                    Call WriteErrorMsg(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor recon�ctate...")
                    Call FlushBuffer(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex)
                End If
            End If
            
            Call CloseSocket(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex)
        End If
    End If
End If

If UserList(UserIndex).flags.Muerto = 1 Then
    Call Empollando(UserIndex)
End If

'Nombre de sistema
UserList(UserIndex).name = name
UserList(UserIndex).Account = Account

UserList(UserIndex).showName = True 'Por default los nombres son visibles

'If in the water, and has a boat, equip it!
If UserList(UserIndex).Invent.BarcoObjIndex > 0 And _
        (HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Or BodyIsBoat(UserList(UserIndex).Char.body)) Then
    Dim Barco As ObjData
    Barco = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
    UserList(UserIndex).Char.Head = 0
    If UserList(UserIndex).flags.Muerto = 0 Then
        If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.body = iBarcaCiuda
        If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGaleraCiuda
        If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleonCiuda
    Else
        UserList(UserIndex).Char.body = iFragataFantasmal
    End If
    
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
    UserList(UserIndex).flags.Navegando = 1
End If


'Info
Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
Call WriteChangeMap(UserIndex, UserList(UserIndex).Pos.Map, MapInfo(UserList(UserIndex).Pos.Map).MapVersion) 'Carga el mapa

If UserList(UserIndex).flags.Privilegios = PlayerType.Dios Then
    UserList(UserIndex).flags.ChatColor = RGB(0, 135, 0)
ElseIf UserList(UserIndex).flags.Privilegios <> PlayerType.User And UserList(UserIndex).flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And UserList(UserIndex).flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
    UserList(UserIndex).flags.ChatColor = RGB(0, 135, 0)
ElseIf UserList(UserIndex).flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
    UserList(UserIndex).flags.ChatColor = RGB(0, 255, 255)
ElseIf UserList(UserIndex).flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
    UserList(UserIndex).flags.ChatColor = RGB(255, 128, 64)
Else
    UserList(UserIndex).flags.ChatColor = vbWhite
End If
 

'Crea  el personaje del usuario
Call MakeUserChar(True, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

Call WriteUserCharIndexInServer(UserIndex)
''[/el oso]

Call WriteUpdateUserStats(UserIndex)

Call WriteUpdateHungerAndThirst(UserIndex)


If haciendoBK Then
    Call WritePauseToggle(UserIndex)
    Call WriteConsoleMsg(UserIndex, "Servidor> Por favor espera algunos segundos, WorldSave esta ejecutandose.", FontTypeNames.FONTTYPE_SERVER)
End If

If EnPausa Then
    Call WritePauseToggle(UserIndex)
    Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar m�s tarde.", FontTypeNames.FONTTYPE_SERVER)
End If

If EnTesting And UserList(UserIndex).Stats.ELV >= 18 Then
    Call WriteErrorMsg(UserIndex, "Servidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
    Call FlushBuffer(UserIndex)
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'Actualiza el Num de usuarios
'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
NumUsers = NumUsers + 1
UserList(UserIndex).flags.UserLogged = True

'usado para borrar Pjs
Call WriteVar(CharPath & UserList(UserIndex).name & ".pjs", "INIT", "Logged", "1")

Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)

MapInfo(UserList(UserIndex).Pos.Map).NumUsers = MapInfo(UserList(UserIndex).Pos.Map).NumUsers + 1

If UserList(UserIndex).Stats.SkillPts > 0 Then
    Call WriteSendSkills(UserIndex)
End If

If NumUsers > recordusuarios Then
    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultaneamente." & "Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
    recordusuarios = NumUsers
    Call WriteVar(IniPath & "IMPAOSV.ini", "INIT", "Record", str(recordusuarios))
    
    Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
End If

If UserList(UserIndex).NroMascotas > 0 And MapInfo(UserList(UserIndex).Pos.Map).Pk Then
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(i) > 0 Then
            UserList(UserIndex).MascotasIndex(i) = SpawnNpc(UserList(UserIndex).MascotasType(i), UserList(UserIndex).Pos, True, True)
            
            If UserList(UserIndex).MascotasIndex(i) > 0 Then
                Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
                Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
            Else
                UserList(UserIndex).MascotasIndex(i) = 0
            End If
        End If
    Next i
End If

If UserList(UserIndex).flags.Navegando = 1 Then
    Call WriteNavigateToggle(UserIndex)
End If

If UserList(UserIndex).flags.Montando = 1 Then
    Call WriteEquitateToggle(UserIndex)
End If


If esRene(UserIndex) Then
    Call WriteSafeModeOff(UserIndex)
    UserList(UserIndex).flags.Seguro = False
Else
    UserList(UserIndex).flags.Seguro = True
    Call WriteSafeModeOn(UserIndex)
End If

If UserList(UserIndex).GuildIndex > 0 Then
    'welcome to the show baby...
    If Not modGuilds.m_ConectarMiembroAClan(UserIndex, UserList(UserIndex).GuildIndex) Then
        Call WriteConsoleMsg(UserIndex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
    End If
End If

Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXIDs.FXWARP, 0))
Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))

Call WriteLoggedMessage(UserIndex)

Call modGuilds.SendGuildNews(UserIndex)

If Lloviendo Then
    Call WriteRainToggle(UserIndex)
End If

tStr = modGuilds.a_ObtenerRechazoDeChar(UserList(UserIndex).name)

If LenB(tStr) <> 0 Then
    Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
End If

'Load the user statistics
Call Statistics.UserConnected(UserIndex)

Call WriteFuerza(UserIndex)
Call WriteAgilidad(UserIndex)

Call MostrarNumUsers


#If SeguridadAlkon Then
    Call Security.UserConnected(UserIndex)
#End If

N = FreeFile
Open App.Path & "\Data\Files logs\numusers.log" For Output As N
Print #N, NumUsers
Close #N

N = FreeFile
'Log
Open App.Path & "\Data\Files logs\Connect.log" For Append Shared As #N
Print #N, UserList(UserIndex).name & " ha entrado al juego. UserIndex:" & UserIndex & " " & time & " " & Date
Close #N
End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************
    With UserList(UserIndex).Faccion
        .ArmadaReal = 0
        .FuerzasCaos = 0
        .Milicia = 0
        .Rango = 0
        
        .CiudadanosMatados = 0
        .RenegadosMatados = 0
        .RepublicanosMatados = 0
        
    End With
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'05/20/2007 Integer - Agregue todas las variables que faltaban.
'*************************************************
    With UserList(UserIndex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pasos = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0
        .bPuedeMeditar = False
        .Lava = 0
        .Mimetismo = 0
        .Saliendo = False
        .Salir = 0
        .TiempoOculto = 0
        .TimerMagiaGolpe = 0
        .TimerGolpeMagia = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeUsarArco = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
    End With
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex)
        .name = vbNullString
        .modName = vbNullString
        .desc = vbNullString
        .DescRM = vbNullString
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .ip = vbNullString
        .Clase = 0
        .email = vbNullString
        .genero = 0
        .Hogar = 0
        .raza = 0
        
        .EmpoCont = 0
        .PartyIndex = 0
        .PartySolicitud = 0
        
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            '.RenegadosMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0
            .GLD = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0
        End With
        
    End With
End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Reputacion
        .AsesinoRep = 0
        .BandidoRep = 0
        .BurguesRep = 0
        .LadronesRep = 0
        .NobleRep = 0
        .PlebeRep = 0
        .NobleRep = 0
        .Promedio = 0
    End With
End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
    If UserList(UserIndex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
        UserList(UserIndex).EscucheClan = 0
    End If
    If UserList(UserIndex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)
    End If
    UserList(UserIndex).GuildIndex = 0
End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 06/28/2008
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'03/29/2006 Maraxus - Reseteo el CentinelaOK tambi�n.
'06/28/2008 NicoNZ - Agrego el flag Inmovilizado
'*************************************************
    With UserList(UserIndex).flags
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = vbNullString
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .ModoCombate = False
        .Vuela = 0
        .Navegando = 0
        .Montando = 0
        .Oculto = 0
        .Envenenado = 0
        .invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = 0
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .EstaEmpo = 0
        .Silenciado = 0
        .CentinelaOK = False
        .AdminPerseguible = False
    End With
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
    Dim LoopC As Long
    
    UserList(UserIndex).NroMascotas = 0
        
    For LoopC = 1 To MAXMASCOTAS
        UserList(UserIndex).MascotasIndex(LoopC) = 0
        UserList(UserIndex).MascotasType(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(UserIndex).BancoInvent.Object(LoopC).amount = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(UserIndex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
    With UserList(UserIndex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(UserIndex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)

UserList(UserIndex).ConnIDValida = False
UserList(UserIndex).ConnID = -1

Call LimpiarComercioSeguro(UserIndex)
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
Call ResetUserExtras(UserIndex)
With UserList(UserIndex).ComUsu
    .Acepto = False
    .cant = 0
    .DestNick = vbNullString
    .DestUsu = 0
    .Objeto = 0
End With

End Sub

Sub CloseUser(ByVal UserIndex As Integer)
'Call LogTarea("CloseUser " & UserIndex)
On Error GoTo errhandler

Dim N As Integer
Dim LoopC As Integer
Dim Map As Integer
Dim name As String
Dim i As Integer

Dim aN As Integer

aN = UserList(UserIndex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = vbNullString
End If
aN = UserList(UserIndex).flags.NPCAtacado
If aN > 0 Then
    If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).name Then
        Npclist(aN).flags.AttackedFirstBy = vbNullString
    End If
End If
UserList(UserIndex).flags.AtacadoPorNpc = 0
UserList(UserIndex).flags.NPCAtacado = 0

Map = UserList(UserIndex).Pos.Map
name = UCase$(UserList(UserIndex).name)

UserList(UserIndex).Char.FX = 0
UserList(UserIndex).Char.loops = 0
Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))


UserList(UserIndex).flags.UserLogged = False
UserList(UserIndex).Counters.Saliendo = False

'Le devolvemos el body y head originales
If UserList(UserIndex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)

'si esta en party le devolvemos la experiencia
If UserList(UserIndex).PartyIndex > 0 Then Call mdParty.SalirDeParty(UserIndex)

'Save statistics
Call Statistics.UserDisconnected(UserIndex)

'Actualizamos los index de los amigos
Call ObtenerIndexAmigos(UserIndex, True)

' Grabamos el personaje del usuario
Call SaveUser(UserIndex, CharPath & name & ".pjs")


'usado para borrar Pjs
Call WriteVar(CharPath & UserList(UserIndex).name & ".pjs", "INIT", "Logged", "0")


'Quitar el dialogo
'If MapInfo(Map).NumUsers > 0 Then
'    Call SendToUserArea(UserIndex, "QDL" & UserList(UserIndex).Char.charindex)
'End If

If MapInfo(Map).NumUsers > 0 Then
    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
End If



'Borrar el personaje
If UserList(UserIndex).Char.CharIndex > 0 Then
    Call EraseUserChar(UserIndex)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).flags.NPCActive Then _
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

'Update Map Users
MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1

If MapInfo(Map).NumUsers < 0 Then
    MapInfo(Map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(UserIndex).name) Then Call Ayuda.Quitar(UserList(UserIndex).name)

Call ResetUserSlot(UserIndex)

Call MostrarNumUsers

N = FreeFile(1)
Open App.Path & "\Data\Files logs\Connect.log" For Append Shared As #N
Print #N, name & " h� dejado el juego. " & "User Index:" & UserIndex & " " & time & " " & Date
Close #N

Exit Sub

errhandler:
Call LogError("Error en CloseUser. N�mero " & Err.Number & " Descripci�n: " & Err.description)

End Sub

Sub ReloadSokcet()
On Error GoTo errhandler
#If UsarQueSocket = 1 Then

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
'       Call apiclosesocket(SockListen)
'       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    End If

#ElseIf UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then

    

#End If

Exit Sub
errhandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.description)

End Sub

Public Sub EnviarNoche(ByVal UserIndex As Integer)
    Call WriteSendNight(UserIndex, IIf(DeNoche And (MapInfo(UserList(UserIndex).Pos.Map).Zona = Campo Or MapInfo(UserList(UserIndex).Pos.Map).Zona = Ciudad), True, False))
    Call WriteSendNight(UserIndex, IIf(DeNoche, True, False))
End Sub

Public Sub EcharPjsNoPrivilegiados()
Dim LoopC As Long

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
        If UserList(LoopC).flags.Privilegios And PlayerType.User Then
            Call CloseSocket(LoopC)
        End If
    End If
Next LoopC

End Sub

Public Sub ResetUserExtras(ByVal UserIndex As Integer)
'***************************************************
'              Sistema de Friends
'***************************************************
  Dim i As Integer
  For i = 1 To MAXAMIGOS
  UserList(UserIndex).Amigos(i).Nombre = vbNullString
  UserList(UserIndex).Amigos(i).Ignorado = 0
UserList(UserIndex).Amigos(i).Index = 0
  Next i
UserList(UserIndex).Quien = vbNullString
End Sub
