Attribute VB_Name = "mod_accounts"
' programado por maTih.-
 
Option Explicit
 
Public Type Char_Acc_Data
       Nick_Name        As String
       Character        As Char
       Nivel            As Byte
       Pos_Map          As String
       Muerto           As Boolean
End Type
 
'Cuenta multi-logeable , 1 activado 0 desactivado.
Private Const MULTI_LOG As Byte = 1
 
'Número de pjs por cuenta.
Private Const MAX_PJS   As Byte = 8
Private Acc_Loggeds()   As String
 
Public Sub Inicializar()
 
'
' @ Inicializa el array de cuentas loggeds.
 
ReDim Acc_Loggeds(1 To 1) As String
 
End Sub
 
Public Sub Quitar_Lista(ByVal acc_Name As String)
 
'
' @ Quita una cuenta de la lista de cuentas logeadas.
 
Dim acc_Slot    As Integer
Dim acc_Loop    As Long
 
For acc_Loop = 1 To UBound(Acc_Loggeds())
    If UCase$(Acc_Loggeds(acc_Loop)) = UCase$(acc_Name) Then Exit For
Next acc_Loop
 
If (acc_Loop > UBound(Acc_Loggeds())) Then Exit Sub
 
Acc_Loggeds(acc_Loop) = vbNullString
 
End Sub
 
Public Sub Crear_Nueva(ByVal UserIndex As Integer, ByVal acc_Name As String, _
    ByVal acc_Password As String, ByVal acc_Pregunta As String, _
          ByVal acc_Respuesta As String, ByVal acc_EMail As String)
         
'
' @ Crea nueva cuenta.
 
Dim tmp_Error   As String
Dim loopX       As Long
 
'Pasa la validación.
If Validar_Creacion(acc_Name, acc_Password, acc_Pregunta, acc_Respuesta, acc_EMail, tmp_Error) Then
   
    'Guarda el password & email.
    WriteVar Acc_Path & acc_Name & ".sdc", "INIT", "EMail", acc_EMail
    WriteVar Acc_Path & acc_Name & ".sdc", "INIT", "Password", acc_Password
   
    'Guarda los datos de pregunta y respuesta.
    WriteVar Acc_Path & acc_Name & ".sdc", "CONTACTO", "Pregunta", acc_Pregunta
    WriteVar Acc_Path & acc_Name & ".sdc", "CONTACTO", "Respuesta", acc_Respuesta
   
    'Llena la cuenta con pjs nulos.
    For loopX = 1 To MAX_PJS
        WriteVar Acc_Path & acc_Name & ".sdc", "PERSONAJES", "PERSONAJE" & CStr(loopX), "NoUsado"
    Next loopX
   
    'Envia el formulario de cuenta vacía.
    Call Protocol.WriteAccountShow(UserIndex)
Else
    Call Protocol.WriteErrorMsg(UserIndex, tmp_Error)
    Call Protocol.FlushBuffer(UserIndex)
    Call TCP.CloseSocket(UserIndex)
End If
 
End Sub
 
Public Sub ConectarNuevoPersonaje(ByVal UserIndex As Integer, ByVal p_name As String, _
    ByVal p_Genero As eGenero, ByVal p_Raza As eRaza, ByVal p_Clase As eClass, _
          ByVal p_Head As Integer, ByVal p_Home As Byte)
         
'
' @ Conecta un nuevo personaje.
 
With UserList(UserIndex)
 
    If Not AsciiValidos(p_name) Or LenB(p_name) = 0 Then
        Call WriteErrorMsg(UserIndex, "Nombre inválido.")
        Exit Sub
    End If
 
    '¿Existe el personaje?
    If FileExist(CharPath & UCase$(p_name) & ".pjs", vbNormal) = True Then
        Call WriteErrorMsg(UserIndex, "Ya existe el personaje.")
        Exit Sub
    End If
   
    'Tiró los dados antes de llegar acá??
    If .Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
        Call WriteErrorMsg(UserIndex, "Debe tirar los dados antes de poder crear un personaje.")
        Exit Sub
    End If
   
    'Para que no chiten la cabeza : p
    If Not ValidarCabeza(p_Raza, p_Genero, p_Head) Then
        Call WriteErrorMsg(UserIndex, "Cabeza inválida, elija una cabeza seleccionable.")
        Exit Sub
    End If
   
    .flags.Muerto = 0
    .flags.Escondido = 0
   
    .Reputacion.AsesinoRep = 0
    .Reputacion.BandidoRep = 0
    .Reputacion.BurguesRep = 0
    .Reputacion.LadronesRep = 0
    .Reputacion.NobleRep = 1000
    .Reputacion.PlebeRep = 30
   
    .Reputacion.Promedio = 30 / 6
   
    .name = p_name
    .Clase = p_Clase
    .raza = p_Raza
    .genero = p_Genero
    .Hogar = p_Home
   
    '[Pablo (Toxic Waste) 9/01/08]
    .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + ModRaza(p_Raza).Fuerza
    .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + ModRaza(p_Raza).Agilidad
    .Stats.UserAtributos(eAtributos.Inteligencia) = .Stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(p_Raza).Inteligencia
    .Stats.UserAtributos(eAtributos.Carisma) = .Stats.UserAtributos(eAtributos.Carisma) + ModRaza(p_Raza).Carisma
    .Stats.UserAtributos(eAtributos.Constitucion) = .Stats.UserAtributos(eAtributos.Constitucion) + ModRaza(p_Raza).Constitucion
    '[/Pablo (Toxic Waste)]
   
    Dim i   As Long
   
    For i = 1 To NUMSKILLS
        .Stats.UserSkills(i) = 0
        Call CheckEluSkill(UserIndex, i, True)
    Next i
   
    .Stats.SkillPts = 10
   
    .Char.heading = eHeading.SOUTH
   
    Call DarCuerpo(UserIndex)
   
    .Char.Head = p_Head
   
    .OrigChar = .Char
   
    Dim MiInt As Long
    MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Constitucion) \ 3)
   
    .Stats.MaxHP = 15 + MiInt
    .Stats.MinHP = 15 + MiInt
   
    MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)
    If MiInt = 1 Then MiInt = 2
   
    .Stats.MaxSta = 20 * MiInt
    .Stats.MinSta = 20 * MiInt
   
   
    .Stats.MaxAGU = 100
    .Stats.MinAGU = 100
   
    .Stats.MaxHam = 100
    .Stats.MinHam = 100
   
   
    '<-----------------MANA----------------------->
    If p_Clase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
        MiInt = .Stats.UserAtributos(eAtributos.Inteligencia) * 3
        .Stats.MaxMAN = MiInt
        .Stats.MinMAN = MiInt
    ElseIf p_Clase = eClass.Cleric Or p_Clase = eClass.Druid _
        Or p_Clase = eClass.Bard Or p_Clase = eClass.Assasin Then
            .Stats.MaxMAN = 50
            .Stats.MinMAN = 50
    ElseIf p_Clase = eClass.Bandit Then 'Mana Inicial del Bandido (ToxicWaste)
            .Stats.MaxMAN = 50
            .Stats.MinMAN = 50
    Else
        .Stats.MaxMAN = 0
        .Stats.MinMAN = 0
    End If
   
    If p_Clase = eClass.Mage Or p_Clase = eClass.Cleric Or _
       p_Clase = eClass.Druid Or p_Clase = eClass.Bard Or _
       p_Clase = eClass.Assasin Then
            .Stats.UserHechizos(1) = 2
       
            If p_Clase = eClass.Druid Then .Stats.UserHechizos(2) = 46
    End If
   
    .Stats.MaxHIT = 2
    .Stats.MinHIT = 1
   
    .Stats.GLD = 0
   
    .Stats.Exp = 0
    .Stats.ELU = 300
    .Stats.ELV = 1
   
    '???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
    Dim Slot As Byte
    Dim IsPaladin As Boolean
   
    IsPaladin = p_Clase = eClass.Paladin
   
    'Pociones Rojas (Newbie)
    Slot = 1
    .Invent.Object(Slot).ObjIndex = 857
    .Invent.Object(Slot).amount = 200
   
    'Pociones azules (Newbie)
    If .Stats.MaxMAN > 0 Or IsPaladin Then
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = 856
        .Invent.Object(Slot).amount = 200
   
    Else
        'Pociones amarillas (Newbie)
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = 855
        .Invent.Object(Slot).amount = 100
   
        'Pociones verdes (Newbie)
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = 858
        .Invent.Object(Slot).amount = 50
   
    End If
   
    ' Ropa (Newbie)
    Slot = Slot + 1
    Select Case p_Raza
        Case eRaza.Humano
            .Invent.Object(Slot).ObjIndex = 463
        Case eRaza.Elfo
            .Invent.Object(Slot).ObjIndex = 464
        Case eRaza.Drow
            .Invent.Object(Slot).ObjIndex = 465
        Case eRaza.Enano
            .Invent.Object(Slot).ObjIndex = 466
        Case eRaza.Gnomo
            .Invent.Object(Slot).ObjIndex = 466
    End Select
   
    ' Equipo ropa
    .Invent.Object(Slot).amount = 1
    .Invent.Object(Slot).Equipped = 1
   
    .Invent.ArmourEqpSlot = Slot
    .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).ObjIndex
 
    'Arma (Newbie)
    Slot = Slot + 1
    Select Case p_Clase
        Case eClass.Hunter
            ' Arco (Newbie)
            .Invent.Object(Slot).ObjIndex = 859
        Case eClass.Worker
            ' Herramienta (Newbie)
            .Invent.Object(Slot).ObjIndex = RandomNumber(561, 565)
        Case Else
            ' Daga (Newbie)
            .Invent.Object(Slot).ObjIndex = 460
    End Select
   
    ' Equipo arma
    .Invent.Object(Slot).amount = 1
    .Invent.Object(Slot).Equipped = 1
   
    .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).ObjIndex
    .Invent.WeaponEqpSlot = Slot
   
    .Char.WeaponAnim = GetWeaponAnim(UserIndex, .Invent.WeaponEqpObjIndex)
 
    ' Municiones (Newbie)
    If p_Clase = eClass.Hunter Then
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = 860
        .Invent.Object(Slot).amount = 150
       
        ' Equipo flechas
        .Invent.Object(Slot).Equipped = 1
        .Invent.MunicionEqpSlot = Slot
        .Invent.MunicionEqpObjIndex = 860
    End If
 
    ' Manzanas (Newbie)
    Slot = Slot + 1
    .Invent.Object(Slot).ObjIndex = 467
    .Invent.Object(Slot).amount = 100
   
    ' Jugos (Nwbie)
    Slot = Slot + 1
    .Invent.Object(Slot).ObjIndex = 468
    .Invent.Object(Slot).amount = 100
   
    ' Sin casco y escudo
    .Char.ShieldAnim = NingunEscudo
    .Char.CascoAnim = NingunCasco
   
    ' Total Items
    .Invent.NroItems = Slot
   
    #If ConUpTime Then
        .LogOnTime = Now
        .UpTime = 0
    #End If
 
End With
 
'Valores Default de facciones al Activar nuevo usuario
Call ResetFacciones(UserIndex)
 
Call SaveUser(UserIndex, CharPath & UCase$(p_name) & ".pjs")
 
'Agrega el personaje a la cuenta.
Call Agregar_Personaje(UserIndex, UserList(UserIndex).acc_User_Name, p_name)
         
'Envia nuevamente los personajes.
Call Enviar_Personajes(UserIndex, UserList(UserIndex).acc_User_Name)
 
 
 
End Sub
 
Public Sub ConectarPersonaje(ByVal UserIndex As Integer, ByVal acc_Name As String, ByVal acc_Personaje As Byte)
 
'
' @ Conecta un personaje.
 
Dim personaje_Name  As String
 
personaje_Name = Personaje(acc_Name, acc_Personaje)
 
'No hay personaje.
If (Not personaje_Name <> vbNullString) Then Exit Sub
 
'Ya está logeado?
If (MULTI_LOG = 1) Then
   If CheckForSameName(personaje_Name) Then
      Call Protocol.WriteErrorMsg(UserIndex, "El personaje ya está logeado.")
      Call Protocol.FlushBuffer(UserIndex)
      Call TCP.CloseSocket(UserIndex)
      Exit Sub
   End If
End If
 
With UserList(UserIndex)
 
    'Reseteamos los FLAGS
    .flags.Escondido = 0
    .flags.TargetNPC = 0
    .flags.TargetNpcTipo = eNPCType.Comun
    .flags.TargetObj = 0
    .flags.TargetUser = 0
    .Char.FX = 0
   
    'Reseteamos los privilegios
    .flags.Privilegios = 0
   
    'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
    If EsAdmin(personaje_Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
        Call LogGM(personaje_Name, "Se conecto con ip:" & .ip)
    ElseIf EsDios(personaje_Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
        Call LogGM(personaje_Name, "Se conecto con ip:" & .ip)
    ElseIf EsSemiDios(personaje_Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
        Call LogGM(personaje_Name, "Se conecto con ip:" & .ip)
    ElseIf EsConsejero(personaje_Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
        Call LogGM(personaje_Name, "Se conecto con ip:" & .ip)
    Else
        .flags.Privilegios = .flags.Privilegios Or PlayerType.User
        .flags.AdminPerseguible = True
    End If
   
    'Add RM flag if needed
    If EsRolesMaster(personaje_Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
    End If
   
    If ServerSoloGMs > 0 Then
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
            Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
   
    'Cargamos el personaje
    Dim Leer As New clsIniReader
   
    Call Leer.Initialize(CharPath & UCase$(personaje_Name) & ".pjs")
   
    'Cargamos los datos del personaje
    Call LoadUserInit(UserIndex, Leer)
   
    Call LoadUserStats(UserIndex, Leer)
   
    If Not ValidateChr(UserIndex) Then
        Call WriteErrorMsg(UserIndex, "Error en el personaje.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
   
    Call LoadUserReputacion(UserIndex, Leer)
   
    Set Leer = Nothing
   
    If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
    If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
    If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
   
    If .Invent.MochilaEqpSlot > 0 Then
        .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(.Invent.Object(.Invent.MochilaEqpSlot).ObjIndex).MochilaType * 5
    Else
        .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
    End If
    If (.flags.Muerto = 0) Then
        .flags.SeguroResu = False
        Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)
    Else
        .flags.SeguroResu = True
        Call WriteMultiMessage(UserIndex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
    End If
   
    Call UpdateUserInv(True, UserIndex, 0)
    Call UpdateUserHechizos(True, UserIndex, 0)
   
    If .flags.Paralizado Then
        Call WriteParalizeOK(UserIndex)
    End If
   
    ''
    'TODO : Feo, esto tiene que ser parche cliente
    If .flags.Estupidez = 0 Then
        Call WriteDumbNoMore(UserIndex)
    End If
   
    'Posicion de comienzo
    If .Pos.Map = 0 Then
        Select Case .Hogar
            Case eCiudad.cNix
                .Pos = Nix
            Case eCiudad.cUllathorpe
                .Pos = Ullathorpe
            Case eCiudad.cBanderbill
                .Pos = Banderbill
            Case eCiudad.cLindos
                .Pos = Lindos
            Case eCiudad.cArghal
                .Pos = Arghal
            Case Else
                .Hogar = eCiudad.cUllathorpe
                .Pos = Ullathorpe
        End Select
    Else
        If Not MapaValido(.Pos.Map) Then
            Call WriteErrorMsg(UserIndex, "El PJ se encuenta en un mapa inválido.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
   
   
    'Nombre de sistema
    .name = personaje_Name
   
    .showName = True 'Por default los nombres son visibles
   
    'If in the water, and has a boat, equip it!
    If .Invent.BarcoObjIndex > 0 And _
            (HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Or BodyIsBoat(.Char.body)) Then
            Dim Barco As ObjData
            Barco = ObjData(.Invent.BarcoObjIndex)
            .Char.Head = 0
            If .flags.Muerto = 0 Then
                Call ToogleBoatBody(UserIndex)
            Else
                .Char.body = iFragataFantasmal
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
            End If
       
        .flags.Navegando = 1
    End If
   
   
    'Info
    Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
    Call WriteChangeMap(UserIndex, .Pos.Map, MapInfo(.Pos.Map).MapVersion) 'Carga el mapa
    Call WritePlayMidi(UserIndex, val(ReadField(1, MapInfo(.Pos.Map).Music, 45)))
   
    If .flags.Privilegios = PlayerType.Dios Then
        .flags.ChatColor = RGB(250, 250, 150)
    ElseIf .flags.Privilegios <> PlayerType.User And .flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
        .flags.ChatColor = RGB(0, 255, 0)
    ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
        .flags.ChatColor = RGB(0, 255, 255)
    ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
        .flags.ChatColor = RGB(255, 128, 64)
    Else
        .flags.ChatColor = vbWhite
    End If
   
   
    ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
    #If ConUpTime Then
        .LogOnTime = Now
    #End If
   
    'Crea  el personaje del usuario
    Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y)
   
    Call WriteUserCharIndexInServer(UserIndex)
    ''[/el oso]
   
    Call CheckUserLevel(UserIndex)
    Call WriteUpdateUserStats(UserIndex)
   
    Call WriteUpdateHungerAndThirst(UserIndex)
    Call WriteUpdateStrenghtAndDexterity(UserIndex)
       
    Call SendMOTD(UserIndex)
 
    'Actualiza el Num de usuarios
    'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
    NumUsers = NumUsers + 1
    .flags.UserLogged = True
   
    'usado para borrar Pjs
    Call WriteVar(CharPath & personaje_Name & ".pjs", "INIT", "Logged", "1")
   
    Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
   
    MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
   
    If .Stats.SkillPts > 0 Then
        Call WriteSendSkills(UserIndex)
        Call WriteLevelUp(UserIndex, .Stats.SkillPts)
    End If
   
    If NumUsers > recordusuarios Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultaneamente." & "Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
        recordusuarios = NumUsers
        Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
       
        Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
    End If
   
    If .NroMascotas > 0 And MapInfo(.Pos.Map).Pk Then
        Dim i As Integer
        For i = 1 To MAXMASCOTAS
            If .MascotasType(i) > 0 Then
                .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, True, True)
               
                If .MascotasIndex(i) > 0 Then
                    Npclist(.MascotasIndex(i)).MaestroUser = UserIndex
                    Call FollowAmo(.MascotasIndex(i))
                Else
                    .MascotasIndex(i) = 0
                End If
            End If
        Next i
    End If
   
    If .flags.Navegando = 1 Then
        Call WriteNavigateToggle(UserIndex)
    End If
   
    If criminal(UserIndex) Then
        Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff) 'Call WriteSafeModeOff(UserIndex)
        .flags.Seguro = False
    Else
        .flags.Seguro = True
        Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)
    End If
   
    If .GuildIndex > 0 Then
        'welcome to the show baby...
        If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
            Call WriteConsoleMsg(UserIndex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End If
   
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
   
    Call WriteLoggedMessage(UserIndex)
   
    Call modGuilds.SendGuildNews(UserIndex)
   
    ' Esta protegido del ataque de npcs por 5 segundos, si no realiza ninguna accion
    Call IntervaloPermiteSerAtacado(UserIndex, True)
   
    If Lloviendo Then
        Call WriteRainToggle(UserIndex)
    End If
   
    Dim tStr    As String
   
    tStr = modGuilds.a_ObtenerRechazoDeChar(personaje_Name)
   
    If LenB(tStr) <> 0 Then
        Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
    End If
   
    'Load the user statistics
    Call Statistics.UserConnected(UserIndex)
   
    Call MostrarNumUsers
   
    #If SeguridadAlkon Then
        Call Security.UserConnected(UserIndex)
    #End If
   
    Dim N As Integer
   
    N = FreeFile
    Open App.Path & "\logs\numusers.log" For Output As N
    Print #N, NumUsers
    Close #N
   
    N = FreeFile
    'Log
    Open App.Path & "\logs\Connect.log" For Append Shared As #N
    Print #N, personaje_Name & " ha entrado al juego. UserIndex:" & UserIndex & " " & time & " " & Date
    Close #N
 
End With
 
   
End Sub
 
Public Sub EliminarPersonaje(ByVal UserIndex As Integer, ByVal acc_Name As String, ByVal acc_Personaje As Byte)
 
'
' @ Elimina un personaje.
 
Dim t_Str       As String
Dim dummy_Char  As Char_Acc_Data
 
If Validar_Eliminacion(acc_Name, acc_Personaje, t_Str) Then
 
    'Si pasó la validación entonces lo borramos y cerramos la conexión.
    Kill CharPath & Personaje(acc_Name, acc_Personaje) & ".pjs"
    WriteVar Acc_Path & acc_Name & ".sdc", "PJS", "PJ" & CStr(acc_Personaje), "NoUsado"
   
    Call CompactarListaPersonajes(acc_Name)
   
    Call Protocol.WriteErrorMsg(UserIndex, "El personaje ha sido borrado.")
    Call Protocol.FlushBuffer(UserIndex)
    Call TCP.CloseSocket(UserIndex)
Else    'No pasa la validación, envio el mensaje de error _
         Y cierro el socket.
    Call Protocol.WriteErrorMsg(UserIndex, t_Str)
    Call Protocol.FlushBuffer(UserIndex)
    Call TCP.CloseSocket(UserIndex)
End If
 
End Sub
 
Private Sub CompactarListaPersonajes(ByVal acc_Name As String)
 
'
' @ Ordena la lista de personajes.
 
Dim free_Slot   As Integer
Dim do_Loop     As Integer
Dim i_Personaje As String
 
do_Loop = 1
 
Do While (do_Loop <= MAX_PJS)
   'El slot está vacio y no tengo un slot.
   If (Personaje(acc_Name, CByte(do_Loop)) = vbNullString) And (free_Slot = 0) Then
       free_Slot = do_Loop
   ElseIf (Personaje(acc_Name, CByte(do_Loop)) <> vbNullString) And (free_Slot <> 0) Then
       'Tengo un personaje y un slot libre, hago el intercambio,
       i_Personaje = Personaje(acc_Name, CByte(do_Loop))
       
       'En el slot libre guardo el temporal.
       WriteVar Acc_Path & acc_Name & ".sdc", "PJS", "PJ" & CStr(free_Slot), i_Personaje
       
       'Reseteo el slot en el que estaba el anterior personaje.
       WriteVar Acc_Path & acc_Name & ".sdc", "PJS", "PJ", ""
       
       'Muevo la posicion del bucle a el ultimo slot encontrado.
       do_Loop = free_Slot
       
       'Reseteo el slot libre
       free_Slot = 0
   End If
   
   'Aumento el contador.
   do_Loop = do_Loop + 1
Loop
 
End Sub
 
Public Sub Conectar(ByVal UserIndex As Integer, ByVal acc_Name As String, ByVal acc_Password As String)
 
'
' @ Conecta la cuenta.
 
Dim tmp_Error   As String
 
'Pasa la validación?
If Validar_Logeo(acc_Name, acc_Password, tmp_Error) Then
   'No está logeada la cuenta?
   If Not Acc_Logged(acc_Name) Then
 
      'Redimensiona el array de cuentas logeadas y guarda el nombre.
      ReDim Preserve Acc_Loggeds(1 To (UBound(Acc_Loggeds()) + 1)) As String
     
      Acc_Loggeds(UBound(Acc_Loggeds())) = acc_Name
       
      'Envia los personajes.
      Call Enviar_Personajes(UserIndex, acc_Name)
     
      UserList(UserIndex).acc_User_Name = acc_Name
     
   Else
      Call Protocol.WriteErrorMsg(UserIndex, "Ya hay alguien conectado en la cuenta.")
      Call Protocol.FlushBuffer(UserIndex)
      Call TCP.CloseSocket(UserIndex)
   End If
Else
    Call Protocol.WriteErrorMsg(UserIndex, tmp_Error)
    Call Protocol.FlushBuffer(UserIndex)
    Call TCP.CloseSocket(UserIndex)
End If
 
End Sub
 
Public Sub Agregar_Personaje(ByVal UserIndex As Integer, ByVal acc_Name As String, ByVal personaje_Name As String)
 
'
' @ Agrega un personaje a la cuenta.
 
Dim slot_Personaje  As Byte
 
'Busca un slot para un personaje.
slot_Personaje = Libre_Slot(acc_Name)
 
'Encuentra slot.
If (slot_Personaje <> 0) Then
   'Agrega.
   WriteVar Acc_Path & acc_Name & ".sdc", "PERSONAJES", "PERSONAJE" & CByte(slot_Personaje), personaje_Name
Else
   Call Protocol.WriteErrorMsg(UserIndex, "No tienes espacio para más personajes.")
   Call FlushBuffer(UserIndex)
   Call CloseSocket(UserIndex)
End If
 
End Sub
 
Public Sub Enviar_Personajes(ByVal UserIndex As Integer, ByVal acc_Name As String)
 
'
' @ Envia los personajes de la cuenta.
 
Dim loopX   As Long
Dim now_Pj  As Char_Acc_Data
 
For loopX = 1 To MAX_PJS
    'Hay un personaje?
    If (Personaje(acc_Name, CByte(loopX)) <> vbNullString) Then
       now_Pj = Obtener_Personaje(acc_Name, CByte(loopX))
       
       Call Protocol.WriteAccountPersonaje(UserIndex, CByte(loopX), now_Pj)
    End If
Next loopX
 
'Termina de enviar los personajes.
Call Protocol.WriteAccountShow(UserIndex)
 
End Sub
 
Public Function Validar_Recuperacion(ByVal acc_Name As String, ByVal acc_Respuesta As String, ByRef f_Error As String) As Boolean
 
'
' @ Checkea un recuperar pass.
 
Validar_Recuperacion = False
 
    If (acc_Name <> vbNullString) Then
       If UCase$(Respuesta(acc_Name)) = UCase$(acc_Respuesta) Then
           Validar_Recuperacion = True
       Else
           f_Error = "La respuesta secreta no es correcta!"
       End If
    Else
        f_Error = "Ingresa el nombre de la cuenta!"
    End If
   
End Function
 
Public Function Validar_Eliminacion(ByVal acc_Name As String, ByVal acc_Pj As Byte, ByRef e_Error As String) As Boolean
 
'
' @ Checkea el borrado de un personaje de la cuenta.
 
Validar_Eliminacion = False
 
    If (acc_Name <> vbNullString) Then
        If (acc_Pj <> 0) And (acc_Pj <= MAX_PJS) Then
            Dim char_Name   As String
            Dim char_UIndex As Integer
            Dim char_GIndex As Integer
           
            char_Name = Personaje(acc_Name, acc_Pj)
           
            If (char_Name <> vbNullString) Then
                If (PersonajeExiste(char_Name) = True) Then
                    char_UIndex = NameIndex(char_Name)
                   
                    'Está logeado?
                    If (char_UIndex <> 0) Then
                       e_Error = "El personaje se encuentra logeado, no puede ser eliminado."
                    Else
                        char_GIndex = val(GetVar(CharPath & char_Name & ".pjs", "GUILD", "GUILDINDEX"))
                       
                        'No tiene clan
                        If Not char_GIndex <> 0 Then
                            Validar_Eliminacion = True
                        Else
                            e_Error = "El personaje se encuentra en un clan, no puede ser eliminado."
                        End If
                    End If
                Else
                    e_Error = "asdklJÑASKDañlsdjÑALSDa"
                End If
            Else
                e_Error = "Selecciona un personaje!"
            End If
        Else
            e_Error = "Selecciona un personaje!"
        End If
    Else
        e_Error = "Ingresa un nombre."
    End If
 
End Function
 
Public Function Validar_Creacion(ByVal acc_Name As String, ByVal acc_Password As String, ByVal acc_Pregunta As String, ByVal acc_Respuesta As String, ByVal acc_EMail As String, ByRef f_Error As String) As Boolean
 
'
' @ Checkea la creación de una nueva cuenta.
 
Validar_Creacion = False
 
    If (acc_Name <> vbNullString) Then
        If (acc_Password <> vbNullString) Then
            If (acc_Pregunta <> vbNullString) Then
                If (acc_Respuesta <> vbNullString) Then
                   If (acc_EMail <> vbNullString) Then
                       If AsciiValidos(acc_Name) Then
                           If AsciiValidos(acc_Pregunta) And AsciiValidos(acc_Respuesta) Then
                              If Not FileExist(acc_Name) Then
                                  Validar_Creacion = True
                              Else
                                  f_Error = "Ya existe una cuenta con ese nombre."
                              End If
                           Else
                               f_Error = "La pregunta secreta o la respuesta poseen carácteres inválidos."
                           End If
                       Else
                           f_Error = "El nombre posee carácteres inválidos."
                       End If
                   Else
                       f_Error = "Ingresa un email!"
                   End If
                Else
                    f_Error = "Ingresa una respuesta!"
                End If
            Else
                f_Error = "Ingresa una pregunta!"
            End If
        Else
            f_Error = "Ingresa una contraseña!"
        End If
    Else
        f_Error = "Ingresa un nombre para la cuenta!"
    End If
   
End Function
 
Public Function Validar_Logeo(ByVal acc_Name As String, ByVal acc_Password As String, ByRef f_Error As String) As Boolean
 
'
' @ Checkea si puede logear.
 
Validar_Logeo = False
 
    If (acc_Name <> vbNullString) Then
       If (acc_Password <> vbNullString) Then
          If AsciiValidos(acc_Name) Then
             If FileExist(Acc_Path & acc_Name & ".sdc") Then
                If Password(acc_Name) = acc_Password Then
                   If Not Banned(acc_Name) Then
                      Validar_Logeo = True
                   Else
                      f_Error = "La cuenta está baneada."
                   End If
                Else
                    f_Error = "La contraseña no es correcta."
                End If
            Else
                f_Error = "La cuenta no existe."
            End If
          Else
             f_Error = "El nombre de la cuenta posee carácteres inválidos."
          End If
       Else
            f_Error = "Ingresa una contraseña."
       End If
    Else
        f_Error = "Ingresa un nombre de cuenta."
    End If
               
End Function
 
Public Function Obtener_Personaje(ByVal acc_Name As String, ByVal acc_Personaje As Byte) As Char_Acc_Data
 
'
' @ Devuelve datos de un personaje en una cuenta.
 
Dim char_Reader     As New clsIniReader
Dim personaje_Char  As String
Dim temp_Num_Map    As Integer
 
personaje_Char = Personaje(acc_Name, acc_Personaje)
 
'Existe el archivo?
If Not FileExist(CharPath & personaje_Char & ".pjs") Then Exit Function
 
'Inicializa la clase.
char_Reader.Initialize CharPath & personaje_Char & ".pjs"
 
With Obtener_Personaje
     
     'Llena los datos.
     .Nick_Name = personaje_Char
     .Muerto = val(char_Reader.GetValue("FLAGS", "Muerto")) = 1
     
     .Nivel = val(char_Reader.GetValue("STATS", "ELV"))
     
     'Obtiene el número de mapa.
     temp_Num_Map = val(ReadField(1, char_Reader.GetValue("INIT", "Position"), Asc("-")))
     
     'Si es un mapa válido.
     If MapaValido(temp_Num_Map) Then
        If MapInfo(temp_Num_Map).name <> vbNullString Then
            .Pos_Map = MapInfo(temp_Num_Map).name
        Else
            .Pos_Map = "Mapa desconocido"
        End If
     End If
     
     'Llena los datos del char.
     With .Character
          .body = val(char_Reader.GetValue("INIT", "Body"))
         
          .Head = val(char_Reader.GetValue("INIT", "Head"))
         
          .WeaponAnim = val(char_Reader.GetValue("INIT", "Arma"))
         
          .CascoAnim = val(char_Reader.GetValue("INIT", "Casco"))
         
          .ShieldAnim = val(char_Reader.GetValue("INIT", "Escudo"))
 
     End With
     
End With
 
End Function
 
Public Function Acc_Logged(ByVal acc_Name As String) As Boolean
 
'
' @ Checkea si la cuenta está logeada.
   
If (MULTI_LOG = 0) Then
   Dim loopX    As Long
   
   For loopX = 1 To UBound(Acc_Loggeds())
       'La cuenta está logeada.
       If UCase$(Acc_Loggeds(loopX)) = UCase$(acc_Name) Then
          Acc_Logged = True
          Exit Function
       End If
   Next loopX
   
   Acc_Logged = False
Else
   Acc_Logged = False
End If
 
End Function
 
Public Function Acc_Path() As String
 
'
' @ Dir de las cuentas.
 
Acc_Path = App.Path & "\Accounts\"
 
'No existe el directorio?
If Not FileExist(Acc_Path, vbDirectory) Then MkDir Acc_Path
 
End Function
 
Private Function Libre_Slot(ByVal acc_Name As String) As Byte
 
'
' @ Busca un slot para un personaje.
 
Dim loopX       As Long
 
For loopX = 1 To MAX_PJS
   
    If Personaje(acc_Name, CByte(loopX)) = vbNullString Then
       Libre_Slot = CByte(loopX)
       Exit Function
    End If
   
Next loopX
 
Libre_Slot = 0
 
End Function
 
Private Function Banned(ByVal acc_Name As String) As Boolean
 
'
' @ Checkea si la cuenta está baneada.
 
Banned = val(GetVar(Acc_Path & acc_Name & ".sdc", "INIT", "Ban")) <> 0
 
End Function
 
Private Function Password(ByVal acc_Name As String) As String
 
'
' @ Devuelve la pass de una cuenta.
 
Password = GetVar(Acc_Path & acc_Name & ".sdc", "INIT", "Password")
 
End Function
 
Public Function Pregunta(ByVal acc_Name As String) As String
 
'
' @ Devuelve la pregunta secreta de la cuenta.
 
Pregunta = GetVar(Acc_Path & acc_Name & ".sdc", "CONTACTO", "Pregunta")
 
End Function
 
Public Function Respuesta(ByVal acc_Name As String) As String
 
'
' @ Devuelve la respuesta de la pregunta secreta.
 
Respuesta = GetVar(Acc_Path & acc_Name & ".sdc", "CONTACTO", "Respuesta")
 
End Function
 
Private Function Personaje(ByVal acc_Name As String, ByVal acc_Personaje As Byte) As String
 
'
' @ Devuelve el nombre de un personaje.
 
Personaje = GetVar(Acc_Path & acc_Name & ".sdc", "PERSONAJES", "PERSONAJE" & CStr(acc_Personaje))
 
If (Personaje = "NoUsado") Then Personaje = vbNullString
 
End Function

