Attribute VB_Name = "modDeclaraciones"
Option Explicit



Public ModoHabla As Byte
Public PrivateTo As String

Public picMouseIcon As Picture

Public FxNavega As Byte
'Caminata Fluida - Dunkansan
Public Movement_Speed As Single

Public MostrarFPS As Boolean
'Velocidad de los grhs al caminar
Public Velocidad As Byte
'Macros--------------------------
Type tBoton
    TipoAccion As Integer
    SendString As String
    hlist As Integer
    invslot As Integer
    amount As Integer
End Type

Public MacroKeys() As tBoton
Public BotonElegido As Integer
'--------------------------------
Public isCapturePending As Boolean
Public Const CentroInventario As Byte = 1
Public Const CentroHechizos As Byte = 2
Public Const CentroMenu As Byte = 3
Public Const Solapas As Byte = 4
Public EstadoLogin As E_MODO
Public day_r_old As Byte
Public day_g_old As Byte
Public day_b_old As Byte

'Climas
Type luzxhora
    r As Long
    g As Long
    b As Long
End Type
Public CurServerIp As String
Public CurServerPort As Integer
'Declaramos las luces
Public luz_dia(0 To 24) As luzxhora '¬¬ la hora 24 dura 1 minuto entre las 24 y las 0
'Fin
Public lblexpactivo As Boolean
Public InfoMapAct As Boolean

Public Windows_Temp_Dir As String

'Objetos públicos
Public Audio As New clsAudio
Public Light As New clsLight
Public Inventario As New clsGrapchicalInventory
Public CustomKeys As New clsCustomKeys
Public CustomMessages As New clsCustomMessages

Public incomingData As New clsByteQueue
Public outgoingData As New clsByteQueue

''
'The main timer of the game.
Public MainTimer As New clsTimer

Public Const CANT_GRH_INDEX As Long = 40000

'Sonidos
Public Const SND_CLICK As String = "190.Wav"
Public Const SND_PASOS1 As String = "23.Wav"
Public Const SND_PASOS2 As String = "24.Wav"
Public Const SND_RETIRAR As String = "172.Wav"
Public Const SND_NAVEGANDO As String = "50.wav"

Public Const SND_PASOS3 As String = "197.Wav" 'Pie 1 de Arena
Public Const SND_PASOS4 As String = "198.Wav" 'Pie 2 de arena
Public Const SND_PASOS5 As String = "199.Wav" 'Pie 3 de  arena
Public Const SND_PASOS6 As String = "200.Wav" 'Pie nieve
Public Const SND_PASOS7 As String = "201.Wav" 'Pie nieve
Public Const SND_PASOS8 As String = "202.Wav" 'Pie  de pasto

Public Const SND_OVER As String = "0.Wav"
Public Const SND_DICE As String = "cupdice.Wav"
Public Const SND_LLUVIAINEND As String = "191.wav"
Public Const SND_LLUVIAOUTEND As String = "194.wav"
Public Const SND_CUERVO As Integer = 126
Public Const SND_TRUENO1 As Integer = 60
Public Const SND_TRUENO2 As Integer = 61
Public Const SND_TRUENO3 As Integer = 62
Public Const SND_TRUENO4 As Integer = 63
Public Const SND_TRUENO5 As Integer = 64
' Head index of the casper. Used to know if a char is killed

' Constantes de intervalo
Public Const INT_MACRO_HECHIS As Integer = 2788
Public Const INT_MACRO_TRABAJO As Integer = 600

Public Const INT_ATTACK As Integer = 1500
Public Const INT_ARROWS As Integer = 1400
Public Const INT_CAST_SPELL As Integer = 1400
Public Const INT_CAST_ATTACK As Integer = 1000
Public Const INT_WORK As Integer = 700
Public Const INT_USEITEMU As Integer = 450
Public Const INT_USEITEMDCK As Integer = 220
Public Const INT_SENTRPU As Integer = 2000

Public MacroBltIndex As Integer

Public Const CASPER_HEAD As Integer = 500

Public Const NUMATRIBUTES As Byte = 5

'Musica
Public Const MIdi_Inicio As Byte = 6

Public MapName As String

Public RawServersList As String

Public Type tColor
    r As Byte
    g As Byte
    b As Byte
End Type

Public ColoresPJ(0 To 50) As tColor

Type tServerInfo
    port As Integer
    Ip As String
    name As String
End Type
Public lServer(1 To 1) As tServerInfo
Public ServersLst() As tServerInfo
Public ServersRecibidos As Boolean

Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer
Public ObjSastreria(0 To 100) As Integer
Public objalquimia(0 To 100) As Integer
Public UsaMacro As Boolean
Public CnTd As Byte

'Musica
Public Const MP3_Inicio As Byte = 54

'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]


Public Tips() As String * 255
Public Const LoopAdEternum As Integer = 999

'Direcciones
Public Enum E_Heading
    north = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS As Integer = 10000
Public Const MAX_INVENTORY_SLOTS As Byte = 25
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50
Public Const MAXHECHI As Byte = 35

Public Const MAXSKILLPOINTS As Byte = 100

Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1

Public Const FOgata As Integer = 1521


Public Enum eClass
    Mage = 1    'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Thief       'Ladrón
    Bard        'Bardo
    Druid       'Druida
    Bandit      'Bandido
    Paladin     'Paladín
    Hunter      'Cazador
    Fisher      'Pescador
    Blacksmith  'Herrero
    Lumberjack  'Leñador
    Miner       'Minero
    Carpenter   'Carpintero
    drakkar       'drakkar
    BountyHunter 'Nigromante
    Sastre 'Sastre
    Nigromante 'Nigromante
End Enum

Public Enum eCiudad
    cUllathorpe = 1
    cNix
    cBanderbill
    CRinkel
    cArghal
End Enum

Enum eRaza
    HUMANO = 1
    ELFO
    ElfoOscuro
    Gnomo
    Enano
    Orco
End Enum

Public Enum eSkill
    Tacticas = 1
    Armas = 2
    Artes = 3
    Apuñalar = 4
    Arrojadizas = 5
    Proyectiles = 6
    DefensaEscudos = 7
    magia = 8
    Resistencia = 9
    Meditar = 10
    Ocultarse = 11
    Domar = 12
    Musica = 13
    Robar = 14
    Comercio = 15
    Supervivencia = 16
    Liderazgo = 17
    Pesca = 18
    Mineria = 19
    Talar = 20
    Botanica = 21
    Herreria = 22
    Carpinteria = 23
    Alquimia = 24
    Sastreria = 25
    Navegacion = 26
    Equitacion = 27
End Enum

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

Public Enum eObjType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otescudo = 16
    otcasco = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otMapa = 38
    otMonturas = 44
    otCualquiera = 1000
    Otnudillos = 46
End Enum
Type tHeadRange
    mStart As Integer
    mEnd As Integer
    fStart As Integer
    fEnd As Integer
End Type
Public Head_Range() As tHeadRange

Public Const FundirMetal As Integer = 88

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "La criatura fallo el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = "La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = "El usuario rechazo el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "Has fallado el golpe!!!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = "¡Has activado el seguro!"
Public Const MENSAJE_SEGURO_DESACTIVADO As String = "¡Has desactivado el Seguro!"
Public Const MENSAJE_PIERDE_NOBLEZA As String = "¡¡Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertirás en uno de ellos y serás perseguido por las tropas de las ciudades."
Public Const MENSAJE_USAR_MEDITANDO As String = "¡Estás meditando! Debes dejar de meditar para usar objetos."
Public Const MENSAJE_ENVIAR_PENA As String = "Tiempo restante en la carcel & .penacarcel"
Public Const MENSAJE_SEGURO_RESU_ON As String = "SEGURO DE RESURRECCION ACTIVADO"
Public Const MENSAJE_SEGURO_RESU_OFF As String = "SEGURO DE RESURRECCION DESACTIVADO"

Public Const MENSAJE_GOLPE_CABEZA As String = "¡¡La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "¡¡La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "¡¡La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "¡¡La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "¡¡La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "¡¡La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "¡¡"
Public Const MENSAJE_2 As String = "!!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "¡¡Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO As String = " te ataco y fallo!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "¡¡Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA As String = "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la victima..."
Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el árbol..."
Public Const MENSAJE_TRABAJO_MINERIA As String = "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_PROYECTILES As String = "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1 As String = "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "

Public Const COLOR_ATAQUE As Long = -65536

'Inventario
Type Inventory
    OBJIndex As Integer
    name As String
    grhindex As Integer
    '[Alejo]: tipo de datos ahora es Long
    amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Single
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
End Type

Type NpCinV
    OBJIndex As Integer
    name As String
    grhindex As Integer
    amount As Integer
    Valor As Single
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
End Type

Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
     Raza As Byte
    Genero As Byte
    VecesMuerto As Integer
    PenaCarcel As Long
End Type

Public Nombres As Boolean

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserHechizos(1 To MAXHECHI) As Integer

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public UserMeditar As Boolean
Public UserName As String
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserMaxAGU As Byte
Public UserMinAGU As Byte
Public UserMaxHAM As Byte
Public UserMinHAM As Byte
Public UserGLD As Long
Public Usuarios As Integer
Public Offline As Byte
Public Sendingtype As Byte
Public OB As Long
Public UserBOVItem As Long
Public UserLvl As Integer
Public UserPort As Integer
Public UserServerIP As String
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public PrimeraVez As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean
Public IScombate As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserMontando As Boolean
Public UserHogar As eCiudad
Public UserAvisado As Boolean
Public MapExt As Byte
'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase As eClass
Public UserSexo As eGenero
Public UserRaza As eRaza
Public UserEmail As String

Public Const NUMCIUDADES As Byte = 5
Public Const NUMSKILLS As Byte = 27
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 19
Public Const NUMRAZAS As Byte = 6
Public Const NUMRESOLUTIONS As Byte = 1

Public UserSkills(1 To NUMSKILLS) As Byte
Public SkillsNames(1 To NUMSKILLS) As String
Public SkillsOrig(1 To NUMSKILLS) As Byte

Public UserAtributos(1 To NUMATRIBUTOS) As Byte
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public Ciudades(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String
Public ListaRes1(NUMRESOLUTIONS) As String

Public ListaRes2 As String
Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
Public logged As Boolean

Public UsingSkill As Integer

Public MD5HushYo As String * 16

Public pingTime As Long

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
End Enum

   
Public Enum FxMeditar
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
    XXGRANDE = 34
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
End Enum

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger
    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer

'Control
Public prgRun As Boolean 'When true the program ends
Public FinPres As Boolean
Public FinGG As Boolean
Public Win2kXP As Boolean

'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el Internet Explorer para el manual
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

'Particle Groups
Public TotalStreams As Integer
Public StreamData() As Stream

Public Type Stream
    name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As tColor
    
    speed As Single
    life_counter As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
Public meteo_particle As Integer
'*****************************
'GAMEINI
Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tSetupMods
    bNoRes      As Boolean ' 24/06/2006 - ^[GS]^
End Type

Public ClientSetup As tSetupMods

Public MiCabecera As tCabecera
'**************************
Public temp_color_array(3) As Long
