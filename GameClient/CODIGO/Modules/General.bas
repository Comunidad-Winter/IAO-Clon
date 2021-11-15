Attribute VB_Name = "modGeneral"
'************************************************* ****************
'ImperiumAO - v1.0
'************************************************* ****************
'Copyright (C) 2015 Gaston Jorge Martinez
'Copyright (C) 2015 Alexis Rodriguez
'Copyright (C) 2015 Luis Merino
'Copyright (C) 2015 Girardi Luciano Valentin
'
'Respective portions copyright by taxpayers below.
'
'This library is free software; you can redistribute it and / or
'Modify it under the terms of the GNU General Public
'License as published by the Free Software Foundation version 2.1
'The License
'
'This library is distributed in the hope that it will be useful,
'But WITHOUT ANY WARRANTY; without even the implied warranty
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'************************************************* ****************
'
'************************************************* ****************
'You can contact me at:
'Gaston Jorge Martinez (Zenitram@Hotmail.com)
'************************************************* ****************
Option Explicit

'Set mouse speed
Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
    ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
 
Private Const SPI_SETMOUSESPEED = 113
Private Const SPI_GETMOUSESPEED = 112

'***************************
'Sinuhe - Map format .CSM
'***************************
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueados
    X As Integer
    Y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    Y As Integer
    grhindex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    Y As Integer
    Trigger As Integer
End Type

Private Type tDatosLuces
    X As Integer
    Y As Integer
    color As tColor
    extra As Byte
    range As Byte
End Type



Private Type tDatosParticulas
    X As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    Y As Integer
    NpcIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    OBJIndex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDat
    map_name As String * 64
    battle_mode As Byte
    backup_mode As Byte
    restrict_mode As String * 4
    music_number As String * 16
    zone As String * 16
    terrain As String * 16
    Ambient As String * 16
    base_light As Long
    letter_grh As Long
    extra1 As Long
    extra2 As Long
    extra3 As String * 32
End Type

Private MapSize As tMapSize
Public MapDat As tMapDat

Public iplst As String
Public bFogata As Boolean
Public bLluvia() As Byte

Private lFrameTimer As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private OSInfo As OSVERSIONINFO

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

'************************
'To get OS version
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type
Private Declare Function GetOSVersion Lib "kernel32" _
Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32s As Long = 0&
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1&
Private Const VER_PLATFORM_WIN32_NT As Long = 2&

Public Function DirGraficos() As String
    DirGraficos = App.path & "\Resources\graphics\"
End Function

Public Function DirSound() As String
    DirSound = App.path & "\Resources\Sounds\"
End Function

Public Function DirMidi() As String
    DirMidi = App.path & "\Resources\Sounds\"
End Function

Public Function DirMapas() As String
    DirMapas = App.path & "\Resources\Maps\"
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

#If SeguridadAlkon Then
Sub InitMI()
    Dim alternativos As Integer
    Dim CualMITemp As Integer
    
    alternativos = RandomNumber(1, 7368)
    CualMITemp = RandomNumber(1, 1233)
    

    Set MI(CualMITemp) = New clsManagerInvisibles
    Call MI(CualMITemp).Inicializar(alternativos, 10000)
    
    If CualMI <> 0 Then
        Call MI(CualMITemp).CopyFrom(MI(CualMI))
        Set MI(CualMI) = Nothing
    End If
    CualMI = CualMITemp
End Sub
#End If

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************r
    With RichTextBox
        .SelFontName = "Tahoma"
        .SelFontSize = 8
        
        If Len(.Text) > 10000 Then .Text = ""
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = IIf(bold, True, False)
        .SelItalic = IIf(italic, True, False)
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        'RichTextBox.Refresh
    End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).CharIndex = loopc
        End If
    Next loopc
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopc
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

#If SeguridadAlkon Then
    Call UnprotectForm
#End If

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
    
    Call Audio.PlayWave(SND_PORTAL)
    
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    'Unload the connect form
    Unload frmConnect
    Unload frmCrearPersonaje
    Unload frmCrearCuenta
    Unload frmPanelAccount
    
    frmMain.Label8.Caption = UserName
    
    'Macros
    Call LoadMacros(UserName)
    
    Dim i As Byte
   
    For i = 1 To 11
    If MacroList(i).mTipe = 0 Or MacroList(i).Grh <= 0 Then
    frmMain.Macros(i).Cls
    Call Engine.DrawGrhToHdc(frmMain.Macros(i).hdc, 1, 0, 0)
    frmMain.Macros(i).Refresh
    Else
    frmMain.Macros(i).Cls
    Call Engine.DrawGrhToHdc(frmMain.Macros(i).hdc, MacroList(i).Grh, 0, 0)
    frmMain.Macros(i).Refresh
    End If
    If MacroList(i).Grh = 0 Then
    frmMain.Macros(i).BackColor = &H0&
    End If
    Next i
    
    'Load main form
    frmMain.Visible = True
    frmMain.Height = 9030

End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.north
            LegalOk = LegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.east
            LegalOk = LegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = LegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk And Not UserParalizado Then
        If Not UserDescansar And Not UserMeditar Then
            Call WriteWalk(Direccion) 'We only walk if we are not meditating or resting
            Engine.Char_Move_by_Head UserCharIndex, Direccion
            MoveScreen Direccion
            Call ActualizarMiniMapa(Direccion)
        Else
            If Not UserAvisado Then
                If UserDescansar Then
                    WriteRest 'Stop resting (we do NOT have the 1 step enforcing anymore) sono como un tema de los guns.
                ElseIf UserMeditar Then
                    WriteMeditate 'Stop meditation
                End If
                UserAvisado = True
            End If
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call MoveTo(RandomNumber(north, WEST))
End Sub

Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
    Static lastMovement As Long
    
    'No input allowed while Argentum is not the active window
    If Not modApplication.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If frmForo.Visible Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                Call MoveTo(north)
                Call InfoMapa
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                Call MoveTo(east)
                Call InfoMapa
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                Call MoveTo(SOUTH)
                Call InfoMapa
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                Call MoveTo(WEST)
                Call InfoMapa
                Exit Sub
            End If
            
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
            End If
            
            Call InfoMapa
        End If
    End If
End Sub

'TODO : Si bien nunca estuvo allí, el mapa es algo independiente o a lo sumo dependiente del engine, no va acá!!!

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function


Sub Main()

'**************************************
' Autor: Gaston Martinez
'**************************************

'**************************************
'Iniciar Apariencia de Windows In Game!
             InitManifest
'**************************************

'Anti Cambiar Nombre - *************************************
'NombreOriginal = "ImperiumAO"
'NoCambiesNombre = App.EXEName
'NoCambiesNombre = App.EXEName

 'If NoCambio Then
  '   Call ErrorCambiName
 '  End
 'End If
'***********************************************************

'Iniciamos Pantalla Completa Segun el Launcher.
    'Call modResolution.SetResolution

'Programs detecting unsuitable to the game from the Second Timer
If Detected("sXe Injected.exe") Then
Call MsgBox("ATENCIÒN: Se detecto el programa: sXe Injected, puede causar el mal funcionamiento del juego, cierrelo.", vbApplicationModal + vbInformation + vbOKOnly, "Seguridad")
End

ElseIf Detected("Cheat Engine.exe") Then
Call MsgBox("ATENCIÒN: Se detecto el programa: Cheat Engine, puede causar el mal funcionamiento del juego, cierrelo.", vbApplicationModal + vbInformation + vbOKOnly, "Seguridad")
End

ElseIf Detected("svchost.exe.exe") Then
Call MsgBox("ATENCIÒN: Se detecto el programa: svchost.exe.exe, puede causar el mal funcionamiento del juego, cierrelo.", vbApplicationModal + vbInformation + vbOKOnly, "Seguridad")
End

ElseIf Detected("processhacker.exe") Then
Call MsgBox("ATENCIÒN: Se detecto el programa: Process Hacker, puede causar el mal funcionamiento del juego, cierrelo.", vbApplicationModal + vbInformation + vbOKOnly, "Seguridad")
End

ElseIf Detected("Sandboxie.exe") Then
Call MsgBox("ATENCIÒN: Se Detecto el Programa: SandBoxie.exe, Puede Causar el mal funcionamiento del Juego, Por favor Cierrelo", vbApplicationModal + vbInformation + vbOKOnly, "Seguridad")
End

End If
    
    
    'ANTI-DOBLE CLIENT'S
    'If FindPreviousInstance Then
    '    Call MsgBox("ImperiumAO ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Seguridad")
    '    End
    'End If
    
    'Read command line. Do it AFTER config file is loaded to prevent this from
    'canceling the effects of "/nores" option.
    Call LeerLineaComandos
    
    'Segurity in MD5 / ANTI-EDIT CLIENT
    ChDrive App.path
    ChDir App.path
    Win2kXP = General_Windows_Is_2000XP
    MD5HushYo = "0123456789abcdef"  'Pass MD5


    Dim CursorDir As String
    Dim Cursor As Long
    
    
    CursorDir = App.path & "\Resources\MAIN.cur"
    hSwapCursor = SetClassLong(frmMain.hWnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
    hSwapCursor = SetClassLong(frmMain.Renderer.hWnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
    hSwapCursor = SetClassLong(frmMain.hlst.hWnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
    
    frmCargando.Show 'Iniciamos el formulario de cargando
    frmCargando.Refresh


    Call InicializarNombres
    Call frmCargando.progresoConDelay(15)
    Call Engine.setup_ambient
    
    ' Initialize FONTTYPES
    Call modProtocol.InitFonts

    Call Engine.Engine_Init
    Call Engine.Engine_Font_Initialize
    
    'Inicializamos el sonido
    Call Audio.Initialize(DX8, frmMain.hWnd, App.path & "\Resources\Sounds\", App.path & "\Resources\MP3\")
    
    
    
    'Enable / Disable audio
    Audio.MusicActivated = True 'Midi y MP3 (Hay que separarlos)
    Audio.SoundActivated = True 'Wavs
    Audio.AmbientActivated = True 'Ambient
    Audio.AmbientVolume = 95 'Volumen de sonidos de ambiente
    Call Audio.Music_Load(1)
    Set Light = New clsLight
    
    UserMap = 1
    
    Call frmCargando.progresoConDelay(45)
    
    'Cargando Engine
    Call LoadGrhData
    Call CargarCabezas
    Call CustomKeys.LoadCustomKeys
    Call CargarCascos
    Call CargarCuerpos
    Call ObjLuz
    Call CargarFxs
    Velocidad = 1
    Call CargarParticulas
    Call CargarAnimArmas
    Call CargarAnimEscudos
    
    Load frmPres
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv)
    frmMain.Socket1.Startup
    Call frmCargando.progresoConDelay(85)
    
    'Give the user enough time to read the welcome text
    Call Sleep(1750)
    
    frmPres.Show
    Unload frmCargando
    
    Do While Not FinPres
        DoEvents
    Loop

    Call frmCargando.progresoConDelay(100)

    frmConnect.Visible = True
    Unload frmPres
    
    'Inicialización de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    
    frmMain.macrotrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.macrotrabajo.Enabled = False
    
Engine.Start
End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As String) As String
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    
    LastPos = 0
    CurrentPos = 0
    
    For i = 1 To field_pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        General_Field_Read = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        General_Field_Read = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).grhindex >= 1505 And MapData(X, Y).Graphic(1).grhindex <= 1520) Or _
            (MapData(X, Y).Graphic(1).grhindex >= 5665 And MapData(X, Y).Graphic(1).grhindex <= 5680) Or _
            (MapData(X, Y).Graphic(1).grhindex >= 13547 And MapData(X, Y).Graphic(1).grhindex <= 13562)) And _
                MapData(X, Y).Graphic(2).grhindex = 0
                
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub
    
Public Sub LeerLineaComandos()
    Dim T() As String
    Dim i As Long
    
    'Parseo los comandos
    T = Split(Command, " ")
    For i = LBound(T) To UBound(T)
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
        End Select
    Next i
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(eCiudad.cBanderbill) = "Banderbill (Imperiales)"
    Ciudades(eCiudad.cRinkel) = "Rinkel (Renegado)"
    Ciudades(eCiudad.cSuramei) = "Suramei (Republica)"
    
    ListaRazas(eRaza.HUMANO) = "Humano"
    ListaRazas(eRaza.ELFO) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    ListaRazas(eRaza.Orco) = "Orco"


    ListaClases(eClass.MAGO) = "Mago"
    ListaClases(eClass.Clerigo) = "Clerigo"
    ListaClases(eClass.Guerrero) = "Guerrero"
    ListaClases(eClass.Asesino) = "Asesino"
    ListaClases(eClass.Ladron) = "Ladron"
    ListaClases(eClass.Bardo) = "Bardo"
    ListaClases(eClass.DRUIDA) = "Druida"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.CAZADOR) = "Cazador"
    ListaClases(eClass.Mercenario) = "Mercenario"
    ListaClases(eClass.Nigromante) = "Nigromante"
    ListaClases(eClass.Gladiador) = "Gladiador"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar arboles"
    SkillsNames(eSkill.Comercio) = "Comercio"
    SkillsNames(eSkill.DefensaEscudos) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
    SkillsNames(eSkill.Artes) = "Artes Marciales"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    SkillsNames(eSkill.Alquimia) = "Alquimia"
    SkillsNames(eSkill.Arrojadizas) = "Armas Arrojadizas"
    SkillsNames(eSkill.Botanica) = "Botanica"
    SkillsNames(eSkill.Equitacion) = "Equitacion"
    SkillsNames(eSkill.Musica) = "Musica"
    SkillsNames(eSkill.Resistencia) = "Resistencia Magica"
    SkillsNames(eSkill.Sastreria) = "Sastreria"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
    
     ReDim Head_Range(1 To NUMRAZAS) As tHeadRange
    
'Male heads
Head_Range(HUMANO).mStart = 1
Head_Range(HUMANO).mEnd = 30
Head_Range(Enano).mStart = 301
Head_Range(Enano).mEnd = 315
Head_Range(ELFO).mStart = 101
Head_Range(ELFO).mEnd = 121
Head_Range(ElfoOscuro).mStart = 202
Head_Range(ElfoOscuro).mEnd = 212
Head_Range(Gnomo).mStart = 401
Head_Range(Gnomo).mEnd = 409
Head_Range(Orco).mStart = 501
Head_Range(Orco).mEnd = 514

'Female heads
Head_Range(HUMANO).fStart = 70
Head_Range(HUMANO).fEnd = 80
Head_Range(Enano).fStart = 370
Head_Range(Enano).fEnd = 373
Head_Range(ELFO).fStart = 170
Head_Range(ELFO).fEnd = 189
Head_Range(ElfoOscuro).fStart = 270
Head_Range(ElfoOscuro).fEnd = 278
Head_Range(Gnomo).fStart = 470
Head_Range(Gnomo).fEnd = 481
Head_Range(Orco).fStart = 570
Head_Range(Orco).fEnd = 573


End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call RemoveAllDialogs
End Sub

Public Sub Auto_Drag(ByVal hWnd As Long)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
End Sub

Public Sub CloseClient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    ' Allow new instances of the client to be opened
    Call modPrevInstance.ReleaseInstance
    
    EngineRun = False
    frmCargando.Show

    Call modResolution.ResetResolution
    
    'Destruimos los objetos públicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    Call UnloadAllForms
    
    End
End Sub

Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, ByVal PartPos As Byte, Optional ByVal particle_life As Long = 0) As Long

On Error Resume Next

If ParticulaInd <= 0 Then Exit Function

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

'General_Char_Particle_Create = engine.Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), PartPos, StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gr, StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)

End Function

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0, Optional ByVal OffsetX As Integer, Optional ByVal OffsetY As Integer) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Particle_Create = Engine.Particle_Group_Create(X, Y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1 + OffsetX, StreamData(ParticulaInd).y1 + OffsetY, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)

End Function

Public Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempint As Integer
Dim f As Integer

ReDim GrhData(0 To CANT_GRH_INDEX) As GrhData
f = FreeFile()
Open App.path & "\Init\Graphics.ind" For Binary Access Read As #f
    
    Seek #f, 1
    
    Get #f, , tempint
    Get #f, , tempint
    Get #f, , tempint
    Get #f, , tempint
    Get #f, , tempint

    'Get first Grh Number
    Get #f, , Grh
    
    Do Until Grh <= 0
        'Get number of frames
        Get #f, , GrhData(Grh).NumFrames
        
        If GrhData(Grh).NumFrames <= 0 Then
            GoTo ErrorHandler
        End If
        
        ReDim GrhData(Grh).Frames(1 To GrhData(Grh).NumFrames)
        
        If GrhData(Grh).NumFrames > 1 Then
        
            'Read a animation GRH set
            For Frame = 1 To GrhData(Grh).NumFrames
                Get #f, , GrhData(Grh).Frames(Frame)
                If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > CANT_GRH_INDEX Then GoTo ErrorHandler
            Next Frame
        
            Get #f, , tempint
            
            If tempint <= 0 Then GoTo ErrorHandler
            GrhData(Grh).speed = CLng(tempint)
            
            'Compute width and height
            GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
            
            GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler

            GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
            If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler

            GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
            If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
        Else
            'Read in normal GRH data
            Get #f, , GrhData(Grh).FileNum
            If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler

            Get #f, , GrhData(Grh).sX
            If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
            
            Get #f, , GrhData(Grh).sY
            If GrhData(Grh).sY < 0 Then GoTo ErrorHandler

            Get #f, , GrhData(Grh).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler

            Get #f, , GrhData(Grh).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler

            'Compute width and height
            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / 32
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / 32
            
            GrhData(Grh).Frames(1) = Grh
        End If
        'Get Next Grh Number
        Get #f, , Grh
    Loop
    
Close #f

    Dim Count As Long
    Open App.path & "\Init\minimap.dat" For Binary As #1
        Seek #1, 1
        For Count = 1 To CANT_GRH_INDEX
            If Grh_Check(Count) Then
                Get #1, , GrhData(Count).mini_map_color
            End If
        Next Count
    Close #1
    
Exit Sub

ErrorHandler:
    Close #f
    MsgBox "Error al cargar el recurso de índice de gráficos: " & Err.Description & " (" & Grh & ")", vbCritical, "Error al cargar"

End Sub

Private Function Grh_Check(ByVal grh_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= CANT_GRH_INDEX Then
        Grh_Check = GrhData(grh_index).NumFrames
    End If
End Function

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo

    N = FreeFile()
    Open App.path & "\Init\Characters.ind " For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
    
End Sub
Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\Init\Heads.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
    
End Sub
Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\Init\Helmets.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
    
    
End Sub
Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer

    N = FreeFile()
    Open App.path & "\Init\fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
     
End Sub
Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim Leer As New clsIniReader
    
    Leer.Initialize App.path & "\Init\weapons.dat"
    
    NumWeaponAnims = Val(Leer.GetValue("INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(Leer.GetValue("ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(Leer.GetValue("ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(Leer.GetValue("ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(Leer.GetValue("ARMA" & loopc, "Dir4")), 0
    Next loopc
    
    Set Leer = Nothing
    
    
End Sub
Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim Leer As New clsIniReader

    Leer.Initialize App.path & "\Init\Shields.dat"
    
    NumEscudosAnims = Val(Leer.GetValue("INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(Leer.GetValue("ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(Leer.GetValue("ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(Leer.GetValue("ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(Leer.GetValue("ESC" & loopc, "Dir4")), 0
    Next loopc
    
    Set Leer = Nothing
    
    
End Sub

Sub SwitchMap(ByVal MapRoute As String, Optional Client_Mode As Boolean = False)

Engine.Char_Clean
Engine.Particle_Group_Remove_All
Engine.Light_Remove_All


On Error GoTo ErrorHandler

Dim fh As Integer
Dim MH As tMapHeader
Dim Blqs() As tDatosBloqueados
Dim L1() As Long
Dim L2() As tDatosGrh
Dim L3() As tDatosGrh
Dim L4() As tDatosGrh
Dim Triggers() As tDatosTrigger
Dim Luces() As tDatosLuces
Dim Particulas() As tDatosParticulas
Dim Objetos() As tDatosObjs
Dim NPCs() As tDatosNPC
Dim TEs() As tDatosTE

Dim i As Long
Dim j As Long

Extract_File map, App.path & "\resources\maps\", "mapa" & MapRoute & ".csm", App.path & "\resources\maps\"

fh = FreeFile
Open App.path & "\resources\maps\mapa" & MapRoute & ".csm" For Binary Access Read As fh
    Get #fh, , MH
    Get #fh, , MapSize
    Get #fh, , MapDat
    
    ReDim MapData(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As MapBlock
    ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
    
    Get #fh, , L1
    
    With MH
        If .NumeroBloqueados > 0 Then
            ReDim Blqs(1 To .NumeroBloqueados)
            Get #fh, , Blqs
            For i = 1 To .NumeroBloqueados
                MapData(Blqs(i).X, Blqs(i).Y).Blocked = 1
            Next i
        End If
        
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2
            For i = 1 To .NumeroLayers(2)
                InitGrh MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).grhindex
            Next i
        End If
        
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3
            For i = 1 To .NumeroLayers(3)
                InitGrh MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).grhindex
            Next i
        End If
        
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4
            For i = 1 To .NumeroLayers(4)
                InitGrh MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).grhindex
            Next i
        End If
        
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers
            For i = 1 To .NumeroTriggers
                MapData(Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
            Next i
        End If
        
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas
            For i = 1 To .NumeroParticulas
                MapData(Particulas(i).X, Particulas(i).Y).particle_group_index = General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).Y)
            Next i
        End If
        
        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces
            For i = 1 To .NumeroLuces
                Call Engine.Light_Create(Luces(i).X, Luces(i).Y, D3DColorXRGB(Luces(i).color.r, Luces(i).color.g, Luces(i).color.b), Luces(i).range)
            Next i
        End If
        
    End With

Close fh


For j = MapSize.YMin To MapSize.YMax
    For i = MapSize.XMin To MapSize.XMax
        If L1(i, j) > 0 Then
            InitGrh MapData(i, j).Graphic(1), L1(i, j)
        End If
    Next i
Next j

Dim r As Integer, g As Integer, b As Integer
'Common light value verify
If MapDat.base_light = 0 Then
   base_light = -1
Else
   General_Long_Color_to_RGB MapDat.base_light, r, g, b
   base_light = D3DColorXRGB(r, g, b)
End If

If Not UserMinHP = 0 Then
Call Engine.setup_ambient
Else
   General_Long_Color_to_RGB MapDat.base_light, 255, 0, 0
   base_light = D3DColorXRGB(255, 0, 0)
End If

'*******************************
'Render lights
Engine.Light_Render_All
'*******************************
Debug.Print "mapa" & MapRoute & ":" & MapDat.extra1
Debug.Print "mapa" & MapRoute & ":" & MapDat.extra2
Debug.Print "mapa" & MapRoute & ":" & MapDat.zone
Debug.Print "mapa" & MapRoute & ":" & MapDat.battle_mode
Debug.Print "mapa" & MapRoute & ":" & MapDat.terrain
frmMain.MiniMap.Cls

    Dim map_x As Long
    Dim map_y As Long
    Dim screen_x As Long
    Dim screen_y As Long
    Dim grh_index As Long
    
    '*********************
    'Layer 1 (Floor level) and walls
    '*********************
    For map_y = 1 To 100
        For map_x = 1 To 100
            screen_x = (map_x - 1)
            screen_y = (map_y - 1)
            '*** Start Layer 1 ***
            If MapData(map_x, map_y).Graphic(1).grhindex Then
                grh_index = MapData(map_x, map_y).Graphic(1).grhindex
                SetPixel frmMain.MiniMap.hdc, map_x, map_y, GrhData(grh_index).mini_map_color
            End If
            '*** End Layer 1 ***
        Next map_x
    Next map_y
    
    For map_y = 1 To 100
        For map_x = 1 To 100
            screen_x = (map_x - 1)
            screen_y = (map_y - 1)
            '*** Start Layer 2 ***
            If MapData(map_x, map_y).Graphic(2).grhindex Then
                grh_index = MapData(map_x, map_y).Graphic(2).grhindex
                SetPixel frmMain.MiniMap.hdc, map_x, map_y, GrhData(grh_index).mini_map_color
            End If
            '*** End Layer 2 ***
        Next map_x
    Next map_y

    For map_y = 1 To 100
        For map_x = 1 To 100
            screen_x = (map_x - 1)
            screen_y = (map_y - 1)
            '*** Start Layer 2 ***
            If MapData(map_x, map_y).Graphic(4).grhindex Then
                grh_index = MapData(map_x, map_y).Graphic(4).grhindex
                SetPixel frmMain.MiniMap.hdc, map_x, map_y, GrhData(grh_index).mini_map_color
            End If
            '*** End Layer 2 ***
        Next map_x
    Next map_y
MapDat.map_name = Trim$(MapDat.map_name)

MapName = MapDat.map_name
Call Audio.Music_Load(MapDat.music_number)

Exit Sub

ErrorHandler:
    If fh <> 0 Then Close fh
End Sub


Public Sub InfoMapa()
    If InfoMapAct = True Then
        frmMain.Coord.Caption = "Posición: " & UserMap & ", " & UserPos.X & ", " & UserPos.Y
    Else
        If Not MapName = "" Then
            frmMain.Coord.Caption = Trim$(MapName)
        Else
            frmMain.Coord.Caption = "Mapa Desconocido"
        End If
    End If
End Sub

Public Sub Make_Transparent_Richtext(ByVal hWnd As Long)

If Win2kXP Then _
    Call SetWindowLong(hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

End Sub

Public Function General_Windows_Is_2000XP() As Boolean
'**************************************************************
'Author: Unknown
'Last Modify Date: Unknown
'Get the windows version
'**************************************************************
On Error GoTo ErrorHandler

Dim RetVal As Long

OSInfo.dwOSVersionInfoSize = Len(OSInfo)
RetVal = GetOSVersion(OSInfo)

If OSInfo.dwPlatformId = VER_PLATFORM_WIN32_NT And OSInfo.dwMajorVersion >= 5 Then
    General_Windows_Is_2000XP = True
Else
    General_Windows_Is_2000XP = False
End If

Exit Function

ErrorHandler:
    General_Windows_Is_2000XP = False

End Function

Public Sub EndGame()
    prgRun = False
    
    'Cerramos los forms y nos vamos
    Call UnloadAllForms
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''clsMeteorologic''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Get_Time_String() As String

Get_Time_String = mid(time, 1, 5) & "... "

Select Case Hour(time)
    Case 5, 6, 7
        Get_Time_String = Get_Time_String & "el sol se asoma lentamente en el horizonte"
    Case 8, 9, 10, 11, 12, 13, 14, 15, 16, 17
        Get_Time_String = Get_Time_String & "¡no pierdas el tiempo!"
    Case 18, 19
        Get_Time_String = Get_Time_String & "lentamente el dia termina"
    Case Else
        Get_Time_String = Get_Time_String & "¿despierto a estas horas? ¡no olvides visitar El Mesón Hostigado!"
End Select

End Function
Public Sub CargarParticulas()
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim Leer As New clsIniReader

    Dim StreamFile As String

    Extract_File Scripts, App.path & "\Init", "Particles.ini", App.path & "\Init\"

    StreamFile = App.path & "\Init\Particles.ini"
    
    Leer.Initialize StreamFile

    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).name = Leer.GetValue(Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = Leer.GetValue(Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = Leer.GetValue(Val(loopc), "X1")
        StreamData(loopc).y1 = Leer.GetValue(Val(loopc), "Y1")
        StreamData(loopc).x2 = Leer.GetValue(Val(loopc), "X2")
        StreamData(loopc).y2 = Leer.GetValue(Val(loopc), "Y2")
        StreamData(loopc).angle = Leer.GetValue(Val(loopc), "Angle")
        StreamData(loopc).vecx1 = Leer.GetValue(Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = Leer.GetValue(Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = Leer.GetValue(Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = Leer.GetValue(Val(loopc), "VecY2")
        StreamData(loopc).life1 = Leer.GetValue(Val(loopc), "Life1")
        StreamData(loopc).life2 = Leer.GetValue(Val(loopc), "Life2")
        StreamData(loopc).friction = Leer.GetValue(Val(loopc), "Friction")
        StreamData(loopc).spin = Leer.GetValue(Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = Leer.GetValue(Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = Leer.GetValue(Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = Leer.GetValue(Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = Leer.GetValue(Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = Leer.GetValue(Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = Leer.GetValue(Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = Leer.GetValue(Val(loopc), "XMove")
        StreamData(loopc).YMove = Leer.GetValue(Val(loopc), "YMove")
        StreamData(loopc).move_x1 = Leer.GetValue(Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = Leer.GetValue(Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = Leer.GetValue(Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = Leer.GetValue(Val(loopc), "move_y2")
        StreamData(loopc).life_counter = Leer.GetValue(Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(Leer.GetValue(Val(loopc), "Speed"))
        
        Dim temp As Integer
        temp = Leer.GetValue(Val(loopc), "resize")
        
        StreamData(loopc).grh_resize = IIf((temp = -1), True, False)
        StreamData(loopc).grh_resizex = Leer.GetValue(Val(loopc), "rx")
        StreamData(loopc).grh_resizey = Leer.GetValue(Val(loopc), "ry")
        
        'Ai ya tenemos un problema de auras
        'tas? si qe paso, nesesito las auras de mi cumpu se pueden pasar por aca?
        ' cuanto pesan? nada osea es particles.ind dije auras jaja mira dale aca
    
        StreamData(loopc).NumGrhs = Leer.GetValue(Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = Leer.GetValue(Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = Field_Read(str(i), GrhListing, ",")
        Next i
        
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        
        For ColorSet = 1 To 4
            TempSet = Leer.GetValue(Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = Field_Read(1, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).g = Field_Read(2, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).b = Field_Read(3, TempSet, ",")
        Next ColorSet

                
    Next loopc
    
    Set Leer = Nothing
    

End Sub

Public Sub RemoveDialog(ByVal CharIndex As Integer)
If charlist(CharIndex).dialog_life > 0 Then charlist(CharIndex).dialog = ""
charlist(CharIndex).dialog_life = 0
charlist(CharIndex).dialog_offset_counter_y = 0
End Sub

Public Sub RemoveAllDialogs()
Dim i As Long
For i = 1 To LastChar
    If charlist(i).dialog <> "" Then
        Engine.Char_Dialog_Set i, "", 0, 0
    End If
Next i
End Sub

Public Function General_RGB_Color_to_Long(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal a As Long) As Long
        
    Dim c As Long
        
    If a > 127 Then
        a = a - 128
        c = a * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    Else
        c = a * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    End If
    
    General_RGB_Color_to_Long = c

End Function

Public Sub General_Var_Write(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, file
End Sub

Public Function Map_NameLoad(ByVal map_num As Integer) As String

On Error GoTo ErrorHandler

If FileExist(App.path & "\resources\maps\mapa" & map_num & ".csm", vbNormal) Then
    SwitchMap map_num
    Map_NameLoad = MapDat.map_name
    If LenB(Map_NameLoad) = 0 Then
        Map_NameLoad = "Mapa Desconocido"
    End If
Else
    Map_NameLoad = "Mapa Desconocido"
End If

Exit Function

ErrorHandler:
    Map_NameLoad = "Mapa Desconocido"

End Function

Public Sub General_Long_Color_to_RGB(ByVal long_color As Long, ByRef red As Integer, ByRef green As Integer, ByRef blue As Integer)
'***********************************
'Coded by Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 2/19/03
'Takes a long value and separates RGB values to the given variables
'***********************************
    Dim temp_color As String
    
    temp_color = Hex(long_color)
    If Len(temp_color) < 6 Then
        'Give is 6 digits for easy RGB conversion.
        temp_color = String(6 - Len(temp_color), "0") + temp_color
    End If
    
    red = CLng("&H" + mid$(temp_color, 1, 2))
    green = CLng("&H" + mid$(temp_color, 3, 2))
    blue = CLng("&H" + mid$(temp_color, 5, 2))
End Sub

Public Function General_Get_Mouse_Speed() As Long
'**************************************************************
'Author: Gaston
'Last Modify Date: 27/09/2015
'
'**************************************************************
 
SystemParametersInfo SPI_GETMOUSESPEED, 0, General_Get_Mouse_Speed, 0
 
End Function
 
Public Sub General_Set_Mouse_Speed(ByVal lngSpeed As Long)
'**************************************************************
'Author: Gaston
'Last Modify Date: 27/09/2015
'
'**************************************************************
 
SystemParametersInfo SPI_SETMOUSESPEED, 0, ByVal lngSpeed, 0
 
End Sub

