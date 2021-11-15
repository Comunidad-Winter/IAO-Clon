Attribute VB_Name = "modTileEngine"
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

Public map_base_light As Long

Public Engine As New clsDX8Engine

Public DX8 As DirectX8
Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8
Public D3DX As D3DX8

Public vertList(3) As TLVERTEX
Public SurfaceDB As clsTexManager

Public Type D3D8Textures
    Texture As Direct3DTexture8
    texwidth As Long
    texheight As Long
End Type

Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Public Const PI As Single = 3.14159265358979

Public base_light As Long

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Private Const GrhFogata As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames() As Integer
    speed As Single 'Integer
    mini_map_color As Long
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    grhindex As Integer
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.north To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.north To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.north To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.north To E_Heading.WEST) As Grh
End Type

Type tAura
    Grh As Grh
    color As Long
End Type

'Apariencia del personaje
Public Type Char
    MinVida As Long
    MaxVida As Long
    active As Byte
    Heading As E_Heading
    Pos As Position
    
    label_color(3) As Long
    
    iHead As Integer
    iBody As Integer
    body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    ShieldOffSetY As Integer
    
    plusGrh(2) As tAura
    
    fX As Grh
    FxIndex As Integer
    
        AlphaX As Integer
    last_tick As Long
    
    Criminal As Byte
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    Priv As Byte
    
    dialog As String
    dialog_color As Long
    dialog_life As Byte
    dialog_font_index As Integer
    dialog_offset_counter_y As Single
    dialog_scroll As Boolean
    
    group_index As Integer
    
    particle_count As Integer
    particle_group() As Long
End Type

'Info de un objeto
Public Type obj
    TieneLuz As Byte
    OBJIndex As Integer
    Amount As Integer
End Type

Public AmbientColor As D3DCOLORVALUE

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    light_value(3) As Long
    particle_group_index As Long
    
    NpcIndex As Integer
    OBJInfo As obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

'DX7 Objects
'Public DirectX As New DirectX7
'Public DirectDraw As DirectDraw7
'Private PrimarySurface As DirectDrawSurface7
'Private PrimaryClipper As DirectDrawClipper
'Private BackBufferSurface As DirectDrawSurface7

Public IniPath As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'********************************************
'*************Configuracion******************
'********************************************
Public Sound As Byte
Public Music As Byte
Public EffectSound As Byte
Public VolumeSound As Integer
Public VolumeMusic As Integer
Public cSombras As Byte
Public cTechos As Byte
Public cLimitarFps As Byte
Public cObjName As Byte
'********************************************
'*************/Configuracion*****************
'********************************************

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer
Private MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer


Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte




'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public charlist(1 To 10000) As Char

#If SeguridadAlkon Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

' Used by GetTextExtentPoint32
Private Type size
    cx As Long
    cy As Long
End Type

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?




'Added by Juan Martín Sotuyo Dodero
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'Added by Barrin


'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As size) As Long

'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, BytesTotal As Currency, FreeBytesTotal As Currency) As Long

'To get free bytes in RAM

Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type



Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ 32 - frmMain.Renderer.ScaleWidth \ 64
    tY = UserPos.Y + viewPortY \ 32 - frmMain.Renderer.ScaleHeight \ 64
    Debug.Print tX; tY
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        Call Engine.Char_Particle_Group_Remove_All(CharIndex)
        .active = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False
        .Moving = 0
        .muerto = False
        .Nombre = ""
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
        .dialog = ""
        .dialog_life = 0
        .dialog_scroll = False
    End With
End Sub
Sub MakeChar(ByVal CharIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = body
        .Head = HeadData(Head)
        .body = BodyData(body)

        If Not Arma = 29 Then .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .active = 1
        
    Select Case .Priv
        Case 1 'Gris
            Engine.Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(128, 128, 128)
        Case 2 'Azul
            Engine.Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(0, 0, 176)
        Case 3 'Verde
            Engine.Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(0, 128, 0)
        Case 4 'Verde
            Engine.Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(0, 128, 0)
        Case 5 'Naranja
            Engine.Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(255, 128, 0)
        Case 6 'Armada Real
            Engine.Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(40, 181, 159)
        Case 7 'Rojo
            Engine.Long_To_RGB_List charlist(CharIndex).label_color, D3DColorXRGB(210, 0, 0)
    End Select
    End With
    
    Call PonerAura(CharIndex, Escudo, Arma, body)
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
End Sub

Sub PonerAura(ByVal CharIndex As Integer, ByVal Escudo As Byte, ByVal Arma As Byte, ByVal body As Integer)
With charlist(CharIndex)
    If body = 255 Then
        InitGrh .plusGrh(2).Grh, 20206
        .plusGrh(2).color = &HFFFD7E
    Else
        .plusGrh(2).Grh.grhindex = 0
    End If

    If Escudo = 27 Then
        InitGrh .plusGrh(1).Grh, 20203
        .plusGrh(1).color = &HFFCC33
    Else
        .plusGrh(1).Grh.grhindex = 0
    End If

    If Arma = 23 Then
        InitGrh .plusGrh(0).Grh, 20128
        .plusGrh(0).color = &HFFCC33
    ElseIf Arma = 24 Then
        InitGrh .plusGrh(0).Grh, 20133
        .plusGrh(0).color = &HFF3300
    ElseIf Arma = 25 Then
        InitGrh .plusGrh(0).Grh, 20152
        .plusGrh(0).color = &HFF0000
    ElseIf Arma = 26 Then
        InitGrh .plusGrh(0).Grh, 20185
        .plusGrh(0).color = -65536
    ElseIf Arma = 31 Then
        InitGrh .plusGrh(0).Grh, 20155
        .plusGrh(0).color = &HFF0000
    ElseIf Arma = 21 Then
        InitGrh .plusGrh(0).Grh, 20151
        .plusGrh(0).color = &HFFFF00
    ElseIf Arma = 28 Then
        InitGrh .plusGrh(0).Grh, 20148
        .plusGrh(0).color = &HFF
    ElseIf Arma = 29 Then
        InitGrh .plusGrh(0).Grh, 20146
        .plusGrh(0).color = &H6B1B
    ElseIf Arma = 30 Then
        InitGrh .plusGrh(0).Grh, 20200
        .plusGrh(0).color = &HCCFF33
    ElseIf Arma = 32 Then
        InitGrh .plusGrh(0).Grh, 20147
        .plusGrh(0).color = &HFF
    Else
        .plusGrh(0).Grh.grhindex = 0
    End If
    
    If body = 291 Then
        .ShieldOffSetY = 30
    ElseIf body = 415 Or body = 384 Or body = 382 Then
        .ShieldOffSetY = 16
    ElseIf body = 416 Then
        .ShieldOffSetY = 32
    ElseIf body = 282 Or body = 292 Then
        .ShieldOffSetY = 20
    ElseIf body = 317 Or body = 292 Then
        .ShieldOffSetY = 20
    ElseIf body = 381 Or body = 383 Then
        .ShieldOffSetY = 24
    Else
        .ShieldOffSetY = 0
    End If
    
    If BodyData(body).HeadOffset.Y = -28 Then
        .ShieldOffSetY = .ShieldOffSetY - 5
    End If
    
End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    charlist(CharIndex).active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call RemoveDialog(CharIndex)
    
    'Reset char
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal grhindex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.grhindex = grhindex
    
    If Started = 2 Then
        If GrhData(Grh.grhindex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.grhindex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.speed = GrhData(Grh.grhindex).speed
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With charlist(CharIndex)
            If Not .muerto And EstaPCarea(CharIndex) Then
                .pie = Not .pie
        'Si esta en una superficie de pasto?
                If MapData(.Pos.X, .Pos.Y).Graphic(1).grhindex >= 6000 And MapData(.Pos.X, .Pos.Y).Graphic(1).grhindex <= 6559 Then
                    If .pie Then
                        Call Audio.PlayWave(SND_PASOS3)
                    Else
                        Call Audio.PlayWave(SND_PASOS4)
                    End If
            'Si esta en una superficie de Arena?
            ElseIf MapData(.Pos.X, .Pos.Y).Graphic(1).grhindex >= 7700 And MapData(.Pos.X, .Pos.Y).Graphic(1).grhindex <= 7719 Then
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS5)
                Else
                    Call Audio.PlayWave(SND_PASOS6)
                End If
            'Si esta en una superficie de Nieve?
            ElseIf MapData(.Pos.X, .Pos.Y).Graphic(1).grhindex >= 7379 And MapData(.Pos.X, .Pos.Y).Graphic(1).grhindex <= 7507 Then
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS7)
                Else
                    Call Audio.PlayWave(SND_PASOS8)
                End If
            Else
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS1)
                Else
                    Call Audio.PlayWave(SND_PASOS2)
                End If
            End If

    'Feo este Sistema****************************
    If UserNavegando Then
    'TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Audio.PlayWave(SND_NAVEGANDO)
    End If
    '********************************************
 
    End If
    End With
    End If
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addx = nX - X
        addy = nY - Y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.east
        End If
        
        If Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        End If
        
        If Sgn(addy) = -1 Then
            nHeading = E_Heading.north
        End If
        
        If Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then Call RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.north
            Y = -1
        
        Case E_Heading.east
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
         AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4 Or _
                MapData(UserPos.X, UserPos.Y).Trigger >= 20, True, False)
    End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.grhindex = GrhFogata Then
                    location.X = j
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopc As Long
    Dim Dale As Boolean
    
    loopc = 1
    Do While charlist(loopc).active And Dale
        loopc = loopc + 1
        Dale = (loopc <= UBound(charlist))
    Loop
    
    NextOpenChar = loopc
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.


Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    If UserMontando = True Then
        If MapData(X, Y).Trigger = 1 Or MapData(X, Y).Trigger = 2 Or MapData(X, Y).Trigger = 4 Or MapData(X, Y).Trigger >= 20 Then
            Exit Function
        End If
    End If
    
    LegalPos = True
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Public Function General_Bytes_To_Megabytes(Bytes As Double) As Double
Dim dblAns As Double
dblAns = (Bytes / 1024) / 1024
General_Bytes_To_Megabytes = format(dblAns, "###,###,##0.00")
End Function

Public Function General_Get_Free_Ram() As Double
    'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    General_Get_Free_Ram = General_Bytes_To_Megabytes(dblAns)
End Function

Public Function General_Get_Free_Ram_Bytes() As Long
    GlobalMemoryStatus pUdtMemStatus
    General_Get_Free_Ram_Bytes = pUdtMemStatus.dwAvailPhys
End Function

Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
    Dim BMHeader As BITMAPFILEHEADER
    Dim BINFOHeader As BITMAPINFOHEADER
    
    Open BmpFile For Binary Access Read As #1
    
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    
    Close #1
    
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight
End Function

Public Sub Grh_Render_To_Hdc(ByVal desthDC As Long, grh_index As Integer, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional transparent As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/30/2004
'This method is SLOW... Don't use in a loop if you care about
'speed!
'Modified by Juan Martín Sotuyo Dodero
'*************************************************************
    
    On Error GoTo ErrorHandler
    
    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long

    If grh_index <= 0 Then Exit Sub

    'If it's animated switch grh_index to first frame
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If

        file_path = App.path & "\Resources\graphics\" & GrhData(grh_index).FileNum & ".bmp"
        
        src_x = GrhData(grh_index).sX
        src_y = GrhData(grh_index).sY
        src_width = GrhData(grh_index).pixelWidth
        src_height = GrhData(grh_index).pixelHeight
            
        hdcsrc = CreateCompatibleDC(desthDC)
        PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
        
        If transparent = False Then
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
        Else
            MaskDC = CreateCompatibleDC(desthDC)
            
            PrevObj2 = SelectObject(MaskDC, LoadPicture(file_path))
            
            Grh_Create_Mask hdcsrc, MaskDC, src_x, src_y, src_width, src_height
            
            'Render tranparently
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, MaskDC, src_x, src_y, vbSrcAnd
            BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcPaint
            
            Call DeleteObject(SelectObject(MaskDC, PrevObj2))
            
            DeleteDC MaskDC
        End If
        
        Call DeleteObject(SelectObject(hdcsrc, PrevObj))
        DeleteDC hdcsrc
        
   
    
    
    Exit Sub
    
ErrorHandler:

    
End Sub

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal grhindex As Integer) As Boolean
    If grhindex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(grhindex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.X <= X + (GrhData(grhindex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(grhindex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.Y <= Y
    End If
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    'Call SurfaceDB.Initialize(DirectDraw, ClientSetup.bUseVideo, DirGraficos, ClientSetup.byMemory)
    
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).bottom = 128: RLluvia(1).bottom = 128: RLluvia(2).bottom = 128: RLluvia(3).bottom = 128
    
    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).bottom = 256: RLluvia(5).bottom = 256: RLluvia(6).bottom = 256: RLluvia(7).bottom = 256
    
    'We are done!
    'Saco esto porque el texto del cargar queda horrible
    'AddtoRichTextBox frmCargando.status, "Hecho.", , , , 1, , False
End Sub

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function


Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If
    End With
End Sub

Private Sub Grh_Create_Mask(ByRef hdcsrc As Long, ByRef MaskDC As Long, ByVal src_x As Integer, ByVal src_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/30/2004
'Creates a Mask hDC, and sets the source hDC to work for trans bliting.
'**************************************************************
    Dim X As Integer
    Dim Y As Integer
    Dim TransColor As Long
    Dim ColorKey As String
    
    'ColorKey = hex(COLOR_KEY)
    
    'Check if it has an alpha component
    'If Len(ColorKey) > 6 Then
         'get rid of alpha
    '    ColorKey = "&H" & Right$(ColorKey, 6)
    'End If
    'piluex prueba
    'TransColor = Val(ColorKey)
    ColorKey = "0"
    TransColor = &H0

    'Make it a mask (set background to black and foreground to white)
    'And set the sprite's background white
    For Y = src_y To src_height + src_y
        For X = src_x To src_width + src_x
            If GetPixel(hdcsrc, X, Y) = TransColor Then
                SetPixel MaskDC, X, Y, vbWhite
                SetPixel hdcsrc, X, Y, vbBlack
            Else
                SetPixel MaskDC, X, Y, vbBlack
            End If
        Next X
    Next Y
End Sub

'**************************************************************
'MiniMapa
Public Sub ActualizarMiniMapa(ByVal tHeading As E_Heading)
'Esta es la forma mas optima que se me ha ocurrido. Solo dibuja una vez.
    frmMain.UserM.Left = UserPos.X - 1
    frmMain.UserM.Top = UserPos.Y - 1
End Sub
Public Sub DibujarMiniMapa()

'Si el usuario esta en piramide, no dibujamos el minimapa
If UserMap = 12 Or UserMap = 13 Then
    frmMain.MiniMap.Cls
    Exit Sub
End If
   
    frmMain.MiniMap.Refresh
    Call ActualizarMiniMapa(0)
End Sub

