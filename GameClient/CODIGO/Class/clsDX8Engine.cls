VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX8Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'MOTOR GRÁFICO ESCRITO(mayormente) POR MENDUZ@NOICODER.COM
Option Explicit

Public bRunning As Boolean

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Private Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2

Private Type tFont
    font_size As Integer
    ascii_code(0 To 255) As Integer 'indice de cada letra
End Type

Private font_types(1 To 2) As tFont

Dim font_count As Long
Dim font_last As Long

Private font_list() As D3DXFont

Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

Dim Texture As Direct3DTexture8
Dim TransTexture As Direct3DTexture8

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public FPS As Integer
Private FramesPerSecCounter As Integer
Private timerElapsedTime As Single
Private timerTicksPerFrame As Double

Public engineBaseSpeed As Single

Private lFrameTimer As Long
Private lFrameLimiter As Long
Private ScrollPixelsPerFrameX As Byte
Private ScrollPixelsPerFrameY As Byte

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

Private MainViewTop As Integer
Private MainViewLeft As Integer
Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte
Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer
Private Const GrhFogata As Integer = 1521

Private Type Light
    active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Integer 'Coordinates
    map_y As Integer
    color As Long 'Start colour
    range As Byte
End Type

'Light list
Dim light_list() As Light
Dim light_count As Long
Dim light_last As Long

Private Type Particle
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Private Type particle_group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    live As Long
    liv1 As Integer
    liveend As Long
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    
    'Added by Juan Martín Sotuyo Dodero
    speed As Single
    life_counter As Long
    
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

'Particle system
Dim particle_group_list() As particle_group
Dim particle_group_count As Long
Dim particle_group_last As Long
#Const HARDCODED = False 'True ' == MÁS FPS ^^

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

Public Sub Text_Render(ByVal font As D3DXFont, ByVal Text As String, ByVal Top As Long, ByVal Left As Long, _
                                ByVal Width As Long, ByVal Height As Long, ByVal color As Long, ByVal format As Long, Optional ByVal shadow As Boolean = False)
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    Dim TextRect As RECT
    Dim ShadowRect As RECT
    
    TextRect.Top = Top
    TextRect.Left = Left
    TextRect.bottom = Top + Height
    TextRect.Right = Left + Width
    
    If shadow Then
        ShadowRect.Top = Top - 1
        ShadowRect.Left = Left - 2
        ShadowRect.bottom = (Top + Height) - 1
        ShadowRect.Right = (Left + Width) - 2
        D3DX.DrawText font, &HFF000000, Text, ShadowRect, format
    End If
    
    D3DX.DrawText font, color, Text, TextRect, format
End Sub
Public Sub Text_Render_ext(ByVal Text As String, ByVal Top As Long, ByVal Left As Long, _
                                ByVal Width As Long, ByVal Height As Long, ByVal color As Long, Optional ByVal shadow As Boolean = False, Optional ByVal center As Boolean = False)
    If center = True Then
        Call Text_Render(font_list(1), Text, Top, Left, Width, Height, color, fa_center, shadow)
    Else
        Call Text_Render(font_list(1), Text, Top, Left, Width, Height, color, DT_TOP Or DT_LEFT, shadow)
    End If
End Sub

Private Sub Font_Make(ByVal font_index As Long, ByVal style As String, ByVal bold As Boolean, _
                        ByVal italic As Boolean, ByVal size As Long)
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    
    Dim font_desc As IFont
    Dim fnt As New StdFont
    fnt.name = style
    fnt.size = size
    fnt.bold = bold
    fnt.italic = italic
    
    Set font_desc = fnt
    Set font_list(font_index) = D3DX.CreateFont(D3DDevice, font_desc.hFont)
End Sub


Public Function Font_Create(ByVal style As String, ByVal size As Long, ByVal bold As Boolean, _
                            ByVal italic As Boolean) As Long
On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    Font_Make Font_Create, style, bold, italic, size
ErrorHandler:
    Font_Create = 0
End Function

Private Function Font_Next_Open() As Long
    Font_Next_Open = font_last + 1
End Function

Private Function Font_Check(ByVal font_index As Long) As Boolean
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True
    End If
End Function

Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
  MakeVector.X = X
  MakeVector.Y = Y
  MakeVector.Z = Z
End Function

Public Sub Engine_Init()
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
'On Error GoTo ErrHandler:

    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim ColorKeyVal As Long
    
    Set SurfaceDB = New clsTexManager
    
    Set DX8 = New DirectX8
    Set D3D = DX8.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.format
        .BackBufferWidth = frmMain.Renderer.ScaleWidth
        .BackBufferHeight = frmMain.Renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.Renderer.hWnd
    End With
    
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.Renderer.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                            D3DWindow)
                                                            
                                                            
    HalfWindowTileHeight = (frmMain.Renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.Renderer.ScaleWidth / 32) \ 2
    
    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    
    D3DDevice.SetVertexShader FVF
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)

    engineBaseSpeed = 0.02
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set FPS value to 60 for startup
    FPS = 1000
    FramesPerSecCounter = 1000
    
    ScrollPixelsPerFrameX = 8
    ScrollPixelsPerFrameY = 8
    
    UserPos.X = 50
    UserPos.Y = 50
    
    MinXBorder = XMinMapSize + (frmMain.Renderer.ScaleWidth / 65)
    MaxXBorder = XMaxMapSize - (frmMain.Renderer.ScaleWidth / 65)
    MinYBorder = YMinMapSize + (frmMain.Renderer.ScaleHeight / 65)
    MaxYBorder = YMaxMapSize - (frmMain.Renderer.ScaleHeight / 65)
    
    'partículas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
    Font_Create "Tahoma", 8, True, 0
    'Font_Create "Verdana", 8, False, 0
    
bRunning = True
Exit Sub
ErrHandler:
Debug.Print "Error Number Returned: " & Err.Number
bRunning = False
End Sub

Public Sub Engine_Deinit()
Erase MapData
Erase charlist
Set D3DDevice = Nothing
Set D3D = Nothing
Set DX8 = Nothing
End
End Sub


Private Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
'*****************************************************
'****** Coded by Menduz (lord.yo.wo@gmail.com) *******
'*****************************************************
    CreateTLVertex.X = X
    CreateTLVertex.Y = Y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
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
Public Sub Draw_GrhIndex(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer)
    If grh_index <= 0 Then Exit Sub
    Dim rgb_list(3) As Long
    
    rgb_list(0) = D3DColorXRGB(255, 255, 255)
    rgb_list(1) = D3DColorXRGB(255, 255, 255)
    rgb_list(2) = D3DColorXRGB(255, 255, 255)
    rgb_list(3) = D3DColorXRGB(255, 255, 255)
    
    Device_Box_Textured_Render grh_index, _
        X, Y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY
End Sub

Private Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByVal Animate As Byte, Optional ByVal alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single)
Dim CurrentGrhIndex As Integer
    If Grh.grhindex = 0 Then Exit Sub
    If GrhData(Grh.grhindex).NumFrames = 0 Then Exit Sub
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + ((timerElapsedTime * engineBaseSpeed) * Grh.speed) * Movement_Speed
            If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.grhindex).NumFrames) + 1
                
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If

        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
        End If
    End If

    Device_Box_Textured_Render CurrentGrhIndex, _
        X, Y, _
        GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
        MapData(map_x, map_y).light_value(), _
        GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
        alpha _
        , angle
'exits:

End Sub

Sub ShowNextFrame()
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                If UserMontando Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame * Velocidad / 10 'Reemplacen este 10 por el valor por el que quieran dividir la vel
            End If
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                   If UserMontando Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame * Velocidad / 10 'Reemplacen este 10 por el valor por el que quieran dividir la vel
            End If
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)

End Sub

Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y                   As Integer     'Keeps track of where on map we are
    Dim X                   As Integer     'Keeps track of where on map we are
    Dim screenminY          As Integer  'Start Y pos on current screen
    Dim screenmaxY          As Integer  'End Y pos on current screen
    Dim screenminX          As Integer  'Start X pos on current screen
    Dim screenmaxX          As Integer  'End X pos on current screen
    Dim minY                As Integer  'Start Y pos on current map
    Dim maxY                As Integer  'End Y pos on current map
    Dim minX                As Integer  'Start X pos on current map
    Dim maxX                As Integer  'End X pos on current map
    Dim ScreenX             As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY             As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset          As Integer
    Dim minYOffset          As Integer
    Dim PixelOffsetXTemp    As Integer 'For centering grhs
    Dim PixelOffsetYTemp    As Integer 'For centering grhs
    Dim CurrentGrhIndex     As Integer
    Dim offx                As Integer
    Dim offy                As Integer

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            'Layer 1 **********************************
            
                Call Draw_Grh(MapData(X, Y).Graphic(1), _
                        (ScreenX - 1) * 32 + PixelOffsetX, _
                        (ScreenY - 1) * 32 + PixelOffsetY, _
                        0, 1, , X, Y)
                
                If MapData(X, Y).Graphic(2).grhindex <> 0 Then
                    Call Draw_Grh(MapData(X, Y).Graphic(2), _
                        (ScreenX - 1) * 32 + PixelOffsetX, _
                        (ScreenY - 1) * 32 + PixelOffsetY, _
                        1, 1, , X, Y)
                End If
            '******************************************
            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            With MapData(X, Y)
                '******************************************

                'Object Layer **********************************
                If .ObjGrh.grhindex <> 0 Then
                    Call Draw_Grh(.ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , X, Y)
                TieneLuz X, Y
                End If
                '***********************************************
                
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call Char_Render(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, X, Y)
                End If
                '*************************************************
                
                
                'Layer 3 *****************************************
                If .Graphic(3).grhindex <> 0 Then
                    Call Draw_Grh(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , X, Y)
                End If
                '************************************************

            End With
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - 5
                    
    Dim ParticleOffsetX As Single
    Dim ParticleOffsetY As Single
    ParticleOffsetX = (PixelOffsetX + (32 \ 2))
    ParticleOffsetY = (PixelOffsetY + (32 \ 2))
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
            'Draw Particle System
            ScreenY = minYOffset - TileBufferSize
            For Y = minY To maxY
                ScreenX = minXOffset - TileBufferSize
                For X = minX To maxX
                    'Particles **********************************
                    If MapData(X, Y).particle_group_index Then
                        Particle_Group_Render MapData(X, Y).particle_group_index, (ScreenX * 32) + ParticleOffsetX, (ScreenY * 32) + ParticleOffsetY
                    End If
                    '**********************************
                    ScreenX = ScreenX + 1
                Next X
                ScreenY = ScreenY + 1
            Next Y
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
        Dim temp_array(3) As Long
      
        'Draw blocked tiles and grid
        ScreenY = minYOffset - TileBufferSize
        For Y = minY To maxY
            ScreenX = minXOffset - TileBufferSize
            For X = minX To maxX
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).grhindex Then
                    If Not bTecho Then
                        Call Draw_Grh(MapData(X, Y).Graphic(4), _
                            ScreenX * 32 + PixelOffsetX, _
                            ScreenY * 32 + PixelOffsetY, _
                            1, 1, , X, Y)
                    
                    End If
                        'bTecho
                End If
                '**********************************
                
                '*** Start Dialogs ***
                If MapData(X, Y).CharIndex Then
                    If charlist(MapData(X, Y).CharIndex).dialog <> "" Then
                        'Set colors
                        temp_array(0) = charlist(MapData(X, Y).CharIndex).dialog_color
                        temp_array(1) = charlist(MapData(X, Y).CharIndex).dialog_color
                        temp_array(2) = charlist(MapData(X, Y).CharIndex).dialog_color
                        temp_array(3) = charlist(MapData(X, Y).CharIndex).dialog_color

                        If charlist(MapData(X, Y).CharIndex).dialog_offset_counter_y < 5 Then
                            charlist(MapData(X, Y).CharIndex).dialog_offset_counter_y = charlist(MapData(X, Y).CharIndex).dialog_offset_counter_y + 1
                        End If
                        
                        Dim screen_y As Integer, screen_x As Integer
                        screen_x = (ScreenX * 32) + (CInt(charlist(MapData(X, Y).CharIndex).MoveOffsetX) + PixelOffsetX)
                        screen_y = (ScreenY * 32) + (CInt(charlist(MapData(X, Y).CharIndex).MoveOffsetY) + PixelOffsetY)
                        
                        'Render dialog
                        Engine_Text_Render charlist(MapData(X, Y).CharIndex).dialog, screen_x + 14 - Engine_Text_Width(charlist(MapData(X, Y).CharIndex).dialog, True) / 2, screen_y + (charlist(MapData(X, Y).CharIndex).body.HeadOffset.Y) - Engine_Text_Height(charlist(MapData(X, Y).CharIndex).dialog, True) - charlist(MapData(X, Y).CharIndex).dialog_offset_counter_y, temp_array, charlist(MapData(X, Y).CharIndex).dialog_font_index, True
                    End If
                End If
                '*** End Dialogs ***
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
         If meteo_particle Then Particle_Group_Render meteo_particle, 250, 0
 'Error LLuvia al Pasar de Mapa
    If bRain = True Then
        meteo_particle = engine.General_Particle_Create(8, -1, -1)
        meteo_particle = 1
    End If
End Sub

Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If

End Sub
Private Sub Device_Box_Textured_Render_Advance(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'Copies the Textures allowing resizing
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long

    If grhindex = 0 Then Exit Sub
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
        
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub

Public Sub Device_Box_Textured_Render(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/12/2004
'Just copies the Textures
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
    
    If grhindex = 0 Then Exit Sub
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    If (light_value(0) = 0) Then light_value(0) = base_light
    If (light_value(1) = 0) Then light_value(1) = base_light
    If (light_value(2) = 0) Then light_value(2) = base_light
    If (light_value(3) = 0) Then light_value(3) = base_light

        
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
    
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub


Public Sub Engine_MoveScreen(ByVal nHeading As E_Heading)
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
        
        Case E_Heading.EAST
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
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
    End If
End Sub

Private Sub Char_Render(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long

    With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                If UserMontando Then
                
                   .MoveOffsetX = .MoveOffsetX + 8 * Sgn(.scrollDirectionX) * timerTicksPerFrame * Velocidad / 10 'Reemplacen este 10 por el valor por el que quieran dividir la vel
                End If
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .body.Walk(.Heading).speed > 0 Then _
                    .body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                     If UserMontando Then
                   .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * Velocidad / 10
                End If
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .body.Walk(.Heading).speed > 0 Then _
                    .body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .body.Walk(.Heading).Started = 0
            .body.Walk(.Heading).FrameCounter = 1
            
            .Arma.WeaponWalk(.Heading).Started = 0
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        If .Head.Head(.Heading).grhindex Then
            If Not .invisible Then
                'Draw Body
      If .body.Walk(.Heading).grhindex Then _
                                Call Draw_Grh(.body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y)
                                
                            If .Head.Head(.Heading).grhindex Then
                                Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 1, 0, , X, Y)
                                If .Casco.Head(.Heading).grhindex Then _
                                    Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 1, 0, , X, Y)
                                If UserMontando = True Then
                            .Arma.WeaponWalk(.Heading).grhindex = 0
                            End If
                               If .Arma.WeaponWalk(.Heading).grhindex Then _
                                Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , X, Y)
                           End If
    If Not .iBody = 84 Then
    
            
            If .Heading <> EAST Then
                If .Escudo.ShieldWalk(.Heading).grhindex Then _
                    Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y + 37, 1, 1, , X, Y)
               
            End If
    
            If .Casco.Head(.Heading).grhindex Then _
                Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .body.HeadOffset.X, PixelOffsetY + .body.HeadOffset.Y, 1, 0, , X, Y)
           
        End If

        Movement_Speed = 0.5
                    'Draw name over head
                    If Nombres Then
                        If Len(.nombre) > 0 Then
                            Pos = InStr(.nombre, "<")
                            If Pos = 0 Then Pos = Len(.nombre) + 2
                            
                            If .priv = 0 Then
                                If .Criminal Then
                                    color = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                Else
                                    color = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                End If
                            Else
                                color = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            End If
                            
                            Dim temp_array(3) As Long
                            temp_array(0) = color
                            temp_array(1) = color
                            temp_array(2) = color
                            temp_array(3) = color
                            
                            'Nick
                            line = Left$(.nombre, Pos - 2)
                            Engine_Text_Render line, PixelOffsetX + 16 - Engine_Text_Width(line) / 2, PixelOffsetY + 30, temp_array()
                            
                            'Clan
                            line = mid$(.nombre, Pos)
                            Engine_Text_Render line, PixelOffsetX + 16 - Engine_Text_Width(line) / 2, PixelOffsetY + 45, temp_array()
                        End If
                    End If
                
                End If
            'End If
        Else
            'Draw Body
            If .body.Walk(.Heading).grhindex Then _
                Call Draw_Grh(.body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
        End If
        Movement_Speed = 1
 'Draw Particles
        Dim i As Integer
        If charlist(CharIndex).particle_count > 0 Then
            For i = 1 To UBound(charlist(CharIndex).particle_group)
                If charlist(CharIndex).particle_group(i) > 0 Then
                    Particle_Group_Render charlist(CharIndex).particle_group(i), PixelOffsetX + 16, PixelOffsetY
                End If
            Next i
        End If
        
        'Draw FX
        If .FxIndex <> 0 Then
            Call Draw_Grh(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1, True)

            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
    End With
End Sub


Public Sub Start()
DoEvents
vertList(0).X = 0: vertList(0).Y = 0
vertList(1).X = 800: vertList(1).Y = 0
vertList(2).X = 0: vertList(2).Y = 600
vertList(3).X = 800: vertList(3).Y = 600
vertList(0).rhw = 1
vertList(1).rhw = 1
vertList(2).rhw = 1
vertList(3).rhw = 1
vertList(0).color = D3DColorXRGB(255, 255, 255)
vertList(1).color = D3DColorXRGB(255, 255, 255)
vertList(2).color = D3DColorXRGB(255, 255, 255)
vertList(3).color = D3DColorXRGB(255, 255, 255)
vertList(0).tu = 0
vertList(0).tv = 0
vertList(1).tu = 1
vertList(1).tv = 0
vertList(2).tu = 0
vertList(2).tv = 1
vertList(3).tu = 1
vertList(3).tv = 1

temp_color_array(0) = D3DColorXRGB(255, 255, 255)
temp_color_array(1) = temp_color_array(0)
temp_color_array(2) = temp_color_array(1)
temp_color_array(3) = temp_color_array(2)

Do While prgRun
Call FlushBuffer
    If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then
        CheckKeys
        engine.DrawInv
        Audio.RenderSounds
        
        'Render
        D3DDevice.BeginScene
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            engine.Light_Render_All
            ShowNextFrame
            Engine_Text_Render FPS & " FPS", 496, 10, temp_color_array()
        
        D3DDevice.EndScene
        D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        
        If GetTickCount >= lFrameTimer + 1000 Then
            lFrameTimer = GetTickCount
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 0
            
            Dim i As Long
            For i = 1 To LastChar
                If charlist(i).dialog <> "" Then
                    charlist(i).dialog_life = charlist(i).dialog_life - 1
                    If charlist(i).dialog_life <= 0 Then
                        charlist(i).dialog = ""
                        charlist(i).dialog_color = 0
                        charlist(i).dialog_life = 0
                        charlist(i).dialog_offset_counter_y = 0
                        charlist(i).dialog_scroll = False
                    End If
                End If
            Next i
        End If
        
        FramesPerSecCounter = FramesPerSecCounter + 1
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = (timerElapsedTime * engineBaseSpeed)
    Else
        Sleep 10&
    End If
DoEvents

Loop
engine.Engine_Deinit
Call CloseClient
End Sub

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

Public Sub Char_Move_by_Head(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addx As Integer
    Dim addy As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.north
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        
        nX = X + addx
        nY = Y + addy
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    
    If UserEstado <> 1 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Public Sub Char_Move_by_Pos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
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
            nHeading = E_Heading.EAST
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
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then
        Call RemoveDialog(CharIndex)
    End If
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Public Sub Char_Clean()
Dim X As Byte
Dim Y As Byte
For X = 1 To 100
    For Y = 1 To 100
        If MapData(X, Y).CharIndex Then
            EraseChar MapData(X, Y).CharIndex
        End If
        If MapData(X, Y).ObjGrh.grhindex Then
            MapData(X, Y).ObjGrh.grhindex = 0
        End If
    Next Y
Next X
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Private Function Engine_FToDW(f As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function

Public Function Light_Remove(ByVal light_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        color_value = light_list(light_index).color
        Light_Color_Value_Get = True
    End If
End Function

Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, Optional ByVal color_value As Long = &HFFFFFF, _
                            Optional ByVal range As Byte = 1, Optional ByVal id As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
        If Map_Light_Get(map_x, map_y) <> 0 Then
            Light_Create = 0
            Exit Function
        End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, color_value, range, id
    End If
End Function

Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If Map_In_Bounds(map_x, map_y) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
    
            Light_Move = True
            
        End If
    End If
End Function

Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If

    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
    
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y
        
        nX = map_x
        nY = map_y
        
        Convert_Heading_to_Direction Heading, nX, nY
        
        'Make sure it's a legal move
        If Map_In_Bounds(nX, nY) Then
        
            'Move it
            Light_Erase light_index

            light_list(light_index).map_x = nX
            light_list(light_index).map_y = nY
    
            Light_Move_By_Head = True
            
        End If
    End If
End Function

Private Sub Convert_Heading_to_Direction(ByVal Heading As Byte, ByRef direction_x As Integer, ByRef direction_y As Integer)
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim addy As Integer
    Dim addx As Integer
    
    'Figure out which way to move
    Select Case Heading
        Case 1
            addy = -1
        
        Case 2
            addy = -1
            addx = 1
        
        Case 3
            addx = 1
        
        Case 4
            addx = 1
            addy = 1
        
        Case 5
            addy = 1
        
        Case 6
            addx = -1
            addy = 1
        
        Case 7
            addx = -1
        
        Case 8
            addx = -1
            addy = -1
    End Select
    
    direction_x = direction_x + addx
    direction_y = direction_y + addy
End Sub

Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, ByVal rgb_value As Long, _
                        ByVal range As Long, Optional ByVal id As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
    
    'Make active
    light_list(light_index).active = True
    
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).color = rgb_value
    light_list(light_index).range = range
    light_list(light_index).id = id
End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).active Then
            Light_Check = True
        End If
    End If
End Function

Public Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        
        If light_list(loop_counter).active Then
            Light_Render loop_counter
        End If
    
    Next loop_counter
End Sub

Private Sub Light_Render(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 11/11/2004
'Renders a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim color As Long
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Set color
    color = light_list(light_index).color
    
    'Arrange corners
    'NE
    If Map_In_Bounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = color
    End If
    'NW
    If Map_In_Bounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = color
    End If
    'SW
    If Map_In_Bounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = color
    End If
    'SE
    If Map_In_Bounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = color
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If Map_In_Bounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = color
            MapData(X, min_y).light_value(2) = color
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If Map_In_Bounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = color
            MapData(X, max_y).light_value(3) = color
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If Map_In_Bounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = color
            MapData(min_x, Y).light_value(3) = color
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If Map_In_Bounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = color
            MapData(max_x, Y).light_value(1) = color
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If Map_In_Bounds(X, Y) Then
                MapData(X, Y).light_value(0) = color
                MapData(X, Y).light_value(1) = color
                MapData(X, Y).light_value(2) = color
                MapData(X, Y).light_value(3) = color
            End If
        Next Y
    Next X
End Sub

Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).active = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function

Public Function Light_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).id = id
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index
        End If
    Next Index
    
    Light_Remove_All = True
End Function

Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Light
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub

Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Arrange corners
    'NE
    If Map_In_Bounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = base_light
    End If
    'NW
    If Map_In_Bounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = base_light
    End If
    'SW
    If Map_In_Bounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = base_light
    End If
    'SE
    If Map_In_Bounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = base_light
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If Map_In_Bounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = base_light
            MapData(X, min_y).light_value(2) = base_light
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If Map_In_Bounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = base_light
            MapData(X, max_y).light_value(3) = base_light
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If Map_In_Bounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = base_light
            MapData(min_x, Y).light_value(3) = base_light
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If Map_In_Bounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = base_light
            MapData(max_x, Y).light_value(1) = base_light
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If Map_In_Bounds(X, Y) Then
                MapData(X, Y).light_value(0) = base_light
                MapData(X, Y).light_value(1) = base_light
                MapData(X, Y).light_value(2) = base_light
                MapData(X, Y).light_value(3) = base_light
            End If
        Next Y
    Next X
End Sub

Public Function Map_Light_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a light_index and return it
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    'We start from the back, to get the last light to be placed on the tile first
    loopc = light_last
    Do Until light_list(loopc).map_x = map_x And light_list(loopc).map_y = map_y
        If loopc = 0 Then
            Map_Light_Get = 0
            Exit Function
        End If
        loopc = loopc - 1
    Loop
    
    Map_Light_Get = loopc
Exit Function
ErrorHandler:
    Map_Light_Get = 0
End Function

Public Sub DrawInv()
If frmMain.picInv.Visible = False Then Exit Sub
Static re As RECT
re.Left = 0
re.Top = 0
re.bottom = 160
re.Right = 160

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
D3DDevice.BeginScene
    Inventario.DrawInventory
D3DDevice.EndScene
D3DDevice.Present re, ByVal 0, frmMain.picInv.hWnd, ByVal 0
End Sub

Public Sub Draw_FilledBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, color As Long, outlinecolor As Long)

    Static box_rect As RECT
    Static Outline As RECT
    Static rgb_list(3) As Long
    Static rgb_list2(3) As Long
    Static Vertex(3) As TLVERTEX
    Static Vertex2(3) As TLVERTEX
    
    rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color
    
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
    
    With box_rect
        .bottom = Y + Height
        .Left = X
        .Right = X + Width
        .Top = Y
    End With
    
    With Outline
        .bottom = Y + Height + 2
        .Left = X - 2
        .Right = X + Width + 2
        .Top = Y - 2
    End With
    
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), 0, 0
    
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
End Sub

Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    'If prgRun = False Then Exit Function
    loopc = 1
    Do Until particle_group_list(loopc).active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function

Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True
        End If
    End If
End Function

Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'Returns the particle_group_index if successful, else 0
'Modified by Juan Martín Sotuyo Dodero
'Modified by Augusto José Rando
'**************************************************************
    
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
        End If
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If

End Function

Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function

Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
    
    Particle_Group_Remove_All = True
End Function

Public Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).id = id
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function

Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As particle_group
    Dim i As Integer
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If particle_group_list(particle_group_index).char_index > 0 And particle_group_list(particle_group_index).char_index <= LastChar Then
            For i = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count
                If charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = particle_group_index Then
                    charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = 0
                    
                    'We don't resize arrays by now, it's really a waste...
                    'If i = UBound(CharList(particle_group_list(particle_group_index).char_index).particle_group) Then
                    '    CharList(particle_group_list(particle_group_index).char_index).particle_count = i - 1
                    '    ReDim Preserve CharList(particle_group_list(particle_group_index).char_index).particle_group(1 To (i - 1)) As Long
                    'End If
                    
                    Exit For
                End If
            Next i
        End If
    ElseIf particle_group_index = meteo_particle Then
       meteo_particle = 0
    End If
    
    particle_group_list(particle_group_index) = temp
    
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Sub

Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).X1 = X1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).X2 = X2
    particle_group_list(particle_group_index).Y2 = Y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'handle
    particle_group_list(particle_group_index).id = id
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on map
    If (map_x <> -1) And (map_y <> -1) Then
        MapData(map_x, map_y).particle_group_index = particle_group_index
    End If
    
End Sub

Private Sub Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Char index
    particle_group_list(particle_group_index).char_index = char_index
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).X1 = X1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).X2 = X2
    particle_group_list(particle_group_index).Y2 = Y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'handle
    particle_group_list(particle_group_index).id = id
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on char
    charlist(char_index).particle_group(particle_char_index) = particle_group_index
End Sub

Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_index As Integer
    
    If char_index > 0 And char_index <= LastChar Then
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)
        If char_part_index = -1 Then Exit Function
        Call Particle_Group_Remove(char_part_index)
    End If

End Function

Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim i As Integer
    
    If char_index > 0 And char_index <= LastChar Then
        For i = 1 To charlist(char_index).particle_count
            If charlist(char_index).particle_group(i) <> 0 Then Call Particle_Group_Remove(charlist(char_index).particle_group(i))
        Next i
    End If
    
End Function

Private Function Char_Particle_Group_Find(ByVal char_index As Integer, ByVal stream_type As Long) As Integer
'*****************************************************************
'Author: Augusto José Rando
'Modified: returns slot or -1
'*****************************************************************

Dim i As Integer

For i = 1 To charlist(char_index).particle_count
    If particle_group_list(charlist(char_index).particle_group(i)).stream_type = stream_type Then
        Char_Particle_Group_Find = charlist(char_index).particle_group(i)
        Exit Function
    End If
Next i

Char_Particle_Group_Find = -1

End Function

Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until charlist(char_index).particle_group(loopc) = 0
        If loopc = charlist(char_index).particle_count Then
            Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
            charlist(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Particle_Group_Next_Open = loopc

Exit Function

ErrorHandler:
    charlist(char_index).particle_count = 1
    ReDim charlist(char_index).particle_group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1

End Function

Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
  
    'Set colors
    temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
    temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
    temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
    temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    
    If particle_group_list(particle_group_index).alive_counter Then
    
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
    
    
        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
        
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).X1, particle_group_list(particle_group_index).Y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).X2, _
                            particle_group_list(particle_group_index).Y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey
        Next loopc
        
        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
    
    Else
        'If it's dead destroy it
        Particle_Group_Destroy particle_group_index
    End If
End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************
    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            'Start new particle
            InitGrh temp_particle.Grh, grh_index, alpha_blend
            temp_particle.X = RandomNumber(X1, X2) - (32 / 2)
            temp_particle.Y = RandomNumber(Y1, Y2) - (32 / 2)
            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.friction = fric
        Else
            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength
                If temp_particle.Y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength
                End If
            End If
            'Do rotation
            If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0
            End If
            
            If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
        End If
        
        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
    'Draw it
    If grh_resize = True Then
        If temp_particle.Grh.grhindex Then
             Grh_Render_Advance temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
            Exit Sub
        End If
    End If
    If temp_particle.Grh.grhindex Then
        Grh_Render temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, rgb_list(), True, True, alpha_blend
    End If
End Sub
Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If Not (map_x < MinXBorder Or map_x > MaxXBorder Or map_y < MinYBorder Or map_y > MaxYBorder) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function

Public Function Map_In_Bounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    Map_In_Bounds = False
    Exit Function
End If

Map_In_Bounds = True
End Function

Private Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Modified by Juan Martín Sotuyo Dodero
'Added centering
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
    
    If Grh.grhindex = 0 Then Exit Sub
        
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.speed) * Movement_Speed
        If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
            Grh.FrameCounter = 1
        End If
    End If

    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1

    
    If Grh.grhindex = 25338 Then Exit Sub
    
    grh_index = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)
    
    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
        
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(Grh.grhindex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(Grh.grhindex).TileWidth * (32 \ 2)) + 32 \ 2
        End If
    End If
    
    If v_centered Then
        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(Grh.grhindex).TileHeight * 32) + 32
        End If
    End If
    
    'Draw it to device
    Device_Box_Textured_Render grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list(), _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        alpha_blend, _
        Grh.angle

End Sub

Private Sub Grh_Render_Advance(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 11/19/2003
'Similar to Grh_Render, but let´s you resize the Grh
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
    
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.speed)
        If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
            Grh.FrameCounter = 1
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    grh_index = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)
    
    'Center Grh over X, Y pos
    If GrhData(Grh.grhindex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(Grh.grhindex).TileWidth * (32 \ 2)) + 32 \ 2
    End If
    
    If GrhData(Grh.grhindex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(Grh.grhindex).TileHeight * 32) + 32
    End If
    
    'Draw it to device
    Device_Box_Textured_Render_Advance grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        Width, Height, alpha_blend, Grh.angle
End Sub

Function Engine_Text_Width(texto As String, Optional multi As Boolean = False) As Integer
Dim a As Integer, b As Integer, d As Integer, e As Integer, f As Integer
Dim graf As Grh

If multi = False Then
    For a = 1 To Len(texto)
        b = Asc(mid$(texto, a, 1))
        graf.grhindex = font_types(1).ascii_code(b)
        If (b <> 32) And (b <> 5) And (b <> 129) And (b <> 9) And (b <> 4) And (b <> 255) And (b <> 2) And (b <> 151) And (b <> 152) Then
            Engine_Text_Width = Engine_Text_Width + GrhData(GrhData(graf.grhindex).Frames(1)).pixelWidth '+ 1
        Else
            Engine_Text_Width = Engine_Text_Width + 4
        End If
    Next a
Else
    e = 0
    f = 0
    For a = 1 To Len(texto)
        b = Asc(mid$(texto, a, 1))
        graf.grhindex = font_types(1).ascii_code(b)
        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then d = d + 4
            End If
        Else
            If graf.grhindex > 12 Then
                d = d + GrhData(GrhData(graf.grhindex).Frames(1)).pixelWidth '+ 1
                If d > Engine_Text_Width Then Engine_Text_Width = d
            End If
        End If
        e = e + 1
    Next a
End If
End Function

Function Engine_Text_Height(texto As String, Optional multi As Boolean = False) As Integer
Dim a As Integer, b As Integer, c  As Integer, d  As Integer, e As Integer, f As Integer
  
If multi = False Then
    Engine_Text_Height = 0
Else
    e = 0
    f = 0
    
    For a = 1 To Len(texto)
        b = Asc(mid$(texto, a, 1))

        If b = 32 Or b = 13 Then
            If e >= 20 Then
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then
                    d = d + 4
                End If
            End If
        End If
        e = e + 1
    Next a
  
Engine_Text_Height = f * 14
  
End If

End Function

Sub Engine_Text_Render(texto As String, X As Integer, Y As Integer, ByRef text_color() As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False)

Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer, f As Integer
Dim graf As Grh

If (Len(texto) = 0) Then Exit Sub

d = 0
If multi_line = False Then
    For a = 1 To Len(texto)
        b = Asc(mid$(texto, a, 1))
        graf.grhindex = font_types(font_index).ascii_code(b)
        If b <> 32 Then
            If graf.grhindex <> 0 Then
                'mega sombra O-matica
                graf.grhindex = font_types(font_index).ascii_code(b) + 100
                Grh_Render graf, (X + d) + 1, Y + 1, text_color, False, False, False
                graf.grhindex = font_types(font_index).ascii_code(b)
                Grh_Render graf, (X + d), Y, text_color, False, False, False
                d = d + GrhData(GrhData(graf.grhindex).Frames(1)).pixelWidth '+ 1
            End If
        Else
            d = d + 4
        End If
    Next a
Else
    e = 0
    f = 0
    For a = 1 To Len(texto)
        b = Asc(mid$(texto, a, 1))
        graf.grhindex = font_types(font_index).ascii_code(b)
        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then d = d + 4
            End If
        Else
            If graf.grhindex > 12 Then
                'mega sombra O-matica
                graf.grhindex = font_types(font_index).ascii_code(b) + 100
                Grh_Render graf, (X + d) + 1, Y + 1 + f * 14, text_color, False, False, False
                graf.grhindex = font_types(font_index).ascii_code(b)
                Grh_Render graf, (X + d), Y + f * 14, text_color, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                d = d + GrhData(GrhData(graf.grhindex).Frames(1)).pixelWidth '+ 1
            End If
       End If
       e = e + 1
    Next a
End If
End Sub

Public Sub Engine_Font_Initialize()

Dim a As Long

font_types(1).font_size = 9
font_types(1).ascii_code(48) = 21452
font_types(1).ascii_code(49) = 21453
font_types(1).ascii_code(50) = 21454
font_types(1).ascii_code(51) = 21455
font_types(1).ascii_code(52) = 21456
font_types(1).ascii_code(53) = 21457
font_types(1).ascii_code(54) = 21458
font_types(1).ascii_code(55) = 21459
font_types(1).ascii_code(56) = 21460
font_types(1).ascii_code(57) = 21461

For a = 0 To 25
    font_types(1).ascii_code(a + 97) = 21400 + a
Next a

For a = 0 To 25
    font_types(1).ascii_code(a + 65) = 21426 + a
Next a

font_types(1).ascii_code(33) = 21462
font_types(1).ascii_code(161) = 21463
font_types(1).ascii_code(34) = 21464
font_types(1).ascii_code(36) = 21465
font_types(1).ascii_code(191) = 21466
font_types(1).ascii_code(35) = 21467
font_types(1).ascii_code(36) = 21468
font_types(1).ascii_code(37) = 21469
font_types(1).ascii_code(38) = 21470
font_types(1).ascii_code(47) = 21471
font_types(1).ascii_code(92) = 21472
font_types(1).ascii_code(40) = 21473
font_types(1).ascii_code(41) = 21474
font_types(1).ascii_code(61) = 21475
font_types(1).ascii_code(39) = 21476
font_types(1).ascii_code(123) = 21477
font_types(1).ascii_code(125) = 21478
font_types(1).ascii_code(95) = 21479
font_types(1).ascii_code(45) = 21480
font_types(1).ascii_code(63) = 21465
font_types(1).ascii_code(64) = 21481
font_types(1).ascii_code(94) = 21482
font_types(1).ascii_code(91) = 21483
font_types(1).ascii_code(93) = 21484
font_types(1).ascii_code(60) = 21485
font_types(1).ascii_code(62) = 21486
font_types(1).ascii_code(42) = 21487
font_types(1).ascii_code(43) = 21488
font_types(1).ascii_code(46) = 21489
font_types(1).ascii_code(44) = 21490
font_types(1).ascii_code(58) = 21491
font_types(1).ascii_code(59) = 21492
font_types(1).ascii_code(124) = 21493
font_types(1).ascii_code(252) = 21800
font_types(1).ascii_code(220) = 21801
font_types(1).ascii_code(225) = 21802
font_types(1).ascii_code(233) = 21803
font_types(1).ascii_code(237) = 21804
font_types(1).ascii_code(243) = 21805
font_types(1).ascii_code(250) = 21806
font_types(1).ascii_code(253) = 21807
font_types(1).ascii_code(193) = 21808
font_types(1).ascii_code(201) = 21809
font_types(1).ascii_code(205) = 21810
font_types(1).ascii_code(211) = 21811
font_types(1).ascii_code(218) = 21812
font_types(1).ascii_code(221) = 21813
font_types(1).ascii_code(224) = 21814
font_types(1).ascii_code(232) = 21815
font_types(1).ascii_code(236) = 21816
font_types(1).ascii_code(242) = 21817
font_types(1).ascii_code(249) = 21818
font_types(1).ascii_code(192) = 21819
font_types(1).ascii_code(200) = 21820
font_types(1).ascii_code(204) = 21821
font_types(1).ascii_code(210) = 21822
font_types(1).ascii_code(217) = 21823
font_types(1).ascii_code(241) = 21824
font_types(1).ascii_code(209) = 21825
font_types(1).ascii_code(196) = 25238
font_types(1).ascii_code(194) = 25239
font_types(1).ascii_code(203) = 25240
font_types(1).ascii_code(207) = 25241
font_types(1).ascii_code(214) = 25242
font_types(1).ascii_code(212) = 25243
font_types(1).ascii_code(172) = 26815

font_types(2).font_size = 9
font_types(2).ascii_code(97) = 21936
font_types(2).ascii_code(108) = 21937
font_types(2).ascii_code(115) = 21938
font_types(2).ascii_code(70) = 21939
font_types(2).ascii_code(48) = 21940
font_types(2).ascii_code(49) = 21941
font_types(2).ascii_code(50) = 21942
font_types(2).ascii_code(51) = 21943
font_types(2).ascii_code(52) = 21944
font_types(2).ascii_code(53) = 21945
font_types(2).ascii_code(54) = 21946
font_types(2).ascii_code(55) = 21947
font_types(2).ascii_code(56) = 21948
font_types(2).ascii_code(57) = 21949
font_types(2).ascii_code(33) = 21950
font_types(2).ascii_code(161) = 21951
font_types(2).ascii_code(42) = 21952

End Sub

Public Sub Char_Dialog_Set(ByVal char_index As Integer, ByVal char_dialog As String, ByVal char_dialog_color As Long, ByVal char_dialog_life As Byte, Optional ByVal font_index As Integer = 1)
charlist(char_index).dialog = char_dialog
charlist(char_index).dialog_color = char_dialog_color
charlist(char_index).dialog_life = char_dialog_life
charlist(char_index).dialog_font_index = font_index
charlist(char_index).dialog_offset_counter_y = 0
End Sub
    Public Sub setup_ambient()
    frmMain.imgHora.Picture = LoadPicture(App.path & "\Recursos\Interface\" & Hour(time) & ".jpg")
    If Hour(time) <= 10 Then
    luz_dia(Hour(time)).r = 160
    luz_dia(Hour(time)).g = 160
    luz_dia(Hour(time)).b = 160
    Else
    luz_dia(Hour(time)).r = Hour(time) * 10
    luz_dia(Hour(time)).g = Hour(time) * 10
    luz_dia(Hour(time)).b = Hour(time) * 10
    End If
    End Sub
 
Public Function change_day_effect(ByVal o_r As Long, ByVal og As Long, ByVal OB As Long, ByVal r_ As Long, ByVal g_ As Long, ByVal b_ As Long) As Long
Dim r As Long
Dim g As Long
Dim b As Long
r = o_r
g = og
b = OB
    
    If o_r > r_ Then
        r = r - 1
    ElseIf o_r < r_ Then
        r = r + 1
    End If

    If og > g_ Then
        g = g - 1
    ElseIf og < r Then
        g = g + 1
    End If

    If OB > b_ Then
        b = b - 1
    ElseIf og < r Then
        b = b + 1
    End If
If r > 255 Then r = 255
If g > 255 Then g = 255
If b > 255 Then b = 255
day_r_old = r
day_g_old = g
day_b_old = b

If r < 0 Or g < 0 Or b < 0 Then r = 1: g = 1: b = 1
change_day_effect = ARGB(r, g, b, 255)

End Function


Public Function ARGB(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal a As Long) As Long
        
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
    
    ARGB = c

End Function
Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long

rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Char_Particle_Create = engine.Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).X1, StreamData(ParticulaInd).Y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).X2, _
    StreamData(ParticulaInd).Y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, _
    StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)

End Function

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Particle_Create = engine.Particle_Group_Create(X, Y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).X1, StreamData(ParticulaInd).Y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).X2, _
    StreamData(ParticulaInd).Y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, _
    StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)

End Function
Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer)

    Dim char_part_free_index As Integer
    
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
    
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If

End Function

Private Function Char_Check(ByVal char_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).Heading > 0)
    End If
    
End Function
    Public Sub MostrarFPSS()
      ' Engine_Text_Render ping & " ping", 496, 10, temp_color_array()
         Engine_Text_Render FPS & " FPS", 496, 10, temp_color_array()
End Sub

Public Function General_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
    Randomize Timer
    General_Random_Number = (UpperBound - LowerBound) * Rnd + LowerBound
End Function
Public Function Light_Remove_From_Pos(ByVal posX As Byte, ByVal posY As Byte, Optional vColor As Long = 0) As Boolean
    Dim i As Long
    
    For i = 1 To light_last
        If light_list(i).map_x = posX Then
            If light_list(i).map_y = posY Then
                If vColor <> 0 Then
                    If light_list(i).color = vColor Then
                        If Light_Check(i) Then
                            Light_Destroy i
                        End If
                    End If
                Else
                    If Light_Check(i) Then
                        Light_Destroy i
                    End If
                End If
            End If
        End If
    Next i
    
    Light_Render_All
End Function

Public Function Map_FileNum_Get(ByVal map_x As Integer, ByVal map_y As Integer, ByVal layer As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando (barrin@imperiumao.com.ar)
'Last Modify Date: 10/6/2005
'*****************************************************************
    On Error GoTo Error_Handler
    
    If layer < 1 Or layer > 4 Then
        Map_FileNum_Get = 0
        Exit Function
    End If
    
    If Map_In_Bounds(map_x, map_y) Then
        If MapData(map_x, map_y).Graphic(layer).grhindex > 0 Then
            Map_FileNum_Get = GrhData(MapData(map_x, map_y).Graphic(layer).grhindex).FileNum
        Else
            Map_FileNum_Get = 0
        End If
    End If

    Exit Function
    
Error_Handler:
    Map_FileNum_Get = 0

End Function

Public Sub DrawGrhtoHdc(ByVal desthDC As Long, Grh As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
 
On Error GoTo Err
 
    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
    Dim grh_index As Integer
   
    grh_index = Grh
 
    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).NumFrames = 0 Then Exit Sub
 
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If
 
    src_x = GrhData(grh_index).sX
    src_y = GrhData(grh_index).sY
    src_width = GrhData(grh_index).pixelWidth
    src_height = GrhData(grh_index).pixelHeight
           
    hdcsrc = CreateCompatibleDC(desthDC)
 
    file_path = App.path & "\recursos\graficos\" & GrhData(grh_index).FileNum & ".bmp"
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
   
    'BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
   TransparentBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, src_width, src_height, RGB(0, 0, 0)
 
   Call DeleteObject(SelectObject(hdcsrc, PrevObj))
   DeleteDC hdcsrc
 
Err:
If Err.Number = 481 Then MsgBox "Imposible cargar recurso error: " & "481"
 
End Sub

