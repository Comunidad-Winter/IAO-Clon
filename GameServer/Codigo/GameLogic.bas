Attribute VB_Name = "Extra"
'ImperiumAO 1.0
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit
Public Function ClaseToEnum(ByVal Clase As String) As eClass
Dim i As Byte
For i = 1 To NUMCLASES
    If UCase$(ListaClases(i)) = UCase$(Clase) Then
        ClaseToEnum = i
    End If
Next i
End Function

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
    EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie
End Function
Public Function esArmada(ByVal UserIndex As Integer) As Boolean
    esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1)
End Function
Public Function esCaos(ByVal UserIndex As Integer) As Boolean
    esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1)
End Function
Public Function esMili(ByVal UserIndex As Integer) As Boolean
    esMili = (UserList(UserIndex).Faccion.Milicia = 1)
End Function
Public Function esFaccion(ByVal UserIndex As Integer) As Boolean
    esFaccion = (UserList(UserIndex).Faccion.ArmadaReal = 1 Or UserList(UserIndex).Faccion.FuerzasCaos = 1 Or UserList(UserIndex).Faccion.Milicia = 1)
End Function
Public Function criminal(ByVal UserIndex As Integer) As Boolean
    criminal = esRene(UserIndex)
End Function
Public Function esRene(ByVal UserIndex As Integer) As Boolean
    esRene = (UserList(UserIndex).Faccion.Renegado)
End Function
Public Function esCiuda(ByVal UserIndex As Integer) As Boolean
    esCiuda = (UserList(UserIndex).Faccion.Ciudadano)
End Function
Public Function esRepu(ByVal UserIndex As Integer) As Boolean
    esRepu = (UserList(UserIndex).Faccion.Republicano)
End Function

Public Function EsGM(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
    EsGM = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))
End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the Map passage of Users. Allows the existance
'of exclusive maps for Newbies, Royal Army and Caos Legion members
'and enables GMs to enter every map without restriction.
'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
'***************************************************
On Error GoTo errhandler

Dim nPos As WorldPos
Dim FxFlag As Boolean
'Controla las salidas
If InMapBounds(Map, X, Y) Then
    
    If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
        FxFlag = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport
    End If
    
    If (MapData(Map, X, Y).TileExit.Map > 0) And (MapData(Map, X, Y).TileExit.Map <= NumMaps) Then
        '¿Es mapa de newbies?
        If UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "NEWBIE" Then
            '¿El usuario es un newbie?
            If EsNewbie(UserIndex) Or EsGM(UserIndex) Then
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                    Else
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, False)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                        Else
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, False)
                        End If
                    End If
                End If
            Else 'No es newbie
                Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                Call ClosestStablePos(UserList(UserIndex).Pos, nPos)

                If nPos.X <> 0 And nPos.Y <> 0 Then
                    Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, False)
                End If
            End If
        ElseIf UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "ARMADA" Then '¿Es mapa de Armadas?
            '¿El usuario es Armada?
            If esArmada(UserIndex) Or EsGM(UserIndex) Then
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                    Else
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                        Else
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                        End If
                    End If
                End If
            Else 'No es armada
                Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para miembros del ejercito Real", FontTypeNames.FONTTYPE_INFO)
                Call ClosestStablePos(UserList(UserIndex).Pos, nPos)

                If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                End If
            End If
        ElseIf UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "CAOS" Then '¿Es mapa de Caos?
            '¿El usuario es Caos?
            If esCaos(UserIndex) Or EsGM(UserIndex) Then
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                    Else
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                        Else
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                        End If
                    End If
                End If
            Else 'No es caos
                Call WriteConsoleMsg(UserIndex, "Mapa exclusivo para miembros del ejercito Oscuro.", FontTypeNames.FONTTYPE_INFO)
                Call ClosestStablePos(UserList(UserIndex).Pos, nPos)

                If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                End If
            End If
        ElseIf UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "FACCION" Then '¿Es mapa de faccionarios?
            '¿El usuario es Armada o Caos?
            If esArmada(UserIndex) Or esCaos(UserIndex) Or EsGM(UserIndex) Then
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                    Else
                        Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                        Else
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                        End If
                    End If
                End If
            Else 'No es Faccionario
                Call WriteConsoleMsg(UserIndex, "Solo se permite entrar al Mapa si eres miembro de alguna Facción", FontTypeNames.FONTTYPE_INFO)
                Call ClosestStablePos(UserList(UserIndex).Pos, nPos)

                If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                End If
            End If
        Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.
            If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                If FxFlag Then
                    Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                Else
                    Call WarpUserChar(UserIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
                End If
            Else
                Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                If nPos.X <> 0 And nPos.Y <> 0 Then
                    If FxFlag Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                    Else
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y)
                    End If
                End If
            End If
        End If
        'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
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
    End If
    
End If



Exit Sub

errhandler:
    Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.description)
End Sub

Function InRangoVision(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

If X > UserList(UserIndex).Pos.X - MinXBorder And X < UserList(UserIndex).Pos.X + MinXBorder Then
    If Y > UserList(UserIndex).Pos.Y - MinYBorder And Y < UserList(UserIndex).Pos.Y + MinYBorder Then
        InRangoVision = True
        Exit Function
    End If
End If
InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean

If X > Npclist(NpcIndex).Pos.X - MinXBorder And X < Npclist(NpcIndex).Pos.X + MinXBorder Then
    If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
        InRangoVisionNPC = True
        Exit Function
    End If
End If
InRangoVisionNPC = False

End Function


Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
            
If (Map <= 0 Or Map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapBounds = False
Else
    InMapBounds = True
End If

End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, Optional PuedeTierra As Boolean = True)
'*****************************************************************
'Author: Unknown (original version)
'Last Modification: 24/01/2007 (ToxicWaste)
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY, PuedeAgua, PuedeTierra) Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub

Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub

Function NameIndex(ByVal name As String) As Integer

Dim UserIndex As Integer
'¿Nombre valido?
If LenB(name) = 0 Then
    NameIndex = 0
    Exit Function
End If

If InStrB(name, "+") <> 0 Then
    name = UCase$(Replace(name, "+", " "))
End If

UserIndex = 1
Do Until UCase$(UserList(UserIndex).name) = UCase$(name)
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        NameIndex = 0
        Exit Function
    End If
    
Loop
 
NameIndex = UserIndex
 
End Function



Function IP_Index(ByVal inIP As String) As Integer
 
Dim UserIndex As Integer
'¿Nombre valido?
If LenB(inIP) = 0 Then
    IP_Index = 0
    Exit Function
End If
  
UserIndex = 1
Do Until UserList(UserIndex).ip = inIP
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        IP_Index = 0
        Exit Function
    End If
    
Loop
 
IP_Index = UserIndex

Exit Function

End Function


Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
Dim LoopC As Integer
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged = True Then
        If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameIP = False
End Function

Function CheckForSameName(ByVal name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
Dim LoopC As Long
For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged Then
        
        'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
        'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
        'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
        'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
        'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
        
        If UCase$(UserList(LoopC).name) = UCase$(name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = Pos.X
Y = Pos.Y

If Head = eHeading.NORTH Then
    nX = X
    nY = Y - 1
End If

If Head = eHeading.SOUTH Then
    nX = X
    nY = Y + 1
End If

If Head = eHeading.EAST Then
    nX = X + 1
    nY = Y
End If

If Head = eHeading.WEST Then
    nX = X - 1
    nY = Y
End If

'Devuelve valores
Pos.X = nX
Pos.Y = nY

End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Checks if the position is Legal.
'***************************************************
'¿Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            LegalPos = False
Else
    If PuedeAgua And PuedeTierra Then
        LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).UserIndex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0)
    ElseIf PuedeTierra And Not PuedeAgua Then
        LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).UserIndex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0) And _
                   (Not HayAgua(Map, X, Y))
    ElseIf PuedeAgua And Not PuedeTierra Then
        LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).UserIndex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0) And _
                   (HayAgua(Map, X, Y))
    Else
        LegalPos = False
    End If
   
End If
End Function

Function MoveToLegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: ZaMa
'Last Modification: 26/03/2009
'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
'***************************************************

Dim UserIndex As Integer
Dim IsDeadChar As Boolean


'¿Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            MoveToLegalPos = False
    Else
        UserIndex = MapData(Map, X, Y).UserIndex
        If UserIndex > 0 Then
            IsDeadChar = UserList(UserIndex).flags.Muerto = 1
        Else
            IsDeadChar = False
        End If
    
    If PuedeAgua And PuedeTierra Then
        MoveToLegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (UserIndex = 0 Or IsDeadChar) And _
                   (MapData(Map, X, Y).NpcIndex = 0)
    ElseIf PuedeTierra And Not PuedeAgua Then
        MoveToLegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (UserIndex = 0 Or IsDeadChar) And _
                   (MapData(Map, X, Y).NpcIndex = 0) And _
                   (Not HayAgua(Map, X, Y))
    ElseIf PuedeAgua And Not PuedeTierra Then
        MoveToLegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (UserIndex = 0 Or IsDeadChar) And _
                   (MapData(Map, X, Y).NpcIndex = 0) And _
                   (HayAgua(Map, X, Y))
    Else
        MoveToLegalPos = False
    End If
  
End If


End Function

Public Sub FindLegalPos(ByVal UserIndex As Integer, ByVal Map As Integer, ByRef X As Integer, ByRef Y As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 26/03/2009
'Search for a Legal pos for the user who is being teleported.
'***************************************************

    If MapData(Map, X, Y).UserIndex <> 0 Or _
        MapData(Map, X, Y).NpcIndex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(Map, X, Y).UserIndex = UserIndex Then Exit Sub
                            
        Dim FoundPlace As Boolean
        Dim tX As Long
        Dim tY As Long
        Dim Rango As Long
        Dim OtherUserIndex As Integer
    
        For Rango = 1 To 5
            For tY = Y - Rango To Y + Rango
                For tX = X - Rango To X + Rango
                    'Reviso que no haya User ni NPC
                    If MapData(Map, tX, tY).UserIndex = 0 And _
                        MapData(Map, tX, tY).NpcIndex = 0 Then
                        
                        If InMapBounds(Map, tX, tY) Then FoundPlace = True
                        
                        Exit For
                    End If

                Next tX
        
                If FoundPlace Then _
                    Exit For
            Next tY
            
            If FoundPlace Then _
                    Exit For
        Next Rango

    
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            X = tX
            Y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            OtherUserIndex = MapData(Map, X, Y).UserIndex
            If OtherUserIndex <> 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(OtherUserIndex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(OtherUserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        Call FlushBuffer(OtherUserIndex)
                    End If
                End If
            
                Call CloseSocket(OtherUserIndex)
            End If
        End If
    End If

End Sub

Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean

If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
    LegalPosNPC = False
Else

 If AguaValida = 0 Then
   LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
     (MapData(Map, X, Y).UserIndex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).Trigger <> eTrigger.POSINVALIDA) _
     And Not HayAgua(Map, X, Y)
 Else
   LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
     (MapData(Map, X, Y).UserIndex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).Trigger <> eTrigger.POSINVALIDA)
 End If
 
End If


End Function

Sub SendHelp(ByVal Index As Integer)
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call WriteConsoleMsg(Index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))
    End If
End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim ft As FontTypeNames

'¿Rango Visión? (ToxicWaste)
If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
    Exit Sub
End If

'¿Posicion valida?
If InMapBounds(Map, X, Y) Then
    UserList(UserIndex).flags.TargetMap = Map
    UserList(UserIndex).flags.TargetX = X
    UserList(UserIndex).flags.TargetY = Y
    '¿Es un obj?
    If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
        'Informa el nombre
        UserList(UserIndex).flags.TargetObjMap = Map
        UserList(UserIndex).flags.TargetObjX = X
        UserList(UserIndex).flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X + 1
            UserList(UserIndex).flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X + 1
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    
    If FoundSomething = 1 Then
        UserList(UserIndex).flags.TargetObj = MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
        If MostrarCantidad(UserList(UserIndex).flags.TargetObj) Then
            Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).name & " - " & MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.amount & "", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).name, FontTypeNames.FONTTYPE_INFO)
        End If
    
    End If
    '¿Es un personaje?
    If Y + 1 <= YMaxMapSize Then
        If MapData(Map, X, Y + 1).UserIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).UserIndex
            FoundChar = 1
        End If
        If MapData(Map, X, Y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(Map, X, Y).UserIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).UserIndex
            FoundChar = 1
        End If
        If MapData(Map, X, Y).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    
    'Reaccion al personaje
    If FoundChar = 1 Then '  ¿Encontro un Usuario?
            
       If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(UserIndex).flags.Privilegios And PlayerType.Dios Then
            
            If LenB(UserList(TempCharIndex).DescRM) = 0 And UserList(TempCharIndex).showName Then 'No tiene descRM y quiere que se vea su nombre.
                Stat = Stat & "(" & ListaClases(UserList(TempCharIndex).Clase) & " " & ListaRazas(UserList(TempCharIndex).raza) & " Nivel " 'No tiene descRM y quiere que se vea su nombre.
                
                If UserList(UserIndex).Stats.ELV + 30 < UserList(TempCharIndex).Stats.ELV Then
                    Stat = Stat & "?? "
                Else
                    Stat = Stat & UserList(TempCharIndex).Stats.ELV & " "
                End If
                
                If UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.05) Then
                    Stat = Stat & "| Muerto"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.1) Then
                    Stat = Stat & "| Casi muerto"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.5) Then
                    Stat = Stat & "| Malherido"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.75) Then
                    Stat = Stat & "| Herido"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP) Then
                    Stat = Stat & "| Levemente Herido"
                Else
                    Stat = Stat & "| Intacto"
                End If
                
                'Si Esta Comerciendo ?
                If UserList(TempCharIndex).flags.Comerciando = True Then
                    Stat = Stat & " | Comerciando)"
                Else
                    Stat = Stat & ")"
                End If
                
                'Si Esta Inmovilizado ?
                If UserList(TempCharIndex).flags.Inmovilizado = True Then
                    Stat = Stat & " | Inmovilizado)"

                End If
                
                'Si Esta Paralizado ?
                If UserList(TempCharIndex).flags.Paralizado = True Then
                    Stat = Stat & " | Paralizado)"
    
                End If
                
                'Si Esta Estupido ? Como Duarte...
                If UserList(TempCharIndex).flags.Estupidez = True Then
                    Stat = Stat & " | Estupidizado)"

                End If
                
                'Si Esta Trabajando ? (Pescando/Talando/Minando)
                If UserList(TempCharIndex).flags.PuedeTrabajar = True Then
                    Stat = Stat & " | Trabajando)"
        
                End If
                
                'Si Esta Descansando ?
                If UserList(TempCharIndex).flags.Descansar = True Then
                    Stat = Stat & " | Descansando)"
      
                End If
                
                'Si Esta Envenenado ?
                If UserList(TempCharIndex).flags.Envenenado = True Then
                    Stat = Stat & " | Envenenado)"
      
                End If
                
                If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                    Stat = Stat & " <Armada Imperial> " & "<" & TituloReal(TempCharIndex) & ">"
                ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                    Stat = Stat & " <Hordas del Caos> " & "<" & TituloCaos(TempCharIndex) & ">"
                ElseIf UserList(TempCharIndex).Faccion.Milicia = 1 Then
                    Stat = Stat & " <Milicia Republicana> " & "<" & TituloMilicia(TempCharIndex) & ">"
                End If
                
                If UserList(TempCharIndex).GuildIndex > 0 Then
                    Stat = Stat & " <" & modGuilds.GuildName(UserList(TempCharIndex).GuildIndex) & ">"
                End If
                
                If Len(UserList(TempCharIndex).desc) > 0 Then
                    Stat = UserList(TempCharIndex).name & " " & Stat & " - " & UserList(TempCharIndex).desc
                Else
                    Stat = UserList(TempCharIndex).name & " " & Stat
                End If
                
                                
                If UserList(TempCharIndex).flags.Privilegios And PlayerType.RoyalCouncil Then
                    Stat = Stat & " [CONSEJO DE BANDERBILL]"
                    ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                ElseIf UserList(TempCharIndex).flags.Privilegios And PlayerType.ChaosCouncil Then
                    Stat = Stat & " [CONSEJO DE LAS SOMBRAS]"
                    ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
                Else
                    If Not UserList(TempCharIndex).flags.Privilegios And PlayerType.User Then
                        Stat = Stat & " <ImperiumAO Game Master>"
                        ft = FontTypeNames.FONTTYPE_GM
                    ElseIf UserList(TempCharIndex).Faccion.Ciudadano = 1 And Not UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                    Stat = Stat & " <Imperial> ~0~0~255~1~0"
                    ft = FontTypeNames.FONTTYPE_CITIZEN
                ElseIf UserList(TempCharIndex).Faccion.Renegado = 1 And Not UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                    Stat = Stat & " <Renegado> ~128~128~128~1~0"
                ElseIf UserList(TempCharIndex).Faccion.Republicano = 1 And Not UserList(TempCharIndex).Faccion.Milicia = 1 Then
                    Stat = Stat & " <Republicano> ~255~128~0~1~0"
                    ft = FontTypeNames.FONTTYPE_CITIZEN
                ElseIf UserList(TempCharIndex).Faccion.Milicia = 1 Then
                    Stat = Stat & "~255~128~0~1~0"
                ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                    Stat = Stat & "~190~0~0~1~0"
                ElseIf UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                    Stat = Stat & "~0~190~200~1~0"
                End If
                End If
            Else  'Si tiene descRM la muestro siempre.
                Stat = UserList(TempCharIndex).DescRM
                ft = FontTypeNames.FONTTYPE_INFOBOLD
            End If
            
            If LenB(Stat) > 0 Then
                Call WriteConsoleMsg(UserIndex, Stat, ft)
            End If
            
            FoundSomething = 1
            UserList(UserIndex).flags.TargetUser = TempCharIndex
            UserList(UserIndex).flags.TargetNPC = 0
            UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
       End If

    End If
    If FoundChar = 2 Then '¿Encontro un NPC?
            Dim estatus As String
            
            If UserList(UserIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
                estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
            Else
                If UserList(UserIndex).flags.Muerto = 0 Then
                If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 0 Then
                    estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
                Else
                    estatus = "¡Error!"
                End If
            End If
        End If
            
            If Len(Npclist(TempCharIndex).desc) > 1 Then
                Call WriteChatOverHead(UserIndex, Npclist(TempCharIndex).desc, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
            ElseIf TempCharIndex = CentinelaNPCIndex Then
                'Enviamos nuevamente el texto del centinela según quien pregunta
                Call modCentinela.CentinelaSendClave(UserIndex)
            Else
                If Npclist(TempCharIndex).MaestroUser > 0 Then
                    Call WriteConsoleMsg(UserIndex, estatus & Npclist(TempCharIndex).name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).name, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, estatus & Npclist(TempCharIndex).name & ".", FontTypeNames.FONTTYPE_INFO)
                    If UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        Call WriteConsoleMsg(UserIndex, "Le pegó primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                
            End If
            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(UserIndex).flags.TargetNPC = TempCharIndex
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
        
    End If
    
    If FoundChar = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If

Else
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If
End If




End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
Dim X As Integer
Dim Y As Integer

X = Pos.X - Target.X
Y = Pos.Y - Target.Y

'NE
If Sgn(X) = -1 And Sgn(Y) = 1 Then
    FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)
    Exit Function
End If

'NW
If Sgn(X) = 1 And Sgn(Y) = 1 Then
    FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
    Exit Function
End If

'SW
If Sgn(X) = 1 And Sgn(Y) = -1 Then
    FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
    Exit Function
End If

'SE
If Sgn(X) = -1 And Sgn(Y) = -1 Then
    FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
    Exit Function
End If

'Sur
If Sgn(X) = 0 And Sgn(Y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If

'norte
If Sgn(X) = 0 And Sgn(Y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If

'oeste
If Sgn(X) = 1 And Sgn(Y) = 0 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'este
If Sgn(X) = -1 And Sgn(Y) = 0 Then
    FindDirection = eHeading.EAST
    Exit Function
End If

'misma
If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal Index As Integer) As Boolean

ItemNoEsDeMapa = ObjData(Index).OBJType <> eOBJType.otPuertas And _
            ObjData(Index).OBJType <> eOBJType.otForos And _
            ObjData(Index).OBJType <> eOBJType.otCarteles And _
            ObjData(Index).OBJType <> eOBJType.otArboles And _
            ObjData(Index).OBJType <> eOBJType.otYacimiento And _
            ObjData(Index).OBJType <> eOBJType.otTeleport
End Function
'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
MostrarCantidad = ObjData(Index).OBJType <> eOBJType.otPuertas And _
            ObjData(Index).OBJType <> eOBJType.otForos And _
            ObjData(Index).OBJType <> eOBJType.otCarteles And _
            ObjData(Index).OBJType <> eOBJType.otArboles And _
            ObjData(Index).OBJType <> eOBJType.otYacimiento And _
            ObjData(Index).OBJType <> eOBJType.otTeleport
End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

'EsObjetoFijo = OBJType = eOBJType.otForos Or _
               OBJType = eOBJType.otCarteles Or _
               OBJType = eOBJType.otArboles Or _
               OBJType = eOBJType.otYacimiento

End Function

Public Function ParticleToLevel(ByVal UserIndex As Integer) As Integer

'Meditaciones Bien Modificadas y Con Facciones By Zenitram

'Meditacion NEWBIE
If UserList(UserIndex).Stats.ELV = 1 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 2 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 3 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 4 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 5 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 6 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 7 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 8 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 9 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 10 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 11 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 12 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 13 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 14 Then
ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV = 15 Then
ParticleToLevel = 81
ElseIf UserList(UserIndex).Stats.ELV = 16 Then
ParticleToLevel = 81
ElseIf UserList(UserIndex).Stats.ELV = 17 Then
ParticleToLevel = 81
ElseIf UserList(UserIndex).Stats.ELV = 18 Then
ParticleToLevel = 81
ElseIf UserList(UserIndex).Stats.ELV = 19 Then
ParticleToLevel = 81
ElseIf UserList(UserIndex).Stats.ELV = 20 Then
ParticleToLevel = 81
ElseIf UserList(UserIndex).Stats.ELV = 21 Then
ParticleToLevel = 81
ElseIf UserList(UserIndex).Stats.ELV = 22 Then
ParticleToLevel = 81
ElseIf UserList(UserIndex).Stats.ELV = 23 Then
ParticleToLevel = 81
ElseIf UserList(UserIndex).Stats.ELV = 24 Then
ParticleToLevel = 81
ElseIf UserList(UserIndex).Stats.ELV = 25 Then
ParticleToLevel = 41
ElseIf UserList(UserIndex).Stats.ELV = 26 Then
ParticleToLevel = 41
ElseIf UserList(UserIndex).Stats.ELV = 27 Then
ParticleToLevel = 41
ElseIf UserList(UserIndex).Stats.ELV = 28 Then
ParticleToLevel = 41
ElseIf UserList(UserIndex).Stats.ELV = 29 Then
ParticleToLevel = 41
ElseIf UserList(UserIndex).Stats.ELV = 30 Then
ParticleToLevel = 41
ElseIf UserList(UserIndex).Stats.ELV = 31 Then
ParticleToLevel = 41
ElseIf UserList(UserIndex).Stats.ELV = 32 Then
ParticleToLevel = 41
ElseIf UserList(UserIndex).Stats.ELV = 33 Then
ParticleToLevel = 41
ElseIf UserList(UserIndex).Stats.ELV = 34 Then
ParticleToLevel = 41
ElseIf UserList(UserIndex).Stats.ELV = 35 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 36 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 37 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 38 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 39 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 40 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 41 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 42 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 43 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 44 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 45 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 46 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 47 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 48 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 49 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then        'By Zenitram
        ParticleToLevel = 39
ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then    'By Zenitram
        ParticleToLevel = 40
ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 71
ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then    'By Zenitram
        ParticleToLevel = 38
ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then     'By Zenitram
        ParticleToLevel = 37
ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then     'By Zenitram
        ParticleToLevel = 71
End If
ElseIf UserList(UserIndex).Stats.ELV = 50 Then
If UserList(UserIndex).Faccion.Renegado = 1 Then      'By Zenitram
        ParticleToLevel = 119
    ElseIf UserList(UserIndex).Faccion.Ciudadano = 1 Then     'By Zenitram
        ParticleToLevel = 36
    ElseIf UserList(UserIndex).Faccion.Republicano = 1 Then     'By Zenitram
        ParticleToLevel = 121
    ElseIf UserList(UserIndex).Faccion.ArmadaReal = 1 Then      'By Zenitram
        ParticleToLevel = 125
    ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then         'By Zenitram
        ParticleToLevel = 120
    ElseIf UserList(UserIndex).Faccion.Milicia = 1 Then              'By Zenitram
        ParticleToLevel = 121
    End If
End If

'Meditacion GM By Zenitram
If UserList(UserIndex).Faccion.Rango = 1 Then
ParticleToLevel = 126
End If


End Function

Public Function NoTieneEspacioAmigos(ByVal Usuario As Integer) As Boolean
  Dim i As Long
  Dim count As Byte
 
  For i = 1 To MAXAMIGOS
  If Not UserList(Usuario).Amigos(i).Nombre = "Nadies" Then
  count = count + 1
  End If
  Next i
 
  If count = MAXAMIGOS Then
  NoTieneEspacioAmigos = True
  End If
 
End Function
Public Function BuscarSlotAmigoVacio(ByVal Usuario As Integer) As Byte
  Dim i As Long
 
  For i = 1 To MAXAMIGOS
  If UserList(Usuario).Amigos(i).Nombre = "Nadies" Then
  BuscarSlotAmigoVacio = i
  Exit Function
  End If
  Next i
 
End Function
Public Function BuscarSlotAmigoName(ByVal Usuario As Integer, ByVal Nombre As String) As Boolean
  Dim i As Long
 
  For i = 1 To MAXAMIGOS
  If UCase$(UserList(Usuario).Amigos(i).Nombre) = UCase$(Nombre) Then
  BuscarSlotAmigoName = True
  Exit Function
  End If
  Next i
 
End Function
 
 
Public Function BuscarSlotAmigoNameSlot(ByVal Usuario As Integer, ByVal Nombre As String) As Byte
  Dim i As Long
 
  For i = 1 To MAXAMIGOS
  If UCase$(UserList(Usuario).Amigos(i).Nombre) = UCase$(Nombre) Then
  BuscarSlotAmigoNameSlot = i
  Exit Function
  End If
  Next i
 
End Function
Public Sub delAmigoOfli(ByVal charName As String, ByVal Amigo As String)
  Dim CharFile As String
  Dim i As Long
  Dim Tiene As Boolean
  CharFile = CharPath & charName & ".pjs"
  If FileExist(CharFile) Then
 
  For i = 1 To MAXAMIGOS
  If UCase$(CStr(GetVar(CharFile, "AMIGOS", "NOMBRE" & i))) = UCase$(Amigo) Then
  Tiene = True
  Exit For
  End If
  Next i
 
  If Tiene Then
  'Lo borramos
  Call WriteVar(CharFile, "AMIGOS", "NOMBRE" & i, "Nadies")
  Call WriteVar(CharFile, "AMIGOS", "IGNORADO" & i, 0)
  End If
 
  End If
End Sub
Public Function IntentarAgregarAmigo(ByVal Usuario As Integer, ByVal Otro As Integer, ByRef razon As String) As Boolean
  With UserList(Usuario)
  If Otro = 0 Or Usuario = 0 Then
  razon = "Usuario Desconectado"
  IntentarAgregarAmigo = False
  Exit Function
 
  ElseIf Usuario = Otro Then
  razon = "Usuario Invalido"
  IntentarAgregarAmigo = False
  Exit Function
 
  ElseIf EsGM(Otro) = True Then
  razon = "Usuario Desconectado"
  IntentarAgregarAmigo = False
  Exit Function
 
  ElseIf EsGM(Usuario) = True Then
  razon = "Los Administradores no pueden agregar a usuarios"
  IntentarAgregarAmigo = False
  Exit Function
 
  ElseIf NoTieneEspacioAmigos(Usuario) = True Then
  razon = "No tienes mas espacio para poder agregar amigos"
  IntentarAgregarAmigo = False
  Exit Function
 
  ElseIf NoTieneEspacioAmigos(Otro) = True Then
  razon = "El otro usuario no tiene mas espacio para aceptar Amigos"
  IntentarAgregarAmigo = False
  Exit Function
 
  ElseIf BuscarSlotAmigoName(Usuario, UserList(Otro).name) = True Then
  razon = "Tu y " & UserList(Otro).name & "Ya son amigos"
  IntentarAgregarAmigo = False
  Exit Function
  End If
 
  IntentarAgregarAmigo = True
  End With
End Function
Public Sub ActualizarSlotAmigo(ByVal Usuario As Integer, ByVal Slot As Byte, Optional ByVal Todo As Boolean = False)

End Sub
Public Function ObtenerIndexLibre(ByVal Usuario As Integer) As Integer
  Dim i As Long
 
  For i = 1 To MAXAMIGOS
  If UserList(Usuario).Amigos(i).Index <= 0 Then
  ObtenerIndexLibre = i
  Exit Function
  End If
  Next i
 
End Function
Public Function ObtenerIndexUsuado(ByVal Usuario As Integer, ByVal Otro As Integer) As Integer
  Dim i As Long
 
  For i = 1 To MAXAMIGOS
  If UserList(Usuario).Amigos(i).Index = Otro Then
  ObtenerIndexUsuado = i
  Exit Function
  End If
  Next i
 
End Function
 
Public Sub ObtenerIndexAmigos(ByVal Usuario As Integer, ByVal Desconectar As Boolean)
  Dim i As Long
  Dim Slot As Byte
  With UserList(Usuario)
 
  If Desconectar = False Then
  For i = 1 To LastUser
  If LenB(UserList(i).name) > 0 Then
  If BuscarSlotAmigoName(Usuario, UserList(i).name) Then
  'Lo encontro y agregamos el index
  Slot = ObtenerIndexLibre(Usuario)
  'Por las dudas
  If Slot > 0 Then _
  .Amigos(Slot).Index = i
  If BuscarSlotAmigoName(i, .name) Then
  'Actualizamos la lista del otro
  Slot = ObtenerIndexLibre(i)
  If Slot > 0 Then
  UserList(i).Amigos(Slot).Index = Usuario
  'Informamos al otro de nuestra presencia
  Call WriteConsoleMsg(i, "Amigos> " & .name & " se ha conectado", FontTypeNames.FONTTYPE_CONSEJO)
  End If
  End If
  End If
  End If
  Next i
  Else
  For i = 1 To MAXAMIGOS
  'Antes q nada
  If .Amigos(i).Index > 0 Then
  Call WriteConsoleMsg(.Amigos(i).Index, "Amigos> " & .name & " se ha desconectado", FontTypeNames.FONTTYPE_CONSEJO)
  'Actualizamos la lista de index de los amigos
  Slot = ObtenerIndexUsuado(.Amigos(i).Index, Usuario)
  If Slot > 0 Then _
  UserList(.Amigos(i).Index).Amigos(Slot).Index = 0
  End If
  Next i
  End If
  End With
End Sub
 
Public Sub SetAreaResuTheNpc(ByVal iNpc As Integer)
 
        ' @@ Miqueas
       ' @@ 17-10-2015
        ' @@ Set Trigger in this NPC area
       Const Range = 4 ' @@ + 4 Tildes a la redonda del npc
 
        Dim X      As Long
 
        Dim Y      As Long
     
        Dim NpcPos As WorldPos
     
        NpcPos.Map = Npclist(iNpc).Pos.Map
        NpcPos.X = Npclist(iNpc).Pos.X
        NpcPos.Y = Npclist(iNpc).Pos.Y
        For X = NpcPos.X - Range To NpcPos.X + Range
                For Y = NpcPos.Y - Range To NpcPos.Y + Range
 
                        If InMapBounds(NpcPos.Map, X, Y) Then
                                If (MapData(NpcPos.Map, X, Y).Trigger <> eTrigger.AutoResu) Or (MapData(NpcPos.Map, X, Y).Trigger = eTrigger.NADA) Then
                                        MapData(NpcPos.Map, X, Y).Trigger = eTrigger.AutoResu
                                End If
                        End If
 
                Next Y
        Next X
 
End Sub
 
Public Function IsAreaResu(ByVal UserIndex As Integer) As Boolean
 
        ' @@ Miqueas
       ' @@ 17/10/2015
        ' @@ Validate Trigger Area
       With UserList(UserIndex)
 
               If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.AutoResu Then
                       IsAreaResu = True
 
                       Exit Function
 
               End If
 
       End With
 
       IsAreaResu = False
End Function
 
Public Sub AutoCurar(ByVal UserIndex As Integer)
 
       ' @@ Miqueas
        ' @@ 17-10-15
       ' @@ Zona de auto curacion
     
        With UserList(UserIndex)
 

                If .flags.Muerto = 1 Then
                        Call RevivirUsuario(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "El sacerdote te ha resucitado", FontTypeNames.FONTTYPE_INFO)
                        GoTo temp
                End If
 
                If .Stats.MinHP < .Stats.MaxHP Then
                        .Stats.MinHP = .Stats.MaxHP
                        Call WriteUpdateHP(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "El sacerdote te ha curado.", FontTypeNames.FONTTYPE_INFO)
                End If
 
temp:
 
                If .flags.Ceguera = 1 Then
                        .flags.Ceguera = 0
                End If
                If .flags.Envenenado = 1 Then
                        .flags.Envenenado = 0
                End If
 
        End With
 
End Sub
 
Public Function isNPCResucitador(ByVal iNpc As Integer) As Boolean
 
       With Npclist(iNpc)
 
               If (.NPCtype = eNPCType.Revividor) Or (.NPCtype = eNPCType.ResucitadorNewbie) Then
                       isNPCResucitador = True
 
                       Exit Function
 
               End If
 
       End With
 
       isNPCResucitador = False
End Function

Public Sub Paso_Por_Fogata(ByVal UserIndex As Integer, Optional ByVal Rango As Byte = 2)
'***************************************************
'Autor:Bateman
'Si el user camina cerca de una fogata y esta invisible
'hay una posibilidad de hacerlo visible.
'***************************************************
    Dim X As Long
    Dim Y As Long
    Dim Pos As WorldPos
    Dim PosF As WorldPos
    Pos.Map = UserList(UserIndex).Pos.Map
    Pos.X = UserList(UserIndex).Pos.X
    Pos.Y = UserList(UserIndex).Pos.Y
 
    With UserList(UserIndex)
        If .flags.invisible = 0 Then Exit Sub
 
        For X = Pos.X - Rango To Pos.X + Rango
            For Y = Pos.Y - Rango To Pos.Y + Rango
                If InMapBounds(Pos.Map, X, Y) Then
                    If MapData(Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
                        If ObjData(MapData(Pos.Map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                            PosF.Map = Pos.Map
                            PosF.X = X
                            PosF.Y = Y
                            If Distancia(Pos, PosF) <= Rango Then
                                If RandomNumber(1, 100) <= 20 Then
                                    'Removemos la invisibilidad
                                    .flags.invisible = 0
                                    .Counters.Invisibilidad = 0
                                    'Call SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
                                    Call WriteConsoleMsg(UserIndex, "Te has acercado demasiado a la fogata,esta misma a revelado tu apariencia!.", FontTypeNames.FONTTYPE_INFO)
                                Else
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            Next Y
        Next X
    End With
End Sub
 
