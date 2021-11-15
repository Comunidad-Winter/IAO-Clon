Attribute VB_Name = "Acciones"
 'ImperiumAO 0.11.6
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

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim tempIndex As Integer
    
On Error Resume Next
    '¿Rango Visión? (ToxicWaste)
    If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub
    End If
    
    '¿Posicion valida?
    If InMapBounds(Map, X, Y) Then
        With UserList(UserIndex)
            If MapData(Map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
                tempIndex = MapData(Map, X, Y).NpcIndex
                
                'Set the target NPC
                .flags.TargetNPC = tempIndex
                
                If Npclist(tempIndex).Comercia = 1 Then
                    '¿Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(UserIndex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then
                    '¿Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Estas muerto", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'A depositar de una
                    Call IniciarDeposito(UserIndex)
                    
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Pirata Then
 
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                       Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del pirata.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                   
                    'Call WriteShowPasajes(UserIndex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Or Npclist(tempIndex).NPCtype = eNPCType.ResucitadorNewbie Then
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Revivimos si es necesario
                    If .flags.Muerto = 1 And (Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex)) Then
                        Call RevivirUsuario(UserIndex)
                    End If
                    
                    If Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(UserIndex) Then
                        'curamos totalmente
                        .Stats.MinHP = .Stats.MaxHP
                        Call WriteUpdateUserStats(UserIndex)
                    End If
                 ElseIf Npclist(tempIndex).NPCtype = 18 Then
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 3 Then
                        Call WriteConsoleMsg(UserIndex, "Estas lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If Npclist(tempIndex).flags.Faccion = 1 Then 'Ciudad
                        If UserList(tempIndex).Faccion.ArmadaMatados > 0 Or UserList(tempIndex).Faccion.CiudadanosMatados > 0 Then
                            Call WriteChatOverHead(UserIndex, "Has asesino ciudadanos del Imperio. Si realmente quieres regresar, una parte de tu alma sera necesaria. Tu experiencia para subir al siguiente nivel sera aumentada en 0.7%, escribiendo /PERDON", Npclist(tempIndex).Char.CharIndex, vbWhite)
                            Exit Sub
                        End If
                        
                        UserList(tempIndex).Faccion.Rango = 0
                        UserList(UserIndex).Faccion.Renegado = 0

                        UserList(UserIndex).Faccion.Ciudadano = 1
                        UserList(UserIndex).Hogar = cBanderbill
                        UserList(UserIndex).Faccion.Republicano = 0
                    
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
                    ElseIf Npclist(tempIndex).flags.Faccion = 2 Then
                        If UserList(UserIndex).Faccion.MilicianosMatados > 0 Or UserList(UserIndex).Faccion.RepublicanosMatados > 0 Then
                            Call WriteChatOverHead(UserIndex, "Has asesino ciudadanos de la Republica. Si realmente quieres regresar, una parte de tu alma sera necesaria. Tu experiencia para subir al siguiente nivel sera aumentada en 0.7%, escribiendo /PERDON", Npclist(tempIndex).Char.CharIndex, vbWhite)
                            Exit Sub
                        End If
                        
                        UserList(UserIndex).Faccion.Rango = 0
                        UserList(UserIndex).Faccion.Renegado = 0
                        UserList(UserIndex).Faccion.Ciudadano = 0
                        UserList(UserIndex).Faccion.Republicano = 1
                        UserList(UserIndex).Hogar = cSuramei
                        
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
                    End If
                End If
                
            '¿Es un obj?
            ElseIf MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X, Y).ObjInfo.ObjIndex
                
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y, UserIndex)
                    Case eOBJType.otCarteles 'Es un cartel
                        Call AccionParaCartel(Map, X, Y, UserIndex)
                    Case eOBJType.otForos 'Foro
                        Call AccionParaForo(Map, X, Y, UserIndex)
                    Case eOBJType.otLeña    'Leña
                        If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                            Call AccionParaRamita(Map, X, Y, UserIndex)
                        End If
                End Select
            '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
            ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X + 1, Y).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X + 1, Y, UserIndex)
                    
                End Select
            
            ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
        
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X + 1, Y + 1, UserIndex)
                End Select
            
            ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
                tempIndex = MapData(Map, X, Y + 1).ObjInfo.ObjIndex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(Map, X, Y + 1, UserIndex)
                End Select
            End If
        End With
    End If
End Sub

Sub AccionParaForo(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Pos As WorldPos
Pos.Map = Map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
Call WriteCorreoList(UserIndex)
End Sub


Sub AccionParaPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim MiObj As Obj
Dim wp As WorldPos

If Not (Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, X, Y) > 2) Then
    If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
        If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                    
                    MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                    
                    Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                    
                    'Desbloquea
                    MapData(Map, X, Y).Blocked = 0
                    MapData(Map, X - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(True, Map, X, Y, 0)
                    Call Bloquear(True, Map, X - 1, Y, 0)
                    
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                    
                Else
                     Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(Map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                                
                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1
                
                
                Call Bloquear(True, Map, X - 1, Y, 1)
                Call Bloquear(True, Map, X, Y, 1)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
        End If
        
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).ObjInfo.ObjIndex
    Else
        Call WriteConsoleMsg(UserIndex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub AccionParaCartel(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next


Dim MiObj As Obj

If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = 8 Then
  
  If Len(ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).texto) > 0 Then
    Call WriteShowSignal(UserIndex, MapData(Map, X, Y).ObjInfo.ObjIndex)
  End If
  
End If

End Sub

Sub AccionParaRamita(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim raise As Integer

Dim Pos As WorldPos
Pos.Map = Map
Pos.X = X
Pos.Y = Y

If Distancia(Pos, UserList(UserIndex).Pos) > 2 Then
    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(Map, X, Y).Trigger = eTrigger.ZONASEGURA Or MapInfo(Map).Pk = False Then
    Call WriteConsoleMsg(UserIndex, "En zona segura no puedes hacer fogatas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).Stats.UserSkills(Supervivencia) > 1 And UserList(UserIndex).Stats.UserSkills(Supervivencia) < 6 Then
            Suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(Supervivencia) <= 10 Then
            Suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(Supervivencia) >= 10 And UserList(UserIndex).Stats.UserSkills(Supervivencia) Then
            Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    If MapInfo(UserList(UserIndex).Pos.Map).Zona <> Ciudad Then
        Obj.ObjIndex = FOGATA
        Obj.amount = 1
        
        Call WriteConsoleMsg(UserIndex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
        
        Call MakeObj(Obj, Map, X, Y)
        
        'Las fogatas prendidas se deben eliminar
        Dim Fogatita As New cGarbage
        Fogatita.Map = Map
        Fogatita.X = X
        Fogatita.Y = Y
        Call TrashCollector.Add(Fogatita)
    Else
        Call WriteConsoleMsg(UserIndex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
Else
    Call WriteConsoleMsg(UserIndex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
End If

Call SubirSkill(UserIndex, Supervivencia)

End Sub
