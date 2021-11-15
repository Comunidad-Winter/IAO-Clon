Attribute VB_Name = "ModFacciones"
'---------------------------------------------------------------------------------------
' Module    : ModFacciones
' Author    : Shermie80
' Date      : 09/03/2015
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public Const ExpAlUnirse As Long = 8600
Public Const ExpX100 As Integer = 4200

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call WriteChatOverHead(UserIndex, "¡Ya perteneces a las tropas reales! Ve a combatir criminales", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If Criminal(UserIndex) Then
    Call WriteChatOverHead(UserIndex, "No aceptamos otros seguidores de otras Facciones en la armada imperial.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call WriteChatOverHead(UserIndex, "Sal de aqui, Asqueroso enemigo.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If (UserList(UserIndex).Faccion.CriminalesMatados + UserList(UserIndex).Faccion.CaosMatados) < 15 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 15 enemigos, solo has matado " & (UserList(UserIndex).Faccion.CriminalesMatados + UserList(UserIndex).Faccion.CaosMatados), str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 40 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes ser al menos de nivel 40", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If
 
With UserList(UserIndex)
    If .guildIndex > 0 Then
        If modGuilds.GuildAlignment(.guildIndex) = .name Then
            If modGuilds.GuildAlignment(.guildIndex) = "Neutro" Then
                Call WriteChatOverHead(UserIndex, "Eres el fundador de un clan neutro", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        End If
    End If
End With

UserList(UserIndex).Faccion.ArmadaReal = 1

UserList(UserIndex).Faccion.Rango = 1

Dim MiObj As obj
Dim bajos As Byte
MiObj.Amount = 1
    
If UserList(UserIndex).raza = Enano Or UserList(UserIndex).raza = Gnomo Then
    bajos = 1
End If
    
Select Case UserList(UserIndex).clase
    Case eClass.Cleric
        MiObj.ObjIndex = 1544 + bajos
    Case eClass.Mage
        MiObj.ObjIndex = 1546 + bajos
    Case eClass.warrior
        MiObj.ObjIndex = 1548 + bajos
    Case eClass.Assasin
        MiObj.ObjIndex = 1550 + bajos
    Case eClass.Bard
        MiObj.ObjIndex = 1552 + bajos
    Case eClass.Druid
        MiObj.ObjIndex = 1554 + bajos
    Case eClass.bountyhunter
        MiObj.ObjIndex = 1556 + bajos
    Case eClass.Paladin
        MiObj.ObjIndex = 1558 + bajos
    Case eClass.hunter
        MiObj.ObjIndex = 1560 + bajos
    Case eClass.drakkar
        MiObj.ObjIndex = 1562 + bajos
    Case eClass.Nigromante
        MiObj.ObjIndex = 1564 + bajos
End Select

If Not MeterItemEnInventario(UserIndex, MiObj) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If

Call WriteChatOverHead(UserIndex, "Bienvenido al Ejército Imperial, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
 
If UserList(UserIndex).guildIndex > 0 Then
    If modGuilds.GuildAlignment(UserList(UserIndex).guildIndex) = "Neutro" Then
        Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).name)
        Call WriteConsoleMsg(UserIndex, "Has sido expulsado del clan por tu nueva facción.", FontTypeNames.FONTTYPE_GUILD)
    End If
End If

If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

End Sub
Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
Dim Matados As Long

If UserList(UserIndex).Faccion.Rango = 10 Then
    Exit Sub
End If

Matados = UserList(UserIndex).Faccion.CriminalesMatados + UserList(UserIndex).Faccion.CaosMatados

If Matados < matadosArmada(UserList(UserIndex).Faccion.Rango) Then
    Call WriteChatOverHead(UserIndex, "Mata " & matadosArmada(UserList(UserIndex).Faccion.Rango) - Matados & " Criminales más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

UserList(UserIndex).Faccion.Rango = UserList(UserIndex).Faccion.Rango + 1
If UserList(UserIndex).Faccion.Rango >= 6 Then ' Segunda jeraquia xD
    Dim MiObj As obj
    MiObj.Amount = 1
    Dim bajos As Byte
    
    If UserList(UserIndex).raza = Enano Or UserList(UserIndex).raza = Gnomo Then
        bajos = 1
    End If
        
    Select Case UserList(UserIndex).clase
        Case eClass.Cleric
            MiObj.ObjIndex = 1566 + bajos
        Case eClass.Mage
            MiObj.ObjIndex = 1568 + bajos
        Case eClass.warrior
            MiObj.ObjIndex = 1570 + bajos
        Case eClass.Assasin
            MiObj.ObjIndex = 1572 + bajos
        Case eClass.Bard
            MiObj.ObjIndex = 1574 + bajos
        Case eClass.Druid
            MiObj.ObjIndex = 1576 + bajos
        Case eClass.bountyhunter
            MiObj.ObjIndex = 1578 + bajos
        Case eClass.Paladin
            MiObj.ObjIndex = 1580 + bajos
        Case eClass.hunter
            MiObj.ObjIndex = 1582 + bajos
        Case eClass.drakkar
            MiObj.ObjIndex = 1584 + bajos
        Case eClass.Nigromante
            MiObj.ObjIndex = 1586 + bajos
    End Select
    
    

    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
End If

End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)


    UserList(UserIndex).Faccion.ArmadaReal = 0
    'Call PerderItemsFaccionarios(UserIndex)
    If Expulsado Then
        Call WriteConsoleMsg(UserIndex, "¡¡¡Has sido expulsado de las tropas reales!!!.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "¡¡¡Te has retirado de las tropas reales!!!.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura real si está equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        'Desequipamos el escudo de caos si está equipado
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpObjIndex)
    End If
    
    If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)
End Sub
Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)

    UserList(UserIndex).Faccion.FuerzasCaos = 0
    'Call PerderItemsFaccionarios(UserIndex)
    If Expulsado Then
        Call WriteConsoleMsg(UserIndex, "¡¡¡Has sido expulsado de la Legión Oscura!!!.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(UserIndex, "¡¡¡Te has retirado de la Legión Oscura!!!.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura de caos si está equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        'Desequipamos el escudo de caos si está equipado
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpObjIndex)
    End If
    
    If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)
End Sub


Public Function TituloReal(ByVal UserIndex As Integer) As String

Select Case UserList(UserIndex).Faccion.RecompensasReal
    Case 0
        TituloReal = "Soldado Real"
    Case 1
        TituloReal = "Sargento Real"
    Case 2
        TituloReal = "Teniente Real"
    Case 3
        TituloReal = "Comandante Real"
    Case 4
        TituloReal = "Teniente Real"
    Case 5
        TituloReal = "General Real"
    Case 6
        TituloReal = "Elite Real"
    Case 7
        TituloReal = "Guardian del Bien"
    Case 8
        TituloReal = "Caballero Imperial"
    Case 9
        TituloReal = "Justiciero"
    Case 10
        TituloReal = "Ejecutor Imperial"
    Case 11
        TituloReal = "Protector del Real"
    Case 12
        TituloReal = "Avatar de la Justicia"
    Case 13
        TituloReal = "Guardián del Bien"
    Case Else
        TituloReal = "Justiciero Del Bien"
End Select


End Function
Public Sub EnlistarCaos(ByVal UserIndex As Integer)
Dim Matados As Integer
Dim NextRecom As Long
Matados = (UserList(UserIndex).Faccion.CriminalesMatados + UserList(UserIndex).Faccion.ArmadaMatados + UserList(UserIndex).Faccion.CiudadanosMatados)


If Not Criminal(UserIndex) Then
Call WriteChatOverHead(UserIndex, "¡¡¡Largate de aqui bufón!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos Then
    Call WriteChatOverHead(UserIndex, "Ya perteneces a la horda del caos Traeme Mas almas.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If Matados < 30 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 30 enemigos, solo has matado " & Matados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 40 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes ser al menos Nivel 40.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

UserList(UserIndex).Faccion.FuerzasCaos = 1
UserList(UserIndex).Faccion.Rango = 1
'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex))
 
'------- Ropa -------
Dim MiObj As obj
Dim bajos As Byte
MiObj.Amount = 1
    
If UserList(UserIndex).raza = Enano Or UserList(UserIndex).raza = Gnomo Then
    bajos = 1
End If
    
Select Case UserList(UserIndex).clase
    Case eClass.Cleric
        MiObj.ObjIndex = 1500 + bajos
    Case eClass.Mage
        MiObj.ObjIndex = 1502 + bajos
    Case eClass.warrior
        MiObj.ObjIndex = 1504 + bajos
    Case eClass.Assasin
        MiObj.ObjIndex = 1506 + bajos
    Case eClass.Bard
        MiObj.ObjIndex = 1508 + bajos
    Case eClass.Druid
        MiObj.ObjIndex = 1510 + bajos
    Case eClass.bountyhunter
        MiObj.ObjIndex = 1512 + bajos
    Case eClass.Paladin
        MiObj.ObjIndex = 1514 + bajos
    Case eClass.hunter
        MiObj.ObjIndex = 1516 + bajos
    Case eClass.drakkar
        MiObj.ObjIndex = 1518 + bajos
    Case eClass.Nigromante
        MiObj.ObjIndex = 1520 + bajos
End Select


If Not MeterItemEnInventario(UserIndex, MiObj) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If

'------- Ropa -------

Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido a la Horda del Caos!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

End Sub
Public Sub RecompensaCaos(ByVal UserIndex As Integer)
Dim Matados As Long

If UserList(UserIndex).Faccion.Rango = 10 Then
    Exit Sub
End If

Matados = UserList(UserIndex).Faccion.CriminalesMatados + UserList(UserIndex).Faccion.CaosMatados + UserList(UserIndex).Faccion.ArmadaMatados + UserList(UserIndex).Faccion.CiudadanosMatados

If Matados < (UserList(UserIndex).Faccion.Rango) Then
    Call WriteChatOverHead(UserIndex, "Mata " & matadosCaos(UserList(UserIndex).Faccion.Rango) - Matados & " enemigos más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

UserList(UserIndex).Faccion.Rango = UserList(UserIndex).Faccion.Rango + 1
If UserList(UserIndex).Faccion.Rango >= 6 Then ' Segunda jeraquia xD
    Dim MiObj As obj
    MiObj.Amount = 1
    Dim bajos As Byte
    
    If UserList(UserIndex).raza = Enano Or UserList(UserIndex).raza = Gnomo Then
        bajos = 1
    End If
        
    Select Case UserList(UserIndex).clase
        Case eClass.Cleric
            MiObj.ObjIndex = 1522 + bajos
        Case eClass.Mage
            MiObj.ObjIndex = 1524 + bajos
        Case eClass.warrior
            MiObj.ObjIndex = 1526 + bajos
        Case eClass.Assasin
            MiObj.ObjIndex = 1528 + bajos
        Case eClass.Bard
            MiObj.ObjIndex = 1530 + bajos
        Case eClass.Druid
            MiObj.ObjIndex = 1532 + bajos
        Case eClass.bountyhunter
            MiObj.ObjIndex = 1534 + bajos
        Case eClass.Paladin
            MiObj.ObjIndex = 1536 + bajos
        Case eClass.hunter
            MiObj.ObjIndex = 1538 + bajos
        Case eClass.drakkar
            MiObj.ObjIndex = 1540 + bajos
        Case eClass.Nigromante
            MiObj.ObjIndex = 1542 + bajos
    End Select

    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    If TituloCaos(UserIndex) = "Avatar del Apocalipsis" Then
    Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores Subditos. Me trajiste " & Matados & ", Almas. Ya no tengo más recompensa para darte la vida eterna. ¡Felicidades!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    End If
    
    Call WriteChatOverHead(UserIndex, "¡¡¡Aqui tienes tu recompensa " + TituloCaos(UserIndex) + "!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
End If
End Sub
Public Function TituloCaos(ByVal UserIndex As Integer) As String
    Select Case UserList(UserIndex).Faccion.Rango
        Case 1
            TituloCaos = "Miembro de las Hordas"
        Case 2
            TituloCaos = "Guerrero del Caos"
        Case 3
            TituloCaos = "Teniente del Caos"
        Case 4
            TituloCaos = "Comandante del Caos"
        Case 5
            TituloCaos = "General del Caos"
        Case 6
            TituloCaos = "Elite del Caos"
        Case 7
            TituloCaos = "Asolador de las Sombras"
        Case 8
            TituloCaos = "Caballero Negro"
        Case 9
            TituloCaos = "Emisario de las Sombras"
        Case 10
            TituloCaos = "Avatar del Apocalipsis"
    End Select
End Function
Public Function matadosArmada(ByVal Rango As Byte) As Integer
    Select Case Rango
        Case 1
            matadosArmada = 25
        Case 2
            matadosArmada = 35
        Case 3
            matadosArmada = 45
        Case 4
            matadosArmada = 55
        Case 5
            matadosArmada = 65
        Case 6
            matadosArmada = 75
        Case 7
            matadosArmada = 85
        Case 8
            matadosArmada = 95
        Case 9
            matadosArmada = 105
    End Select
End Function
Public Function matadosCaos(ByVal Rango As Byte) As Integer
    Select Case Rango
        Case 1
            matadosCaos = 40
        Case 2
            matadosCaos = 50
        Case 3
            matadosCaos = 60
        Case 4
            matadosCaos = 70
        Case 5
            matadosCaos = 80
        Case 6
            matadosCaos = 90
        Case 7
            matadosCaos = 100
        Case 8
            matadosCaos = 120
        Case 9
            matadosCaos = 130
    
    End Select
End Function

