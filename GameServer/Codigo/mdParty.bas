Attribute VB_Name = "mdParty"


Option Explicit

Public Const MAX_PARTIES As Integer = 300

''
'nivel minimo para crear party
Public Const MINPARTYLEVEL As Byte = 15

''
'Cantidad maxima de gente en la party
Public Const PARTY_MAXMEMBERS As Byte = 5

Public Const PARTY_EXPERIENCIAPORGOLPE As Boolean = True

''
'maxima diferencia de niveles permitida en una party
Public Const MAXPARTYDELTALEVEL As Byte = 7

''
'distancia al leader para que este acepte el ingreso
Public Const MAXDISTANCIAINGRESOPARTY As Byte = 2

''
'maxima distancia a un exito para obtener su experiencia
Public Const PARTY_MAXDISTANCIA As Byte = 18

''
'restan las muertes de los miembros?
Public Const CASTIGOS As Boolean = False

Public ExponenteNivelParty As Single

Public Type tPartyMember
    UserIndex As Integer
    Experiencia As Double
End Type


Public Function NextParty() As Integer
Dim i As Integer
NextParty = -1
For i = 1 To MAX_PARTIES
    If Parties(i) Is Nothing Then
        NextParty = i
        Exit Function
    End If
Next i
End Function

Public Function PuedeCrearParty(ByVal UserIndex As Integer) As Boolean
    'If UserList(UserIndex).Stats.ELV < MINPARTYLEVEL Then

    'If CInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma)) * UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) < 100 Then
    '    Call WriteConsoleMsg( UserIndex, "Tu carisma y liderazgo no son suficientes para liderar un grupo.", FontTypeNames.FONTTYPE_INFO)
    '    PuedeCrearParty = False
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "Est�s muerto!", FontTypeNames.FONTTYPE_INFO)
       PuedeCrearParty = False
    End If
    
    PuedeCrearParty = True
End Function

Public Sub CrearParty(ByVal UserIndex As Integer)
Dim tInt As Integer
If UserList(UserIndex).PartyIndex = 0 Then
    If UserList(UserIndex).flags.Muerto = 0 Then
        'If UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) >= 5 Then
            tInt = mdParty.NextParty
            If tInt = -1 Then
                Call WriteConsoleMsg(UserIndex, "Por el momento no se pueden crear mas grupos", FontTypeNames.FONTTYPE_INFOBOLD)
                Exit Sub
            Else
                Set Parties(tInt) = New clsParty
                If Not Parties(tInt).NuevoMiembro(UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "El grupo est� lleno, no puedes entrar", FontTypeNames.FONTTYPE_INFOBOLD)
                    Set Parties(tInt) = Nothing
                    Exit Sub
                Else
                    Call WriteConsoleMsg(UserIndex, "�Has formado el grupo!", FontTypeNames.FONTTYPE_INFOBOLD)
                    UserList(UserIndex).PartyIndex = tInt
                    UserList(UserIndex).PartySolicitud = 0
                    If Not Parties(tInt).HacerLeader(UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "No puedes hacerte l�der.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "� Te has convertido en l�der del grupo !", FontTypeNames.FONTTYPE_INFOBOLD)
                    End If
                End If
           End If
       'Else
            'Call WriteConsoleMsg( UserIndex, " No tienes suficientes puntos de liderazgo para liderar una party.", FontTypeNames.FONTTYPE_INFO)
        'End If
    Else
        Call WriteConsoleMsg(UserIndex, "Est�s muerto!", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(UserIndex, " Ya perteneces a una party.", FontTypeNames.FONTTYPE_INFOBOLD)
End If
End Sub

Public Sub SolicitarIngresoAParty(ByVal UserIndex As Integer)
'ESTO ES enviado por el PJ para solicitar el ingreso a la party
Dim tInt As Integer

    If UserList(UserIndex).PartyIndex > 0 Then
        'si ya esta en una party
        Call WriteConsoleMsg(UserIndex, " Ya perteneces a una party, escribe /SALIRGRUPO para abandonarla", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).PartySolicitud = 0
        Exit Sub
    End If
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, " �Est�s muerto!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).PartySolicitud = 0
        Exit Sub
    End If
    tInt = UserList(UserIndex).flags.TargetUser
    If tInt > 0 Then
        If UserList(tInt).PartyIndex > 0 Then
            UserList(UserIndex).PartySolicitud = UserList(tInt).PartyIndex
            Call WriteConsoleMsg(UserIndex, " El fundador decidir� si te acepta en la party", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, UserList(tInt).name & " no es fundador de ninguna party.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).PartySolicitud = 0
            Exit Sub
        End If
    Else
        Call WriteConsoleMsg(UserIndex, " Para ingresar a una party debes hacer click sobre el fundador y luego escribir /PARTY", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).PartySolicitud = 0
    End If
End Sub

Public Sub SalirDeParty(ByVal UserIndex As Integer)
Dim PI As Integer
PI = UserList(UserIndex).PartyIndex
If PI > 0 Then
    If Parties(PI).SaleMiembro(UserIndex) Then
        'sale el leader
        Set Parties(PI) = Nothing
    Else
        UserList(UserIndex).PartyIndex = 0
    End If
Else
    Call WriteConsoleMsg(UserIndex, " No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Public Sub ExpulsarDeParty(ByVal leader As Integer, ByVal OldMember As Integer)
Dim PI As Integer
PI = UserList(leader).PartyIndex

If PI = UserList(OldMember).PartyIndex Then
    If Parties(PI).SaleMiembro(OldMember) Then
        'si la funcion me da true, entonces la party se disolvio
        'y los partyindex fueron reseteados a 0
        Set Parties(PI) = Nothing
    Else
        UserList(OldMember).PartyIndex = 0
    End If
Else
    Call WriteConsoleMsg(leader, LCase(UserList(OldMember).name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

''
' Determines if a user can use party commands like /acceptparty or not.
'
' @param User Specifies reference to user
' @return  True if the user can use party commands, false if not.
Public Function UserPuedeEjecutarComandos(ByVal User As Integer) As Boolean
'*************************************************
'Author: Marco Vanotti(Marco)
'Last modified: 05/05/09
'
'*************************************************
    Dim PI As Integer
    
    PI = UserList(User).PartyIndex
    
    If PI > 0 Then
        If Parties(PI).EsPartyLeader(User) Then
            UserPuedeEjecutarComandos = True
        Else
            Call WriteConsoleMsg(User, "�No eres el l�der de tu Party!", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    Else
        Call WriteConsoleMsg(User, "No eres miembro de ninguna party.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
End Function

Public Sub AprobarIngresoAParty(ByVal leader As Integer, ByVal NewMember As Integer)
'el UI es el leader
Dim PI As Integer
Dim razon As String

PI = UserList(leader).PartyIndex

If UserList(NewMember).PartySolicitud = PI Then
    If Not UserList(NewMember).flags.Muerto = 1 Then
        If UserList(NewMember).PartyIndex = 0 Then
            If Parties(PI).PuedeEntrar(NewMember, razon) Then
                If Parties(PI).NuevoMiembro(NewMember) Then
                    Call Parties(PI).MandarMensajeAConsola(UserList(leader).name & " ha aceptado a " & UserList(NewMember).name & " en la party.", "Servidor")
                    UserList(NewMember).PartyIndex = PI
                    UserList(NewMember).PartySolicitud = 0
                Else
                    'no pudo entrar
                    'ACA UNO PUEDE CODIFICAR OTRO TIPO DE ERRORES...
                    Call SendData(SendTarget.ToAdmins, leader, PrepareMessageConsoleMsg(" Servidor> CATASTROFE EN PARTIES, NUEVOMIEMBRO DIO FALSE! :S ", FontTypeNames.FONTTYPE_INFO))
                    End If
                Else
                'no debe entrar
                Call WriteConsoleMsg(leader, razon, FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(leader, UserList(NewMember).name & " ya es miembro de otra party.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        Call WriteConsoleMsg(leader, "�Est� muerto, no puedes aceptar miembros en ese estado!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
Else
    Call WriteConsoleMsg(leader, LCase(UserList(NewMember).name) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

End Sub

Public Sub BroadCastParty(ByVal UserIndex As Integer, ByRef texto As String)
Dim PI As Integer
    
    PI = UserList(UserIndex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).MandarMensajeAConsola(texto, UserList(UserIndex).name)
    End If

End Sub

Public Sub OnlineParty(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 11/27/09 (Budi)
'Adapte la funci�n a los nuevos m�todos de clsParty
'*************************************************
Dim i As Integer
Dim PI As Integer
Dim Text As String
Dim MembersOnline(1 To PARTY_MAXMEMBERS) As Integer
    PI = UserList(UserIndex).PartyIndex
    
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(MembersOnline)
        Text = "Nombre(Exp): "
        For i = 1 To PARTY_MAXMEMBERS
            If MembersOnline(i) > 0 Then
                Text = Text & " - " & UserList(MembersOnline(i)).name & " (" & Fix(Parties(PI).MiExperiencia(MembersOnline(i))) & ")"
            End If
        Next i
        Text = Text & ". Experiencia total: " & Parties(PI).ObtenerExperienciaTotal
        Call WriteConsoleMsg(UserIndex, Text, FontTypeNames.FONTTYPE_INFO)
    End If
    
End Sub


Public Sub TransformarEnLider(ByVal OldLeader As Integer, ByVal NewLeader As Integer)
Dim PI As Integer

If OldLeader = NewLeader Then Exit Sub

PI = UserList(OldLeader).PartyIndex

If PI = UserList(NewLeader).PartyIndex Then
    If UserList(NewLeader).flags.Muerto = 0 Then
        If Parties(PI).HacerLeader(NewLeader) Then
            Call Parties(PI).MandarMensajeAConsola("El nuevo l�der de la party es " & UserList(NewLeader).name, UserList(OldLeader).name)
        Else
            Call WriteConsoleMsg(OldLeader, "�No se ha hecho el cambio de mando!", FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(OldLeader, "�Est� muerto!", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(OldLeader, LCase(UserList(NewLeader).name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub


Public Sub ActualizaExperiencias()
'esta funcion se invoca antes de worlsaves, y apagar servidores
'en caso que la experiencia sea acumulada y no por golpe
'para que grabe los datos en los charfiles
Dim i As Integer

If Not PARTY_EXPERIENCIAPORGOLPE Then
    
    haciendoBK = True
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Distribuyendo experiencia en parties.", FontTypeNames.FONTTYPE_SERVER))
    For i = 1 To MAX_PARTIES
        If Not Parties(i) Is Nothing Then
            Call Parties(i).FlushExperiencia
        End If
    Next i
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Experiencia distribuida.", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    haciendoBK = False

End If

End Sub
Public Sub ObtenerExito(ByVal UserIndex As Integer, ByVal Exp As Long, mapa As Integer, x As Integer, Y As Integer)
    If Exp <= 0 Then
        If Not CASTIGOS Then Exit Sub
    End If
    
    Call Parties(UserList(UserIndex).PartyIndex).ObtenerExito(Exp, mapa, x, Y)

End Sub

Public Function CantMiembros(ByVal UserIndex As Integer) As Integer
CantMiembros = 0
If UserList(UserIndex).PartyIndex > 0 Then
    CantMiembros = Parties(UserList(UserIndex).PartyIndex).CantMiembros
End If

End Function

Public Sub ActualizarSumaNivelesElevados(ByVal UserIndex As Integer)
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 28/10/08
'
'*************************************************
    If UserList(UserIndex).PartyIndex > 0 Then
        Call Parties(UserList(UserIndex).PartyIndex).UpdateSumaNivelesElevados(UserList(UserIndex).Stats.ELV)
    End If
End Sub




