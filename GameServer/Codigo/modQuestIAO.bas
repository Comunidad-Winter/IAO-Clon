Attribute VB_Name = "modQuestIAO"
'Sistema de Quest (Estilo IAO 1.2)
'Fuente: Mithrandir
 
'Adaptado por Crip
 
Public Type UserQuest
    QuestIndex As Byte
    CantidadMatada As Integer
   
    IndexNix As Byte
    IndexUllathorpe As Byte
End Type
 
Public Type Quest
    desc As String
    Tipo As Byte
    NumPremio As Integer
    CantidadPremio As Integer
    TargetNPC As Integer
    CantObjetivos As Integer
    TargetUser As Integer
    MapaCiudad As Byte
    IndexDarle As Byte
End Type
 
Public Enum eFormQuest
    FormInformacionQuest
    FormQuest
End Enum
 
Public QuestList() As Quest
Public Sub LoadQuest()
'**************************************************************
'Author:
'Last Modify Date:
'
'**************************************************************
On Error GoTo ErrHandler
 
Dim numQuest As Integer
Dim i As Long
 
    
    Set Leer = New clsIniReader
 
Call Leer.Initialize(DatPath & "Quests.dat")
numQuest = val(Leer.GetValue("INIT", "NumQuests"))
 
ReDim Preserve QuestList(1 To numQuest) As Quest
 
For i = 1 To numQuest
   With QuestList(i)
       .desc = Leer.GetValue("Quest" & i, "Desc")
       .Tipo = val(Leer.GetValue("Quest" & i, "Tipo"))
 
       .NumPremio = val(Leer.GetValue("Quest" & i, "Premio"))
       .CantidadPremio = val(Leer.GetValue("Quest" & i, "Cantidad"))
 
       .TargetNPC = val(Leer.GetValue("Quest" & i, "TargetNPC"))
       .TargetUser = val(Leer.GetValue("Quest" & i, "TargetUser"))
       .CantObjetivos = val(Leer.GetValue("Quest" & i, "CantObjetivos"))
 
       .MapaCiudad = val(ReadField(1, Leer.GetValue("Quest" & i, "Ciudad"), 45))
       .IndexDarle = val(ReadField(2, Leer.GetValue("Quest" & i, "Ciudad"), 45))
   End With
Next i
Set Leer = Nothing
Exit Sub
ErrHandler:
   MsgBox "error cargando quest " & Err.Number & ": " & Err.description
End Sub
Public Sub AbandonarQuest(ByVal UserIndex As Integer)
'**************************************************************
'Author:
'Last Modify Date:
'
'**************************************************************
 
With UserList(UserIndex)
 
    If .Quest.QuestIndex <> 0 Then
        Call ResetiarQuestUsuario(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has decidido abandonar tu búsqueda.", FontTypeNames.FONTTYPE_QUEST)
    Else
        Call WriteConsoleMsg(UserIndex, "Actualmente no estas en una quest.", FontTypeNames.FONTTYPE_QUEST)
    End If
   
End With
End Sub
 
Private Function PuedeAceptarQuest(ByVal UserIndex As Integer) As Boolean
'**************************************************************
'Author:
'Last Modify Date:
'
'**************************************************************
Dim Puede As Boolean
 
PuedeAceptarQuest = UserList(UserIndex).Quest.QuestIndex = 0
 
End Function
 
Public Sub AceptarQuest(ByVal UserIndex As Integer)
'**************************************************************
'Author:
'Last Modify Date:
'
'**************************************************************
 
With UserList(UserIndex)
   
    Dim NpcIndex As Integer
   
    NpcIndex = .flags.TargetNPC
   
    If .Quest.QuestIndex <> 0 Then
        Call WriteConsoleMsg(UserIndex, "No puedes aceptar la propuesta ya que aún tienes objetivos pendientes.", FontTypeNames.FONTTYPE_QUEST)
        Exit Sub
    End If
   
    'esto esta de sobra pero por las dudas..
    If NpcIndex = 0 Or Npclist(NpcIndex).NPCtype <> eNPCType.Quest Then
        Exit Sub
    End If
   
    If Npclist(NpcIndex).QuestIndex < 1 Or Npclist(NpcIndex).QuestIndex > UBound(QuestList) Then
        Call WriteConsoleMsg(UserIndex, "El Npc Tiene una quest invalida", FontTypeNames.FONTTYPE_QUEST)
        Exit Sub
    End If
   
    Call ResetiarQuestUsuario(UserIndex)
   
    .Quest.QuestIndex = Npclist(NpcIndex).QuestIndex
    Call WriteConsoleMsg(UserIndex, "Has aceptado la propuesta. Revisa tus objetivos.", FontTypeNames.FONTTYPE_QUEST)
   
End With
End Sub
 
Public Sub EnviarFormularioQuestInfo(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'**************************************************************
'Author:
'Last Modify Date:
'
'**************************************************************
With UserList(UserIndex)
   
    If .Quest.QuestIndex <> 0 Then
        Call WriteConsoleMsg(UserIndex, "Ya estas en una quest", FontTypeNames.FONTTYPE_QUEST)
        Exit Sub
    End If
   
    'esto esta de sobra pero por las dudas..
    If NpcIndex = 0 Or Npclist(NpcIndex).NPCtype <> eNPCType.Quest Then
        Exit Sub
    End If
   
    If Npclist(NpcIndex).QuestIndex < 1 Or Npclist(NpcIndex).QuestIndex > UBound(QuestList) Then
        Call WriteConsoleMsg(UserIndex, "El Npc Tiene una quest invalida", FontTypeNames.FONTTYPE_QUEST)
        Exit Sub
    End If
   
    If Distancia(.Pos, Npclist(NpcIndex).Pos) > 10 Then
       Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_QUEST)
       Exit Sub
    End If
 
    Call WriteQuestForm(UserIndex, FormInformacionQuest)
    Call WriteQuestList(UserIndex, True, Npclist(NpcIndex).QuestIndex)
 
End With
End Sub
 
Public Sub InfoQuest(ByVal UserIndex As Integer)
'**************************************************************
'Author:
'Last Modify Date:
'
'**************************************************************
 
With UserList(UserIndex)
   
     If .Quest.QuestIndex <> 0 Then
        Call WriteQuestForm(UserIndex, FormInformacionQuest)
        Call WriteQuestList(UserIndex, False, .Quest.QuestIndex)
    Else
        Call WriteConsoleMsg(UserIndex, "Actualmente no estas en una quest.", FontTypeNames.FONTTYPE_QUEST)
    End If
 
End With
End Sub
 
Public Sub ResetiarQuestUsuario(ByVal UserIndex As Integer)
'**************************************************************
'Author:
'Last Modify Date:
'
'**************************************************************
 
With UserList(UserIndex)
 
    .Quest.QuestIndex = 0
    .Quest.CantidadMatada = 0
    .Quest.IndexNix = 0
    .Quest.IndexUllathorpe = 0
 
End With
End Sub
 
Public Sub FinalizarQuest(ByVal UserIndex As Integer)
'**************************************************************
'Author:
'Last Modify Date:
'
'**************************************************************
Dim item As obj
 
With UserList(UserIndex)
 
    If .Quest.QuestIndex = 0 Then
        Exit Sub
    End If
   
    item.ObjIndex = QuestList(.Quest.QuestIndex).NumPremio
    item.Amount = QuestList(.Quest.QuestIndex).CantidadPremio
   
    Call WriteConsoleMsg(UserIndex, "¡Felicidades! Has completado tus objetivos.", FontTypeNames.FONTTYPE_QUEST)
   
    If item.ObjIndex = 12 Then
        .Stats.GLD = .Stats.GLD + item.Amount
        Call WriteUpdateGold(UserIndex)
        Call WriteConsoleMsg(UserIndex, "¡Has ganado " & item.Amount & " Monedas de oro", FontTypeNames.FONTTYPE_QUEST)
    Else
        If Not MeterItemEnInventario(UserIndex, item) Then
            Call TirarItemAlPiso(.Pos, item)
        End If
       
        Call WriteConsoleMsg(UserIndex, "Has obtenido " & ObjData(item.ObjIndex).name & " (" & item.Amount & ")", FontTypeNames.FONTTYPE_QUEST)
    End If
   
    Call ResetiarQuestUsuario(UserIndex)
   
End With
End Sub

