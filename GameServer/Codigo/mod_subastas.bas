Attribute VB_Name = "mod_subastas"
Option Explicit
 
Private Type tSubasta
    HaySubasta As Boolean
    OfertadorNombre As String
    SubastadorNombre As String
    ObjetoSubastado As obj
    PujaAnterio As Long
    PrecioInicial As Long
End Type
 
Private Const NPCSubastas As Byte = 164
Private Const Duracion_Subasta As Byte = 2
Private Subasta As tSubasta
Private SegundosSubasta As Integer
Private MinutosSubasta As Integer
Public Sub PasarMinutoSubasta()
With Subasta
If .HaySubasta = True Then
MinutosSubasta = MinutosSubasta + 1
If SegundosSubasta >= MinutosSubasta Then 'Termino la subasta
If .PujaAnterio <> 0 Then 'alguien oferto
Call EntregarObjetoAOfertador
Call ReniciarSubasta
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("La subasta finalizo", FontTypeNames.FONTTYPE_INFO))
Else 'nadie oferto
Call EntregarObjetoASubastador
Call ReniciarSubasta
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("La subasta finalizo", FontTypeNames.FONTTYPE_INFO))
End If
Else
'Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El usuario " & .SubastadorNombre & " esta subastando " & .ObjetoSubastado.Amount & "-" & ObjData(.ObjetoSubastado_ObjIndex).name & " a un precio inicial de " & Subasta.PrecioInicial & ".", FontTypeNames.FONTTYPE_INFO))
End If
End If
End With
End Sub
 
Private Sub EntregarObjetoAOfertador()
 
    Dim OfertadorIndex As Integer
    Dim SubastadorIndex As Integer
 
    OfertadorIndex = NameIndex(Subasta.OfertadorNombre)
    SubastadorIndex = NameIndex(Subasta.SubastadorNombre)
   
    If (OfertadorIndex <= 0 And SubastadorIndex <= 0) Then
        Call LogCriticEvent("Error: Subasta...")
        Exit Sub
    End If
   
    If (OfertadorIndex <> 0) Then 'Esta conectado el usuario que oferto.
        If Not MeterItemEnInventario(OfertadorIndex, Subasta.ObjetoSubastado) Then _
            Call TirarItemAlPiso(UserList(OfertadorIndex).Pos, Subasta.ObjetoSubastado)
    End If
   
    If (SubastadorIndex <> 0) Then 'Esta conectado el usuario que subasto.
        UserList(SubastadorIndex).Stats.GLD = UserList(SubastadorIndex).Stats.GLD + Subasta.PujaAnterio
        Call WriteUpdateGold(SubastadorIndex)
    End If
   
End Sub
Private Sub EntregarObjetoASubastador()
 
Dim UserIndex As Integer
 
UserIndex = NameIndex(Subasta.SubastadorNombre)
 
If (UserIndex <= 0) Then
'Call LogCriticEvent("No se logro entregar el objeto [" & Subasta.ObjetoSubastado & " cantidad: " & Subasta.ObjetoSubastado.ObjIndex & " a->" & Subasta.SubastadorNombre)
Else
    If Not MeterItemEnInventario(UserIndex, Subasta.ObjetoSubastado) Then _
         Call TirarItemAlPiso(UserList(UserIndex).Pos, Subasta.ObjetoSubastado)
End If
 
End Sub
Public Sub Subastar(ByVal UserIndex As Integer, ByVal PrecioInicial As Long)
   Dim NpcIndex As Integer
   Dim ObjTile As obj
   With UserList(UserIndex)
   NpcIndex = UserList(UserIndex).flags.TargetNPC
   ObjTile = MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo
   
   If UserList(UserIndex).flags.Muerto Then
       Call WriteConsoleMsg(UserIndex, "Estas Muerto", FontTypeNames.FONTTYPE_INFO)
       Exit Sub
   End If
   
   If NpcIndex = 0 Then
       Call WriteConsoleMsg(UserIndex, "Tiene que seleccionar al npc subastador", FontTypeNames.FONTTYPE_INFO)
       Exit Sub
   End If
   
   If Npclist(NpcIndex).NPCtype <> eNPCType.Subastador Then
       Call WriteConsoleMsg(UserIndex, "El npc seleccionado no es un subastador", FontTypeNames.FONTTYPE_INFO)
       Exit Sub
   End If
   
   If Subasta.HaySubasta Then
       Call WriteConsoleMsg(UserIndex, "Hay una subasta actualmente, espere hasta que esta termine", FontTypeNames.FONTTYPE_INFO)
       Exit Sub
   End If
   
   If ObjTile.ObjIndex = 0 Then
       Call WriteConsoleMsg(UserIndex, "No hay un item en el piso.", FontTypeNames.FONTTYPE_INFO)
       Exit Sub
   End If
   
   If PrecioInicial <= 0 Then
       Call WriteConsoleMsg(UserIndex, "No puede subastar por cantidades negativas", FontTypeNames.FONTTYPE_INFO)
       Exit Sub
   End If
   
   If Distancia(Npclist(NpcIndex).Pos, UserList(UserIndex).Pos) > 1 Then
       Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos", FontTypeNames.FONTTYPE_INFO)
       Exit Sub
   End If
   
   Call IniciarSubasta(UserIndex, PrecioInicial, ObjTile)
End With
End Sub
Private Sub IniciarSubasta(ByVal UserIndex As Integer, ByVal PrecioInicial As Long, ByRef Objeto As obj)
 
With UserList(UserIndex)
 
   Call ReniciarSubasta
   Call EraseObj(10000, .Pos.map, .Pos.X, .Pos.Y)
   Subasta.HaySubasta = True
   Subasta.ObjetoSubastado = Objeto
   Subasta.PrecioInicial = PrecioInicial
   Subasta.SubastadorNombre = .name
   Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El usuario " & .name & " esta subastando " & Objeto.Amount & "-" & ObjData(Objeto.ObjIndex).name & " a un precio inicial de " & PrecioInicial & ".", FontTypeNames.FONTTYPE_INFO))
   Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Escribar /OFERTAR ...", FontTypeNames.FONTTYPE_INFO))
   
End With
End Sub
Private Sub ReniciarSubasta()
   Subasta.HaySubasta = False
   Subasta.ObjetoSubastado.Amount = 0
   Subasta.ObjetoSubastado.ObjIndex = 0
   Subasta.PujaAnterio = 0
   Subasta.PrecioInicial = 0
   Subasta.SubastadorNombre = vbNullString
   Subasta.OfertadorNombre = vbNullString
   SegundosSubasta = 0
   MinutosSubasta = 0
End Sub
Public Sub ofertar(UserIndex As Integer, ByVal oferta As Long)
   With Subasta
       Dim UserName As String
       If .HaySubasta = False Then
           Call WriteConsoleMsg(UserIndex, "¡No hay subastas en este momento!", FontTypeNames.FONTTYPE_INFO)
           Exit Sub
       End If
     
     
       If UserList(UserIndex).Stats.GLD < oferta Then
           Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad", FontTypeNames.FONTTYPE_INFO)
           Exit Sub
       End If
     
     
       If oferta <= .PujaAnterio Then
           Call WriteConsoleMsg(UserIndex, "Tu Oferta debe ser mayor a " & .PujaAnterio & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
           Exit Sub
       End If
     
     
       If .OfertadorNombre <> 0 Then
           UserList(.SubastadorNombre).Stats.GLD = UserList(Subasta.SubastadorNombre).Stats.GLD - Subasta.SubastadorNombre
           Call SendUserOROTxtFromChar(Subasta.SubastadorNombre, UserName)
       End If
     
       Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El usuario " & UserList(UserIndex).name & " ha ofertado " & oferta & " monedas de oro.", FontTypeNames.FONTTYPE_INFO))
       .PujaAnterio = UserIndex
       .OfertadorNombre = UserIndex
     
   End With
   UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - oferta
   Call SendUserOROTxtFromChar(UserIndex, UserName)
End Sub
 

