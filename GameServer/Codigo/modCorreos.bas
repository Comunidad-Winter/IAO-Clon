Attribute VB_Name = "modCorreos"
Option Explicit
 
 
Public Const Max_Correos As Byte = 10
Public Type tCorreos
 
      Carta As String
      Emisor As String
      Leida As Byte
      ObjetoIndex As Integer
      ObjetoCantidad As Integer
 
End Type
Public Sub EnviarCorreo(ByVal UserIndex As Integer, ByVal Destinatario As String, ByVal Mensaje As String, ByVal ObjIndex As Integer, ByVal AmountIndex As Integer)
   
    Dim Slot As Byte
        Slot = getFreeSlotCorreo(Destinatario)
       
        Call WriteVar(App.Path & "\Data\Characters Created\" & Destinatario & ".pjs", "CORREO", "EMISOR" & Slot, UserList(UserIndex).name)
        Call WriteVar(App.Path & "\Data\Characters Created\" & Destinatario & ".pjs", "CORREO", "CARTA" & Slot, Mensaje)
        Call WriteVar(App.Path & "\Data\Characters Created\" & Destinatario & ".pjs", "CORREO", "OBJETO" & Slot, ObjIndex & "-" & AmountIndex)
       
   
End Sub
Public Function CantSendCorreo(ByVal UserIndex As Integer, _
                               ByVal Destinatario As String) As Boolean
 
      CantSendCorreo = False
     
      If Not (getFreeSlotCorreo(Destinatario) <> -1) Then
            ' @@ El target no tiene suficiente espacio en la lista de correos :P
            Exit Function
      End If
 
      CantSendCorreo = True
End Function
Private Function getFreeSlotCorreo(ByVal UserName As String) As Integer
      ' @@ damos un slot del correo para generar un nuevo correo
 
      Dim Index As Long
      Dim Carta As String
   
      For Index = 1 To Max_Correos
        Carta = GetVar(CharPath & UserName & ".pjs", "CORREO", "CARTA" & Index)
       
            If Carta = vbNullString Then
                getFreeSlotCorreo = Index
                Exit Function
            End If
           
      Next Index
 
      getFreeSlotCorreo = -1
End Function
Sub ResetCorreos(ByVal UserIndex As Integer, Optional ByVal Index As Byte = 0)
      '//Shak - Sistema de correos
 
      Dim C As Long
 
      With UserList(UserIndex)
 
 
        If Index <> 0 Then
            With .Correos(Index)
                .Carta = vbNullString
                .Emisor = vbNullString 'Nombre de quien manda la carta.
                .Leida = 0
                .ObjetoCantidad = 0
                .ObjetoIndex = 0
            End With
        Else
            For C = 1 To Max_Correos
 
                  With .Correos(C)
                        .Carta = vbNullString
                        .Emisor = vbNullString 'Nombre de quien manda la carta.
                        .Leida = 0
                        .ObjetoCantidad = 0
                        .ObjetoIndex = 0
                  End With
 
            Next C
        End If
 
      End With
 
End Sub
Public Sub LeerCorreos(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
' Leemos los correos
      Dim C As Long
 
      With UserList(UserIndex)
 
            For C = 1 To Max_Correos
 
                  With .Correos(C)
                     
                        .Carta = UserFile.GetValue("CORREO", "Carta" & C)
                        .Emisor = UserFile.GetValue("CORREO", "Emisor" & C)
                        .Leida = UserFile.GetValue("CORREO", "Leida" & C)
                        .ObjetoIndex = CInt(ReadField(1, UserFile.GetValue("CORREO", "Objeto" & C), 45))
                        .ObjetoCantidad = CInt(ReadField(2, UserFile.GetValue("CORREO", "Objeto" & C), 45))
 
                  End With
 
            Next C
 
      End With
 
End Sub
Public Sub GuardarCorreos(ByVal UserIndex As Integer, ByVal UserFile As String)
'//Guardamos los correos
      Dim C As Long
 
      With UserList(UserIndex)
 
            For C = 1 To Max_Correos
 
                  With .Correos(C)
                 
                        Call WriteVar(UserFile, "CORREO", "Carta" & C, CStr(.Carta))
                        Call WriteVar(UserFile, "CORREO", "Emisor" & C, CStr(.Emisor))
                        Call WriteVar(UserFile, "CORREO", "Leida" & C, CStr(.Leida))
                        Call WriteVar(UserFile, "CORREO", "Objeto" & C, .ObjetoIndex & "-" & .ObjetoCantidad)
     
                  End With
 
            Next C
 
      End With
 
End Sub
