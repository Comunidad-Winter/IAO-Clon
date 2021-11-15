Attribute VB_Name = "Mod_Viajes"
Option Explicit
 
Public Sub DondeViaje(ByVal UserIndex As Integer, ByVal P As Byte)
 
    With UserList(UserIndex)
        Select Case P
            Case 0 'Nix
               If .Stats.GLD < 1700 Then
               Call WriteConsoleMsg(UserIndex, "No tienes suficiente oro.", FontTypeNames.FONTTYPE_INFO)
               Exit Sub
               End If
               Call WarpUserChar(UserIndex, 1, 27, 78, True)
               .Stats.GLD = .Stats.GLD - 1700
               Call WriteUpdateGold(UserIndex)
            Case 1 ' Banderbill
                If .Stats.GLD < 1850 Then
               Call WriteConsoleMsg(UserIndex, "No tienes suficiente oro.", FontTypeNames.FONTTYPE_INFO)
               Exit Sub
               End If
               Call WarpUserChar(UserIndex, 1, 51, 59, True)
               .Stats.GLD = .Stats.GLD - 1850
               Call WriteUpdateGold(UserIndex)
            Case 2 ' Lindos
              If .Stats.GLD < 1900 Then
               Call WriteConsoleMsg(UserIndex, "No tienes suficiente oro.", FontTypeNames.FONTTYPE_INFO)
               Exit Sub
               End If
               Call WarpUserChar(UserIndex, 1, 53, 83, True)
               .Stats.GLD = .Stats.GLD - 1900
               Call WriteUpdateGold(UserIndex)
            Case 3 ' Ullathorpe
                If .Stats.GLD < 1200 Then
               Call WriteConsoleMsg(UserIndex, "No tienes suficiente oro.", FontTypeNames.FONTTYPE_INFO)
               Exit Sub
               End If
               Call WarpUserChar(UserIndex, 1, 50, 50, True)
               .Stats.GLD = .Stats.GLD - 1200
               Call WriteUpdateGold(UserIndex)
        End Select
   End With
   
End Sub
