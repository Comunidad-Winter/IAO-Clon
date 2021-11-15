Attribute VB_Name = "ModClimas"
'********************************Modulo Climas*********************************
'Author: Manuel (Lorwik)
'Last Modification: 23/11/2014
'Controla el clima y lo envia al cliente.
'******************************************************************************

Public ClimaColor As Byte

Option Explicit

'******************************************************************************
'Sorteo Del Clima
'******************************************************************************
Public Function SortearClima()
'Sorteamos el clima, si hay tormenta y es de Mañana o de Dia ponemos el efecto de tarde,
'pero si es de Tarde o de Noche no ponemos nigun efecto.

    If Hour(Now) >= 6 And Hour(Now) < 12 Then
        Call SetearClima(0)
        frmMain.Clima.Caption = "Clima: Mañana"
    ElseIf Hour(Now) >= 12 And Hour(Now) < 18 Then
        Call SetearClima(1)
        frmMain.Clima.Caption = "Clima: MedioDia"
    ElseIf Hour(Now) >= 18 And Hour(Now) < 20 Then
        Call SetearClima(2)
        frmMain.Clima.Caption = "Clima: Tarde"
    ElseIf Hour(Now) >= 20 And Hour(Now) < 6 Then
        Call SetearClima(3)
        frmMain.Clima.Caption = "Clima: Noche"
    End If
End Function

Public Function SetearClima(Clima As Byte)
Dim UserIndex As Integer
Dim i As Long

    Select Case Clima
        Case 0
            'mañana
            ClimaColor = 0
            
        Case 1
            'mediodia
            ClimaColor = 1
    
        Case 2
            'tarde
            ClimaColor = 2
            
        Case 3
            'noche
            ClimaColor = 3
    End Select
    
    For i = 1 To LastUser
        Call writeNoche(i, ClimaColor)
    Next i
End Function

