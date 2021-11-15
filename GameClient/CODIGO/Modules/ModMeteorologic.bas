Attribute VB_Name = "ModMeteorologic"
Option Explicit

Private Type RGBClimax
    r As Byte
    g As Byte
    b As Byte
End Type
 
Public ColorClimax As RGBClimax
Public ColorClima As Byte

Public Sub ClimaX()
frmMain.imgHora.Picture = LoadPicture(App.path & "\Recursos\Interface\" & Hour(time) & ".jpg")

If UserEstado = 1 Or RTrim$(MapDat.zone) = "DUNGEON" Then
    base_light = General_RGB_Color_to_Long(160, 160, 160, 255)
Exit Sub
End If

Select Case ColorClima
    'Mañana
    Case 0
        base_light = General_RGB_Color_to_Long(200, 200, 230, 255)
    'MedioDia
    Case 1
        base_light = General_RGB_Color_to_Long(255, 255, 255, 255)
    'Tarde
    Case 2
        base_light = General_RGB_Color_to_Long(230, 200, 200, 255)
    'Noche
    Case 3
        base_light = General_RGB_Color_to_Long(170, 170, 170, 255)
       
End Select
End Sub
