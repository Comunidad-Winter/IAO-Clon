Attribute VB_Name = "ModMeteorologic"
'************************************************* ****************
'ImperiumAO - v1.0
'************************************************* ****************
'Copyright (C) 2015 Gaston Jorge Martinez
'Copyright (C) 2015 Alexis Rodriguez
'Copyright (C) 2015 Luis Merino
'Copyright (C) 2015 Girardi Luciano Valentin
'
'Respective portions copyright by taxpayers below.
'
'This library is free software; you can redistribute it and / or
'Modify it under the terms of the GNU General Public
'License as published by the Free Software Foundation version 2.1
'The License
'
'This library is distributed in the hope that it will be useful,
'But WITHOUT ANY WARRANTY; without even the implied warranty
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'************************************************* ****************
'
'************************************************* ****************
'You can contact me at:
'Gaston Jorge Martinez (Zenitram@Hotmail.com)
'************************************************* ****************

Option Explicit

Private Type RGBClimax
    r As Byte
    g As Byte
    b As Byte
End Type
 
Public ColorClimax As RGBClimax
Public ColorClima As Byte

Public Sub ClimaX()
frmMain.imgHora.Picture = LoadPicture(App.path & "\Resources\Interface\" & Hour(time) & ".jpg")

If UserEstado = 1 Or RTrim$(MapDat.zone) = "DUNGEON" Then

Exit Sub
End If

Select Case ColorClima
    'Mañana
    Case 0

    'MedioDia
    Case 1

    'Tarde
    Case 2

    'Noche
    Case 3

End Select
End Sub
