VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Dim InventoryOffset As Long
Dim InvSelectedItem As Long
Dim MouseDownSelectedItem As Long

Dim WithEvents InventoryWindow As PictureBox
Attribute InventoryWindow.VB_VarHelpID = -1

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu
#End If

Public Sub ScrollInventory(ByVal up As Boolean)

    If up Then
        If (InventoryOffset + (InventoryWindow.ScaleHeight \ TilePixelHeight)) * (InventoryWindow.ScaleWidth \ TilePixelWidth) < MAX_INVENTORY_SLOTS Then _
            InventoryOffset = InventoryOffset + 1
    Else
        If InventoryOffset > 0 Then _
            InventoryOffset = InventoryOffset - 1
    End If
    
End Sub
Public Sub SelectGold()

    Dim prevSelectedItem As Long
    
    prevSelectedItem = InvSelectedItem
    
    InvSelectedItem = FLAGORO
    
End Sub
Public Property Get SelectedItem() As Long

    SelectedItem = InvSelectedItem
    
End Property
Public Property Get MaxHit(ByVal slot As Byte) As Integer

    MaxHit = UserInventory(slot).MaxHit
    
End Property
Public Property Get MinHit(ByVal slot As Byte) As Integer

    MinHit = UserInventory(slot).MinHit
    
End Property
Public Property Get Def(ByVal slot As Byte) As Integer

    Def = UserInventory(slot).Def
    
End Property
Public Property Get grhindex(ByVal slot As Byte) As Integer

    grhindex = UserInventory(slot).grhindex
    
End Property
Public Property Get Valor(ByVal slot As Byte) As Single

    Valor = UserInventory(slot).Valor
    
End Property
Public Property Get Amount(ByVal slot As Byte) As Long

    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(UserInventory) And slot <= UBound(UserInventory) Then
        Amount = UserInventory(slot).Amount
    End If
    
End Property
Public Property Get OBJIndex(ByVal slot As Byte) As Integer

    OBJIndex = UserInventory(slot).OBJIndex
    
End Property
Public Property Get OBJType(ByVal slot As Byte) As Integer

    OBJType = UserInventory(slot).OBJType
    
End Property
Public Property Get ItemName(ByVal slot As Byte) As String

    ItemName = UserInventory(slot).name
    
End Property
Public Property Get Equipped(ByVal slot As Byte) As Boolean

    Equipped = UserInventory(slot).Equipped
    
End Property
Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Single, ByVal eName As String, ByVal Puede As Byte)

    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserInventory(slot)
        .Amount = eAmount
        .Def = eDef
        .Equipped = eEquipped
        .grhindex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
        .PuedeUsar = Puede
    End With

End Sub
Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long

    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = X \ 32
    temp_y = Y \ 32
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
    
    If TempItem <= MAX_INVENTORY_SLOTS Then
            ClickItem = TempItem
    End If
End Function
Public Sub DrawInventory()
    Dim i As Byte
    Dim X As Integer
    Dim Y As Integer
    Dim temp_array(3) As Long
    
    For i = 1 To UBound(UserInventory)
        If UserInventory(i).grhindex Then
            X = ((i - 1) Mod (InventoryWindow.Width / 32)) * 32
            Y = ((i - 1) \ (InventoryWindow.Width / 32)) * 32
            
            Engine.Long_To_RGB_List temp_array, -1
            
            If Not UserInventory(i).PuedeUsar = 1 Then
                Call Engine.Draw_GrhIndex(UserInventory(i).grhindex, X - 1, Y - 1, -65536)
            Else
                Call Engine.Draw_GrhIndex(UserInventory(i).grhindex, X - 1, Y - 1, -1)
            End If

            Call Engine.Engine_Text_Render(CStr(UserInventory(i).Amount), X, Y, temp_color_array())
            
            Engine.Long_To_RGB_List temp_array, D3DColorXRGB(255, 0, 0)
            If UserInventory(i).Equipped Then
                Call Engine.Engine_Text_Render("+", X + 20, Y, temp_array)
            End If
        End If
    Next i
    
    If InvSelectedItem > 0 Then
        Engine.Draw_GrhIndex 2, (((InvSelectedItem - 1) Mod (InventoryWindow.Width / 32)) * 32) - 1, (((InvSelectedItem - 1) \ (InventoryWindow.Width / 32)) * 32) - 1, -1
    End If

End Sub

#If ConMenuesConextuales = 1 Then
    Public Sub Initialize(ByRef DirectDraw As DirectDraw7, ByRef InvPic As PictureBox, ByRef PopUpMenu As Menu)
#Else
    Public Sub Initialize(ByRef InvPic As PictureBox)
#End If

    Set InventoryWindow = InvPic
    
    InvSelectedItem = ClickItem(1, 1)
    
#If ConMenuesConextuales = 1 Then
    Set ItemMenu = PopUpMenu
#End If

End Sub
Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
    
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    temp_x = X \ 32
    temp_y = Y \ 32
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
    
    If TempItem <= MAX_INVENTORY_SLOTS Then
        frmMain.LBLITEMS.Visible = True
        If UserInventory(TempItem).name = "" Then
            frmMain.LBLITEMS.Caption = "Nada"
        Else
            frmMain.LBLITEMS.Caption = UserInventory(TempItem).name '& " || " & "Cantidad: " & UserInventory(TempItem).Amount & vbNewLine & "Def:" & UserInventory(TempItem).Def & " MinHit:" & UserInventory(TempItem).MinHit & " MaxHit:" & UserInventory(TempItem).MaxHit
        End If
    End If
End Sub
Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim prevSelItem As Long
    
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
    
    prevSelItem = InvSelectedItem

    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
    
    If InvSelectedItem <> MouseDownSelectedItem And MouseDownSelectedItem <> 0 And InvSelectedItem <> 0 Then
        Call WriteSwapObjects(InvSelectedItem, MouseDownSelectedItem)
    End If
    
    
#If ConMenuesConextuales = 1 Then
    If Button = vbRightButton And prevSelItem = InvSelectedItem Then
        Call InventoryWindow.Parent.PopUpMenu(ItemMenu, , X, Y)
    End If
#End If

End Sub
Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TempItem As Long
 
    If Button = 2 Then
        TempItem = ClickItem(X, Y)
       
        If UserInventory(TempItem).grhindex Then
            MouseDownSelectedItem = TempItem
            Exit Sub
        End If
    End If
   
    MouseDownSelectedItem = 0
End Sub
