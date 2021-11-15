VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   900
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   1620
      Width           =   480
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   3960
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   2580
      Width           =   2490
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   3960
      Index           =   0
      Left            =   780
      TabIndex        =   1
      Top             =   2580
      Width           =   2460
   End
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Text            =   "1"
      Top             =   6960
      Width           =   510
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   1
      Left            =   3840
      Tag             =   "1"
      Top             =   6870
      Width           =   195
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   0
      Left            =   2940
      Tag             =   "1"
      Top             =   6870
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   5100
      TabIndex        =   7
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5520
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   5
      Top             =   1530
      Width           =   2985
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   1
      Left            =   4230
      Tag             =   "1"
      Top             =   6855
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   585
      Tag             =   "1"
      Top             =   6855
      Width           =   2175
   End
   Begin VB.Image Command2 
      Height          =   345
      Left            =   6480
      Tag             =   "1"
      Top             =   180
      Width           =   345
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public LasActionBuy As Boolean
Private lIndex As Byte

Private Sub Cantidad_Change()
    If Val(Cantidad.Text) < 1 Then
        Cantidad.Text = 1
    End If
    
    If Val(Cantidad.Text) > MAX_INVENTORY_OBJS Then
        Cantidad.Text = 1
    End If
    
    If lIndex = 0 Then
        Label1(1).Caption = Round(NPCInventory(List1(0).ListIndex + 1).Valor * Val(Cantidad.Text), 0) 'No mostramos numeros reales
    Else
        Label1(1).Caption = Round(Inventario.Valor(List1(1).ListIndex + 1) * Val(Cantidad.Text), 0) 'No mostramos numeros reales
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub cmdMasMenos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        cmdMasMenos(Index).Picture = LoadPicture(App.path & "\Recursos\Interface\menos-down.jpg")
        cmdMasMenos(Index).Tag = "1"
        Cantidad.Text = Val(Cantidad.Text - 1)
    Case 1
        cmdMasMenos(Index).Picture = LoadPicture(App.path & "\Recursos\Interface\mas-down.jpg")
        cmdMasMenos(Index).Tag = "1"
        Cantidad.Text = Val(Cantidad.Text + 1)
End Select

End Sub

Private Sub cmdMasMenos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
    Case 0
        If cmdMasMenos(Index).Tag = "0" Then
            cmdMasMenos(Index).Picture = LoadPicture(App.path & "\Recursos\Interface\menos-over.jpg")
            cmdMasMenos(Index).Tag = "1"
        End If
    Case 1
        If cmdMasMenos(Index).Tag = "0" Then
            cmdMasMenos(Index).Picture = LoadPicture(App.path & "\Recursos\Interface\mas-over.jpg")
            cmdMasMenos(Index).Tag = "1"
        End If
End Select

End Sub

Private Sub Command2_Click()
    Call WriteCommerceEnd
End Sub

Private Sub Form_Load()
'Cargamos la interfase
Me.Picture = LoadPicture(App.path & "\Recursos\Interface\comercio.jpg")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.path & "\Recursos\Graficos\BotónComprar.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.path & "\Recursos\Graficos\Botónvender.jpg")
    Image1(1).Tag = 1
End If
End Sub

Private Sub Image1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

If List1(Index).List(List1(Index).ListIndex) = "" Or _
   List1(Index).ListIndex < 0 Then Exit Sub

If Not IsNumeric(Cantidad.Text) Or Cantidad.Text = 0 Then Exit Sub

Select Case Index
    Case 0
        frmComerciar.List1(0).SetFocus
        LastIndex1 = List1(0).ListIndex
        LasActionBuy = True
        If UserGLD >= Round(NPCInventory(List1(0).ListIndex + 1).Valor * Val(Cantidad), 0) Then
            Call WriteCommerceBuy(List1(0).ListIndex + 1, Cantidad.Text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1
            Exit Sub
        End If
   
   Case 1
        LastIndex2 = List1(1).ListIndex
        LasActionBuy = False
        
        Call WriteCommerceSell(List1(1).ListIndex + 1, Cantidad.Text)
End Select

End Sub

Private Sub list1_Click(Index As Integer)
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.bottom = 32

lIndex = Index

Select Case Index
    Case 0
        
        Label1(0).Caption = NPCInventory(List1(0).ListIndex + 1).name
        Label1(1).Caption = Round(NPCInventory(List1(0).ListIndex + 1).Valor * Val(Cantidad.Text), 0) 'No mostramos numeros reales
        Label1(2).Caption = NPCInventory(List1(0).ListIndex + 1).amount
        
        If Label1(2).Caption <> 0 Then
        
        Select Case NPCInventory(List1(0).ListIndex + 1).OBJType
            Case eObjType.otWeapon
                Label1(3).Caption = "Golpe:" & NPCInventory(List1(0).ListIndex + 1).MaxHit & "/" & NPCInventory(List1(0).ListIndex + 1).MinHit
                Label1(3).Visible = True
            Case eObjType.otArmadura
                Label1(3).Caption = "Defensa:" & NPCInventory(List1(0).ListIndex + 1).Def
                Label1(3).Visible = True
            Case Else
                Label1(3).Visible = False
        End Select
        
        Call Grh_Render_To_Hdc(Picture1.hdc, (NPCInventory(List1(0).ListIndex + 1).grhindex), 0, 0)
        
        End If
    
    Case 1
        Label1(0).Caption = Inventario.ItemName(List1(1).ListIndex + 1)
        Label1(1).Caption = Round(Inventario.Valor(List1(1).ListIndex + 1) * Val(Cantidad.Text), 0) 'No mostramos numeros reales
        Label1(2).Caption = Inventario.amount(List1(1).ListIndex + 1)
        
        If Label1(2).Caption <> 0 Then
        
        Select Case Inventario.OBJType(List1(1).ListIndex + 1)
            Case eObjType.otWeapon
                Label1(3).Caption = "Golpe:" & Inventario.MaxHit(List1(1).ListIndex + 1) & "/" & Inventario.MinHit(List1(1).ListIndex + 1)
                Label1(3).Visible = True
            Case eObjType.otArmadura
                Label1(3).Caption = "Defensa:" & Inventario.Def(List1(1).ListIndex + 1)
                Label1(3).Visible = True
            Case Else
                Label1(3).Visible = False
        End Select
        
        Call Grh_Render_To_Hdc(Picture1.hdc, Inventario.grhindex(List1(1).ListIndex + 1), 0, 0)
        
        End If
        
End Select

If Label1(2).Caption = 0 Then ' 27/08/2006 - GS > No mostrar imagen ni nada, cuando no ahi nada que mostrar.
    Label1(3).Visible = False
    Picture1.Visible = False
Else
    Picture1.Visible = True
    Picture1.Refresh
End If

End Sub

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Image1(0).Tag = 0 Then
    Image1(0).Picture = LoadPicture(App.path & "\Recursos\Graficos\BotónComprar.jpg")
    Image1(0).Tag = 1
End If
If Image1(1).Tag = 0 Then
    Image1(1).Picture = LoadPicture(App.path & "\Recursos\Graficos\Botónvender.jpg")
    Image1(1).Tag = 1
End If
End Sub
