VERSION 5.00
Begin VB.Form frmPanelAccount 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Panel de Cuenta"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   Icon            =   "frmPanelAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPanelAccount.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   3000
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   0
      Left            =   2415
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   9
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   1
      Left            =   3915
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   8
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   2
      Left            =   5415
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   3
      Left            =   6915
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   4
      Left            =   8415
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   3930
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   5
      Left            =   2415
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   6
      Left            =   3915
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   7
      Left            =   5415
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   8
      Left            =   6915
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   1
      Top             =   5730
      Width           =   1140
   End
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   9
      Left            =   8415
      ScaleHeight     =   1200
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   5730
      Width           =   1140
   End
   Begin VB.Image BorrarPJ 
      Height          =   615
      Left            =   4200
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   0
      Left            =   2250
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   1
      Left            =   3750
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   2
      Left            =   5250
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   3
      Left            =   6750
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   4
      Left            =   8250
      Top             =   3510
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   5
      Left            =   2250
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   6
      Left            =   3750
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   7
      Left            =   5250
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   8
      Left            =   6750
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image imgAcc 
      Height          =   1755
      Index           =   9
      Left            =   8250
      Top             =   5310
      Width           =   1455
   End
   Begin VB.Image pass 
      Height          =   615
      Index           =   1
      Left            =   6120
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2085
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   0
      Left            =   2205
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   2
      Left            =   7995
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   2085
      Width           =   1755
   End
   Begin VB.Image imgAccion 
      Height          =   615
      Index           =   4
      Left            =   7995
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   7575
      Width           =   1755
   End
   Begin VB.Label lblAccData 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la Cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   23
      Top             =   2400
      Width           =   3825
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   2295
      TabIndex        =   22
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   3810
      TabIndex        =   21
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   5295
      TabIndex        =   20
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   6810
      TabIndex        =   19
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   8310
      TabIndex        =   18
      Top             =   3570
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   2295
      TabIndex        =   17
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   3810
      TabIndex        =   16
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   5295
      TabIndex        =   15
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   6810
      TabIndex        =   14
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   10
      Left            =   8310
      TabIndex        =   13
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   6180
      TabIndex        =   12
      Top             =   7620
      Width           =   1605
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicaci�n"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   6180
      TabIndex        =   11
      Top             =   7770
      Width           =   675
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clase"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   6180
      TabIndex        =   10
      Top             =   7920
      Width           =   390
   End
End
Attribute VB_Name = "frmPanelAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Public Seleccionado As Byte

Private Sub Borrar_Click()
UserName = lblAccData(1 + Seleccionado).Caption
frmDELETEPJ.Show
End Sub

Private Sub Cerrar_Click()
End
End Sub

Private Sub BorrarPJ_Click()
frmDELETEPJ.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
    Unload frmConnect
    Me.Icon = frmMain.Icon
    
    Dim i As Byte
    For i = 1 To 10
        lblAccData(i).Caption = ""
    Next i
    
    
End Sub

Private Sub Image1_Click()
Dim i As Byte
    For i = 0 To 7
        If lblAccData(i + 1).Caption = "" Then
            frmCrearPersonaje.Show
            Exit Sub
        End If
    Next i
End Sub

Private Sub Image2_Click()
    MsgBox "No habilitado"
End Sub

Private Sub Image3_Click()
    frmMain.Socket1.Disconnect
    Unload Me
    frmConnect.Show
End Sub

Private Sub Image4_Click()
MsgBox "No habilitado"
End Sub

Private Sub Image5_Click()
    If Not lblAccData(Index + 1).Caption = "" Then
        UserName = lblAccData(Index + 1).Caption
        WriteLoginExistingChar
    End If
End Sub

Private Sub lblName_Click(Index As Integer)
    Seleccionado = Index
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) And (RunWindowed = 1) Then Call Auto_Drag(Me.hWnd)
End Sub

Private Sub imgAcc_Click(Index As Integer)
Call Audio.MusicMP3Stop
End Sub

Private Sub imgAccion_Click(Index As Integer)
Call Audio.MusicMP3Stop

Dim i As Byte

    Select Case Index
        Case 0

            For i = 0 To 7
                If lblAccData(i + 1).Caption = "" Then
                    frmCrearPersonaje.Show
                    Exit Sub
                End If
            Next i
        Case 1, 3
            MsgBox "No habilitado"
            
        Case 2
            frmMain.Socket1.Disconnect
            frmConnect.Show
            
        Case 4
            UserName = lblAccData(1 + Seleccionado).Caption
            WriteLoginExistingChar
            
    End Select
End Sub

Private Sub pass_Click(Index As Integer)
    Call frmNewPassword.Show(vbModal, Me)
End Sub

Private Sub picChar_Click(Index As Integer)
    Seleccionado = Index
    If cPJ(Index).Nombre <> "" Then
        lblCharData(0) = "Nivel " & cPJ(Index).Nivel
        lblCharData(1) = Map_NameLoad(cPJ(intSelChar).Mapa)  'Ubicacion
        lblCharData(2) = ListaClases(cPJ(Index).Clase)
        lblCharData(0).Visible = True
        lblCharData(1).Visible = True
        lblCharData(2).Visible = True
    Else
        lblCharData(0) = ""
        lblCharData(1) = ""
        lblCharData(2) = ""
    End If
End Sub

Private Sub picChar_DblClick(Index As Integer)
Call Audio.MusicMP3Stop
    Seleccionado = Index
    If Not lblAccData(Index + 1).Caption = "" Then
        UserName = lblAccData(1 + Index).Caption
        WriteLoginExistingChar
    Else
        frmCrearPersonaje.Show
    End If
End Sub

Private Sub salir_Click(Index As Integer)

frmMain.Socket1.Disconnect
Unload Me
frmConnect.Show

End Sub


Private Sub Timer1_Timer()
    Dim i As Byte
    For i = 1 To 10
    Call Engine.DrawPJ(i)
    Next i
End Sub

