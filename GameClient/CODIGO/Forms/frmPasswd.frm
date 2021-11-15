VERSION 5.00
Begin VB.Form frmPasswd 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5025
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextCorreoCheck 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   720
      TabIndex        =   11
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
      Height          =   420
      Left            =   105
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3825
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   420
      Left            =   3885
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3825
      Width           =   1080
   End
   Begin VB.TextBox txtPasswdCheck 
      BorderStyle     =   0  'None
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   765
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3455
      Width           =   3510
   End
   Begin VB.TextBox txtPasswd 
      BorderStyle     =   0  'None
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   765
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2835
      Width           =   3510
   End
   Begin VB.TextBox txtCorreo 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   765
      TabIndex        =   3
      Top             =   1710
      Width           =   3510
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Verificación del correo electronico"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label lblstatus 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   35
      TabIndex        =   10
      Top             =   4500
      Width           =   4935
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5040
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Verifiación del password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   750
      TabIndex        =   6
      Top             =   3255
      Width           =   3555
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   750
      TabIndex        =   4
      Top             =   2625
      Width           =   3555
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Dirección de correo electronico:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   750
      TabIndex        =   2
      Top             =   1455
      Width           =   3555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"frmPasswd.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   60
      TabIndex        =   1
      Top             =   405
      Width           =   4890
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "¡CUIDADO!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   1965
      TabIndex        =   0
      Top             =   105
      Width           =   1035
   End
End
Attribute VB_Name = "frmPasswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.12.1 MENDUZ DX8 VERSION www.noicoder.com
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

Function CheckDatos() As Boolean

If txtPasswd.text <> txtPasswdCheck.text Then
   lblstatus.Caption = "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
    Exit Function
End If
If TextCorreoCheck.text <> txtCorreo.text Then
lblstatus.Caption = "La dirección de correo no coincide "
End If

CheckDatos = True
End Function

Private Sub Command1_Click()

If CheckDatos() Then
    UserPassword = txtPasswd.text
    UserEmail = txtCorreo.text
    
    If Not CheckMailString(UserEmail) Then
        lblstatus.Caption = "La dirección de correo no se puede reconocer como válida. Por favor, complete el formulario con una dirección de correo real."
        Command1.Enabled = False
        Exit Sub
        
    End If
    
    EstadoLogin = E_MODO.CrearNuevoPj
    
    If frmMain.Winsock1.State <> sckConnected Then
       lblstatus.Caption = "Advertencia: por favor espere, se está realizando la conexión con el servidor."
        Unload Me
        
    Else
        Call Login
    End If
End If
Me.MousePointer = 11
End Sub

Private Sub Command2_Click()
    EstadoLogin = E_MODO.Dados
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Auto_Drag(Me.hWnd)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Auto_Drag(Me.hWnd)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Auto_Drag(Me.hWnd)
End Sub

Private Sub txtCorreo_Change()
Call VerificarDatos
End Sub

Private Sub textCorreoCheck_Change()
Call VerificarDatos
End Sub

Private Sub txtPasswd_Change()
Call VerificarDatos
End Sub

Private Sub txtPasswdCheck_Change()
Call VerificarDatos
End Sub

Private Sub VerificarDatos()
Command1.Enabled = ((txtPasswd.text <> "" And txtCorreo.text <> "") And (txtPasswd.text = txtPasswdCheck.text) And (txtCorreo.text = TextCorreoCheck.text))
End Sub

