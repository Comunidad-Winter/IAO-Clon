VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Crear nueva cuenta"
   ClientHeight    =   4650
   ClientLeft      =   10005
   ClientTop       =   3930
   ClientWidth     =   8100
   Icon            =   "frmCrearCuenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearCuenta.frx":8D25A
   ScaleHeight     =   4650
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   1455
      Left            =   3960
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   200
   End
   Begin VB.TextBox mailTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000B&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4950
      TabIndex        =   5
      Top             =   2885
      Width           =   2295
   End
   Begin VB.ComboBox CuentQuestions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      DragIcon        =   "frmCrearCuenta.frx":B8C21
      ForeColor       =   &H80000005&
      Height          =   315
      ItemData        =   "frmCrearCuenta.frx":145E7B
      Left            =   1080
      List            =   "frmCrearCuenta.frx":145E8B
      Sorted          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2385
   End
   Begin VB.TextBox answerTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   1055
      TabIndex        =   3
      Top             =   3760
      Width           =   2385
   End
   Begin VB.TextBox pass1Txt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4970
      PasswordChar    =   "x"
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox passTxt 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000004&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   4950
      PasswordChar    =   "x"
      TabIndex        =   1
      Top             =   1240
      Width           =   2295
   End
   Begin VB.TextBox nameTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000004&
      Height          =   240
      Left            =   4950
      TabIndex        =   0
      Top             =   730
      Width           =   2265
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCrearCuenta.frx":145EEA
      ForeColor       =   &H8000000B&
      Height          =   1455
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Verificando..."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   4320
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5880
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "frmCrearCuenta"
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

Private Sub Form_Load()
    Me.Icon = frmMain.Icon

End Sub

Private Sub Image1_Click()


    UserAccount = nameTxt.Text
    UserPassword = passTxt.Text
    UserEmail = mailTxt.Text
    
    If mailTxt.Text = "a@a.c" Then
    MsgBox "Hotmain Inexistente"
    Exit Sub
    End If
    
    If Not UserPassword = pass1Txt.Text Then
        MsgBox "Las Contraseñas No Coinciden Porfabor Rescribanla denuevo."
        Label1.Caption = "Las Contraseñas No Coinciden"
        Exit Sub
    End If
    
    If Not CheckMailString(UserEmail) Then
        MsgBox "Direccion de Email Esta Mal Escrita."
        Exit Sub
    End If
    
    If Check1.value = 0 Then
        MsgBox "No Acepto las Reglas y Condiciones"
        Exit Sub
    End If
    
    UserAnswer = answerTxt.Text
    UserQuestion = CuentQuestions.ListIndex
    If Len(UserAnswer) < 10 Then
        MsgBox "Respuesta Muy Corta, Porfabor escriba al menos más de 10 Letras."
        Exit Sub
    End If
    

    EstadoLogin = CrearNuevaCuenta
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
 
    
    Unload Me
    
End Sub

Private Sub Label5_Click()

End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Slider1_Change()
Label2.Caption = Slider1.value
End Sub

Private Sub Label1_Change()
    If UserPassword = pass1Txt.Text Then
        Label1.Caption = "Las Contraseñas No Coinciden"
        Label1.ForeColor = &HFF&
    Else
    Label1.Caption = "Las Contraseñas Coinciden"
    Label1.ForeColor = &HFF00&
    End If
End Sub

Private Sub pass1Txt_Change()
    If Not UserPassword = pass1Txt.Text Then
        Label1.Caption = "Las Contraseñas Coinciden"
        Label1.ForeColor = &HFF00&
    End If
    
    If UserPassword = pass1Txt.Text Then
        Label1.Caption = "Las Contraseñas No Coinciden"
        Label1.ForeColor = &HFF&
    End If
End Sub
