VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmPassSeguridad 
   Caption         =   "Nombre de usuario y contraseña"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Internet 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Txtpasswd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Txtnomusu 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton CmdAcept 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Sistema de seguridad para servidor de AO creado por Pablo Seibelt (a.k.a. ""Sicarul"") Para ""AFRODITA-AO"""
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña:"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de usuario:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FrmPassSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAcept_Click()
Dim PassYUserCorrectos As String
Me.MousePointer = 11
Internet.URL = "http://ar.geocities.com/pwagma/server/private.txt" 'Aca hay que cambiar por el archivo a usar y su lugar en la web(Exactisimo)
Internet.protocol = icHTTP ' cambiar si es otro protocolo(FTP:\\) (Https:\\)
Internet.Password = "" 'Cambiar si se nescesita para acceder al archivo
Internet.UserName = "" 'Cambiar si se nescesita para acceder al archivo
PassYUserCorrectos = Internet.OpenURL 'Accedemos al archivo :D
CorrectPass = ReadField(2, PassYUserCorrectos, Asc(";"))
CorrectUser = ReadField(1, PassYUserCorrectos, Asc(";"))

'¿No se encontro el archivo? :(
If CorrectPass = "" Or CorrectUser = "" Then
    MsgBox "No se encontro el password y la contraseña correctos en su lugar correspondiente, puede tener que ver con algo de programacion mal tipeado ;) o una caida de la web de Spring", vbCritical, "ERROR"
    Cago = True
    Me.MousePointer = 1
    Exit Sub
End If
':@ ¿a si? ¿¿queres abrir el server sin el nombre de usuario y contraseña?? ¡¡NO PODES!! MUAHAHAHAHAHA =P
If Txtnomusu.Text <> CorrectUser Or Txtpasswd.Text <> CorrectPass Then
    MsgBox "El nombre de usuario o la contraseña son incorrectos, intentelo nuevamente", vbCritical, "INVALIDOOOO"
    Me.MousePointer = 1
    Exit Sub
End If
Entrando = True


End Sub

Private Sub CmdCancel_Click()
Cago = True
End
End Sub
