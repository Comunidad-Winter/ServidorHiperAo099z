VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "cswsk32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   3465
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   5295
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3465
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   30
      Top             =   -60
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin SocketWrenchCtrl.Socket Socket2 
      Index           =   0
      Left            =   15
      Top             =   495
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<---Contar de nuevo users"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   240
      Width           =   2535
   End
   Begin ComctlLib.StatusBar BarraEstado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3210
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enviar SMSG a los GM's"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Timer tTraficStat 
      Interval        =   6000
      Left            =   570
      Top             =   105
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1425
      Top             =   630
   End
   Begin VB.Timer CmdExec 
      Interval        =   1
      Left            =   1260
      Top             =   90
   End
   Begin VB.Timer GameTimer 
      Interval        =   40
      Left            =   1800
      Top             =   60
   End
   Begin VB.Timer tPiqueteC 
      Interval        =   6000
      Left            =   495
      Top             =   510
   End
   Begin VB.Timer tLluviaEvent 
      Interval        =   60000
      Left            =   2040
      Top             =   720
   End
   Begin VB.Timer tLluvia 
      Interval        =   500
      Left            =   60
      Top             =   1035
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   570
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
      Begin VB.Timer TimerCartelito 
         Interval        =   5000
         Left            =   3720
         Top             =   120
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Enviar RMSG a TODOS"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   4695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Enviar RMSG a los GM's"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   4695
      End
      Begin VB.Timer Auditoria 
         Interval        =   1000
         Left            =   3180
         Top             =   120
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enviar SMSG a TODOS"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Timer FX 
      Interval        =   200
      Left            =   2520
      Top             =   60
   End
   Begin VB.Timer npcataca 
      Interval        =   4000
      Left            =   4440
      Top             =   0
   End
   Begin VB.Timer KillLog 
      Interval        =   60000
      Left            =   3465
      Top             =   45
   End
   Begin VB.Timer TIMER_AI 
      Interval        =   100
      Left            =   3960
      Top             =   60
   End
   Begin VB.Label lblDicho 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Lo ultimo que dijeron"
      Top             =   2880
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MOD Hiper-Ao Version 1.0 Cortesia de Sicarul y WAG Agradecemos a www.glaine.com.ar por el hosting."
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios online:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2070
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez


Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
Dim iUserIndex As Integer

For iUserIndex = 1 To MaxUsers
   
   'Conexion activa? y es un usuario loggeado?
   If UserList(iUserIndex).ConnID <> -1 And UserList(iUserIndex).flags.UserLogged Then
        'Actualiza el contador de inactividad
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        If UserList(iUserIndex).Counters.IdleCount >= IdleLimit Then
            Call SendData(ToIndex, iUserIndex, 0, "!!Demasiado tiempo inactivo. Has sido desconectado..")
            Call Cerrar_Usuario(iUserIndex)
        End If
  End If
  
Next iUserIndex

End Sub



Private Sub Auditoria_Timer()
On Error GoTo errhand

Dim k As Integer
For k = 1 To LastUser
    If UserList(k).ConnID <> -1 Then
        DayStats.Segundos = DayStats.Segundos + 1
    End If
Next k

Call PasarSegundo

Static Andando As Boolean
Static Contador As Long
Dim Tmp As Boolean

Contador = Contador + 1

If Contador >= 10 Then
    Contador = 0
    Tmp = EstadisticasWeb.EstadisticasAndando()
    
    If Andando = False And Tmp = True Then
        Call InicializaEstadisticas
    End If
    
    Andando = Tmp
End If

Exit Sub

errhand:
Call LogError("Error en Timer Auditoria (sistema de desconexion de 10 segundos). Err: " & Err.Description & " - " & Err.number)
End Sub

Private Sub AutoSave_Timer()

On Error GoTo errhandler
'fired every minute

Static Minutos As Long
Static MinutosLatsClean As Long

Static MinsSocketReset As Long

Static MinsPjesSave As Long

MinsRunning = MinsRunning + 1

If MinsRunning = 60 Then
    Horas = Horas + 1
    If Horas = 24 Then
        Call SaveDayStats
        DayStats.MaxUsuarios = 0
        DayStats.Segundos = 0
        DayStats.Promedio = 0
        Call DayElapsed
        'Dias = Dias + 1
        Horas = 0
    End If
    MinsRunning = 0
End If

Dim i As Integer
    
Minutos = Minutos + 1

'MinsSocketReset = MinsSocketReset + 1
'for debug purposes
'If MinsSocketReset > 1 Then
'    MinsSocketReset = 0
'    For i = 1 To MaxUsers
'        If UserList(i).ConnID <> -1 And Not UserList(i).flags.UserLogged Then Call CloseSocket(i)
'    Next i
'    Call ReloadSokcet
'End If
    
If Minutos >= MinutosWs Then
    Call DoBackUp
    Call aClon.VaciarColeccion
    Minutos = 0
End If

If MinutosLatsClean >= 15 Then
        MinutosLatsClean = 0
        Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
        Call LimpiarMundo
Else
        MinutosLatsClean = MinutosLatsClean + 1
End If

'[Consejeros]
'If MinsPjesSave >= 30 Then
'    MinsPjesSave = 0
'    Call GuardarUsuarios
'Else
'    MinsPjesSave = MinsPjesSave + 1
'End If

Call PurgarPenas
Call CheckIdleUser

'<<<<<-------- Log the number of users online ------>>>
Dim N As Integer
N = FreeFile(1)
Open App.Path & "\logs\numusers.log" For Output Shared As N
Print #N, NumUsers
Close #N
'<<<<<-------- Log the number of users online ------>>>

Exit Sub
errhandler:
    Call LogError("Error en TimerAutoSave")

End Sub
Private Sub Cheats_Timer()
Dim LoopC As Integer
For LoopC = 1 To LastUser
    If UserList(LoopC).CheatCont >= 2 Then
        UserList(LoopC).Epa = UserList(LoopC).Epa + 1
    End If
    UserList(LoopC).CheatCont = 0
Next LoopC
End Sub






Private Sub CmdExec_Timer()
Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).ConnID <> -1 Then
        If Not UserList(i).CommandsBuffer.IsEmpty Then
            Call HandleData(i, UserList(i).CommandsBuffer.Pop)
        End If
    End If
Next i

End Sub

Private Sub Command1_Click()
Call SendData(ToAll, 0, 0, "!!" & BroadMsg.Text & ENDC)
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub




Private Sub Command2_Click()
Call SendData(ToAdmins, 0, 0, "||" & "GM's: " & BroadMsg.Text & FONTTYPE_FIGHT & ENDC)
End Sub

Private Sub Command3_Click()
Call SendData(ToAdmins, 0, 0, "!!" & BroadMsg.Text & ENDC)
End Sub

Private Sub Command4_Click()
Call SendData(ToAll, 0, 0, "||" & BroadMsg.Text & FONTTYPE_FIGHT & ENDC)
End Sub

Private Sub Command5_Click()
Dim h As Long
Dim numeroJuaZ As Integer

For h = 1 To LastUser
    If UserList(h).ConnID <> -1 Then numeroJuaZ = numeroJuaZ + 1
Next h
NumUsers = numeroJuaZ
Call MostrarNumUsers
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case x \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Cancel = 1
'Me.Hide
End Sub

Private Sub Form_Resize()
'If WindowState = vbMinimized Then Command2_Click
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call QuitarIconoSystray

#If UsarAPI Then
Call LimpiaWsApi
#Else
Socket1.Cleanup
#End If

Call DescargaNpcsDat


Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " server cerrado."
Close #N

End



End Sub




Private Sub FX_Timer()
Dim MapIndex As Integer
Dim N As Integer
For MapIndex = 1 To NumMaps
    Randomize
    If RandomNumber(1, 150) < 12 Then
        If MapInfo(MapIndex).NumUsers > 0 Then

                Select Case MapInfo(MapIndex).Terreno
                   'Bosque
                   Case Bosque
                        N = RandomNumber(1, 100)
                        Select Case MapInfo(MapIndex).Zona
                            Case Campo
                              If Not Lloviendo Then
                                If N < 30 And N >= 15 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE)
                                ElseIf N < 30 And N < 15 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE2)
                                ElseIf N >= 30 And N <= 35 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_GRILLO)
                                ElseIf N >= 35 And N <= 40 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_GRILLO2)
                                ElseIf N >= 40 And N <= 45 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE3)
                                End If
                               End If
                            Case Ciudad
                               If Not Lloviendo Then
                                If N < 30 And N >= 15 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE)
                                ElseIf N < 30 And N < 15 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE2)
                                ElseIf N >= 30 And N <= 35 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_GRILLO)
                                ElseIf N >= 35 And N <= 40 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_GRILLO2)
                                ElseIf N >= 40 And N <= 45 Then
                                  Call SendData(ToMap, 0, MapIndex, "TW" & SND_AVE3)
                                End If
                               End If
                        End Select

                End Select

        End If
    End If
Next

End Sub

Private Sub GameTimer_Timer()
Dim iUserIndex As Integer
Dim bEnviarStats As Boolean
Dim bEnviarAyS As Boolean
Dim iNpcIndex As Integer

Static lTirarBasura As Long
Static lPermiteAtacar As Long
Static lPermiteCast As Long
Static lPermiteTrabajar As Long

'[Alejo]
If lPermiteAtacar < IntervaloUserPuedeAtacar Then
    lPermiteAtacar = lPermiteAtacar + 1
End If

If lPermiteCast < IntervaloUserPuedeCastear Then
    lPermiteCast = lPermiteCast + 1
End If

If lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     lPermiteTrabajar = lPermiteTrabajar + 1
End If
'[/Alejo]


 '<<<<<< Procesa eventos de los usuarios >>>>>>
 For iUserIndex = 1 To MaxUsers
   'Conexion activa?
   If UserList(iUserIndex).ConnID <> -1 Then
      '�User valido?
      If UserList(iUserIndex).flags.UserLogged Then
         
         '[Alejo-18-5]
         bEnviarStats = False
         bEnviarAyS = False
         
         UserList(iUserIndex).NumeroPaquetesPorMiliSec = 0

         '<<<<<<<<<<<< Allow attack >>>>>>>>>>>>>
'         If lPermiteAtacar < IntervaloUserPuedeAtacar Then
'                lPermiteAtacar = lPermiteAtacar + 1
'         Else
         If Not lPermiteAtacar < IntervaloUserPuedeAtacar Then
                UserList(iUserIndex).flags.PuedeAtacar = 1
'                lPermiteAtacar = 0
         End If
         '<<<<<<<<<<<< Allow attack >>>>>>>>>>>>>

         '<<<<<<<<<<<< Allow Cast spells >>>>>>>>>>>
'         If lPermiteCast < IntervaloUserPuedeCastear Then
'              lPermiteCast = lPermiteCast + 1
'         Else
         If Not lPermiteCast < IntervaloUserPuedeCastear Then
              UserList(iUserIndex).flags.PuedeLanzarSpell = 1
'              lPermiteCast = 0
         End If
         '<<<<<<<<<<<< Allow Cast spells >>>>>>>>>>>

         '<<<<<<<<<<<< Allow Work >>>>>>>>>>>
'         If lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
'              lPermiteTrabajar = lPermiteTrabajar + 1
'         Else
         If Not lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
              UserList(iUserIndex).flags.PuedeTrabajar = 1
'              lPermiteTrabajar = 0
         End If
         '<<<<<<<<<<<< Allow Work >>>>>>>>>>>


         Call DoTileEvents(iUserIndex, UserList(iUserIndex).Pos.Map, UserList(iUserIndex).Pos.x, UserList(iUserIndex).Pos.y)
         
                
         If UserList(iUserIndex).flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
         If UserList(iUserIndex).flags.Ceguera = 1 Or _
            UserList(iUserIndex).flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
          
         If UserList(iUserIndex).flags.Muerto = 0 Then
               
               '[Consejeros]
               If UserList(iUserIndex).flags.Desnudo And UserList(iUserIndex).flags.Privilegios = 0 Then Call EfectoFrio(iUserIndex)
               If UserList(iUserIndex).flags.Meditando Then Call DoMeditar(iUserIndex)
               If UserList(iUserIndex).flags.Envenenado = 1 And UserList(iUserIndex).flags.Privilegios = 0 Then Call EfectoVeneno(iUserIndex, bEnviarStats)
               If UserList(iUserIndex).flags.AdminInvisible <> 1 And UserList(iUserIndex).flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
          
               Call DuracionPociones(iUserIndex)
               Call HambreYSed(iUserIndex, bEnviarAyS)

               If Lloviendo Then
                    If Not Intemperie(iUserIndex) Then
                                 If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                                 'No esta descansando
                                          Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                                          Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                                 ElseIf UserList(iUserIndex).flags.Descansar Then
                                 'esta descansando
                                          Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                                          Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                          'termina de descansar automaticamente
                                          If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And _
                                             UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                                                    Call SendData(ToIndex, iUserIndex, 0, "DOK")
                                                    Call SendData(ToIndex, iUserIndex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                                    UserList(iUserIndex).flags.Descansar = False
                                          End If
                                 End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
                    End If
               Else
                    If Not UserList(iUserIndex).flags.Descansar And (UserList(iUserIndex).flags.Hambre = 0 And UserList(iUserIndex).flags.Sed = 0) Then
                    'No esta descansando
                             Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                             Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
                    ElseIf UserList(iUserIndex).flags.Descansar Then
                    'esta descansando
                             Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                             Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                             'termina de descansar automaticamente
                             If UserList(iUserIndex).Stats.MaxHP = UserList(iUserIndex).Stats.MinHP And _
                                UserList(iUserIndex).Stats.MaxSta = UserList(iUserIndex).Stats.MinSta Then
                                     Call SendData(ToIndex, iUserIndex, 0, "DOK")
                                     Call SendData(ToIndex, iUserIndex, 0, "||Has terminado de descansar." & FONTTYPE_INFO)
                                     UserList(iUserIndex).flags.Descansar = False
                             End If
                    End If 'Not UserList(UserIndex).Flags.Descansar And (UserList(UserIndex).Flags.Hambre = 0 And UserList(UserIndex).Flags.Sed = 0)
               End If

               If bEnviarStats Then Call SendUserStatsBox(iUserIndex)
               If bEnviarAyS Then Call EnviarHambreYsed(iUserIndex)

               If UserList(iUserIndex).NroMacotas > 0 Then Call TiempoInvocacion(iUserIndex)
       End If 'Muerto
     Else 'no esta logeado?
     'UserList(iUserIndex).Counters.IdleCount = 0
     '[Gonzalo]: deshabilitado para el nuevo sistema de tiraje
     'de dados :)
        UserList(iUserIndex).Counters.IdleCount = UserList(iUserIndex).Counters.IdleCount + 1
        If UserList(iUserIndex).Counters.IdleCount > IntervaloParaConexion Then
              UserList(iUserIndex).Counters.IdleCount = 0
              Call CloseSocket(iUserIndex)
        End If
     End If 'UserLogged
        
   End If

   Next iUserIndex

'[Alejo]
If Not lPermiteAtacar < IntervaloUserPuedeAtacar Then
    lPermiteAtacar = 0
End If

If Not lPermiteCast < IntervaloUserPuedeCastear Then
    lPermiteCast = 0
End If

If Not lPermiteTrabajar < IntervaloUserPuedeTrabajar Then
     lPermiteTrabajar = 0
End If

'[/Alejo]
  'DoEvents
End Sub





Private Sub mnuCerrar_Click()

Call SaveGuildsDB

If MsgBox("��Atencion!! Si cierra el servidor puede provocar la perdida de datos. �Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    Dim f
    For Each f In Forms
        Unload f
    Next
End If

End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog_Timer()
On Error Resume Next

If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\connect.log"
If FileExist(App.Path & "\logs\haciendo.log", vbNormal) Then Kill App.Path & "\logs\haciendo.log"
If FileExist(App.Path & "\logs\stats.log", vbNormal) Then Kill App.Path & "\logs\stats.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If FileExist(App.Path & "\logs\wsapi.log", vbNormal) Then Kill App.Path & "\logs\wsapi.log"

End Sub

Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim S As String
Dim nid As NOTIFYICONDATA

S = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub






Private Sub npcataca_Timer()

Dim npc As Integer

For npc = 1 To LastNPC
    Npclist(npc).CanAttack = 1
Next npc


End Sub




Private Sub Socket1_Accept(SocketId As Integer)
#If Not (UsarAPI = 1) Then

'=========================================================
'USO DEL CONTROL SOCKET WRENCH
'=============================

If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Pedido de conexion SocketID:" & SocketId & vbCrLf

On Error Resume Next
    
    Dim NewIndex As Integer
    
    
    If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "NextOpenUser" & vbCrLf
    
    NewIndex = NextOpenUser ' Nuevo indice
    If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "UserIndex asignado " & NewIndex & vbCrLf
    
    If NewIndex <= MaxUsers Then
            If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Cargando Socket " & NewIndex & vbCrLf
            
            Unload Socket2(NewIndex)
            Load Socket2(NewIndex)
            
            Socket2(NewIndex).AddressFamily = AF_INET
            Socket2(NewIndex).protocol = IPPROTO_IP
            Socket2(NewIndex).SocketType = SOCK_STREAM
            Socket2(NewIndex).Binary = False
            Socket2(NewIndex).BufferSize = SOCKET_BUFFER_SIZE
            Socket2(NewIndex).Blocking = False
            Socket2(NewIndex).Linger = 1
            
            Socket2(NewIndex).accept = SocketId
            
            
            If aDos.MaxConexiones(Socket2(NewIndex).PeerAddress) Then
            
                UserList(NewIndex).ConnID = -1
                If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "User slot reseteado " & NewIndex & vbCrLf
            
               
                
                If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Socket unloaded" & NewIndex & vbCrLf
                
                'Call LogCriticEvent(Socket2(NewIndex).PeerAddress & " intento crear mas de 3 conexiones.")
                Call aDos.RestarConexion(Socket2(NewIndex).PeerAddress)
                'Socket2(NewIndex).Disconnect
                Unload frmMain.Socket2(NewIndex)
                
                Exit Sub
            End If
            
            UserList(NewIndex).ConnID = SocketId
            UserList(NewIndex).ip = Socket2(NewIndex).PeerAddress
            
            If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & Socket2(NewIndex).PeerAddress & " logged." & vbCrLf
    Else
        Call LogCriticEvent("No acepte conexion porque no tenia slots")
    End If
    
Exit Sub

#End If
End Sub


Private Sub Socket1_Blocking(Status As Integer, Cancel As Integer)
Cancel = True
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
' solo para depurar
'Call LogError("Socket1:" & ErrorString)

If DebugSocket Then frmDebugSocket.Text2.Text = frmDebugSocket.Text2.Text & Time & " " & ErrorString & vbCrLf

frmDebugSocket.Label3.Caption = Socket1.State
End Sub

Private Sub Socket2_Blocking(Index As Integer, Status As Integer, Cancel As Integer)
'Cancel = True
End Sub

Private Sub Socket2_Connect(Index As Integer)
'If DebugSocket Then frmDebugSocket.Text1.Text = frmDebugSocket.Text1.Text & "Conectado" & vbCrLf
Set UserList(Index).CommandsBuffer = New CColaArray
End Sub

Private Sub Socket2_Disconnect(Index As Integer)
    If UserList(Index).flags.UserLogged And _
        UserList(Index).Counters.Saliendo = False Then
        Call Cerrar_Usuario(Index)
    Else
        Call CloseSocket(Index)
    End If
End Sub

'Private Sub Socket2_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
''24004   WSAEINTR    Blocking function was canceled
''24009   WSAEBADF    Invalid socket descriptor passed to function
''24013   WSAEACCES   Access denied
''24014   WSAEFAULT   Invalid address passed to function
''24022   WSAEINVAL   Invalid socket function call
''24024   WSAEMFILE   No socket descriptors are available
''24035   WSAEWOULDBLOCK  Socket would block on this operation
''24036   WSAEINPROGRESS  Blocking function in progress
''24037   WSAEALREADY Function being canceled has already completed
''24038   WSAENOTSOCK Invalid socket descriptor passed to function
''24039   WSAEDESTADDRREQ Destination address is required
''24040   WSAEMSGSIZE Datagram was too large to fit in specified buffer
''24041   WSAEPROTOTYPE   Specified protocol is the wrong type for this socket
''24042   WSAENOPROTOOPT  Socket option is unknown or unsupported
''24043   WSAEPROTONOSUPPORT  Specified protocol is not supported
''24044   WSAESOCKTNOSUPPORT  Specified socket type is not supported in this address family
''24045   WSAEOPNOTSUPP   Socket operation is not supported
''24046   WSAEPFNOSUPPORT Specified protocol family is not supported
''24047   WSAEAFNOSUPPORT Specified address family is not supported by this protocol
''24048   WSAEADDRINUSE   Specified address is already in use
''24049   WSAEADDRNOTAVAIL    Specified address is not available
''24050   WSAENETDOWN Network subsystem has failed
''24051   WSAENETUNREACH  Network cannot be reached from this host
''24052   WSAENETRESET    Network dropped connection on reset
''24053   WSAECONNABORTED Connection was aborted due to timeout or other failure
''24054   WSAECONNRESET   Connection was reset by remote network
''24055   WSAENOBUFS  No buffer space is available
''24056   WSAEISCONN  Socket is already connected
''24057   WSAENOTCONN Socket Is Not Connected
''24058   WSAESHUTDOWN    Socket connection has been shut down
''24060   WSAETIMEDOUT    Operation timed out before completion
''24061   WSAECONNREFUSED Connection refused by remote network
''24064   WSAEHOSTDOWN    Remote host is down
''24065   WSAEHOSTUNREACH Remote host is unreachable
''24091   WSASYSNOTREADY  Network subsystem is not ready for communication
''24092   WSAVERNOTSUPPORTED  Requested version is not available
''24093   WSANOTINITIALIZED   Windows sockets library not initialized
''25001   WSAHOST_NOT_FOUND   Authoritative Answer Host not found
''25002   WSATRY_AGAIN    Non-authoritative Answer Host not found
''25003   WSANO_RECOVERY  Non-recoverable error
''25004   WSANO_DATA  No data record of requested type
''Response = SOCKET_ERRIGNORE
'If ErrorCode = 24053 Then Call CloseSocket(Index)
'End Sub

Private Sub Socket2_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)

#If Not (UsarAPI = 1) Then

On Error GoTo ErrorHandler

'*********************************************
'Separamos las lineas con ENDC y las enviamos a HandleData()
'*********************************************
Dim LoopC As Integer
Dim RD As String
Dim rBuffer(1 To COMMAND_BUFFER_SIZE) As String
Dim CR As Integer
Dim tChar As String
Dim sChar As Integer
Dim eChar As Integer
Dim aux$
Dim OrigCad As String

Dim LenRD As Long

'<<<<<<<<<<<<<<<<<< Evitamos DoS >>>>>>>>>>>>>>>>>>>>>>>>>>>
'Call AddtoVar(UserList(Index).NumeroPaquetesPorMiliSec, 1, 1000)
'
'If UserList(Index).NumeroPaquetesPorMiliSec > 700 Then
'   'UserList(Index).Flags.AdministrativeBan = 1
'   Call LogCriticalHackAttemp(UserList(Index).Name & " " & frmMain.Socket2(Index).PeerAddress & " alcanzo el max paquetes por iteracion.")
'   Call SendData(ToIndex, Index, 0, "ERRSe ha perdido la conexion, por favor vuelva a conectarse.")
'   Call CloseSocket(Index)
'   Exit Sub
'End If

Call Socket2(Index).Read(RD, DataLength)

OrigCad = RD
LenRD = Len(RD)

'Call AddtoVar(UserList(Index).BytesTransmitidosUser, LenB(RD), 100000)

'[��BUCLE INFINITO!!]'
If LenRD = 0 Then
    UserList(Index).AntiCuelgue = UserList(Index).AntiCuelgue + 1
    If UserList(Index).AntiCuelgue >= 150 Then
        UserList(Index).AntiCuelgue = 0
        Call LogError("!!!! Detectado bucle infinito de eventos socket2_read. cerrando indice " & Index)
        Socket2(Index).Disconnect
        Call CloseSocket(Index)
        Exit Sub
    End If
Else
    UserList(Index).AntiCuelgue = 0
End If
'[��BUCLE INFINITO!!]'

'Verificamos por una comando roto y le agregamos el resto
If UserList(Index).RDBuffer <> "" Then
    RD = UserList(Index).RDBuffer & RD
    UserList(Index).RDBuffer = ""
End If

'Verifica por mas de una linea
sChar = 1
For LoopC = 1 To LenRD

    tChar = Mid$(RD, LoopC, 1)

    If tChar = ENDC Then
        CR = CR + 1
        eChar = LoopC - sChar
        rBuffer(CR) = Mid$(RD, sChar, eChar)
        sChar = LoopC + 1
    End If
        
Next LoopC

'Verifica una linea rota y guarda
If Len(RD) - (sChar - 1) <> 0 Then
    UserList(Index).RDBuffer = Mid$(RD, sChar, Len(RD))
End If

'Enviamos el buffer al manejador
For LoopC = 1 To CR
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    '%%% SI ESTA OPCION SE ACTIVA SOLUCIONA %%%
    '%%% EL PROBLEMA DEL SPEEDHACK          %%%
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    If ClientsCommandsQueue = 1 Then
        If rBuffer(LoopC) <> "" Then If Not UserList(Index).CommandsBuffer.Push(rBuffer(LoopC)) Then Call Cerrar_Usuario(Index)
    
    Else ' SH tiebe efecto
          If UserList(Index).ConnID <> -1 Then
            Call HandleData(Index, rBuffer(LoopC))
          Else
            Exit Sub
          End If
    End If
        
Next LoopC

Exit Sub


ErrorHandler:
    Call LogError("Error en Socket read." & Err.Description & " Numero paquetes:" & UserList(Index).NumeroPaquetesPorMiliSec & " . Rdata:" & OrigCad)

#End If
End Sub



Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler

Dim NpcIndex As Integer
Dim x As Integer
Dim y As Integer
Dim UseAI As Integer
Dim mapa As Integer

If Not haciendoBK Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
        
        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
           If Npclist(NpcIndex).flags.Paralizado = 1 Then
                 Call EfectoParalisisNpc(NpcIndex)
           Else
                'Usamos AI si hay algun user en el mapa
                mapa = Npclist(NpcIndex).Pos.Map
                If mapa > 0 Then
                     If MapInfo(mapa).NumUsers > 0 Then
                             If Npclist(NpcIndex).Movement <> ESTATICO Then
                                   Call NPCAI(NpcIndex)
                             End If
                     End If
                End If
                
           End If
                   
        End If
    
    Next NpcIndex

End If


Exit Sub

ErrorHandler:
 Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.Map)
 Call MuereNpc(NpcIndex, 0)

End Sub

Private Sub Timer1_Timer()

Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).flags.UserLogged Then _
        If UserList(i).flags.Oculto = 1 Then Call DoPermanecerOculto(i)
Next i

End Sub

Private Sub TimerCartelito_Timer()
If FrmStat.Visible = True And haciendoBK = False Then FrmStat.Visible = False

End Sub

Private Sub tLluvia_Timer()
On Error GoTo errhandler

Dim iCount As Integer

If Lloviendo Then
   For iCount = 1 To LastUser
    Call EfectoLluvia(iCount)
   Next iCount
End If

Exit Sub
errhandler:
Call LogError("tLluvia")
End Sub

Private Sub tLluviaEvent_Timer()

On Error GoTo ErrorHandler

Static MinutosLloviendo As Long
Static MinutosSinLluvia As Long

If Not Lloviendo Then
    MinutosSinLluvia = MinutosSinLluvia + 1
    If MinutosSinLluvia >= 15 And MinutosSinLluvia < 1440 Then
            If RandomNumber(1, 100) <= 10 Then
                Lloviendo = True
                MinutosSinLluvia = 0
                Call SendData(ToAll, 0, 0, "LLU")
            End If
    ElseIf MinutosSinLluvia >= 1440 Then
                Lloviendo = True
                MinutosSinLluvia = 0
                Call SendData(ToAll, 0, 0, "LLU")
    End If
Else
    MinutosLloviendo = MinutosLloviendo + 1
    If MinutosLloviendo >= 5 Then
            Lloviendo = False
            Call SendData(ToAll, 0, 0, "LLU")
            MinutosLloviendo = 0
    Else
            If RandomNumber(1, 100) <= 7 Then
                Lloviendo = False
                MinutosLloviendo = 0
                Call SendData(ToAll, 0, 0, "LLU")
            End If
    End If
End If


Exit Sub
ErrorHandler:
Call LogError("Error tLluviaTimer")



End Sub


Private Sub tPiqueteC_Timer()
On Error GoTo errhandler

Static Segundos As Integer

Segundos = Segundos + 6

Dim i As Integer

For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
            
            If MapData(UserList(i).Pos.Map, UserList(i).Pos.x, UserList(i).Pos.y).trigger = 5 Then
                    UserList(i).Counters.PiqueteC = UserList(i).Counters.PiqueteC + 1
                    Call SendData(ToIndex, i, 0, "||Estas obstruyendo la via publica, muevete o seras encarcelado!!!" & FONTTYPE_INFO)
                    If UserList(i).Counters.PiqueteC > 23 Then
                            UserList(i).Counters.PiqueteC = 0
                            Call Encarcelar(i, 3)
                    End If
            Else
                    If UserList(i).Counters.PiqueteC > 0 Then UserList(i).Counters.PiqueteC = 0
            End If
            
            If Segundos >= 18 Then
'                Dim nfile As Integer
'                nfile = FreeFile ' obtenemos un canal
'                Open App.Path & "\logs\maxpasos.log" For Append Shared As #nfile
'                Print #nfile, UserList(i).Counters.Pasos
'                Close #nfile
                If Segundos >= 18 Then UserList(i).Counters.Pasos = 0
            End If
            
    End If
Next i

If Segundos >= 18 Then Segundos = 0
   
Exit Sub

errhandler:
    Call LogError("Error en tPiqueteC_Timer")
End Sub


Private Sub tTraficStat_Timer()

'Dim i As Integer
'
'If frmTrafic.Visible Then frmTrafic.lstTrafico.Clear
'
'For i = 1 To LastUser
'    If UserList(i).Flags.UserLogged Then
'        If frmTrafic.Visible Then
'            frmTrafic.lstTrafico.AddItem UserList(i).Name & " " & UserList(i).BytesTransmitidosUser + UserList(i).BytesTransmitidosSvr & " bytes per second"
'        End If
'        UserList(i).BytesTransmitidosUser = 0
'        UserList(i).BytesTransmitidosSvr = 0
'    End If
'Next i

End Sub

Private Sub Userslst_Click()

End Sub
