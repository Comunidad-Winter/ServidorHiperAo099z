Attribute VB_Name = "Admin"
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

Public Type tMotd
    texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public NPCs As Long
Public DebugSocket As Boolean

Public Horas As Long
Public Dias As Long
Public MinsRunning As Long

Public tInicioServer As Long
Public EstadisticasWeb As New clsEstadisticasIPC

Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloMover As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloUserPuedeAtacar As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]
Public MinutosWs As Long
Public Puerto As Integer

Public MAXPASOS As Long

Public BootDelBackUp As Byte
Public Lloviendo As Boolean

Public IpList As New Collection
Public ClientsCommandsQueue As Byte

'Public ResetThread As New clsThreading

Function VersionOK(ByVal Ver As String) As Boolean
VersionOK = (Ver = ULTIMAVERSION)
End Function


Public Function ValidarLoginMSG(ByVal N As Integer) As Integer
On Error Resume Next
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(N)
AuxInteger2 = SDM(N)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function


Sub ReSpawnOrigPosNpcs()
On Error Resume Next

Dim i As Integer
Dim MiNPC As npc
   
For i = 1 To LastNPC
   'OJO
   If Npclist(i).flags.NPCActive Then
        
        If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.x, Npclist(i).Orig.y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
        End If
        
        If Npclist(i).Contadores.TiempoExistencia > 0 Then
                Call MuereNpc(i, 0)
        End If
   End If
   
Next i

End Sub

Sub WorldSave()
On Error Resume Next
'Call LogTarea("Sub WorldSave")

Dim loopX As Integer
Dim Porc As Long

Call SendData(ToAll, 0, 0, "||INCIANDO WORLDSAVE, POR FAVOR ESPERE." & FONTTYPE_WORDL)

Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales

Dim j As Integer, k As Integer

For j = 1 To NumMaps
    If MapInfo(j).BackUp = 1 Then k = k + 1
Next j

FrmStat.ProgressBar1.Min = 0
FrmStat.ProgressBar1.max = k
FrmStat.ProgressBar1.Value = 0

For loopX = 1 To NumMaps
    'DoEvents
    
    If MapInfo(loopX).BackUp = 1 Then
    
            Call SaveMapData(loopX)
            FrmStat.ProgressBar1.Value = FrmStat.ProgressBar1.Value + 1
    End If

Next loopX

FrmStat.Visible = False

If FileExist(DatPath & "\bkNpc.dat", vbNormal) Then Kill (DatPath & "bkNpc.dat")
If FileExist(DatPath & "\bkNPCs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNPCs-HOSTILES.dat")

For loopX = 1 To LastNPC
    If Npclist(loopX).flags.BackUp = 1 Then
            Call BackUPnPc(loopX)
    End If
Next

Call SendData(ToAll, 0, 0, "||WORLD SAVE TERMINADO, GRACIAS POR ESPERAR." & FONTTYPE_WORDL)


End Sub

Public Sub PurgarPenas()
Dim i As Integer
For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
    
        If UserList(i).Counters.Pena > 0 Then
                
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                Call SendData(ToIndex, i, 0, "||Te quedan " & UserList(i).Counters.Pena & " minutos en la carcel." & FONTTYPE_WARNING)
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.x, Libertad.y, True)
                    Call SendData(ToIndex, i, 0, "||Has sido liberado!" & FONTTYPE_INFO)
                End If
                
        End If
        
    End If
Next i
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = "")
        
        UserList(UserIndex).Counters.Pena = Minutos
       
        
        Call WarpUserChar(UserIndex, Prision.Map, Prision.x, Prision.y, True)
        
        If GmName = "" Then
            Call SendData(ToIndex, UserIndex, 0, "||Has sido encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & GmName & " te ha encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & FONTTYPE_INFO)
        End If
        
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
On Error Resume Next
If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
    Kill CharPath & UCase$(UserName) & ".chr"
End If
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean
If Inbaneable(Name) Then Exit Function

BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1) 'Or _
(val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "AdminBan")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean

PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal Name As String) As Boolean
'Unban the character
Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")
'Remove it from the banned people database
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NOONE")
End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(MD5s)
        If (md5formateado = MD5s(i)) Then
            MD5ok = True
            Exit Function
        End If
    Next i
    MD5ok = True
End Function
