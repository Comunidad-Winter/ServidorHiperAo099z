Attribute VB_Name = "modHechizos"
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

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(UserIndex).flags.Invisible = 1 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim da�o As Integer

If Hechizos(Spell).SubeHP = 1 Then

    da�o = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + da�o
    If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
    Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    da�o = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    If UserList(UserIndex).flags.Privilegios = 0 Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - da�o
    
    Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
    
    'Muere
    If UserList(UserIndex).Stats.MinHP < 1 Then
        UserList(UserIndex).Stats.MinHP = 0
        Call UserDie(UserIndex)
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Then
     If UserList(UserIndex).flags.Paralizado = 0 Then
          UserList(UserIndex).flags.Paralizado = 1
          UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
          Call SendData(ToIndex, UserIndex, 0, "PARADOK")
     End If
End If


End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

On Error GoTo errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
errhandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, UserIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No tenes espacio para mas hechizos." & FONTTYPE_INFO)
    Else
        UserList(UserIndex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, UserIndex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||Ya tenes ese hechizo." & FONTTYPE_INFO)
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal S As String, ByVal UserIndex As Integer)
On Error Resume Next

    Dim ind As String
    ind = UserList(UserIndex).Char.CharIndex
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbCyan & "�" & S & "�" & ind)
    Exit Sub
End Sub
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean


If UserList(UserIndex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(UserIndex).flags.TargetMap
    wp2.x = UserList(UserIndex).flags.TargetX
    wp2.y = UserList(UserIndex).flags.TargetY
    
    If Distancia(UserList(UserIndex).Pos, wp2) > 18 Then
            'UserList(UserIndex).Flags.AdministrativeBan = 1
            'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
            'Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de lanzar un spell desde otro mapa.")
            'Call Cerrar_Usuario(UserIndex)
            Exit Function
    End If
    
    If UserList(UserIndex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(UserIndex).Stats.UserSkills(Magia) >= Hechizos(HechizoIndex).MinSkill Then
            PuedeLanzar = (UserList(UserIndex).Stats.MinSta > 0)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes puntos de magia para lanzar este hechizo." & FONTTYPE_INFO)
            PuedeLanzar = False
        End If
    Else
            Call SendData(ToIndex, UserIndex, 0, "||No tenes suficiente mana." & FONTTYPE_INFO)
            PuedeLanzar = False
    End If
Else
   Call SendData(ToIndex, UserIndex, 0, "||No podes lanzar hechizos porque estas muerto." & FONTTYPE_INFO)
   PuedeLanzar = False
End If

End Function

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)

'Call LogTarea("HechizoInvocacion")
If UserList(UserIndex).NroMacotas >= MAXMASCOTAS Then Exit Sub

Dim H As Integer, j As Integer, ind As Integer, Index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(UserIndex).flags.TargetMap
TargetPos.x = UserList(UserIndex).flags.TargetX
TargetPos.y = UserList(UserIndex).flags.TargetY

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
For j = 1 To Hechizos(H).Cant
    
    If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).NumNpc, TargetPos, True, False)
        If ind <= MAXNPCS Then
            UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
            
            Index = FreeMascotaIndex(UserIndex)
            
            UserList(UserIndex).MascotasIndex(Index) = ind
            UserList(UserIndex).MascotasType(Index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = UserIndex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(UserIndex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case uInvocacion '
       Call HechizoInvocacion(UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    Call SendUserStatsBox(UserIndex)
End If


End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)

Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(UserIndex, b)
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    Call SendUserStatsBox(UserIndex)
    Call SendUserStatsBox(UserList(UserIndex).flags.TargetUser)
    UserList(UserIndex).flags.TargetUser = 0
End If

End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)



Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNpc, uh, b, UserIndex)
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNpc, UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    UserList(UserIndex).flags.TargetNpc = 0
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    Call SendUserStatsBox(UserIndex)
End If

End Sub
Sub LanzarHechizo(Index As Integer, UserIndex As Integer)



Dim uh As Integer
Dim exito As Boolean

uh = UserList(UserIndex).Stats.UserHechizos(Index)

If PuedeLanzar(UserIndex, uh) Then
    Select Case Hechizos(uh).Target
        
        Case uUsuarios
            If UserList(UserIndex).flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Este hechizo actua solo sobre usuarios." & FONTTYPE_INFO)
            End If
        Case uNPC
            If UserList(UserIndex).flags.TargetNpc > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Este hechizo solo afecta a los npcs." & FONTTYPE_INFO)
            End If
        Case uUsuariosYnpc
            If UserList(UserIndex).flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)
            ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Target invalido." & FONTTYPE_INFO)
            End If
        Case uTerreno
            Call HandleHechizoTerreno(UserIndex, uh)
    End Select
    
End If
                

End Sub
Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)



Dim H As Integer, TU As Integer
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
TU = UserList(UserIndex).flags.TargetUser

If Hechizos(H).Invisibilidad = 1 Then
   UserList(TU).flags.Invisible = 1
   Call SendData(ToMap, 0, UserList(TU).Pos.Map, "NOVER" & UserList(TU).Char.CharIndex & ",1")
   Call InfoHechizo(UserIndex)
   b = True
End If

If Hechizos(H).Envenena = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Envenenado = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).CuraVeneno = 1 Then
        UserList(TU).flags.Envenenado = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Maldicion = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).RemoverMaldicion = 1 Then
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Paraliza = 1 Then
     If UserList(TU).flags.Paralizado = 0 Then
            If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
            
            If UserIndex <> TU Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TU)
            End If
            
            UserList(TU).flags.Paralizado = 1
            UserList(TU).Counters.Paralisis = IntervaloParalizado
            Call SendData(ToIndex, TU, 0, "PARADOK")
            Call InfoHechizo(UserIndex)
            b = True
    End If
End If

If Hechizos(H).RemoverParalisis = 1 Then
    If UserList(TU).flags.Paralizado = 1 Then
                UserList(TU).flags.Paralizado = 0
                Call SendData(ToIndex, TU, 0, "PARADOK")
                Call InfoHechizo(UserIndex)
                b = True
    End If
End If

If Hechizos(H).Revivir = 1 Then
    If UserList(TU).flags.Muerto = 1 Then
        If Not Criminal(TU) Then
                If TU <> UserIndex Then
                    Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, 500, MAXREP)
                    Call SendData(ToIndex, UserIndex, 0, "||�Los Dioses te sonrien, has ganado 500 puntos de nobleza!." & FONTTYPE_INFO)
                End If
        End If
        
        Call RevivirUsuario(TU)
    End If
    Call InfoHechizo(UserIndex)
    b = True
End If

If Hechizos(H).Ceguera = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
        Call SendData(ToIndex, TU, 0, "CEGU")
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Estupidez = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
        Call SendData(ToIndex, TU, 0, "DUMB")
        Call InfoHechizo(UserIndex)
        b = True
End If

End Sub
Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)



If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Invisible = 1
   b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        Exit Sub
   End If
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Envenenado = 1
   b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Envenenado = 0
   b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        Exit Sub
   End If
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Maldicion = 1
   b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Maldicion = 0
   b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).flags.Bendicion = 1
   b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
   If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 1
            Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            b = True
   Else
      Call SendData(ToIndex, UserIndex, 0, "||El npc es inmune a este hechizo." & FONTTYPE_FIGHT)
   End If
End If

If Hechizos(hIndex).RemoverParalisis = 1 Then
   If Npclist(NpcIndex).flags.Paralizado = 1 Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
   Else
      Call SendData(ToIndex, UserIndex, 0, "||El npc no esta paralizado." & FONTTYPE_FIGHT)
   End If
End If

 


End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)
Dim Calculo As Long
Dim da�o As Integer


'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    da�o = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)
    Call AddtoVar(Npclist(NpcIndex).Stats.MinHP, da�o, Npclist(NpcIndex).Stats.MaxHP)
    Call SendData(ToIndex, UserIndex, 0, "||Has curado " & da�o & " puntos de salud a la criatura." & FONTTYPE_FIGHT)
    b = True
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    
    If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    da�o = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)
    b = True
    Call NpcAtacado(NpcIndex, UserIndex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    
    '[Wag]
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    
    Calculo = (da�o / Npclist(NpcIndex).Stats.MaxHP * MiNPC.GiveEXP)
    '[/Wag]
    
    
    '[Wag](Elim by Sicarul XD)

    If da�o >= Npclist(NpcIndex).Stats.MinHP Then
    '    If da�o >= Npclist(NpcIndex).Stats.MaxHP And Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP Then
        Calculo = (Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP * MiNPC.GiveEXP)
    '    Else
    '    Calculo = MiNPC.GiveEXP / 2 + (Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP * MiNPC.GiveEXP / 2)
    '    End If
    End If

    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - da�o
    
    Call SendData(ToIndex, UserIndex, 0, "U2" & da�o)

    Call AddtoVar(UserList(UserIndex).Stats.Exp, Calculo, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & Calculo & " puntos de experiencia." & FONTTYPE_FIGHT)
    '[Sicarul]
    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)
    '[/Sicarul]
    '[/Wag]
    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Calculo = 0
        Call MuereNpc(NpcIndex, UserIndex)
    End If
End If

End Sub
Sub InfoHechizo(ByVal UserIndex As Integer)


    Dim H As Integer
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    
    Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, UserIndex)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(H).WAV)
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
        Call SendData(ToPCArea, UserIndex, Npclist(UserList(UserIndex).flags.TargetNpc).Pos.Map, "CFX" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    End If
    
    If UserList(UserIndex).flags.TargetUser > 0 Then
        If UserIndex <> UserList(UserIndex).flags.TargetUser Then
            Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & UserList(UserList(UserIndex).flags.TargetUser).Name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, UserList(UserIndex).flags.TargetUser, 0, "||" & UserList(UserIndex).Name & " " & Hechizos(H).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).PropioMsg & FONTTYPE_FIGHT)
        End If
    ElseIf UserList(UserIndex).flags.TargetNpc > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & "la criatura." & FONTTYPE_FIGHT)
    End If
    
End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)

Dim H As Integer
Dim da�o As Integer
Dim tempChr As Integer
    
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
tempChr = UserList(UserIndex).flags.TargetUser
      
'Hambre
If Hechizos(H).SubeHam = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    da�o = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    Call AddtoVar(UserList(tempChr).Stats.MinHam, _
         da�o, UserList(tempChr).Stats.MaxHam)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & da�o & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & da�o & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & da�o & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeHam = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    da�o = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - da�o
    
    If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & da�o & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & da�o & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & da�o & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    
    b = True
    
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).flags.Hambre = 1
    End If
    
End If

'Sed
If Hechizos(H).SubeSed = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    Call AddtoVar(UserList(tempChr).Stats.MinAGU, da�o, _
         UserList(tempChr).Stats.MaxAGU)
         
    If UserIndex <> tempChr Then
      Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & da�o & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
      Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & da�o & " puntos de sed." & FONTTYPE_FIGHT)
    Else
      Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & da�o & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeSed = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - da�o
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & da�o & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & da�o & " puntos de sed." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & da�o & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1
    End If
    
    b = True
End If

' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then
    
    Call InfoHechizo(UserIndex)
    da�o = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), da�o, MAXATRIBUTOS)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    da�o = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(Agilidad) = UserList(tempChr).Stats.UserAtributos(Agilidad) - da�o
    If UserList(tempChr).Stats.UserAtributos(Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Agilidad) = MINATRIBUTOS
    b = True
    
End If

' <-------- Fuerza ---------->
If Hechizos(H).SubeFuerza = 1 Then
    
    Call InfoHechizo(UserIndex)
    da�o = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Fuerza), da�o, MAXATRIBUTOS)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeFuerza = 2 Then

    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    da�o = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(Fuerza) = UserList(tempChr).Stats.UserAtributos(Fuerza) - da�o
    If UserList(tempChr).Stats.UserAtributos(Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Fuerza) = MINATRIBUTOS
    b = True
    
End If

'Salud
If Hechizos(H).SubeHP = 1 Then
    da�o = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)
    
    Call AddtoVar(UserList(tempChr).Stats.MinHP, da�o, _
         UserList(tempChr).Stats.MaxHP)
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & da�o & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(H).SubeHP = 2 Then
    
    If UserIndex = tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacarte a vos mismo." & FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    da�o = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    da�o = da�o + Porcentaje(da�o, 3 * UserList(UserIndex).Stats.ELV)
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - da�o
    
    Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & da�o & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
    Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & da�o & " puntos de vida." & FONTTYPE_FIGHT)
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        Call ContarMuerte(tempChr, UserIndex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, UserIndex)
        Call UserDie(tempChr)
    End If
    
    b = True
End If

'Mana
If Hechizos(H).SubeMana = 1 Then
    
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinMAN, da�o, UserList(tempChr).Stats.MaxMAN)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & da�o & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & da�o & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & da�o & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & da�o & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & da�o & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & da�o & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - da�o
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
    
End If

'Stamina
If Hechizos(H).SubeSta = 1 Then
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinSta, da�o, _
         UserList(tempChr).Stats.MaxSta)
    If UserIndex <> tempChr Then
         Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & da�o & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
         Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & da�o & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & da�o & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & da�o & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & da�o & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & da�o & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - da�o
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If


End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(UserIndex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(UserIndex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then

    Call SendData(ToIndex, UserIndex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).Nombre)

Else

    Call SendData(ToIndex, UserIndex, 0, "SHS" & Slot & "," & "0" & "," & "(None)")

End If


End Sub
