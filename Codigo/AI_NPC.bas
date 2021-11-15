Attribute VB_Name = "AI"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
Option Explicit

Public Const ESTATICO = 1
Public Const MUEVE_AL_AZAR = 2
Public Const NPC_MALO_ATACA_USUARIOS_BUENOS = 3
Public Const NPCDEFENSA = 4
Public Const GUARDIAS_ATACAN_CRIMINALES = 5
Public Const GUARDIAS_ATACAN_CIUDADANOS = 6
Public Const SIGUE_AMO = 8
Public Const NPC_ATACA_NPC = 9
Public Const NPC_PATHFINDING = 10



'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo AI_NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'AI de los NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Private Sub GuardiasAI(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim UI As Integer

For HeadingLoop = NORTH To WEST
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(HeadingLoop, nPos)
    If InMapBounds(nPos.Map, nPos.x, nPos.y) Then
        UI = MapData(nPos.Map, nPos.x, nPos.y).UserIndex
        If UI > 0 Then
              If UserList(UI).flags.Muerto = 0 Then
                     '¿ES CRIMINAL?
                     If Criminal(UI) Then
                            Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop)
                            Call NpcAtacaUser(NpcIndex, UI)
                            Exit Sub
                     ElseIf Npclist(NpcIndex).flags.AttackedBy = UserList(UI).Name _
                               And Not Npclist(NpcIndex).flags.Follow Then
                           Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop)
                           Call NpcAtacaUser(NpcIndex, UI)
                           Exit Sub
                     End If
              End If
        End If
    End If
Next HeadingLoop

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim UI As Integer
For HeadingLoop = NORTH To WEST
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(HeadingLoop, nPos)
    If InMapBounds(nPos.Map, nPos.x, nPos.y) Then
        UI = MapData(nPos.Map, nPos.x, nPos.y).UserIndex
        If UI > 0 Then
            If UserList(UI).flags.Muerto = 0 Then
                If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then
                    Dim k As Integer
                    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                    Call NpcLanzaUnSpell(NpcIndex, UI)
                End If
                Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop)
                Call NpcAtacaUser(NpcIndex, MapData(nPos.Map, nPos.x, nPos.y).UserIndex)
                Exit Sub
            End If
        End If
    End If
Next HeadingLoop

Call RestoreOldMovement(NpcIndex)

End Sub


Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim UI As Integer
For HeadingLoop = NORTH To WEST
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(HeadingLoop, nPos)
    If InMapBounds(nPos.Map, nPos.x, nPos.y) Then
        UI = MapData(nPos.Map, nPos.x, nPos.y).UserIndex
        If UI > 0 Then
            If UserList(UI).Name = Npclist(NpcIndex).flags.AttackedBy Then
                If UserList(UI).flags.Muerto = 0 Then
                        If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                          Dim k As Integer
                          k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                          Call NpcLanzaUnSpell(NpcIndex, UI)
                        End If
                        Call ChangeNPCChar(ToMap, 0, nPos.Map, NpcIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, HeadingLoop)
                        Call NpcAtacaUser(NpcIndex, UI)
                        Exit Sub
                End If
            End If
        End If
    End If
Next HeadingLoop

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim UI As Integer
For y = Npclist(NpcIndex).Pos.y - 10 To Npclist(NpcIndex).Pos.y + 10
    For x = Npclist(NpcIndex).Pos.x - 10 To Npclist(NpcIndex).Pos.x + 10
        If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
               UI = MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex
               If UI > 0 Then
                  If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
                       If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                       tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex).Pos)
                       Call MoveNPCChar(NpcIndex, tHeading)
                       Exit Sub
                  End If
               End If
        End If
    Next x
Next y

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim UI As Integer

For y = Npclist(NpcIndex).Pos.y - 10 To Npclist(NpcIndex).Pos.y + 10
    For x = Npclist(NpcIndex).Pos.x - 10 To Npclist(NpcIndex).Pos.x + 10
        If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
            UI = MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex
            If UI > 0 Then
                If UserList(UI).Name = Npclist(NpcIndex).flags.AttackedBy Then
                    If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
                         If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Dim k As Integer
                              k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                         End If
                         tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex).Pos)
                         Call MoveNPCChar(NpcIndex, tHeading)
                         Exit Sub
                    End If
                End If
            End If
        End If
    Next x
Next y

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).MaestroUser = 0 Then
    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    Npclist(NpcIndex).flags.AttackedBy = ""
End If

End Sub


Private Sub PersigueCriminal(ByVal NpcIndex As Integer)
Dim UI As Integer
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
For y = Npclist(NpcIndex).Pos.y - 10 To Npclist(NpcIndex).Pos.y + 10
    For x = Npclist(NpcIndex).Pos.x - 10 To Npclist(NpcIndex).Pos.x + 10
        If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
           UI = MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex
           If UI > 0 Then
                If Criminal(UI) Then
                   If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
                        If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Dim k As Integer
                              k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                        End If
                        tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
           End If
        End If
    Next x
Next y

Call RestoreOldMovement(NpcIndex)

End Sub
Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)
Dim UI As Integer
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
For y = Npclist(NpcIndex).Pos.y - 10 To Npclist(NpcIndex).Pos.y + 10
    For x = Npclist(NpcIndex).Pos.x - 10 To Npclist(NpcIndex).Pos.x + 10
        If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
           UI = MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex
           If UI > 0 Then
                If Criminal(UI) Then
                   If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.Invisible = 0 Then
                        If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Dim k As Integer
                              k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                        End If
                        tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
           End If
        End If
    Next x
Next y

Call RestoreOldMovement(NpcIndex)

End Sub
Private Sub SeguirAmo(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim UI As Integer
For y = Npclist(NpcIndex).Pos.y - 10 To Npclist(NpcIndex).Pos.y + 10
    For x = Npclist(NpcIndex).Pos.x - 10 To Npclist(NpcIndex).Pos.x + 10
        If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
            If Npclist(NpcIndex).Target = 0 And Npclist(NpcIndex).TargetNpc = 0 Then
                UI = MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex
                If UI > 0 Then
                   If UserList(UI).flags.Muerto = 0 _
                   And UserList(UI).flags.Invisible = 0 _
                   And UI = Npclist(NpcIndex).MaestroUser _
                   And Distancia(Npclist(NpcIndex).Pos, UserList(UI).Pos) > 3 Then
                        tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
            End If
        End If
    Next x
Next y

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer
Dim NI As Integer
Dim bNoEsta As Boolean
For y = Npclist(NpcIndex).Pos.y - 10 To Npclist(NpcIndex).Pos.y + 10
    For x = Npclist(NpcIndex).Pos.x - 10 To Npclist(NpcIndex).Pos.x + 10
        If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
           NI = MapData(Npclist(NpcIndex).Pos.Map, x, y).NpcIndex
           If NI > 0 Then
                If Npclist(NpcIndex).TargetNpc = NI Then
                     bNoEsta = True
                     tHeading = FindDirection(Npclist(NpcIndex).Pos, Npclist(MapData(Npclist(NpcIndex).Pos.Map, x, y).NpcIndex).Pos)
                     Call MoveNPCChar(NpcIndex, tHeading)
                     Call NpcAtacaNpc(NpcIndex, NI)
                     Exit Sub
                End If
           End If
           
        End If
    Next x
Next y

If Not bNoEsta Then
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call FollowAmo(NpcIndex)
    Else
        Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
        Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    End If
End If
    
End Sub

Function NPCAI(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If Npclist(NpcIndex).MaestroUser = 0 Then
            'Busca a alguien para atacar
            '¿Es un guardia?
            If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then
                    Call GuardiasAI(NpcIndex)
            ElseIf Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion <> 0 Then
                    Call HostilMalvadoAI(NpcIndex)
            ElseIf Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Call HostilBuenoAI(NpcIndex)
            End If
        Else
            'Evitamos que ataque a su amo, a menos
            'que el amo lo ataque.
            'Call HostilBuenoAI(NpcIndex)
        End If
        
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
        Select Case Npclist(NpcIndex).Movement
            Case MUEVE_AL_AZAR
                If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then
                    If Int(RandomNumber(1, 12)) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                    End If
                    Call PersigueCriminal(NpcIndex)
                ElseIf Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIASKAOS Then
                    If Int(RandomNumber(1, 12)) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                    End If
                    Call PersigueCiudadano(NpcIndex)
                Else
                    If Int(RandomNumber(1, 12)) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                    End If
                End If
            'Va hacia el usuario cercano
            Case NPC_MALO_ATACA_USUARIOS_BUENOS
                Call IrUsuarioCercano(NpcIndex)
            'Va hacia el usuario que lo ataco(FOLLOW)
            Case NPCDEFENSA
                Call SeguirAgresor(NpcIndex)
            'Persigue criminales
            Case GUARDIAS_ATACAN_CRIMINALES
                Call PersigueCriminal(NpcIndex)
            Case GUARDIAS_ATACAN_CIUDADANOS
                Call PersigueCiudadano(NpcIndex)
            Case SIGUE_AMO
                Call SeguirAmo(NpcIndex)
                If Int(RandomNumber(1, 12)) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(1, 4)))
                End If
            Case NPC_ATACA_NPC
                Call AiNpcAtacaNpc(NpcIndex)
            Case NPC_PATHFINDING
                
                If ReCalculatePath(NpcIndex) Then
                    Call PathFindingAI(NpcIndex)
                    'Existe el camino?
                    If Npclist(NpcIndex).PFINFO.NoPath Then 'Si no existe nos movemos al azar
                        'Move randomly
                        Call MoveNPCChar(NpcIndex, Int(RandomNumber(1, 4)))
                    End If
                Else
                    If Not PathEnd(NpcIndex) Then
                        Call FollowPath(NpcIndex)
                    Else
                        Npclist(NpcIndex).PFINFO.PathLenght = 0
                    End If
                End If

        End Select


Exit Function


ErrorHandler:
    Call LogError("NPCAI " & Npclist(NpcIndex).Name & " " & Npclist(NpcIndex).MaestroUser & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).Pos.Map & " x:" & Npclist(NpcIndex).Pos.x & " y:" & Npclist(NpcIndex).Pos.y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNpc)
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)
    
End Function


Function UserNear(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Returns True if there is an user adjacent to the npc position.
'#################################################################
UserNear = Not Int(Distance(Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.y, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.x, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.y)) > 1
End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Returns true if we have to seek a new path
'#################################################################
If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
    ReCalculatePath = True
ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
    ReCalculatePath = True
End If
End Function

Function SimpleAI(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Old Ore4 AI function
'#################################################################
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer

For y = Npclist(NpcIndex).Pos.y - 5 To Npclist(NpcIndex).Pos.y + 5    'Makes a loop that looks at
    For x = Npclist(NpcIndex).Pos.x - 5 To Npclist(NpcIndex).Pos.x + 5   '5 tiles in every direction
           'Make sure tile is legal
            If x > MinXBorder And x < MaxXBorder And y > MinYBorder And y < MaxYBorder Then
                'look for a user
                If MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex > 0 Then
                    'Move towards user
                    tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex).Pos)
                    MoveNPCChar NpcIndex, tHeading
                    'Leave
                    Exit Function
                End If
            End If
    Next x
Next y

End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Returns if the npc has arrived to the end of its path
'#################################################################
PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght
End Function

Function FollowPath(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Moves the npc.
'#################################################################

Dim tmpPos As WorldPos
Dim tHeading As Byte

tmpPos.Map = Npclist(NpcIndex).Pos.Map
tmpPos.x = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).y ' invertí las coordenadas
tmpPos.y = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).x

'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"

tHeading = FindDirection(Npclist(NpcIndex).Pos, tmpPos)

MoveNPCChar NpcIndex, tHeading

Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.CurPos + 1

End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock / 11-07-02
'www.geocities.com/gmorgolock
'morgolock@speedy.com.ar
'This function seeks the shortest path from the Npc
'to the user's location.
'#################################################################
Dim nPos As WorldPos
Dim HeadingLoop As Byte
Dim tHeading As Byte
Dim y As Integer
Dim x As Integer

For y = Npclist(NpcIndex).Pos.y - 10 To Npclist(NpcIndex).Pos.y + 10    'Makes a loop that looks at
     For x = Npclist(NpcIndex).Pos.x - 10 To Npclist(NpcIndex).Pos.x + 10   '5 tiles in every direction

         'Make sure tile is legal
         If x > MinXBorder And x < MaxXBorder And y > MinYBorder And y < MaxYBorder Then
         
             'look for a user
             If MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex > 0 Then
                 'Move towards user
                  Dim tmpUserIndex As Integer
                  tmpUserIndex = MapData(Npclist(NpcIndex).Pos.Map, x, y).UserIndex
                  'We have to invert the coordinates, this is because
                  'ORE refers to maps in converse way of my pathfinding
                  'routines.
                  Npclist(NpcIndex).PFINFO.Target.x = UserList(tmpUserIndex).Pos.y
                  Npclist(NpcIndex).PFINFO.Target.y = UserList(tmpUserIndex).Pos.x 'ops!
                  Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
                  Call SeekPath(NpcIndex)
                  Exit Function
             End If
             
         End If
              
     Next x
 Next y
End Function


Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

If UserList(UserIndex).flags.Invisible = 1 Then Exit Sub

Dim k As Integer
k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))

End Sub

