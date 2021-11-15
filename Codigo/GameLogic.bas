Attribute VB_Name = "Extra"
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

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie
End Function



Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)

On Error GoTo errhandler

Dim nPos As WorldPos
Dim FxFlag As Boolean
'Controla las salidas
If InMapBounds(Map, x, y) Then
    
    If MapData(Map, x, y).OBJInfo.ObjIndex > 0 Then
        FxFlag = ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).ObjType = OBJTYPE_TELEPORT
    End If
    
    If MapData(Map, x, y).TileExit.Map > 0 Then
        '¿Es mapa de newbies?
        If UCase$(MapInfo(MapData(Map, x, y).TileExit.Map).Restringir) = "SI" Then
            '¿El usuario es un newbie?
            If EsNewbie(UserIndex) Then
                If LegalPos(MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y, PuedeAtravesarAgua(UserIndex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(UserIndex, MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y, True)
                    Else
                        Call WarpUserChar(UserIndex, MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, x, y).TileExit, nPos)
                    If nPos.x <> 0 And nPos.y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, True)
                        Else
                            Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y)
                        End If
                    End If
                End If
            Else 'No es newbie
                Call SendData(ToIndex, UserIndex, 0, "||Mapa exclusivo para newbies." & FONTTYPE_INFO)
                
                Call ClosestLegalPos(UserList(UserIndex).Pos, nPos)
                If nPos.x <> 0 And nPos.y <> 0 Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y)
                End If
            End If
        Else 'No es un mapa de newbies
            If LegalPos(MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y, PuedeAtravesarAgua(UserIndex)) Then
                If FxFlag Then
                    Call WarpUserChar(UserIndex, MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y, True)
                Else
                    Call WarpUserChar(UserIndex, MapData(Map, x, y).TileExit.Map, MapData(Map, x, y).TileExit.x, MapData(Map, x, y).TileExit.y)
                End If
            Else
                Call ClosestLegalPos(MapData(Map, x, y).TileExit, nPos)
                If nPos.x <> 0 And nPos.y <> 0 Then
                    If FxFlag Then
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y, True)
                    Else
                        Call WarpUserChar(UserIndex, nPos.Map, nPos.x, nPos.y)
                    End If
                End If
            End If
        End If
    End If
    
End If

Exit Sub

errhandler:
    Call LogError("Error en DotileEvents")
    Call SendData(ToAdmins, 0, 0, "ERROR en mapa " & Map)
End Sub

Function InRangoVision(ByVal UserIndex As Integer, x As Integer, y As Integer) As Boolean

If x > UserList(UserIndex).Pos.x - MinXBorder And x < UserList(UserIndex).Pos.x + MinXBorder Then
    If y > UserList(UserIndex).Pos.y - MinYBorder And y < UserList(UserIndex).Pos.y + MinYBorder Then
        InRangoVision = True
        Exit Function
    End If
End If
InRangoVision = False

End Function

Function InMapBounds(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean

If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    InMapBounds = False
Else
    InMapBounds = True
End If

End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.x, nPos.y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.y - LoopC To Pos.y + LoopC
        For tX = Pos.x - LoopC To Pos.x + LoopC
            
            If LegalPos(nPos.Map, tX, tY) Then
                nPos.x = tX
                nPos.y = tY
                '¿Hay objeto?
                
                tX = Pos.x + LoopC
                tY = Pos.y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.x = 0
    nPos.y = 0
End If

End Sub

Function NameIndex(ByVal Name As String) As Integer

Dim UserIndex As Integer
'¿Nombre valido?
If Name = "" Then
    NameIndex = 0
    Exit Function
End If
  
UserIndex = 1
Do Until UCase$(Left$(UserList(UserIndex).Name, Len(Name))) = UCase$(Name)
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        UserIndex = 0
        Exit Do
    End If
    
Loop
NameIndex = UserIndex
End Function


Function IP_Index(ByVal inIP As String) As Integer
On Error GoTo local_errHand

Dim UserIndex As Integer
'¿Nombre valido?
If inIP = "" Then
    IP_Index = 0
    Exit Function
End If
  
UserIndex = 1
Do Until UserList(UserIndex).ip = inIP
    
    UserIndex = UserIndex + 1
    
    If UserIndex > MaxUsers Then
        IP_Index = 0
        Exit Do
    End If
    
Loop

local_errHand:
    
    IP_Index = UserIndex

End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
Dim LoopC As Integer
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged = True Then
        If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameIP = False
End Function

Function CheckForSameName(ByVal UserIndex As Integer, ByVal Name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
Dim LoopC As Integer
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged Then
        If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameName = False
End Function

Sub HeadtoPos(Head As Byte, ByRef Pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
Dim x As Integer
Dim y As Integer
Dim tempVar As Single
Dim nX As Integer
Dim nY As Integer

x = Pos.x
y = Pos.y

If Head = NORTH Then
    nX = x
    nY = y - 1
End If

If Head = SOUTH Then
    nX = x
    nY = y + 1
End If

If Head = EAST Then
    nX = x + 1
    nY = y
End If

If Head = WEST Then
    nX = x - 1
    nY = y
End If

'Devuelve valores
Pos.x = nX
Pos.y = nY

End Sub

Function LegalPos(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal PuedeAgua = False) As Boolean

'¿Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or _
   (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
            LegalPos = False
Else
  
  If Not PuedeAgua Then
        LegalPos = (MapData(Map, x, y).Blocked <> 1) And _
                   (MapData(Map, x, y).UserIndex = 0) And _
                   (MapData(Map, x, y).NpcIndex = 0) And _
                   (Not HayAgua(Map, x, y))
  Else
        LegalPos = (MapData(Map, x, y).Blocked <> 1) And _
                   (MapData(Map, x, y).UserIndex = 0) And _
                   (MapData(Map, x, y).NpcIndex = 0) And _
                   (HayAgua(Map, x, y))
  End If
   
End If

End Function



Function LegalPosNPC(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal AguaValida As Byte) As Boolean

If (Map <= 0 Or Map > NumMaps) Or _
   (x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder) Then
    LegalPosNPC = False
Else

 If AguaValida = 0 Then
   LegalPosNPC = (MapData(Map, x, y).Blocked <> 1) And _
     (MapData(Map, x, y).UserIndex = 0) And _
     (MapData(Map, x, y).NpcIndex = 0) And _
     (MapData(Map, x, y).trigger <> POSINVALIDA) _
     And Not HayAgua(Map, x, y)
 Else
   LegalPosNPC = (MapData(Map, x, y).Blocked <> 1) And _
     (MapData(Map, x, y).UserIndex = 0) And _
     (MapData(Map, x, y).NpcIndex = 0) And _
     (MapData(Map, x, y).trigger <> POSINVALIDA)
 End If
 
End If


End Function

Sub SendHelp(ByVal Index As Integer)
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call SendData(ToIndex, Index, 0, "||" & GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC) & FONTTYPE_INFO)
Next LoopC
End Sub
Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

If Npclist(NpcIndex).NroExpresiones > 0 Then
    Dim randomi
    randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & Npclist(NpcIndex).Expresiones(randomi) & "°" & Npclist(NpcIndex).Char.CharIndex & FONTTYPE_INFO)
End If
                    
End Sub
Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)

'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim NpcIndex As Integer

'¿Posicion valida?
If InMapBounds(Map, x, y) Then
    UserList(UserIndex).flags.TargetMap = Map
    UserList(UserIndex).flags.TargetX = x
    UserList(UserIndex).flags.TargetY = y
    '¿Es un obj?
    If MapData(Map, x, y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
        UserList(UserIndex).flags.TargetObj = MapData(Map, x, y).OBJInfo.ObjIndex
        UserList(UserIndex).flags.TargetObjMap = Map
        UserList(UserIndex).flags.TargetObjX = x
        UserList(UserIndex).flags.TargetObjY = y
        FoundSomething = 1
    ElseIf MapData(Map, x + 1, y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, x + 1, y).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
            UserList(UserIndex).flags.TargetObj = MapData(Map, x + 1, y).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = x + 1
            UserList(UserIndex).flags.TargetObjY = y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
            'Informa el nombre
            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
            UserList(UserIndex).flags.TargetObj = MapData(Map, x + 1, y + 1).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = x + 1
            UserList(UserIndex).flags.TargetObjY = y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, x, y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
            'Informa el nombre
            Call SendData(ToIndex, UserIndex, 0, "||" & ObjData(MapData(Map, x, y + 1).OBJInfo.ObjIndex).Name & FONTTYPE_INFO)
            UserList(UserIndex).flags.TargetObj = MapData(Map, x, y).OBJInfo.ObjIndex
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = x
            UserList(UserIndex).flags.TargetObjY = y + 1
            FoundSomething = 1
        End If
    End If
    '¿Es un personaje?
    If y + 1 <= YMaxMapSize Then
        If MapData(Map, x, y + 1).UserIndex > 0 Then
            TempCharIndex = MapData(Map, x, y + 1).UserIndex
            FoundChar = 1
        End If
        If MapData(Map, x, y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, x, y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(Map, x, y).UserIndex > 0 Then
            TempCharIndex = MapData(Map, x, y).UserIndex
            FoundChar = 1
        End If
        If MapData(Map, x, y).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, x, y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    
    'Reaccion al personaje
    If FoundChar = 1 Then '  ¿Encontro un Usuario?
            
       If UserList(TempCharIndex).flags.AdminInvisible = 0 Then
            
            If EsNewbie(TempCharIndex) Then
                Stat = " <NEWBIE>"
            End If

            If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                Stat = Stat & " <Ejercito real> " & "<" & TituloReal(TempCharIndex) & ">"
            ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                Stat = Stat & " <Fuerzas del caos> " & "<" & TituloCaos(TempCharIndex) & ">"
            End If
            If UserList(TempCharIndex).flags.Casado <> "" Then
            Dim ReRaRo$
            ReRaRo$ = IIf(UCase$(UserList(TempCharIndex).Genero) = "HOMBRE", "Casado", "Casada")
                Stat = Stat & " <" & ReRaRo$ & " con " & UserList(TempCharIndex).flags.Casado & ">"
            End If
            If UserList(TempCharIndex).GuildInfo.GuildName <> "" Then
                Stat = Stat & " <" & UserList(TempCharIndex).GuildInfo.GuildName & ">"
            End If
            
            If Len(UserList(TempCharIndex).Desc) > 1 Then
                Stat = "||Ves a " & UserList(TempCharIndex).Name & Stat & " - " & UserList(TempCharIndex).Desc & " [Nivel: " & str(UserList(TempCharIndex).Stats.ELV)
            Else
                Stat = "||Ves a " & UserList(TempCharIndex).Name & Stat & " [Nivel: " & str(UserList(TempCharIndex).Stats.ELV)
            End If
            If UserList(TempCharIndex).Faccion.CiudadanosMatados > 0 And UserList(TempCharIndex).Faccion.CriminalesMatados > 0 Then
                Stat = Stat & " /Ciudas Matados:" & str(UserList(TempCharIndex).Faccion.CiudadanosMatados) & " /Crimis Matados:" & str(UserList(TempCharIndex).Faccion.CriminalesMatados) & " ]"
            ElseIf UserList(TempCharIndex).Faccion.CiudadanosMatados > 0 Then
                Stat = Stat & " /Ciudas Matados:" & str(UserList(TempCharIndex).Faccion.CiudadanosMatados) & " ]"
            ElseIf UserList(TempCharIndex).Faccion.CriminalesMatados > 0 Then
                Stat = Stat & " /Ciudas Matados:" & str(UserList(TempCharIndex).Faccion.CriminalesMatados) & " ]"
            Else
                Stat = Stat & "]"
            End If
            
            If UserList(TempCharIndex).flags.Privilegios > 0 Then
                Stat = Stat & " <GAME MASTER> ~0~185~0~1~0"
            ElseIf Criminal(TempCharIndex) Then
                Stat = Stat & " <CRIMINAL> ~255~0~0~1~0"
            Else
                Stat = Stat & " <CIUDADANO> ~0~0~200~1~0"
            End If
            
            Call SendData(ToIndex, UserIndex, 0, Stat)
                
            
            FoundSomething = 1
            UserList(UserIndex).flags.TargetUser = TempCharIndex
            UserList(UserIndex).flags.TargetNpc = 0
            UserList(UserIndex).flags.TargetNpcTipo = 0
       
       End If
       
    End If
    If FoundChar = 2 Then '¿Encontro un NPC?
            
            If Len(Npclist(TempCharIndex).Desc) > 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).Char.CharIndex & FONTTYPE_INFO)
            Else
                
                If Npclist(TempCharIndex).MaestroUser > 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "|| " & Npclist(TempCharIndex).Name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).Name & FONTTYPE_INFO)
                Else
                    'Call SendData(ToIndex, UserIndex, 0, "|| " & Npclist(TempCharIndex).Name & "." & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "|| " & Npclist(TempCharIndex).Name & " - Vida: " & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & "." & FONTTYPE_INFO)
                    'Call SendData(ToIndex, UserIndex, 0, "|| " & Npclist(TempCharIndex).Name & " - Vida: " & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & " - Exp Total: " & Npclist(NpcIndex).GiveEXP & " - ORO: " & Npclist(NpcIndex).GiveEXP & "." & FONTTYPE_INFO)
                End If
                
            End If
            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(UserIndex).flags.TargetNpc = TempCharIndex
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
        
    End If
    
    If FoundChar = 0 Then
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
        Call SendData(ToIndex, UserIndex, 0, "||No ves nada interesante." & FONTTYPE_INFO)
    End If

Else
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
        Call SendData(ToIndex, UserIndex, 0, "||No ves nada interesante." & FONTTYPE_INFO)
    End If
End If


End Sub
Function FindDirection(Pos As WorldPos, Target As WorldPos) As Byte
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
Dim x As Integer
Dim y As Integer

x = Pos.x - Target.x
y = Pos.y - Target.y

'NE
If Sgn(x) = -1 And Sgn(y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

'NW
If Sgn(x) = 1 And Sgn(y) = 1 Then
    FindDirection = WEST
    Exit Function
End If

'SW
If Sgn(x) = 1 And Sgn(y) = -1 Then
    FindDirection = WEST
    Exit Function
End If

'SE
If Sgn(x) = -1 And Sgn(y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

'Sur
If Sgn(x) = 0 And Sgn(y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

'norte
If Sgn(x) = 0 And Sgn(y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

'oeste
If Sgn(x) = 1 And Sgn(y) = 0 Then
    FindDirection = WEST
    Exit Function
End If

'este
If Sgn(x) = -1 And Sgn(y) = 0 Then
    FindDirection = EAST
    Exit Function
End If

'misma
If Sgn(x) = 0 And Sgn(y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function



