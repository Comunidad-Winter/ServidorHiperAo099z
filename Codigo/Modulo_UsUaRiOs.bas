Attribute VB_Name = "UsUaRiOs"
'Argentum Online 0.9.0.2
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

'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'                        Modulo Usuarios
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'Rutinas de los usuarios
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)

Dim DaExp As Integer
DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)

Call AddtoVar(UserList(AttackerIndex).Stats.Exp, DaExp, MAXEXP)
     
'Lo mata
Call SendData(ToIndex, AttackerIndex, 0, "||Has matado " & UserList(VictimIndex).Name & "!" & FONTTYPE_FIGHT)
Call SendData(ToIndex, AttackerIndex, 0, "||Has ganado " & DaExp & " puntos de experiencia." & FONTTYPE_FIGHT)
      
Call SendData(ToIndex, VictimIndex, 0, "||" & UserList(AttackerIndex).Name & " te ha matado!" & FONTTYPE_FIGHT)

If Not Criminal(VictimIndex) Then
     Call AddtoVar(UserList(AttackerIndex).Reputacion.AsesinoRep, vlASESINO * 2, MAXREP)
     UserList(AttackerIndex).Reputacion.BurguesRep = 0
     UserList(AttackerIndex).Reputacion.NobleRep = 0
     UserList(AttackerIndex).Reputacion.PlebeRep = 0
Else
     Call AddtoVar(UserList(AttackerIndex).Reputacion.NobleRep, vlNoble, MAXREP)
End If

Call UserDie(VictimIndex)

Call AddtoVar(UserList(AttackerIndex).Stats.UsuariosMatados, 1, 31000)

'Log
Call LogAsesinato(UserList(AttackerIndex).Name & " asesino a " & UserList(VictimIndex).Name)

End Sub


Sub RevivirUsuario(ByVal UserIndex As Integer)

UserList(UserIndex).flags.Muerto = 0
UserList(UserIndex).Stats.MinHP = 10

Call DarCuerpoDesnudo(UserIndex)
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).OrigChar.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call SendUserStatsBox(UserIndex)

End Sub


Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal UserIndex As Integer, _
ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

UserList(UserIndex).Char.Body = Body
UserList(UserIndex).Char.Head = Head
UserList(UserIndex).Char.Heading = Heading
UserList(UserIndex).Char.WeaponAnim = Arma
UserList(UserIndex).Char.ShieldAnim = Escudo
UserList(UserIndex).Char.CascoAnim = Casco

Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(UserIndex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(UserIndex).Char.FX & "," & UserList(UserIndex).Char.loops & "," & Casco)

End Sub


Sub EnviarSubirNivel(ByVal UserIndex As Integer, ByVal Puntos As Integer)
Call SendData(ToIndex, UserIndex, 0, "SUNI" & Puntos)
End Sub

Sub EnviarSkills(ByVal UserIndex As Integer)

Dim i As Integer
Dim cad$
For i = 1 To NUMSKILLS
   cad$ = cad$ & UserList(UserIndex).Stats.UserSkills(i) & ","
Next
SendData ToIndex, UserIndex, 0, "SKILLS" & cad$
End Sub
Sub EnviarMatados(ByVal UserIndex As Integer)
Dim cad$

   cad$ = BuscaMatados(UserIndex, "MUERTES", "UserMuertes") - BuscaMatados(UserIndex, "FACCIONES", "CiudMatados") - BuscaMatados(UserIndex, "FACCIONES", "CrimMatados") & "," & BuscaMatados(UserIndex, "FACCIONES", "CiudMatados") & "," & BuscaMatados(UserIndex, "FACCIONES", "CrimMatados") & "," & BuscaMatados(UserIndex, "MUERTES", "NpcsMuertes") & ","
   
SendData ToIndex, UserIndex, 0, "MATADOS" & cad$

End Sub

Sub EnviarFama(ByVal UserIndex As Integer)
Dim cad$
cad$ = cad$ & UserList(UserIndex).Reputacion.AsesinoRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.BandidoRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.BurguesRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.LadronesRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.NobleRep & ","
cad$ = cad$ & UserList(UserIndex).Reputacion.PlebeRep & ","

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6

UserList(UserIndex).Reputacion.Promedio = L

cad$ = cad$ & UserList(UserIndex).Reputacion.Promedio

SendData ToIndex, UserIndex, 0, "FAMA" & cad$

End Sub

Sub EnviarAtrib(ByVal UserIndex As Integer)
Dim i As Integer
Dim cad$
For i = 1 To NUMATRIBUTOS
  cad$ = cad$ & UserList(UserIndex).Stats.UserAtributos(i) & ","
Next
Call SendData(ToIndex, UserIndex, 0, "ATR" & cad$)
End Sub

Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer)

On Error GoTo ErrorHandler
   
    CharList(UserList(UserIndex).Char.CharIndex) = 0
    
    If UserList(UserIndex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).UserIndex = 0
    
    'Le mandamos el mensaje para que borre el personaje a los clientes que este en el mismo mapa
    Call SendData(ToMap, UserIndex, UserList(UserIndex).Pos.Map, "BP" & UserList(UserIndex).Char.CharIndex)
    
    UserList(UserIndex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
    
    Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar")

End Sub

Sub MakeUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
On Error Resume Next

Dim CharIndex As Integer

If InMapBounds(Map, x, y) Then

       'If needed make a new character in list
       If UserList(UserIndex).Char.CharIndex = 0 Then
           CharIndex = NextOpenCharIndex
           UserList(UserIndex).Char.CharIndex = CharIndex
           CharList(CharIndex) = UserIndex
       End If
       
       'Place character on map
       MapData(Map, x, y).UserIndex = UserIndex
       
       'Send make character command to clients
       Dim klan$
       klan$ = UserList(UserIndex).GuildInfo.GuildName
       Dim bCr As Byte
       Dim bGm As Byte
       bCr = Criminal(UserIndex)
       '[Sicarul]
       If UserList(UserIndex).flags.Privilegios > 1 Then bGm = 1 Else bGm = 0
       '[/Sicarul]
       If klan$ <> "" Then
            Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & x & "," & y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & " <" & klan$ & ">" & "," & bCr & "," & bGm)
       Else
            Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(UserIndex).Char.Body & "," & UserList(UserIndex).Char.Head & "," & UserList(UserIndex).Char.Heading & "," & UserList(UserIndex).Char.CharIndex & "," & x & "," & y & "," & UserList(UserIndex).Char.WeaponAnim & "," & UserList(UserIndex).Char.ShieldAnim & "," & UserList(UserIndex).Char.FX & "," & 999 & "," & UserList(UserIndex).Char.CascoAnim & "," & UserList(UserIndex).Name & "," & bCr & "," & bGm)
       End If
End If

End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)

On Error GoTo errhandler

Dim Pts As Integer
Dim AumentoHIT As Integer
Dim AumentoST As Integer
Dim AumentoMANA As Integer
Dim WasNewbie As Boolean

'[Wag]
Dim Powa As Integer

Powa = val(GetVar(IniPath & "Config.ini", "NOTOKAR", "NOPSD"))

'[/wag]
'�Alcanzo el maximo nivel?
If UserList(UserIndex).Stats.ELV >= STAT_MAXELV Then
    UserList(UserIndex).Stats.Exp = 0
    UserList(UserIndex).Stats.ELU = 0
    Exit Sub
End If
If UserList(UserIndex).Stats.ELU = 0 Then Exit Sub

WasNewbie = EsNewbie(UserIndex)

'Si exp >= then Exp para subir de nivel entonce subimos el nivel
If UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU Then
    
    
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_NIVEL)
    Call SendData(ToIndex, UserIndex, 0, "||�Has subido de nivel!" & FONTTYPE_INFO)
    
    
    If UserList(UserIndex).Stats.ELV = 1 Then
      Pts = 20
      
    Else
      Pts = 10
    End If
    
    UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + Pts
    
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & Pts & " skillpoints." & FONTTYPE_INFO)
       
    UserList(UserIndex).Stats.ELV = UserList(UserIndex).Stats.ELV + 1
    
    UserList(UserIndex).Stats.Exp = 0
    
    If Not EsNewbie(UserIndex) And WasNewbie Then Call QuitarNewbieObj(UserIndex)
    
    If UserList(UserIndex).Stats.ELV < 11 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.3
    ElseIf UserList(UserIndex).Stats.ELV < 25 Then
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.4
    Else
        UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.1
    End If
    
    Dim AumentoHP As Integer
    Select Case UserList(UserIndex).Clase
        Case "Guerrero"
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            
            '�?�?�?�?�?�?� HitPoints �?�?�?�?�?�?�
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            
            '�?�?�?�?�?�?� Stamina �?�?�?�?�?�?�
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            
            '�?�?�?�?�?�?� Golpe �?�?�?�?�?�?�
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
        
        Case "Cazador"
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            
            '�?�?�?�?�?�?� HitPoints �?�?�?�?�?�?�
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            
            '�?�?�?�?�?�?� Stamina �?�?�?�?�?�?�
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            
            '�?�?�?�?�?�?� Golpe �?�?�?�?�?�?�
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
            
        Case "Pirata"
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            
            '�?�?�?�?�?�?� HitPoints �?�?�?�?�?�?�
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            
            '�?�?�?�?�?�?� Stamina �?�?�?�?�?�?�
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            
            '�?�?�?�?�?�?� Golpe �?�?�?�?�?�?�
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
            
        Case "Paladin"
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) + AdicionalHPGuerrero
            AumentoST = 15
            AumentoHIT = 3
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            
            'HP
            Call AddtoVar(UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
            'Mana
            Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN)
            
            'STA
            Call AddtoVar(UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA)
            
            'Golpe
            Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT)
            Call AddtoVar(UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT)
                        
        Case "Ladron"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTLadron
            AumentoHIT = 1
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
            
        Case "Mago"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2) - AdicionalHPGuerrero / 2
            If AumentoHP < 1 Then AumentoHP = 4
            AumentoST = 15 - AdicionalSTLadron / 2
            If AumentoST < 1 Then AumentoST = 5
            AumentoHIT = 1
            AumentoMANA = 3 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Le�ador"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTLe�ador
            AumentoHIT = 2
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Minero"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTMinero
            AumentoHIT = 2
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Pescador"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15 + AdicionalSTPescador
            AumentoHIT = 1
            
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
                   
        Case "Clerigo"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Druida"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case "Asesino"
            
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 3
            AumentoMANA = UserList(UserIndex).Stats.UserAtributos(Inteligencia)
                
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
            
        Case "Bardo"
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(Inteligencia)
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Mana
            AddtoVar UserList(UserIndex).Stats.MaxMAN, AumentoMANA, STAT_MAXMAN
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
        Case Else
            AumentoHP = RandomNumber(4, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 2)
            AumentoST = 15
            AumentoHIT = 2
            'HP
            AddtoVar UserList(UserIndex).Stats.MaxHP, AumentoHP, STAT_MAXHP
            'STA
            AddtoVar UserList(UserIndex).Stats.MaxSta, AumentoST, STAT_MAXSTA
            'Golpe
            AddtoVar UserList(UserIndex).Stats.MaxHIT, AumentoHIT, STAT_MAXHIT
            AddtoVar UserList(UserIndex).Stats.MinHIT, AumentoHIT, STAT_MAXHIT
    End Select
    
    'AddtoVar UserList(UserIndex).Stats.MaxHIT, 2, STAT_MAXHIT
    'AddtoVar UserList(UserIndex).Stats.MinHIT, 2, STAT_MAXHIT
    'AddtoVar UserList(UserIndex).Stats.Def, 2, STAT_MAXDEF
    
    If AumentoHP > 0 Then SendData ToIndex, UserIndex, 0, "||Has ganado " & AumentoHP & " puntos de vida." & FONTTYPE_INFO
    If AumentoST > 0 Then SendData ToIndex, UserIndex, 0, "||Has ganado " & AumentoST & " puntos de vitalidad." & FONTTYPE_INFO
    If AumentoMANA > 0 Then SendData ToIndex, UserIndex, 0, "||Has ganado " & AumentoMANA & " puntos de magia." & FONTTYPE_INFO
    If AumentoHIT > 0 Then
        SendData ToIndex, UserIndex, 0, "||Tu golpe maximo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
        SendData ToIndex, UserIndex, 0, "||Tu golpe minimo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
    End If
    
    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
    Call EnviarSkills(UserIndex)
    Call EnviarSubirNivel(UserIndex, Pts)
   
        SendUserStatsBox UserIndex
    If UserList(UserIndex).Stats.ELV > Powa And UserList(UserIndex).flags.Privilegios < 1 Then
    Call WriteVar(IniPath & "Config.ini", "NOTOKAR", "NOPSD", val(UserList(UserIndex).Stats.ELV))
    Call WriteVar(IniPath & "Config.ini", "NOTOKAR", "NPSDO", val(UserList(UserIndex).Name))
    Call SendData(ToAll, UserIndex, 0, "||Ahora " & UserList(UserIndex).Name & " es el nivel mas alto de el servidor." & FONTTYPE_INFO)
    End If

    
End If


Exit Sub

errhandler:
    LogError ("Error en la subrutina CheckUserLevel")
End Sub




Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean

PuedeAtravesarAgua = _
  UserList(UserIndex).flags.Navegando = 1 Or _
  UserList(UserIndex).flags.Vuela = 1

End Function

Private Sub EnviaNuevaPosUsuarioPj(ByVal UserIndex As Integer, ByVal Quien As Integer, ByVal Heading As Integer)
'       Dim klan$
'       klan$ = UserList(UserIndex).GuildInfo.GuildName
'       Dim bCr As Byte
'       bCr = Criminal(UserIndex)
'
'       'Call SendData(ToIndex, UserIndex, 0, "BP" & UserList(Quien).Char.CharIndex)
'
'       If klan$ <> "" Then
'            Call SendData(ToIndex, UserIndex, 0, "CC" & UserList(Quien).Char.Body & "," & UserList(Quien).Char.Head & "," & UserList(Quien).Char.Heading & "," & UserList(Quien).Char.CharIndex & "," & UserList(Quien).Pos.X & "," & UserList(Quien).Pos.Y & "," & UserList(Quien).Char.WeaponAnim & "," & UserList(Quien).Char.ShieldAnim & "," & UserList(Quien).Char.FX & "," & 999 & "," & UserList(Quien).Char.CascoAnim & "," & UserList(Quien).Name & " <" & klan$ & ">" & "," & bCr)
'       Else
'            Call SendData(ToIndex, UserIndex, 0, "CC" & UserList(Quien).Char.Body & "," & UserList(Quien).Char.Head & "," & UserList(Quien).Char.Heading & "," & UserList(Quien).Char.CharIndex & "," & UserList(Quien).Pos.X & "," & UserList(Quien).Pos.Y & "," & UserList(Quien).Char.WeaponAnim & "," & UserList(Quien).Char.ShieldAnim & "," & UserList(Quien).Char.FX & "," & 999 & "," & UserList(Quien).Char.CascoAnim & "," & UserList(Quien).Name & "," & bCr)
'       End If

'Call SendData(ToIndex, UserIndex, 0, "MP" & UserList(Quien).Char.CharIndex  & "," & UserList(Quien).Pos.X & "," & UserList(Quien).Pos.Y)
Call SendData(ToIndex, UserIndex, 0, "MP" & UserList(Quien).Char.CharIndex & "," & UserList(Quien).Pos.x & "," & UserList(Quien).Pos.y)

End Sub

Private Sub EnviaNuevaPosNPC(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Heading As Integer)
'Dim cX As Integer, cY As Integer
'
'Select Case Heading
'Case NORTH: cX = 0: cY = -1
'Case SOUTH: cX = 0: cY = 1
'Case WEST:  cX = -1: cY = 0
'Case EAST:  cX = 1: cY = 0
'End Select
'
''Call SendData(ToIndex, UserIndex, 0, "BP" & Npclist(NpcIndex).Char.CharIndex)
''Call SendData(ToIndex, UserIndex, 0, "CC" & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading & "," & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Pos.X & "," & Npclist(NpcIndex).Pos.Y)
'Call SendData(ToIndex, UserIndex, 0, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Pos.X + cX & "," & Npclist(NpcIndex).Pos.Y + cY)
Call SendData(ToIndex, UserIndex, 0, "MP" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Pos.x & "," & Npclist(NpcIndex).Pos.y)
'Call SendData(ToIndex, UserIndex, 0, "CP" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).Char.Body & "," & Npclist(NpcIndex).Char.Head & "," & Npclist(NpcIndex).Char.Heading)

End Sub

Private Sub EnviaGenteEnnuevoRango(ByVal UserIndex As Integer, ByVal nHeading As Byte)
Dim x As Integer, y As Integer
Dim M As Integer

M = UserList(UserIndex).Pos.Map

Select Case nHeading
Case NORTH, SOUTH
    '***** GENTE NUEVA *****
    If nHeading = NORTH Then
        y = UserList(UserIndex).Pos.y - MinYBorder
    Else 'SOUTH
        y = UserList(UserIndex).Pos.y + MinYBorder
    End If
    For x = UserList(UserIndex).Pos.x - MinXBorder + 1 To UserList(UserIndex).Pos.x + MinXBorder - 1
        If MapData(M, x, y).UserIndex > 0 Then
            Call EnviaNuevaPosUsuarioPj(UserIndex, MapData(M, x, y).UserIndex, nHeading)
        ElseIf MapData(M, x, y).NpcIndex > 0 Then
            Call EnviaNuevaPosNPC(UserIndex, MapData(M, x, y).NpcIndex, nHeading)
        End If
    Next x
'    '***** GENTE VIEJA *****
'    If nHeading = NORTH Then
'        Y = UserList(UserIndex).Pos.Y + MinYBorder
'    Else 'SOUTH
'        Y = UserList(UserIndex).Pos.Y - MinYBorder
'    End If
'    For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
'        If MapData(M, X, Y).UserIndex > 0 Then
'            Call SendData(ToIndex, UserIndex, 0, "BP" & UserList(MapData(M, X, Y).UserIndex).Char.CharIndex)
'        ElseIf MapData(M, X, Y).NpcIndex > 0 Then
'            Call SendData(ToIndex, UserIndex, 0, "BP" & Npclist(MapData(M, X, Y).NpcIndex).Char.CharIndex)
'        End If
'    Next X
Case EAST, WEST
    '***** GENTE NUEVA *****
    If nHeading = EAST Then
        x = UserList(UserIndex).Pos.x + MinXBorder
    Else 'SOUTH
        x = UserList(UserIndex).Pos.x - MinXBorder
    End If
    For y = UserList(UserIndex).Pos.y - MinYBorder + 1 To UserList(UserIndex).Pos.y + MinYBorder - 1
        If MapData(M, x, y).UserIndex > 0 Then
            Call EnviaNuevaPosUsuarioPj(UserIndex, MapData(M, x, y).UserIndex, nHeading)
        ElseIf MapData(M, x, y).NpcIndex > 0 Then
            Call EnviaNuevaPosNPC(UserIndex, MapData(M, x, y).NpcIndex, nHeading)
        End If
    Next y
'    '****** GENTE VIEJA *****
'    If nHeading = EAST Then
'        X = UserList(UserIndex).Pos.X - MinXBorder
'    Else 'SOUTH
'        X = UserList(UserIndex).Pos.X + MinXBorder
'    End If
'    For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
'        If MapData(M, X, Y).UserIndex > 0 Then
'            Call SendData(ToIndex, UserIndex, 0, "BP" & UserList(MapData(M, X, Y).UserIndex).Char.CharIndex)
'        ElseIf MapData(M, X, Y).NpcIndex > 0 Then
'            Call SendData(ToIndex, UserIndex, 0, "BP" & Npclist(MapData(M, X, Y).NpcIndex).Char.CharIndex)
'        End If
'    Next Y
End Select

End Sub

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As Byte)

On Error Resume Next

Dim nPos As WorldPos

'Move
nPos = UserList(UserIndex).Pos
Call HeadtoPos(nHeading, nPos)



If LegalPos(UserList(UserIndex).Pos.Map, nPos.x, nPos.y, PuedeAtravesarAgua(UserIndex)) Then
    
    '[Alejo-18-5]
    Call SendData(ToMapButIndex, UserIndex, UserList(UserIndex).Pos.Map, "MP" & UserList(UserIndex).Char.CharIndex & "," & nPos.x & "," & nPos.y & "," & "1")
    'Call SendData(ToPCAreaButIndex, UserIndex, UserList(UserIndex).Pos.Map, "MP" & UserList(UserIndex).Char.CharIndex & "," & nPos.X & "," & nPos.Y & "," & "1")

    'Call EnviaGenteEnnuevoRango(UserIndex, nHeading)
    
    'Update map and user pos
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).UserIndex = 0
    UserList(UserIndex).Pos = nPos
    UserList(UserIndex).Char.Heading = nHeading
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y).UserIndex = UserIndex
    
Else
    'else correct user's pos
    Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.x & "," & UserList(UserIndex).Pos.y)
End If

End Sub

Sub ChangeUserInv(UserIndex As Integer, Slot As Byte, Object As UserOBJ)


UserList(UserIndex).Invent.Object(Slot) = Object

If Object.ObjIndex > 0 Then

    Call SendData(ToIndex, UserIndex, 0, "CSI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
    & ObjData(Object.ObjIndex).ObjType & "," _
    & ObjData(Object.ObjIndex).MaxHIT & "," _
    & ObjData(Object.ObjIndex).MinHIT & "," _
    & ObjData(Object.ObjIndex).MaxDef & "," _
    & ObjData(Object.ObjIndex).Valor \ 3)

Else

    Call SendData(ToIndex, UserIndex, 0, "CSI" & Slot & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")

End If


End Sub

Function NextOpenCharIndex() As Integer

Dim LoopC As Integer

For LoopC = 1 To LastChar + 1
    If CharList(LoopC) = 0 Then
        NextOpenCharIndex = LoopC
        NumChars = NumChars + 1
        If LoopC > LastChar Then LastChar = LoopC
        Exit Function
    End If
Next LoopC

End Function

Function NextOpenUser() As Integer

Dim LoopC As Integer
  
For LoopC = 1 To MaxUsers + 1
  If LoopC > MaxUsers Then Exit For
  If (UserList(LoopC).ConnID = -1) Then Exit For
Next LoopC
  
NextOpenUser = LoopC

End Function

Sub SendUserStatsBox(ByVal UserIndex As Integer)
Call SendData(ToIndex, UserIndex, 0, "EST" & UserList(UserIndex).Stats.MaxHP & "," & UserList(UserIndex).Stats.MinHP & "," & UserList(UserIndex).Stats.MaxMAN & "," & UserList(UserIndex).Stats.MinMAN & "," & UserList(UserIndex).Stats.MaxSta & "," & UserList(UserIndex).Stats.MinSta & "," & UserList(UserIndex).Stats.GLD & "," & UserList(UserIndex).Stats.ELV & "," & UserList(UserIndex).Stats.ELU & "," & UserList(UserIndex).Stats.Exp)
End Sub

Sub EnviarHambreYsed(ByVal UserIndex As Integer)
Call SendData(ToIndex, UserIndex, 0, "EHYS" & UserList(UserIndex).Stats.MaxAGU & "," & UserList(UserIndex).Stats.MinAGU & "," & UserList(UserIndex).Stats.MaxHam & "," & UserList(UserIndex).Stats.MinHam)
End Sub

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

Call SendData(ToIndex, sendIndex, 0, "||Estadisticas de: " & UserList(UserIndex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & UserList(UserIndex).Stats.ELU & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Vitalidad: " & UserList(UserIndex).Stats.FIT & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Salud: " & UserList(UserIndex).Stats.MinHP & "/" & UserList(UserIndex).Stats.MaxHP & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta & FONTTYPE_INFO)

If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHIT & ")" & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHIT & FONTTYPE_INFO)
End If

If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
    Call SendData(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: 0" & FONTTYPE_INFO)
End If

If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
    Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)
End If

If UserList(UserIndex).GuildInfo.GuildName <> "" Then
    Call SendData(ToIndex, sendIndex, 0, "||Clan: " & UserList(UserIndex).GuildInfo.GuildName & FONTTYPE_INFO)
    If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
       If UserList(UserIndex).GuildInfo.ClanFundado = UserList(UserIndex).GuildInfo.GuildName Then
            Call SendData(ToIndex, sendIndex, 0, "||Status:" & "Fundador/Lider" & FONTTYPE_INFO)
       Else
            Call SendData(ToIndex, sendIndex, 0, "||Status:" & "Lider" & FONTTYPE_INFO)
       End If
    Else
        Call SendData(ToIndex, sendIndex, 0, "||Status:" & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    Call SendData(ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
End If


Call SendData(ToIndex, sendIndex, 0, "||Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).Pos.x & "," & UserList(UserIndex).Pos.y & " en mapa " & UserList(UserIndex).Pos.Map & FONTTYPE_INFO)

End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & UserList(UserIndex).Invent.NroItems & " objetos." & FONTTYPE_INFO)
For j = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
        Call SendData(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).Amount & FONTTYPE_INFO)
    End If
Next
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(ToIndex, sendIndex, 0, "||" & UserList(UserIndex).Name & FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call SendData(ToIndex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j) & FONTTYPE_INFO)
Next
End Sub


Sub UpdateUserMap(ByVal UserIndex As Integer)

Dim Map As Integer
Dim x As Integer
Dim y As Integer

Map = UserList(UserIndex).Pos.Map

For y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize

        If MapData(Map, x, y).UserIndex > 0 And UserIndex <> MapData(Map, x, y).UserIndex Then
            Call MakeUserChar(ToIndex, UserIndex, 0, MapData(Map, x, y).UserIndex, Map, x, y)
            If UserList(MapData(Map, x, y).UserIndex).flags.Invisible = 1 Then Call SendData(ToIndex, UserIndex, 0, "NOVER" & UserList(MapData(Map, x, y).UserIndex).Char.CharIndex & ",1")
        End If

        If MapData(Map, x, y).NpcIndex > 0 Then
            Call MakeNPCChar(ToIndex, UserIndex, 0, MapData(Map, x, y).NpcIndex, Map, x, y)
        End If

        If MapData(Map, x, y).OBJInfo.ObjIndex > 0 Then
            Call MakeObj(ToIndex, UserIndex, 0, MapData(Map, x, y).OBJInfo, Map, x, y)
            
            If ObjData(MapData(Map, x, y).OBJInfo.ObjIndex).ObjType = OBJTYPE_PUERTAS Then
                      Call Bloquear(ToIndex, UserIndex, 0, Map, x, y, MapData(Map, x, y).Blocked)
                      Call Bloquear(ToIndex, UserIndex, 0, Map, x - 1, y, MapData(Map, x - 1, y).Blocked)
            End If
        End If
        
    Next x
Next y

End Sub

Function DameUserindex(SocketId As Integer) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserindex = 0
        Exit Function
    End If
    
Loop
  
DameUserindex = LoopC

End Function

Function DameUserIndexConNombre(ByVal Nombre As String) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Nombre = UCase$(Nombre)

Do Until UCase$(UserList(LoopC).Name) = Nombre

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserIndexConNombre = 0
        Exit Function
    End If
    
Loop
  
DameUserIndexConNombre = LoopC

End Function


Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not Criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then Call SendData(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "||��" & UserList(UserIndex).Name & " esta atacando tu mascota!!" & FONTTYPE_FIGHT)
End If

End Function

Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)


'Guardamos el usuario que ataco el npc
Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name

If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)

If EsMascotaCiudadano(NpcIndex, UserIndex) Then
            Call VolverCriminal(UserIndex)
            Npclist(NpcIndex).Movement = NPCDEFENSA
            Npclist(NpcIndex).Hostile = 1
Else
    'Reputacion
    If Npclist(NpcIndex).Stats.Alineacion = 0 Then
       If Npclist(NpcIndex).NPCtype = NPCTYPE_GUARDIAS Then
                Call VolverCriminal(UserIndex)
       Else
                Call AddtoVar(UserList(UserIndex).Reputacion.BandidoRep, vlASALTO, MAXREP)
       End If
    ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
       Call AddtoVar(UserList(UserIndex).Reputacion.PlebeRep, vlCAZADOR / 2, MAXREP)
    End If
    
    'hacemos que el npc se defienda
           Npclist(NpcIndex).Movement = NPCDEFENSA
           Npclist(NpcIndex).Hostile = 1
    
End If


End Sub

Function PuedeApu�alar(ByVal UserIndex As Integer) As Boolean

If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApu�alar = _
 ((UserList(UserIndex).Stats.UserSkills(Apu�alar) >= MIN_APU�ALAR) _
 And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apu�ala = 1)) _
 Or _
  ((UserList(UserIndex).Clase = "Asesino") And _
  (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apu�ala = 1))
Else
 PuedeApu�alar = False
End If
End Function
Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
Dim Aumenta As Integer
Aumenta = RandomNumber(1, 15)
If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
If Aumenta = 4 Then
Call AddtoVar(UserList(UserIndex).Stats.UserSkills(Skill), 1, MAXSKILLPOINTS)
Call SendData(ToIndex, UserIndex, 0, "||�Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts." & FONTTYPE_INFO)
Call AddtoVar(UserList(UserIndex).Stats.Exp, 300, MAXEXP)
Call SendData(ToIndex, UserIndex, 0, "||�Has ganado 300 puntos de experiencia!" & FONTTYPE_FIGHT)
Call CheckUserLevel(UserIndex)
End If
End Sub


Sub UserDie(ByVal UserIndex As Integer)
'Call LogTarea("Sub UserDie")
On Error GoTo ErrorHandler

'Sonido
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_USERMUERTE)


'Quitar el dialogo del user muerto
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "QDL" & UserList(UserIndex).Char.CharIndex)

UserList(UserIndex).Stats.MinHP = 0
UserList(UserIndex).flags.AtacadoPorNpc = 0
UserList(UserIndex).flags.AtacadoPorUser = 0
UserList(UserIndex).flags.Envenenado = 0
UserList(UserIndex).flags.Muerto = 1



Dim aN As Integer

aN = UserList(UserIndex).flags.AtacadoPorNpc

If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = ""
End If

'<<<< Paralisis >>>>
If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).flags.Paralizado = 0
    Call SendData(ToIndex, UserIndex, 0, "PARADOK")
End If

'<<<< Descansando >>>>
If UserList(UserIndex).flags.Descansar Then
    UserList(UserIndex).flags.Descansar = False
    Call SendData(ToIndex, UserIndex, 0, "DOK")
End If

'<<<< Meditando >>>>
If UserList(UserIndex).flags.Meditando Then
    UserList(UserIndex).flags.Meditando = False
    Call SendData(ToIndex, UserIndex, 0, "MEDOK")
End If

' << Si es newbie no pierde el inventario >>
If Not EsNewbie(UserIndex) Or Criminal(UserIndex) Then
    Call TirarTodo(UserIndex)
Else
    If EsNewbie(UserIndex) Then Call TirarTodosLosItemsNoNewbies(UserIndex)
End If

' DESEQUIPA TODOS LOS OBJETOS
'desequipar armadura
If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
End If
'desequipar arma
If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
End If
'desequipar casco
If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
End If
'desequipar herramienta
If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
End If
'desequipar municiones
If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
End If

' << Reseteamos los posibles FX sobre el personaje >>
If UserList(UserIndex).Char.loops = LoopAdEternum Then
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
End If

'<< Cambiamos la apariencia del char >>
If UserList(UserIndex).flags.Navegando = 0 Then
    UserList(UserIndex).Char.Body = iCuerpoMuerto
    UserList(UserIndex).Char.Head = iCabezaMuerto
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
Else
    UserList(UserIndex).Char.Body = iFragataFantasmal ';)
End If

Dim i As Integer
For i = 1 To MAXMASCOTAS
    
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
           If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
           Else
                Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = 0
                Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldMovement
                Npclist(UserList(UserIndex).MascotasIndex(i)).Hostile = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.OldHostil
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0
           End If
    End If
    
Next i

UserList(UserIndex).NroMacotas = 0


'If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
'        Dim MiObj As Obj
'        Dim nPos As WorldPos
'        MiObj.ObjIndex = RandomNumber(554, 555)
'        MiObj.Amount = 1
'        nPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
'        Dim ManchaSangre As New cGarbage
'        ManchaSangre.Map = nPos.Map
'        ManchaSangre.X = nPos.X
'        ManchaSangre.Y = nPos.Y
'        Call TrashCollector.Add(ManchaSangre)
'End If

'<< Actualizamos clientes >>
Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, val(UserIndex), UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
Call SendUserStatsBox(UserIndex)


Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE")

End Sub


Sub ContarMuerte(ByVal Muerto As Integer, ByVal atacante As Integer)

If EsNewbie(Muerto) Then Exit Sub

If Criminal(Muerto) Then
        If UserList(atacante).flags.LastCrimMatado <> UserList(Muerto).Name Then
            UserList(atacante).flags.LastCrimMatado = UserList(Muerto).Name
            Call AddtoVar(UserList(atacante).Faccion.CriminalesMatados, 1, 65000)
        End If
        
        If UserList(atacante).Faccion.CriminalesMatados > MAXUSERMATADOS Then
            UserList(atacante).Faccion.CriminalesMatados = 0
            UserList(atacante).Faccion.RecompensasReal = 0
        End If
Else
        If UserList(atacante).flags.LastCiudMatado <> UserList(Muerto).Name Then
            UserList(atacante).flags.LastCiudMatado = UserList(Muerto).Name
            Call AddtoVar(UserList(atacante).Faccion.CiudadanosMatados, 1, 65000)
        End If
        
        If UserList(atacante).Faccion.CiudadanosMatados > MAXUSERMATADOS Then
            UserList(atacante).Faccion.CiudadanosMatados = 0
            UserList(atacante).Faccion.RecompensasCaos = 0
        End If
End If


End Sub

Sub Tilelibre(Pos As WorldPos, nPos As WorldPos)
'Call LogTarea("Sub Tilelibre")

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
hayobj = False
nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.x, nPos.y) Or hayobj
    
    If LoopC > 15 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.y - LoopC To Pos.y + LoopC
        For tX = Pos.x - LoopC To Pos.x + LoopC
        
            If LegalPos(nPos.Map, tX, tY) = True Then
               hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex > 0)
               If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                     nPos.x = tX
                     nPos.y = tY
                     tX = Pos.x + LoopC
                     tY = Pos.y + LoopC
                End If
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

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal FX As Boolean = False)

'Quitar el dialogo
Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "QDL" & UserList(UserIndex).Char.CharIndex)

Call SendData(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "QTDL")

Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

OldMap = UserList(UserIndex).Pos.Map
OldX = UserList(UserIndex).Pos.x
OldY = UserList(UserIndex).Pos.y

Call EraseUserChar(ToMap, 0, OldMap, UserIndex)

UserList(UserIndex).Pos.x = x
UserList(UserIndex).Pos.y = y
UserList(UserIndex).Pos.Map = Map


If OldMap <> Map Then
    Call SendData(ToIndex, UserIndex, 0, "CM" & Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion)
    Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(Map).Music)
    
    Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y)
    
    Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)

    'Update new Map Users
    MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1

    'Update old Map Users
    MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
    If MapInfo(OldMap).NumUsers < 0 Then
        MapInfo(OldMap).NumUsers = 0
    End If

Else
    
    Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.y)
    Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)

End If


Call UpdateUserMap(UserIndex)

        'Seguis invisible al pasar de mapa
        If (UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1) And (Not UserList(UserIndex).flags.AdminInvisible = 1) Then
            Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",1")
        End If

If FX And UserList(UserIndex).flags.AdminInvisible = 0 Then 'FX
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_WARP)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 0)
End If


Call WarpMascotas(UserIndex)

End Sub

Sub WarpMascotas(ByVal UserIndex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer

NroPets = UserList(UserIndex).NroMacotas

For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        PetRespawn(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).flags.Respawn = 0
        PetTypes(i) = UserList(UserIndex).MascotasType(i)
        PetTiempoDeVida(i) = Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia
        Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

For i = 1 To MAXMASCOTAS
    If PetTypes(i) > 0 Then
        UserList(UserIndex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(UserIndex).Pos, False, PetRespawn(i))
        UserList(UserIndex).MascotasType(i) = PetTypes(i)
        'Controlamos que se sumoneo OK
        If UserList(UserIndex).MascotasIndex(i) = MAXNPCS Then
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0
                If UserList(UserIndex).NroMacotas > 0 Then UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
                Exit Sub
        End If
        Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
        Npclist(UserList(UserIndex).MascotasIndex(i)).Movement = SIGUE_AMO
        Npclist(UserList(UserIndex).MascotasIndex(i)).Target = 0
        Npclist(UserList(UserIndex).MascotasIndex(i)).TargetNpc = 0
        Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
        Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

UserList(UserIndex).NroMacotas = NroPets

End Sub


Sub RepararMascotas(ByVal UserIndex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

For i = 1 To MAXMASCOTAS
  If UserList(UserIndex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
Next i

If MascotasReales <> UserList(UserIndex).NroMacotas Then UserList(UserIndex).NroMacotas = 0


End Sub

Sub Cerrar_Usuario(UserIndex As Integer)
    If UserList(UserIndex).flags.Paralizado = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||No Podes Salir del AO Estando paralizado." & FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
        UserList(UserIndex).Counters.Saliendo = True
        UserList(UserIndex).Counters.Salir = IntervaloCerrarConexion
        
        Call SendData(ToIndex, UserIndex, 0, "||Cerrando...Se cerrar� el juego en " & IntervaloCerrarConexion & " segundos..." & FONTTYPE_INFO)
    'ElseIf Not UserList(UserIndex).Counters.Saliendo Then
    '    If NumUsers <> 0 Then NumUsers = NumUsers - 1
    '    Call SendData(ToIndex, UserIndex, 0, "||Gracias por jugar Argentum Online" & FONTTYPE_INFO)
    '    Call SendData(ToIndex, UserIndex, 0, "FINOK")
    '
    '    Call CloseUser(UserIndex)
    '    UserList(UserIndex).ConnID = -1: UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    '    frmMain.Socket2(UserIndex).Cleanup
    '    Unload frmMain.Socket2(UserIndex)
    '    Call ResetUserSlot(UserIndex)
    End If
End Sub

