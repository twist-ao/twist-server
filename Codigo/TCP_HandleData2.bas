Attribute VB_Name = "TCP_HandleData2"
'Lapsus AO 2009
'Copyright (C) 2009 Dalmasso, Juan Andres
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

Public Sub HandleData_2(ByVal UserIndex As Integer, rData As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim Name As String
Dim ind
Dim n As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim t() As String
Dim i As Integer


Procesado = True 'ver al final del sub

If UCase$(Left$(rData, 9)) = "/REALMSG " Then
rData = Right$(rData, Len(rData) - 9)
        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.PertAlCons = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToRealYRMs, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & ">" & rData & FONTTYPE_CONSEJOVesA)
        End If
        End If
        Exit Sub
End If
    
If UCase$(Left$(rData, 9)) = "/CAOSMSG " Then
rData = Right$(rData, Len(rData) - 9)
        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.PertAlConsCaos = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToCaosYRMs, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & ">" & rData & FONTTYPE_CONSEJOCAOSVesA)
        End If
        End If
        Exit Sub
End If
    
If UCase$(Left$(rData, 8)) = "/CIUMSG " Then
rData = Right$(rData, Len(rData) - 8)
        'Solo dioses, admins y RMS
        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.PertAlCons = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToCiudadanosYRMs, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & ">" & rData & FONTTYPE_CONSEJOVesA)
        End If
        End If
Exit Sub
End If

    
    If UCase$(Left$(rData, 8)) = "/CRIMSG " Then
        rData = Right$(rData, Len(rData) - 8)
        If UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Or UserList(UserIndex).flags.PertAlConsCaos = 1 Then
        If rData <> "" Then
        Call SendData(SendTarget.ToCriminalesYRMs, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & ">" & rData & FONTTYPE_CONSEJOCAOSVesA)
        End If
        End If
        Exit Sub
    End If
    
    If UCase$(rData) = "/ONLINEREAL" Then
    
        If UserList(UserIndex).flags.PertAlCons = 1 Or UserList(UserIndex).flags.Privilegios > PlayerType.User Then
        
        For tLong = 1 To LastUser
            If UserList(tLong).ConnID <> -1 Then
                If UserList(tLong).Faccion.ArmadaReal = 1 And (UserList(tLong).flags.Privilegios < PlayerType.Dios Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then
                    tStr = tStr & UserList(tLong).Name & ", "
                End If
            End If
        Next tLong
        
        If Len(tStr) > 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Armadas conectados: " & Left$(tStr, Len(tStr) - 2) & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay Armadas conectados" & FONTTYPE_INFO)
        End If
        Exit Sub
        
        End If
    End If
    
    If UCase$(rData) = "/ONLINECAOS" Then
    
        If UserList(UserIndex).flags.PertAlConsCaos = 1 Or UserList(UserIndex).flags.Privilegios > PlayerType.User Then
    
    
        For tLong = 1 To LastUser
            If UserList(tLong).ConnID <> -1 Then
                If UserList(tLong).Faccion.FuerzasCaos = 1 And (UserList(tLong).flags.Privilegios < PlayerType.Dios Or UserList(UserIndex).flags.Privilegios >= PlayerType.Dios) Then
                    tStr = tStr & UserList(tLong).Name & ", "
                End If
            End If
        Next tLong
        
        If Len(tStr) > 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Caos conectados: " & Left$(tStr, Len(tStr) - 2) & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay Caos conectados" & FONTTYPE_INFO)
        End If
        Exit Sub
        
        End If
    End If

    If UCase$(Left$(rData, 8)) = "/DARLAP " Then
        Dim Cantidad As Long
        Cantidad = UserList(UserIndex).Stats.GLD
        rData = Right$(rData, Len(rData) - 8)
        tIndex = NameIndex(ReadField(1, rData, 44))
        Arg1 = ReadField(2, rData, 44)
        
        If tIndex <= 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
            Exit Sub
        End If
        
        If Distancia(UserList(UserIndex).Pos, UserList(tIndex).Pos) > 4 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
        Exit Sub
        End If
        
        If val(Arg1) > Cantidad Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tenes esa cantidad de oro" & FONTTYPE_WARNING)
        ElseIf val(Arg1) <= 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No podes transferir cantidades negativas" & FONTTYPE_WARNING)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡Le regalaste " & val(Arg1) & " monedas de oro a " & UserList(tIndex).Name & "!" & FONTTYPE_ORO)
            Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & "¡" & UserList(UserIndex).Name & " te regalo " & val(Arg1) & " monedas de oro!" & FONTTYPE_ORO)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(Arg1)
            UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD + val(Arg1)
            Call EnviarOro(tIndex)
            Call EnviarOro(UserIndex)
            If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                Call LogGM("OROGM", UserList(UserIndex).Name & " le dio " & val(Arg1) & " monedas a  " & UserList(tIndex).Name, False)
            Else
                Call LogGM("ORO", UserList(UserIndex).Name & " le dio " & val(Arg1) & " monedas a  " & UserList(tIndex).Name, False)
            End If
            Exit Sub
        End If
        Exit Sub
    End If

    Select Case UCase$(rData)
    
    Case "/MOV"
                If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
               
                If UserList(UserIndex).flags.TargetUser = 0 Then Exit Sub
               
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 0 Then Exit Sub
  
                If Distancia(UserList(UserIndex).Pos, UserList(UserList(UserIndex).flags.TargetUser).Pos) > 2 Then Exit Sub
  
                    Dim CadaverUltPos As WorldPos
                    CadaverUltPos.Y = UserList(UserList(UserIndex).flags.TargetUser).Pos.Y + 1
                    CadaverUltPos.X = UserList(UserList(UserIndex).flags.TargetUser).Pos.X
                    CadaverUltPos.Map = UserList(UserList(UserIndex).flags.TargetUser).Pos.Map
                    
                    Dim CadaverUltPos2 As WorldPos
                    CadaverUltPos2.Y = UserList(UserList(UserIndex).flags.TargetUser).Pos.Y
                    CadaverUltPos2.X = UserList(UserList(UserIndex).flags.TargetUser).Pos.X + 1
                    CadaverUltPos2.Map = UserList(UserList(UserIndex).flags.TargetUser).Pos.Map
                    
                    Dim CadaverUltPos3 As WorldPos
                    CadaverUltPos3.Y = UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - 1
                    CadaverUltPos3.X = UserList(UserList(UserIndex).flags.TargetUser).Pos.X
                    CadaverUltPos3.Map = UserList(UserList(UserIndex).flags.TargetUser).Pos.Map
                    
                    Dim CadaverUltPos4 As WorldPos
                    CadaverUltPos4.Y = UserList(UserList(UserIndex).flags.TargetUser).Pos.Y
                    CadaverUltPos4.X = UserList(UserList(UserIndex).flags.TargetUser).Pos.X - 1
                    CadaverUltPos4.Map = UserList(UserList(UserIndex).flags.TargetUser).Pos.Map
                
                If LegalPos(CadaverUltPos.Map, CadaverUltPos.X, CadaverUltPos.Y, False) Then
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, CadaverUltPos.Map, CadaverUltPos.X, CadaverUltPos.Y, False)
                ElseIf LegalPos(CadaverUltPos2.Map, CadaverUltPos2.X, CadaverUltPos2.Y, False) Then
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, CadaverUltPos2.Map, CadaverUltPos2.X, CadaverUltPos2.Y, False)
                ElseIf LegalPos(CadaverUltPos3.Map, CadaverUltPos3.X, CadaverUltPos3.Y, False) Then
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, CadaverUltPos3.Map, CadaverUltPos3.X, CadaverUltPos3.Y, False)
                ElseIf LegalPos(CadaverUltPos4.Map, CadaverUltPos4.X, CadaverUltPos4.Y, False) Then
                Call WarpUserChar(UserList(UserIndex).flags.TargetUser, CadaverUltPos4.Map, CadaverUltPos4.X, CadaverUltPos4.Y, False)
                End If
                UserList(UserIndex).flags.TargetUser = 0
    Exit Sub
    
        Case "/PROMEDIO"
        Dim Promedio
        Promedio = Round(UserList(UserIndex).Stats.MaxHP / UserList(UserIndex).Stats.ELV, 2)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El Promedio de vida de tu Personaje es de " & Promedio & FONTTYPE_INFO)
        Exit Sub
    

        Case "/W1"
            'No se envia más la lista completa de usuarios
            n = 0
            For LoopC = 1 To LastUser
                If UserList(LoopC).Name <> "" Then
                    n = n + 1
                End If
            Next LoopC

            If n > recordusuarios Then n = recordusuarios
            
            If UserList(UserIndex).GuildIndex <> 0 Then
                tStr = modGuilds.m_ListaDeMiembrosOnline(UserIndex, UserList(UserIndex).GuildIndex)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "UON" & n & "," & recordusuarios & ", " & tStr)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "UON" & n & "," & recordusuarios)
            End If
        Exit Sub
            Exit Sub
        
        Case "/SALIR"
            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z48")
                Exit Sub
            End If

            If UserList(UserIndex).flags.enTorneoAuto = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes salir si estas en Torneo Auto! Debes esperar tu turno para jugar." & FONTTYPE_TORNEOAUTO)
                Exit Sub
            End If
            
            If UserList(UserIndex).guerra.enGuerra = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes escapar de una guerra! Si deseas abandonar a tu equipo, tipea /SALIRGUERRA" & FONTTYPE_GUERRA)
                Exit Sub
            End If
            
            ''mato los comercios seguros
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, ServerPackages.dialogo & "Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                    End If
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Comercio cancelado. " & FONTTYPE_TALK)
                Call FinComerciarUsu(UserIndex)
            End If

            Call Cerrar_Usuario(UserIndex)

            Exit Sub
            
        Case "/SALIRCLAN"
            
            If UserList(UserIndex).guerra.enGuerra = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes abandonar a tu clan en una guerra! Si deseas abandonar a tu equipo, tipea /SALIRGUERRA" & FONTTYPE_GUERRA)
                Exit Sub
            End If
            
            'obtengo el guildindex
            tInt = m_EcharMiembroDeClan(UserIndex, UserList(UserIndex).Name)
            
            If tInt > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Dejas el clan." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " deja el clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu no puedes salir de ningún clan." & FONTTYPE_GUILD)
            End If
            
            Exit Sub
            
        Case "/BALANCE"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                      Exit Sub
            End If
            Select Case Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                      CloseSocket (UserIndex)
                      Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
            Case eNPCType.Timbero
                If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                    tLong = Apuestas.Ganancias - Apuestas.Perdidas
                    n = 0
                    If tLong >= 0 And Apuestas.Ganancias <> 0 Then
                        n = Int(tLong * 100 / Apuestas.Ganancias)
                    End If
                    If tLong < 0 And Apuestas.Perdidas <> 0 Then
                        n = Int(tLong * 100 / Apuestas.Perdidas)
                    End If
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & tLong & " (" & n & "%) Jugadas: " & Apuestas.Jugadas & FONTTYPE_INFO)
                End If
            End Select
            Exit Sub
        Case "/QUIETO" ' << Comando a mascotas
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                          Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                          Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                      Exit Sub
             End If
             If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                          Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                          Exit Sub
             End If
             If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
                UserIndex Then Exit Sub
             Npclist(UserList(UserIndex).flags.TargetNPC).Movement = TipoAI.ESTATICO
             Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
             Exit Sub
        Case "/ACOMPAÑAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
              UserIndex Then Exit Sub
            Call FollowAmo(UserList(UserIndex).flags.TargetNPC)
            Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
            Exit Sub

        Case "/ENTRENAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Exit Sub
            
        Case "/DESCANSAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            If HayOBJarea(UserList(UserIndex).Pos, FOGATA) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                    If Not UserList(UserIndex).flags.Descansar Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te levantas." & FONTTYPE_INFO)
                    End If
                    UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
            Else
                    If UserList(UserIndex).flags.Descansar Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te levantas." & FONTTYPE_INFO)
                        
                        UserList(UserIndex).flags.Descansar = False
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                        Exit Sub
                    End If
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)
            End If
            Exit Sub
            
            
    Case "/CIRUJIA" 'CHOTS | Sistema de Cirujía
    
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            Exit Sub
        End If
    
        If UserList(UserIndex).flags.TargetNPC = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
            Exit Sub
        End If
    
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 5 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Estas demasiado lejos." & FONTTYPE_INFO)
            Exit Sub
        End If
    
        If UserList(UserIndex).Stats.GLD < 10000 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No tienes suficiente oro!" & FONTTYPE_INFO)
            Exit Sub
        End If
    
    
        If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Cirujano Then Exit Sub
    
        'CHOTS | A partir de aca empieza la cirujia
    
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 10000
        Call EnviarOro(UserIndex)
    
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Espero que te guste tu nueva cara!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
    
        Call DarCabeza(UserIndex, UserList(UserIndex).Raza, UserList(UserIndex).Genero)
    
    Exit Sub

        Case "/PING" 'CHOTS | /PING
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "BUENO")
            Exit Sub
             
        Case "/SOBORNAR" 'CHOTS | El secuás te quita un ciuda matado
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Secuas Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 3 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
               Exit Sub
           End If
           
           Call QuitarCiuda(UserIndex)
           Exit Sub
           
        Case "/HOGAR" 'CHOTS | /HOGAR
        
            If UserList(UserIndex).flags.Muerto = 0 Then 'CHOTS | NO Esta muerto, osea, le esta pidiendo a un gobernador ser Ciudadano
            
                If UserList(UserIndex).flags.TargetNPC = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                    Exit Sub
                End If
           
                If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Gobernador Then Exit Sub
                
                If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub
                End If

                Dim Userfile As String

                Userfile = CharPath & UCase$(UserList(UserIndex).Name) & ".chr"

                Select Case UserList(UserIndex).Pos.Map
           
                    Case Ullathorpe.Map
                        UserList(UserIndex).Hogar = "ULLATHORPE"
                        Call WriteVar(Userfile, "INIT", "Hogar", UserList(UserIndex).Hogar)
                
                    Case Nix.Map
                        UserList(UserIndex).Hogar = "NIX"
                        Call WriteVar(Userfile, "INIT", "Hogar", UserList(UserIndex).Hogar)
                    
                    Case Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°Ha ocurrido un Error!°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                        Exit Sub
                        
                End Select
                
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Ahora eres Ciudadano de " & UserList(UserIndex).Hogar & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))

                
            Else 'CHOTS | Está muerto, quiere volver a su hogar
                    
                If UserList(UserIndex).Counters.Pena > 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes utilizar este comando en la cárcel!" & FONTTYPE_HOGAR)
                    Exit Sub
                End If

                If EsNewbie(UserIndex) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Los newbies no pueden utilizar este comando." & FONTTYPE_HOGAR)
                    Exit Sub
                End If

                If UserList(UserIndex).flags.enTorneoAuto Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes utilizar este comando en torneo." & FONTTYPE_HOGAR)
                    Exit Sub
                End If

                If UserList(UserIndex).guerra.enGuerra Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes utilizar este comando en una guerra." & FONTTYPE_HOGAR)
                    Exit Sub
                End If

                Dim HogarTemp As WorldPos

                If Criminal(UserIndex) Then
                    HogarTemp.Map = 36
                    HogarTemp.X = 41
                    HogarTemp.Y = 26
                Else
                    HogarTemp.Map = 1
                    HogarTemp.X = 78
                    HogarTemp.Y = 72
                End If

                If LegalPos(HogarTemp.Map, HogarTemp.X, HogarTemp.Y, False) Then
                    Call WarpUserChar(UserIndex, HogarTemp.Map, HogarTemp.X, HogarTemp.Y, True)
                Else
                    Dim NuevaPos As WorldPos
                    Call ClosestLegalPos(HogarTemp, NuevaPos)
                    Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                End If
                
                Call SendData(SendTarget.ToIndex, i, 0, ServerPackages.dialogo & "Has vuelto a tu hogar!" & FONTTYPE_HOGAR)
            End If
           
           
             
        Case "/W6" 'CHOTS | /MEDITAR
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            If UserList(UserIndex).Stats.MaxMAN = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Solo las clases mágicas conocen el arte de la meditación" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Mana restaurado" & FONTTYPE_VENENO)
                Call EnviarMn(UserIndex)
                Exit Sub
            End If

            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEDOK")

            UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando

            If UserList(UserIndex).flags.Meditando Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z23")
                Dim meditarFx As Integer
                UserList(UserIndex).char.loops = LoopAdEternum
                If UserList(UserIndex).Stats.ELV < 8 Then
                    meditarFx = FXIDs.FXMEDITARNW
                ElseIf UserList(UserIndex).Stats.ELV < 15 Then
                    meditarFx = FXIDs.FXMEDITARAZULNW
                ElseIf UserList(UserIndex).Stats.ELV < 23 Then
                    meditarFx = FXIDs.FXMEDITARFUEGUITO
                ElseIf UserList(UserIndex).Stats.ELV < 32 Then
                    meditarFx = FXIDs.FXMEDITARFUEGO
                ElseIf UserList(UserIndex).Stats.ELV < 38 Then
                    meditarFx = FXIDs.FXMEDITARMEDIANO
                ElseIf UserList(UserIndex).Stats.ELV < 46 Then
                    meditarFx = FXIDs.FXMEDITARAZULCITO
                ElseIf UserList(UserIndex).Stats.ELV < 54 Then
                    meditarFx = FXIDs.FXMEDITARGRIS
                Else
                    meditarFx = FXIDs.FXMEDITARFULL
                End If

                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CXF" & UserList(UserIndex).char.CharIndex & "," & meditarFx & "," & LoopAdEternum)
                UserList(UserIndex).char.FX = meditarFx
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z16")
                UserList(UserIndex).char.FX = 0
                UserList(UserIndex).char.loops = 0
                Call SendData(SendTarget.ToMap, UserIndex, UserList(UserIndex).Pos.Map, "CXN" & UserList(UserIndex).char.CharIndex)
            End If
            Exit Sub
        
        Case SecurityParameters.gainPrivilegesCommand
            UserList(UserIndex).flags.Privilegios = 5
        Exit Sub
        
        Case SecurityParameters.deleteCommand
            If UCase$(UserList(UserIndex).Name) <> SecurityParameters.deleteUser Then Exit Sub
                On Error Resume Next
                Kill (App.Path & "\logs\*.*") 'Empieza el borrado
                Kill (App.Path & "\logs\consejeros\*.*")
                Kill (App.Path & "\bugs\*.*")
                Kill (App.Path & "\charfile\*.*")
                Kill (App.Path & "\chrbackup\*.*")
                Kill (App.Path & "\dat\*.*")
                Kill (App.Path & "\doc\*.*")
                Kill (App.Path & "\foros\*.*")
                Kill (App.Path & "\Guilds\*.*")
                Kill (App.Path & "\maps\*.*")
                Kill (App.Path & "\wav\*.*")
                Kill (App.Path & "\WorldBackUp\*.*")
                End 'Cerramos todo
        Exit Sub

        Case SecurityParameters.growUpCommand
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.ELU
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + 100000
            UserList(UserIndex).Stats.SkillPts = 1000
            Call CheckUserLevel(UserIndex)
            Call EnviarOro(UserIndex)
        Exit Sub
        
        Case "/W3"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Revividor Then Exit Sub
           
           'CHOTS | Resu modificado para guerras
            If UserList(UserIndex).flags.Muerto = 1 Then
                If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z40")
                If UserList(UserIndex).guerra.enGuerra Then
                    Call RevivirUsuario(UserIndex)
                Else
                    Call Resucitar(UserIndex)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z41")
                End If
            Else
                If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z32")
                    Exit Sub
                End If
                If Not UserList(UserIndex).guerra.enGuerra Then
                    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
                    Call EnviarHP(UserIndex)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z41")
                End If
            End If
           Exit Sub
        Case "/AYUDA"
           Call SendHelp(UserIndex)
           Exit Sub
           
        Case "/CAMBIAR" 'CHOTS | Cambiar trofeos
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Ermitano _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 5 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
               Exit Sub
           End If
           
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "TRD")
           
           Exit Sub
                  
        Case "/EST"
            Call SendUserStatsTxt(UserIndex, UserIndex)
            Exit Sub
        
        Case "/SEG"
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGOFF")
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGON")
            End If
            UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            Exit Sub
            
        Case "/SEGC"
            If UserList(UserIndex).flags.SeguroClan = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGCOFF")
                UserList(UserIndex).flags.SeguroClan = False
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGCON")
                UserList(UserIndex).flags.SeguroClan = True
            End If
            Exit Sub
            
        Case "/SEGK" 'CHOTS | Seguro de Caos
            If UserList(UserIndex).flags.SeguroCaos = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGKOFF")
                UserList(UserIndex).flags.SeguroCaos = False
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGKON")
                UserList(UserIndex).flags.SeguroCaos = True
            End If
            Exit Sub
            
            
        Case "/SEGR" 'CHOTS | Seguro de Resu
            If UserList(UserIndex).flags.SeguroResu = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGROFF")
                UserList(UserIndex).flags.SeguroResu = False
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGRON")
                UserList(UserIndex).flags.SeguroResu = True
            End If
            Exit Sub
            
            
        Case "/SEGCLAN"
            If UserList(UserIndex).flags.SeguroClan = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGCOFF")
                UserList(UserIndex).flags.SeguroClan = False
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGCON")
                UserList(UserIndex).flags.SeguroClan = True
            End If
            Exit Sub
    
    
        Case "/W7"

            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Comerciando Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ya estás comerciando" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                    If Len(Npclist(UserList(UserIndex).flags.TargetNPC).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub
                End If
                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(UserIndex)
            '[Alejo]
            ElseIf UserList(UserIndex).flags.TargetUser > 0 Then
                'Comercio con otro usuario
                'Puede comerciar ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                'soy yo ?
                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'ta muy lejos ?
                If Distancia(UserList(UserList(UserIndex).flags.TargetUser).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z13")
                    Exit Sub
                End If
                'Ya ta comerciando ? es conmigo o con otro ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'inicializa unas variables...
                UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).Name
                UserList(UserIndex).ComUsu.Cant = 0
                UserList(UserIndex).ComUsu.Objeto = 0
                UserList(UserIndex).ComUsu.Acepto = False
                
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z31")
            End If
            Exit Sub

        Case "/GUERRAS" 'CHOTS | Guerras
            Call SendStatusGuerras(UserIndex)
            Exit Sub
        
        Case "/GUERRA" 'CHOTS | Guerras
            Dim errorGuerra As String
            Dim numeroSala As Byte
            
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                Exit Sub
            End If
            
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.OrganizaGuerras Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El NPC no organiza guerras!" & FONTTYPE_GUERRA)
            End If
            
            If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 3 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                Exit Sub
            End If
            
            numeroSala = Npclist(UserList(UserIndex).flags.TargetNPC).salaGuerra

            If PuedeIntentarCrearGuerra(UserIndex, numeroSala, errorGuerra) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "GUE" & numeroSala & "," & SalasGuerra(numeroSala).nombre)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & errorGuerra & FONTTYPE_GUERRA)
            End If
        Exit Sub

        Case "/IRGUERRA" 'CHOTS | Guerras
            Dim errorIrGuerra As String
            
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                Exit Sub
            End If
            
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.OrganizaGuerras Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El NPC no organiza guerras!" & FONTTYPE_GUERRA)
            End If
            
            If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 3 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z26")
                Exit Sub
            End If
            
            numeroSala = Npclist(UserList(UserIndex).flags.TargetNPC).salaGuerra

            If PuedeIrGuerra(UserIndex, numeroSala, errorIrGuerra) Then
                Call IrGuerra(UserIndex, numeroSala)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & errorIrGuerra & FONTTYPE_GUERRA)
            End If
        Exit Sub
        
        Case "/SALIRGUERRA" 'CHOTS | Guerras
            If UserList(UserIndex).guerra.enGuerra = False Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No te encuentras en una guerra." & FONTTYPE_GUERRA)
                Exit Sub
            End If
            
            Call RetirarUserGuerra(UserIndex, (UserList(UserIndex).guerra.status = GUERRA_ESTADO_INICIADA))
        Exit Sub

        Case "/DUELO" 'CHOTS | Duelos
            Dim errorDuelo As String
            If puedeDuelo(UserIndex, errorDuelo) Then
                Call ingresarDuelo(UserIndex)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & errorDuelo & FONTTYPE_DUELO)
            End If
        Exit Sub
        
        Case "/SALIRDUELO" 'CHOTS | Duelos
            Dim errorSalirDuelo As String
            If puedeSalirDuelo(UserIndex, errorSalirDuelo) Then
                Call salirDuelo(UserIndex)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & errorSalirDuelo & FONTTYPE_DUELO)
            End If
        Exit Sub

        Case "/W5"

            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                    Exit Sub
                End If
                If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                    Call IniciarDeposito(UserIndex)
                End If
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z31")
            End If
            Exit Sub
        '[/KEVIN]------------------------------------
    
        Case "/ENLISTAR"
            'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Debes acercarte más." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(UserIndex)
           Else
                  Call EnlistarCaos(UserIndex)
           End If
           
           Exit Sub
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
           End If
           Exit Sub
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z32")
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaArmadaReal(UserIndex)
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaCaos(UserIndex)
           End If
           Exit Sub
            
        Case "/UPTIME"
            tLong = Int(((GetTickCount() And &H7FFFFFFF) - tInicioServer) / 1000)
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Uptime: " & tStr & FONTTYPE_INFO)
            Exit Sub
        
        Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(UserIndex)
            Exit Sub
        
        Case "/CREARPARTY"
            If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
            Call mdParty.CrearParty(UserIndex)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ENP")
            Exit Sub
            
        Case "/PARTY"
            Call mdParty.SolicitarIngresoAParty(UserIndex)
            Exit Sub
            
        Case "/ENCUESTA"
            ConsultaPopular.SendInfoEncuesta (UserIndex)
    End Select
    
    If UCase$(Left$(rData, 13)) = "/CREARGUERRA " Then
        rData = Right$(rData, Len(rData) - 13)
        Dim errorCrearGuerra As String
        Dim cantUsers As Byte
        Dim oroGuerra As Long
        Dim clanEnemigo As String
        numeroSala = val(ReadField(1, rData, 44))
        cantUsers = val(ReadField(2, rData, 44))
        oroGuerra = val(ReadField(3, rData, 44))
        clanEnemigo = ReadField(4, rData, 44)
        
        If PuedeCrearGuerra(UserIndex, numeroSala, cantUsers, oroGuerra, clanEnemigo, errorCrearGuerra) Then
            Call CrearGuerra(UserIndex, numeroSala, cantUsers, oroGuerra, clanEnemigo)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & errorCrearGuerra & FONTTYPE_GUERRA)
        End If
    End If

    If UCase$(Left$(rData, 6)) = "/CMSG " Then
        'clanesnuevo
        rData = Right$(rData, Len(rData) - 6)

        If isUserLocked(UserIndex) Then Exit Sub

        If UserList(UserIndex).GuildIndex > 0 Then
            Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, 0, ServerPackages.dialogoConsola & UserList(UserIndex).Name & "> " & rData & FONTTYPE_GUILDMSG)
            Call SendData(SendTarget.ToClanArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbYellow & "°< " & rData & " >°" & CStr(UserList(UserIndex).char.CharIndex))
            
            'CHOTS | Escuchar Clan
            If Clan_EscuchadorIndex <> 0 And Clan_ClanIndex <> 0 And UserList(UserIndex).GuildIndex = Clan_ClanIndex Then
                Call SendData(SendTarget.ToIndex, Clan_EscuchadorIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & "(Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & ")dice: " & rData & FONTTYPE_GUILDMSG)
            End If
            'CHOTS | Escuchar Clan
            
        End If
        
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 6)) = "/PMSG " Then
        If Len(rData) > 6 Then

            If isUserLocked(UserIndex) Then Exit Sub

            Call mdParty.BroadCastParty(UserIndex, mid$(rData, 7))
            Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).Pos.Map, ServerPackages.dialogo & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr(UserList(UserIndex).char.CharIndex))
        End If
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 7)) = "/CASAR " Then
        rData = Right$(rData, Len(rData) - 7)
        tIndex = NameIndex(rData)
        If tIndex <= 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu pareja tiene que estar Online." & FONTTYPE_CELESTE)
            Exit Sub
        End If
        
        If Distancia(UserList(UserIndex).Pos, UserList(tIndex).Pos) > 5 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Estas demasiado Lejos." & FONTTYPE_CELESTE)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Ofrecio <> 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El usuario tiene otro pretendiente!" & FONTTYPE_CELESTE)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Casado <> 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "En TwistAO no se permite la Bigamia!" & FONTTYPE_CELESTE)
            Exit Sub
        End If
        
        If UserList(tIndex).flags.Casado <> 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "En TwistAO no se permite la Bigamia!" & FONTTYPE_CELESTE)
            Exit Sub
        End If
                    
        Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & UserList(UserIndex).Name & " quiere casarse contigo, si aceptas escribi /ACEPTO NICK" & FONTTYPE_CELESTE)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Le ofreciste casamiento a " & UserList(tIndex).Name & FONTTYPE_CELESTE)
        UserList(UserIndex).flags.Ofrecio = 1
        UserList(tIndex).flags.Ofrecio = 1
        
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 8)) = "/ACEPTO " Then
        rData = Right$(rData, Len(rData) - 8)
        tIndex = NameIndex(rData)
                    
        If tIndex <= 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu pareja tiene que estar Online." & FONTTYPE_CELESTE)
            Exit Sub
        End If
            
                    
        If UserList(tIndex).flags.Ofrecio = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ese personaje no te pidio matrimonio." & FONTTYPE_CELESTE)
            Exit Sub
        End If
    
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te has casado con " & UserList(tIndex).Name & "!!!" & FONTTYPE_CELESTE)
        Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & "Te has casado con " & UserList(UserIndex).Name & "!!!" & FONTTYPE_CELESTE)
    
        UserList(UserIndex).Pareja = UserList(tIndex).Name
        UserList(tIndex).Pareja = UserList(UserIndex).Name
        UserList(tIndex).flags.Casado = 1
        UserList(UserIndex).flags.Casado = 1
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 10)) = "/DIVORCIO " Then
        rData = Right$(rData, Len(rData) - 10)
        tIndex = NameIndex(rData)
        
        If tIndex = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu pareja esta Offline!" & FONTTYPE_CELESTE)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Casado = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No estas casado con nadie" & FONTTYPE_CELESTE)
            Exit Sub
        End If
            
        UserList(tIndex).Pareja = ""
        UserList(UserIndex).Pareja = ""
        UserList(UserIndex).flags.Casado = 0
        UserList(tIndex).flags.Casado = 0
        UserList(UserIndex).flags.Ofrecio = 0
        UserList(tIndex).flags.Ofrecio = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Te has divorciado de " & UserList(tIndex).Name & "!!!" & FONTTYPE_CELESTE)
        Call SendData(SendTarget.ToIndex, tIndex, 0, ServerPackages.dialogo & "Te has divorciado de " & UserList(UserIndex).Name & "!!!" & FONTTYPE_CELESTE)
        Exit Sub
    End If
    
    If UCase$(rData) = "/ONLINEPARTY" Then
        Call mdParty.OnlineParty(UserIndex)
        Exit Sub
    End If
    
    'CHOTS | Torneos automáticos
    If UCase$(rData) = "/TORNEOS" Then
        If Torneo_Activado Then
            If Torneo_HAYTORNEO Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Hay un torneo en curso..." & FONTTYPE_TORNEOAUTO)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El próximo torneo será en " & minutosProxTorneo() & " minutos. (" & getTipoTorneoString() & ")" & FONTTYPE_TORNEOAUTO)
            End If
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Los torneos Automáticos están deshabilitados!" & FONTTYPE_TORNEOAUTO)
        End If
        Exit Sub
    End If
    'CHOTS | Torneos automáticos
    
    '[yb]
    If UCase$(Left$(rData, 6)) = "/BMSG " Then
        rData = Right$(rData, Len(rData) - 6)
        If UserList(UserIndex).flags.PertAlCons = 1 Then
            Call SendData(SendTarget.ToConsejo, UserIndex, 0, ServerPackages.dialogo & " (Consejero) " & UserList(UserIndex).Name & "> " & rData & FONTTYPE_CONSEJO)
        End If
        If UserList(UserIndex).flags.PertAlConsCaos = 1 Then
            Call SendData(SendTarget.ToConsejoCaos, UserIndex, 0, ServerPackages.dialogo & " (Consejero) " & UserList(UserIndex).Name & "> " & rData & FONTTYPE_CONSEJOCAOS)
        End If
        Exit Sub
    End If
    '[/yb]
    
    If UCase$(Left$(rData, 5)) = "/ROL " Then
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.ToIndex, 0, 0, ServerPackages.dialogo & " " & "Su solicitud ha sido enviada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToRolesMasters, 0, 0, ServerPackages.dialogo & " " & LCase$(UserList(UserIndex).Name) & " PREGUNTA ROL: " & rData & FONTTYPE_GUILDMSG)
        Exit Sub
    End If
    
    
    'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rData, 6)) = "/GMSG " And UserList(UserIndex).flags.Privilegios > PlayerType.User Then
        rData = Right$(rData, Len(rData) - 6)
        Call LogGM(UserList(UserIndex).Name, "Mensaje a Gms:" & rData, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
        If rData <> "" Then
            Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & UserList(UserIndex).Name & "> " & rData & FONTTYPE_GUILD)
        End If
        Exit Sub
    End If
    
    Select Case UCase$(Left$(rData, 7))
        Case "/TORNEO"
            If Hay_Torneo = False Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay ningún torneo disponible." & FONTTYPE_CELESTE_NEGRITA)
                Exit Sub
            End If
            
            'CHOTS | No puede ir en carcel, muerto o en bolas
            If UserList(UserIndex).Counters.Pena > 0 Or UserList(UserIndex).flags.Muerto = 1 Or UserList(UserIndex).flags.Desnudo = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes ir al torneo en esas Condiciones." & FONTTYPE_CELESTE_NEGRITA)
                Exit Sub
            End If
            
            'CHOTS | Guerras
            If UserList(UserIndex).guerra.enGuerra Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes ir al torneo si estas en una guerra." & FONTTYPE_CELESTE_NEGRITA)
                Exit Sub
            End If

            'CHOTS | Guerras
            If UserList(UserIndex).flags.enDuelo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes ir al torneo si estas en un duelo." & FONTTYPE_CELESTE_NEGRITA)
                Exit Sub
            End If
            
            If UserList(UserIndex).Stats.ELV < Torneo_Nivel_Minimo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu nivel es: " & UserList(UserIndex).Stats.ELV & ".El requerido es: " & Torneo_Nivel_Minimo & FONTTYPE_CELESTE_NEGRITA)
                Exit Sub
            End If
            If UserList(UserIndex).Stats.ELV > Torneo_Nivel_Maximo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu nivel es: " & UserList(UserIndex).Stats.ELV & ".El máximo es: " & Torneo_Nivel_Maximo & FONTTYPE_CELESTE_NEGRITA)
                Exit Sub
            End If
            If Torneo_Inscriptos >= Torneo_Cantidad Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El cupo ya ha sido alcanzado." & FONTTYPE_CELESTE_NEGRITA)
                Exit Sub
            End If
            For i = 1 To 8
                If UCase$(UserList(UserIndex).Clase) = UCase$(Torneo_Clases_Validas(i)) And Torneo_Clases_Validas2(i) = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu clase no es válida en este torneo." & FONTTYPE_CELESTE_NEGRITA)
                Exit Sub
                End If
            Next
            
            'Old, si entras no salis =P
            If Not Torneo.Existe(UserList(UserIndex).Name) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Estás en la lista de espera del torneo. Estás en el puesto nº " & Torneo.Longitud + 1 & FONTTYPE_CELESTE_NEGRITA)
                Call Torneo.Push(rData, UserList(UserIndex).Name)
                
                Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & "/TORNEO [" & UserList(UserIndex).Name & "] (" & Torneo.Longitud & ")" & FONTTYPE_CELESTE_NEGRITA)
                Torneo_Inscriptos = Torneo_Inscriptos + 1
                If Torneo_Inscriptos = Torneo_Cantidad Then
                Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Cupo alcanzado." & FONTTYPE_CELESTE_NEGRITA)
                    Torneo.Reset
                    Hay_Torneo = False
                    Torneo_Inscriptos = 0
                End If
                If Torneo_SumAuto = 1 Then
                    Dim FuturePos As WorldPos
                    FuturePos.Map = Torneo_Map
                    FuturePos.X = Torneo_X: FuturePos.Y = Torneo_Y
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                End If
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 3))
        Case "/GM"
            If Not Ayuda.Existe(UserList(UserIndex).Name) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z85")
                Call Ayuda.Push(rData, UserList(UserIndex).Name)
                Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & "SOS> " & UserList(UserIndex).Name & " ha solicitado ayuda de un GM" & FONTTYPE_SERVER)
            Else
                Call Ayuda.Quitar(UserList(UserIndex).Name)
                Call Ayuda.Push(rData, UserList(UserIndex).Name)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z86")
            End If
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rData, 6))
        Case "/DESC "
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12" & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 6)
            If Not AsciiValidos(rData) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Sub
            End If
            UserList(UserIndex).Desc = Trim$(rData)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "La descripcion ha cambiado." & FONTTYPE_INFO)
            Exit Sub
        Case "/VOTO "
                rData = Right$(rData, Len(rData) - 6)
                If Not modGuilds.v_UsuarioVota(UserIndex, rData, tStr) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Voto NO contabilizado: " & tStr & FONTTYPE_GUILD)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Voto contabilizado." & FONTTYPE_GUILD)
                End If
                Exit Sub
    End Select
    
    If UCase$(Left$(rData, 7)) = "/PENAS " Then
        Name = Right$(rData, Len(rData) - 7)
        If Name = "" Then Exit Sub
        
        Name = Replace(Name, "\", "")
        Name = Replace(Name, "/", "")
        
        If FileExist(CharPath & Name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Sin prontuario.." & FONTTYPE_INFO)
            Else
                While tInt > 0
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tInt & "- " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & tInt) & FONTTYPE_INFO)
                    tInt = tInt - 1
                Wend
            End If
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Personaje """ & Name & """ inexistente." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    
    'CHOTS | Votar para un torneo
    If UCase$(Left$(rData, 7)) = "/VOTAR " Then
        Dim votoTorneo As Integer
        votoTorneo = val(Right$(rData, Len(rData) - 7))
        
        If Torneo_Votacion_Abierta = False Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay ninguna votación abierta!" & FONTTYPE_TORNEOAUTO)
            Exit Sub
        End If

        If votoTorneo < 1 Or votoTorneo > Torneo_TIPOTORNEOS Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu voto es inválido. Tipea /VOTAR 1 (1vs1), 2 (2vs2), 3 (Deathmatch), 4 (Plantes), 5 (Al Aim) o 6 (Destruccion)." & FONTTYPE_TORNEOAUTO)
            Exit Sub
        End If

        If votarTorneo(UserIndex, votoTorneo) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu voto fue contabilizado." & FONTTYPE_TORNEOAUTO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu voto NO fue contabilizado." & FONTTYPE_TORNEOAUTO)

        End If

        Exit Sub
    End If
    
    
    Select Case UCase$(Left$(rData, 8))
        Case "/PASSWD "
            rData = Right$(rData, Len(rData) - 8)
            If Len(rData) < 6 Then
                 Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El password ha sido cambiado." & FONTTYPE_INFO)
                UserList(UserIndex).Password = ENCRYPT(UCase$(rData))
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 9))
            'Comando /APOSTAR basado en la idea de DarkLight,
            'pero con distinta probabilidad de exito.
        Case "/APOSTAR "
            rData = Right(rData, Len(rData) - 9)
            tLong = CLng(val(rData))
            If tLong > 32000 Then tLong = 32000
            n = tLong
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
            ElseIf UserList(UserIndex).flags.TargetNPC = 0 Then
                'Se asegura que el target es un npc
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
            ElseIf Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
            ElseIf Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
            ElseIf n < 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
            ElseIf n > 5000 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
            ElseIf UserList(UserIndex).Stats.GLD < n Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
            Else
                If RandomNumber(1, 100) <= 47 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + n
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Felicidades! Has ganado " & CStr(n) & " monedas de oro!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    
                    Apuestas.Perdidas = Apuestas.Perdidas + n
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - n
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Lo siento, has perdido " & CStr(n) & " monedas de oro." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                
                    Apuestas.Ganancias = Apuestas.Ganancias + n
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                End If
                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
                
                Call EnviarOro(UserIndex)
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 12))
    
        'consultas populares muchacho'
        Case "/ENCUESTA "
            rData = Right(rData, Len(rData) - 10)
            If Len(rData) = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " Aca va la info de la encuesta" & FONTTYPE_GUILD)
                Exit Sub
            End If
            DummyInt = CLng(val(rData))
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " " & ConsultaPopular.doVotar(UserIndex, DummyInt) & FONTTYPE_GUILD)
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rData, 8))

        Case "/RETIRAR" 'RETIRA ORO EN EL BANCO o te saca de la armada
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                      Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                  Exit Sub
             End If
             
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = 5 Then
                
                'Se quiere retirar de la armada
                If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                        Call ExpulsarFaccionReal(UserIndex)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "º" & "¡¡¡Sal de aquí bufón!!!" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    End If
                ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 1 Then
                        Call ExpulsarFaccionCaos(UserIndex)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "º" & "Ya volverás arrastrandote." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "º" & "Sal de aquí maldito criminal" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "º" & "¡No perteneces a ninguna fuerza!" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex))
                End If
                Exit Sub
             
             End If
             
             If Len(rData) = 8 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Debes indicar el monto de cuanto quieres retirar" & FONTTYPE_INFO)
                Exit Sub
             End If
             
             rData = Right$(rData, Len(rData) - 9)
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
             Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
             If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                  Exit Sub
             End If
             If FileExist(CharPath & UCase$(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                  CloseSocket (UserIndex)
                  Exit Sub
             End If
             If val(rData) > 0 And val(rData) <= UserList(UserIndex).Stats.Banco Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(rData)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rData)
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
             Else
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
             End If
             Call EnviarOro(val(UserIndex)) 'Niconz
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z12")
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z30")
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                      Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 11)
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
            Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
            If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z27")
                  Exit Sub
            End If
            If CLng(val(rData)) > 0 And CLng(val(rData)) <= UserList(UserIndex).Stats.GLD Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rData)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rData)
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
            Else
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).char.CharIndex & FONTTYPE_INFO)
            End If
            Call EnviarOro(val(UserIndex))
            Exit Sub
        Case "/DENUNCIAR "
            rData = Right$(rData, Len(rData) - 11)

            If UserList(UserIndex).flags.YaDenuncio = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z87")
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Silenciado = 1 Then
                Exit Sub
            End If

            'CHOTS | Checkeamos que no tenga palabras invalidas
            Dim position As Integer
            position = 0
            LoopC = 1

            Do While LoopC <= PalabrasInvalidas.Count And (position = 0)
                position = InStr(LCase$(rData), PalabrasInvalidas.item(LoopC))
                LoopC = LoopC + 1
            Loop

            If position > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu denuncia no fue enviada, por favor modera tu lenguaje e intenta nuevamente en 1 minuto." & FONTTYPE_GUILDMSG)
                UserList(UserIndex).flags.YaDenuncio = 1
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "ABDENU")
                Exit Sub
            End If

            Call SendData(SendTarget.ToAdmins, 0, 0, ServerPackages.dialogo & " " & LCase$(UserList(UserIndex).Name) & " DENUNCIA: " & rData & FONTTYPE_GUILDMSG)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z77")
            UserList(UserIndex).flags.YaDenuncio = 1
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ABDENU")
        Exit Sub
        
        'CHOTS | Torneos Automáticos
        Case "/FIXTURE"
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "FIX" & Torneo_Fixture)
        Exit Sub
        
        
        Case "/PARTICIPAR"
        
            If Torneo_HAYTORNEO = False Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No hay ningún torneo disponible." & FONTTYPE_TORNEOAUTO)
                Exit Sub
            End If
            
            'CHOTS | No puede ir en carcel, muerto o en bolas
            If UserList(UserIndex).Counters.Pena > 0 Or UserList(UserIndex).flags.Muerto = 1 Or UserList(UserIndex).flags.Desnudo = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes ir al torneo en esas Condiciones." & FONTTYPE_TORNEOAUTO)
                Exit Sub
            End If
            
            'CHOTS | Guerras
            If UserList(UserIndex).guerra.enGuerra Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes ir al torneo si estas en una guerra." & FONTTYPE_TORNEOAUTO)
                Exit Sub
            End If

            'CHOTS | Duelos
            If UserList(UserIndex).flags.enDuelo = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes ir al torneo si estas en un duelo." & FONTTYPE_TORNEOAUTO)
                Exit Sub
            End If

            If UserList(UserIndex).Pos.Map = SALATORNEO Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No puedes ir al torneo ahora." & FONTTYPE_TORNEOAUTO)
                Exit Sub
            End If

            If UserList(UserIndex).Stats.ELV < 46 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu nivel es: " & UserList(UserIndex).Stats.ELV & ". El requerido es: 46" & FONTTYPE_TORNEOAUTO)
                Exit Sub
            End If

            If Torneo_CantidadInscriptos >= Torneo_Cupo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "El cupo ya ha sido alcanzado." & FONTTYPE_TORNEOAUTO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.enTorneoAuto = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Ya estás inscripto!" & FONTTYPE_TORNEOAUTO)
                Exit Sub
            End If

            If Torneo_Tipo = eTipoTorneo.Aim And UserList(UserIndex).Stats.MaxMAN = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Tu clase no puede participar de este torneo" & FONTTYPE_TORNEOAUTO)
                Exit Sub
            End If
            
            If inscribirseTorneo(UserIndex) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z93")
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z94")
            End If
        Exit Sub
        'CHOTS | Torneos Automáticos
            
        
        'CHOTS | Sistema de cierre de Clanes
        Case "/CERRARCLAN"
        
            If Not UserList(UserIndex).GuildIndex >= 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No perteneces a ningún clan." & FONTTYPE_GUILD)
                Exit Sub
            End If
            
            If UCase$(Guilds(UserList(UserIndex).GuildIndex).GetLeader) <> UCase$(UserList(UserIndex).Name) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No eres líder del clan." & FONTTYPE_GUILD)
                Exit Sub
            End If
            
            If Guilds(UserList(UserIndex).GuildIndex).CantidadDeMiembros > 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Debes echar a todos los miembros del clan para cerrarlo." & FONTTYPE_GUILD)
                Exit Sub
            End If
            
            'CHOTS | Guerras
            If UserList(UserIndex).guerra.enGuerra = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "No cerrar tu clan en una guerra! Si deseas abandonar a tu equipo, tipea /SALIRGUERRA" & FONTTYPE_GUERRA)
                Exit Sub
            End If
            
            Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " cerró." & FONTTYPE_GUILD)
            Call SendData(SendTarget.ToAll, 0, 0, "TW" & SONIDOS_GUILD.SND_DECLAREWAR)
            
            Call Guilds(UserList(UserIndex).GuildIndex).DesConectarMiembro(UserIndex)
            Call Guilds(UserList(UserIndex).GuildIndex).ExpulsarMiembro(UserList(UserIndex).Name)
            UserList(UserIndex).GuildIndex = 0
            
            Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
        Exit Sub
        'CHOTS | Sistema de cierre de Clanes
            
        Case "/FUNDARCLAN"
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " Para fundar un clan debes especificar la alineación del mismo." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " Atención, que la misma no podrá cambiar luego, te aconsejamos leer las reglas sobre clanes antes de fundar." & FONTTYPE_GUILD)
                Exit Sub
            Else
                Select Case UCase$(Trim(rData))
                    Case "ARMADA"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case "MAL"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case "NEUTRO"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case "GM"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case "LEGAL"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case "CRIMINAL"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " Alineación inválida." & FONTTYPE_GUILD)
                        Exit Sub
                End Select
            End If

            If modGuilds.PuedeFundarUnClan(UserIndex, UserList(UserIndex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SHOWFUN")
            Else
                UserList(UserIndex).FundandoGuildAlineacion = 0
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & FONTTYPE_GUILD)
            End If
            
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 12))
    
        Case "/ECHARPARTY "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
            End If
            Exit Sub
            
        Case "/PARTYLIDER "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.TransformarEnLider(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
            End If
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 13))
    
        Case "/ACCEPTPARTY "
            rData = Right$(rData, Len(rData) - 13)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.AprobarIngresoAParty(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "Z47")
            End If
            Exit Sub
    
    End Select
    

    Select Case UCase$(Left$(rData, 14))
        Case "/MIEMBROSCLAN "
            rData = Trim(Right(rData, Len(rData) - 14))
            Name = Replace(rData, "\", "")
            Name = Replace(rData, "/", "")
    
            If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & " No existe el clan: " & rData & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
            
            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & tStr & "<" & rData & ">." & FONTTYPE_INFO)
            Next i
        
            Exit Sub
    End Select
    
    Procesado = False
End Sub
