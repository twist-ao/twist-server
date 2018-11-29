Attribute VB_Name = "Torneos_Auto"
'Módulo de Torneos Automáticos
'Creado por Juan Andrés Dalmasso (CHOTS)
'CHOTS_AO@HOTMAIL.COM
'Para Lapsus 3
'05/10/2011
'Modificado y adaptado por CHOTS
'j_dalmasso@outlook.com
'Para TwistAO 2018
'Agregado torneo de parejas

Public Torneo_Activado As Boolean
Public Torneo_Cupo As Byte
Public Torneo_CantidadInscriptos As Byte
Public Torneo_HAYTORNEO As Boolean
Public Torneo_Fixture As String

Public Enum eTipoTorneo
    t1vs1 = 1
    t2vs2 = 2
    Deathmatch = 3
    Plantes = 4
    Aim = 5
    Destruccion = 6
End Enum

Public Torneo_Tipo As eTipoTorneo

Public Torneo_Votacion_Abierta As Boolean
Public Torneo_Votos_1 As Integer
Public Torneo_Votos_2 As Integer
Public Torneo_Votos_3 As Integer
Public Torneo_Votos_4 As Integer
Public Torneo_Votos_5 As Integer
Public Torneo_Votos_6 As Integer
Public Torneo_Votantes() As String

Public Const Torneo_MAPATORNEO As Byte = 66
Public Const Torneo_MAPAMUERTE As Byte = 65
Public Const Torneo_TIPOTORNEOS As Byte = 6
Public Const Torneo_EQUIPOSDESTRUCCION As Byte = 4
Public Const Torneo_NPCDESTRUCCION As Integer = 616

Public Type tAreasTorneo
    mapa As Byte
    MinX As Byte
    MaxX As Byte
    MinY As Byte
    MaxY As Byte
End Type

Public Type tCuenta
    segundos As Byte
    razon As String
    next As eRonda
End Type

Public Type tDuelo
    usuario1 As Integer
    usuario2 As Integer
    ganador As Integer
End Type

Public Type tEquipoDestruccion
    Numero As Integer
    NpcIndex As Integer
    usuarios() As Integer
End Type

Public Enum eRonda
    Ronda_Dieciseisavos = 1
    Ronda_Octavos = 2
    Ronda_Cuartos = 3
    Ronda_Semi = 4
    Ronda_Final = 5
    Ronda_Deathmatch = 6
    Ronda_Destruccion = 7
End Enum

Public Torneo_AreaDescanso As tAreasTorneo
Public Torneo_AreasDuelo() As tAreasTorneo
Public Torneo_AreasPlante() As tAreasTorneo
Public Torneo_AreaDeathmatch As tAreasTorneo

Public Torneo_CR As tCuenta
Public Torneo_CuentaPelea As Byte

Public Torneo_RondaActual As eRonda

Public Torneo_UsuariosInscriptos() As Integer
Public Torneo_Final As tDuelo
Public Torneo_Semifinal() As tDuelo
Public Torneo_Cuartos() As tDuelo
Public Torneo_Octavos() As tDuelo
Public Torneo_Dieciseisavos() As tDuelo
Public Torneo_DestruccionEquipos() As tEquipoDestruccion

Public Sub activarTorneos()
    Torneo_Activado = True
    Call inicializarTorneo
    Call SendData(SendTarget.ToAll, 0, 0, "Z92")
End Sub

Public Sub desactivarTorneos()
    Torneo_Activado = False
    Call inicializarTorneo
    Call SendData(SendTarget.ToAll, 0, 0, "Z91")
End Sub

Public Sub crearTorneo()
On Error GoTo chotserror
    Dim i As Byte
    If Not Torneo_Activado Then Exit Sub

    Torneo_Cupo = calcularCantidad()
    Torneo_CuentaPelea = 0
    
    If Torneo_Cupo = 0 Then
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto> El torneo no se abrirá dado que la cantidad de usuarios no es suficiente!" & FONTTYPE_TORNEOAUTO)
        Call LogGM("TORNEOAUTO", "El torneo no se abrio por falta de usuarios.", False)
        Exit Sub
    End If
    
    ReDim Torneo_UsuariosInscriptos(1 To Torneo_Cupo) As Integer
    
    For i = 1 To Torneo_Cupo
        Torneo_UsuariosInscriptos(i) = 0
    Next i
    
    Call inicializarAreas
    
    Torneo_CantidadInscriptos = 0
    
    Torneo_HAYTORNEO = True
    
    Call SendData(SendTarget.ToAll, 0, 0, "TAU" & Torneo_Cupo & "," & getTipoTorneoString())

    Call LogGM("TORNEOAUTO", "Se abrio un torneo auto de tipo " & Torneo_Tipo & " para " & Torneo_Cupo & " participantes.", False)
    Exit Sub
chotserror:
    Call LogError("Error en crearTorneo " & Err.number & " " & Err.Description)
End Sub

Private Sub inicializarAreas()
    On Error GoTo chotserror
    ReDim Torneo_AreasDuelo(1 To 4) As tAreasTorneo
    ReDim Torneo_AreasPlante(1 To 4) As tAreasTorneo
    
    'CHOTS | Salas 1vs1
    Torneo_AreasDuelo(1).mapa = Torneo_MAPATORNEO
    Torneo_AreasDuelo(1).MinX = 18
    Torneo_AreasDuelo(1).MaxX = 33
    Torneo_AreasDuelo(1).MinY = 14
    Torneo_AreasDuelo(1).MaxY = 29
    
    Torneo_AreasDuelo(2).mapa = Torneo_MAPATORNEO
    Torneo_AreasDuelo(2).MinX = 69
    Torneo_AreasDuelo(2).MaxX = 84
    Torneo_AreasDuelo(2).MinY = 14
    Torneo_AreasDuelo(2).MaxY = 29
    
    Torneo_AreasDuelo(3).mapa = Torneo_MAPATORNEO
    Torneo_AreasDuelo(3).MinX = 17
    Torneo_AreasDuelo(3).MaxX = 32
    Torneo_AreasDuelo(3).MinY = 72
    Torneo_AreasDuelo(3).MaxY = 87
    
    Torneo_AreasDuelo(4).mapa = Torneo_MAPATORNEO
    Torneo_AreasDuelo(4).MinX = 70
    Torneo_AreasDuelo(4).MaxX = 85
    Torneo_AreasDuelo(4).MinY = 71
    Torneo_AreasDuelo(4).MaxY = 86

    'CHOTS | Salas plante
    Torneo_AreasPlante(1).mapa = Torneo_MAPATORNEO
    Torneo_AreasPlante(1).MinX = 24
    Torneo_AreasPlante(1).MaxX = 25
    Torneo_AreasPlante(1).MinY = 45
    Torneo_AreasPlante(1).MaxY = 45
    
    Torneo_AreasPlante(2).mapa = Torneo_MAPATORNEO
    Torneo_AreasPlante(2).MinX = 49
    Torneo_AreasPlante(2).MaxX = 50
    Torneo_AreasPlante(2).MinY = 30
    Torneo_AreasPlante(2).MaxY = 30
    
    Torneo_AreasPlante(3).mapa = Torneo_MAPATORNEO
    Torneo_AreasPlante(3).MinX = 79
    Torneo_AreasPlante(3).MaxX = 80
    Torneo_AreasPlante(3).MinY = 58
    Torneo_AreasPlante(3).MaxY = 58
    
    Torneo_AreasPlante(4).mapa = Torneo_MAPATORNEO
    Torneo_AreasPlante(4).MinX = 49
    Torneo_AreasPlante(4).MaxX = 50
    Torneo_AreasPlante(4).MinY = 88
    Torneo_AreasPlante(4).MaxY = 88
    
    'CHOTS | Sala de deathmatch
    Torneo_AreaDeathmatch.mapa = Torneo_MAPATORNEO
    Torneo_AreaDeathmatch.MinX = 41
    Torneo_AreaDeathmatch.MaxX = 63
    Torneo_AreaDeathmatch.MinY = 43
    Torneo_AreaDeathmatch.MaxY = 60
    
    'CHOTS | Area de descanso
    Torneo_AreaDescanso.mapa = Torneo_MAPAMUERTE
    Torneo_AreaDescanso.MinX = 35
    Torneo_AreaDescanso.MaxX = 46
    Torneo_AreaDescanso.MinY = 71
    Torneo_AreaDescanso.MaxY = 79
    
    Exit Sub
chotserror:
    Call LogError("Error en inicializarAreas " & Err.number & " " & Err.Description)

End Sub

Private Function calcularCantidad() As Byte
    On Error GoTo chotserror
    calcularCantidad = 0

    If Torneo_Tipo = eTipoTorneo.Deathmatch Then
        calcularCantidad = 32
        Exit Function
    End If
    
    If NumUsers < 50 Then
        calcularCantidad = 8
    ElseIf NumUsers < 100 Then
        calcularCantidad = 16
    Else
        calcularCantidad = 32
    End If
    Exit Function
chotserror:
    Call LogError("Error en calcularCantidad " & Err.number & " " & Err.Description)
End Function

Public Function inscribirseTorneo(ByVal UserIndex As Integer) As Boolean
    Dim i As Integer
    
    If Torneo_CantidadInscriptos >= Torneo_Cupo Then
        inscribirseTorneo = False
        Exit Function
    End If
    
    For i = 1 To Torneo_Cupo
        If Torneo_UsuariosInscriptos(i) = 0 Then
            Torneo_UsuariosInscriptos(i) = UserIndex
            UserList(UserIndex).flags.enTorneoAuto = True
            Torneo_CantidadInscriptos = Torneo_CantidadInscriptos + 1
            inscribirseTorneo = True
            Call telepToAreaDescanso(UserIndex)
            Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " se inscribio a un torneo Auto.", False)
            If Torneo_CantidadInscriptos = Torneo_Cupo Then
                Call armarFixture
                Call SendData(SendTarget.ToAll, 0, 0, "Z94")
            End If
            Exit Function
        End If
    Next i
End Function

Public Sub telepToAreaDescanso(ByVal UserIndex As Integer)
    On Error GoTo chotserror
    Dim dPos2 As WorldPos
    Dim nPos2 As WorldPos
    dPos2.Map = Torneo_AreaDescanso.mapa
    dPos2.X = RandomNumber(Torneo_AreaDescanso.MinX, Torneo_AreaDescanso.MaxX)
    dPos2.Y = RandomNumber(Torneo_AreaDescanso.MinY, Torneo_AreaDescanso.MaxY)
    Call ClosestLegalPos(dPos2, nPos2)
    Call WarpUserChar(UserIndex, nPos2.Map, nPos2.X, nPos2.Y, True)
    UserList(UserIndex).flags.enDueloTorneoAuto = False

    'CHOTS | Si es 2vs2 llevamos a su pareja tambien
    If Torneo_Tipo = eTipoTorneo.t2vs2 Then
        If UserList(UserIndex).torneoPareja > 0 Then
            If UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                Call ClosestLegalPos(dPos2, nPos2)
                Call WarpUserChar(UserList(UserIndex).torneoPareja, nPos2.Map, nPos2.X, nPos2.Y, True)
                UserList(UserList(UserIndex).torneoPareja).flags.enDueloTorneoAuto = False
            End If
            
            'CHOTS | Si estaba muerto lo revivimos
            If UserList(UserIndex).flags.Muerto = 1 Then Call Resucitar(UserIndex)
            If UserList(UserList(UserIndex).torneoPareja).flags.Muerto = 1 Then Call Resucitar(UserList(UserIndex).torneoPareja)
        End if
    End If
    Exit Sub
    
chotserror:
    Call LogError("Error en TelepAreaDescanso " & Err.number & " " & Err.Description)
End Sub

Public Sub telepToAreaDuelo(ByVal UserIndex As Integer, ByVal area As Byte)
    On Error GoTo chotserror
    Dim dPos As WorldPos
    Dim nPos As WorldPos
    If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.t2vs2 Or Torneo_Tipo = eTipoTorneo.Aim Or Torneo_Tipo = eTipoTorneo.Destruccion Then
        dPos.Map = Torneo_AreasDuelo(area).mapa
        dPos.X = RandomNumber(Torneo_AreasDuelo(area).MinX, Torneo_AreasDuelo(area).MaxX)
        dPos.Y = RandomNumber(Torneo_AreasDuelo(area).MinY, Torneo_AreasDuelo(area).MaxY)
    ElseIf Torneo_Tipo = eTipoTorneo.Deathmatch Then
        dPos.Map = Torneo_AreaDeathmatch.mapa
        dPos.X = RandomNumber(Torneo_AreaDeathmatch.MinX, Torneo_AreaDeathmatch.MaxX)
        dPos.Y = RandomNumber(Torneo_AreaDeathmatch.MinY, Torneo_AreaDeathmatch.MaxY)
        UserList(UserIndex).showName = False
    ElseIf Torneo_Tipo = eTipoTorneo.Plantes Then
        dPos.Map = Torneo_AreasPlante(area).mapa
        dPos.X = RandomNumber(Torneo_AreasPlante(area).MinX, Torneo_AreasPlante(area).MaxX)
        dPos.Y = RandomNumber(Torneo_AreasPlante(area).MinY, Torneo_AreasPlante(area).MaxY)
    End If
    Call ClosestLegalPos(dPos, nPos)
    UserList(UserIndex).flags.enDueloTorneoAuto = True
    Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)

    'CHOTS | Si es 2vs2 llevamos a su pareja tambien
    If Torneo_Tipo = eTipoTorneo.t2vs2 Then
        If UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
            Call ClosestLegalPos(dPos, nPos)
            Call WarpUserChar(UserList(UserIndex).torneoPareja, nPos.Map, nPos.X, nPos.Y, True)
            UserList(UserList(UserIndex).torneoPareja).flags.enDueloTorneoAuto = True
        End If
    End If

    Exit Sub
    
chotserror:
    Call LogError("Error en TelepAreaDuelo " & Err.number & " " & Err.Description)
End Sub

Private Sub armarFixture()
On Error GoTo chotserror

    'CHOTS | Si es torneo de parejas, armamos las parejas antes
    If Torneo_Tipo = eTipoTorneo.t2vs2 Then
        Call armarParejasTorneo
    End If

    'CHOTS | Si es torneo de destruccion, armamos los equipos
    If Torneo_Tipo = eTipoTorneo.Destruccion Then
        Call armarEquiposDestruccion
    End If

    Dim i As Integer

    If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.t2vs2 Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Then
        If Torneo_CantidadInscriptos = 32 Then
            ReDim Torneo_Dieciseisavos(1 To 16) As tDuelo
            For i = 1 To 16
                Torneo_Dieciseisavos(i).usuario1 = 0
                Torneo_Dieciseisavos(i).usuario2 = 0
                Torneo_Dieciseisavos(i).ganador = 0
            Next i
        End If
        
        If Torneo_CantidadInscriptos >= 16 Then
            ReDim Torneo_Octavos(1 To 8) As tDuelo
            For i = 1 To 8
                Torneo_Octavos(i).usuario1 = 0
                Torneo_Octavos(i).usuario2 = 0
                Torneo_Octavos(i).ganador = 0
            Next i
        End If
        
        ReDim Torneo_Cuartos(1 To 4) As tDuelo
        For i = 1 To 4
            Torneo_Cuartos(i).usuario1 = 0
            Torneo_Cuartos(i).usuario2 = 0
            Torneo_Cuartos(i).ganador = 0
        Next i
            
        ReDim Torneo_Semifinal(1 To 2) As tDuelo
        For i = 1 To 2
            Torneo_Semifinal(i).usuario1 = 0
            Torneo_Semifinal(i).usuario2 = 0
            Torneo_Semifinal(i).ganador = 0
        Next i

        Torneo_Final.ganador = 0
        Torneo_Final.usuario1 = 0
        Torneo_Final.usuario2 = 0

        Dim contador As Byte
        contador = 1

        Torneo_Fixture = Torneo_CantidadInscriptos & "@"

        For i = 1 To Torneo_CantidadInscriptos

            Select Case Torneo_CantidadInscriptos

                Case 32:
                    Torneo_Dieciseisavos(contador).usuario1 = Torneo_UsuariosInscriptos(i)
                    Torneo_Dieciseisavos(contador).usuario2 = Torneo_UsuariosInscriptos(i + 1)
                    Torneo_Fixture = Torneo_Fixture & Torneo_Dieciseisavos(contador).usuario1 & "," & Torneo_Dieciseisavos(contador).usuario2 & ","
                    contador = contador + 1

                Case 16:
                    Torneo_Octavos(contador).usuario1 = Torneo_UsuariosInscriptos(i)
                    Torneo_Octavos(contador).usuario2 = Torneo_UsuariosInscriptos(i + 1)
                    Torneo_Fixture = Torneo_Fixture & Torneo_Octavos(contador).usuario1 & "," & Torneo_Octavos(contador).usuario2 & ","
                    contador = contador + 1

                Case 8:
                    Torneo_Cuartos(contador).usuario1 = Torneo_UsuariosInscriptos(i)
                    Torneo_Cuartos(contador).usuario2 = Torneo_UsuariosInscriptos(i + 1)
                    Torneo_Fixture = Torneo_Fixture & Torneo_Cuartos(contador).usuario1 & "," & Torneo_Cuartos(contador).usuario2 & ","
                    contador = contador + 1

                Case 4:
                    Torneo_Semifinal(contador).usuario1 = Torneo_UsuariosInscriptos(i)
                    Torneo_Semifinal(contador).usuario2 = Torneo_UsuariosInscriptos(i + 1)
                    Torneo_Fixture = Torneo_Fixture & Torneo_Semifinal(contador).usuario1 & "," & Torneo_Semifinal(contador).usuario2 & ","
                    contador = contador + 1
                    
                Case 2:
                    Torneo_Final.usuario1 = Torneo_UsuariosInscriptos(i)
                    Torneo_Final.usuario2 = Torneo_UsuariosInscriptos(i + 1)
                    Torneo_Fixture = Torneo_Fixture & Torneo_Final.usuario1 & "," & Torneo_Final.usuario2 & ","
                    contador = contador + 1
                    
            End Select
            
            i = i + 1
            
        Next i
    ElseIf Torneo_Tipo = eTipoTorneo.Deathmatch Or Torneo_Tipo = eTipoTorneo.Destruccion Then
        For i = 1 To Torneo_CantidadInscriptos
            Torneo_Fixture = Torneo_Fixture & Torneo_UsuariosInscriptos(i) & ","
        Next i
    End If
    
    Call comenzarTorneo
    Exit Sub
chotserror:
    Call LogError("Error en armarFixture " & Err.number & " " & Err.Description)
End Sub

Public Sub comenzarTorneo()

    MapInfo(Torneo_MAPATORNEO).Pk = False
    If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.t2vs2 Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Then
        Select Case Torneo_CantidadInscriptos
            Case 32:
                Call setearCuenta(10, eRonda.Ronda_Dieciseisavos)
                Exit Sub
                
            Case 16:
                Call setearCuenta(10, eRonda.Ronda_Octavos)
                Exit Sub
                
            Case 8:
                Call setearCuenta(10, eRonda.Ronda_Cuartos)
                Exit Sub
                
            Case 4:
                Call setearCuenta(10, eRonda.Ronda_Semi)
                Exit Sub
            
            Case 2:
                Call setearCuenta(10, eRonda.Ronda_Final)
                Exit Sub
        End Select
    ElseIf Torneo_Tipo = eTipoTorneo.Deathmatch Then
        Call setearCuenta(10, eRonda.Ronda_Deathmatch)
        Exit Sub
    ElseIf Torneo_Tipo = eTipoTorneo.Destruccion Then
        Call setearCuenta(10, eRonda.Ronda_Destruccion)
        Exit Sub
    End If
    
End Sub

Public Sub comenzarDieciseisavos()
    Torneo_RondaActual = eRonda.Ronda_Dieciseisavos
    Dim i As Byte
    Dim j As Byte
    j = 0

    MapInfo(Torneo_MAPATORNEO).Pk = False
    
    For i = 1 To 16
        If Torneo_Dieciseisavos(i).ganador = 0 Then
        
            j = j + 1
            
            If Torneo_Dieciseisavos(i).usuario1 = 0 And Torneo_Dieciseisavos(i).usuario2 > 0 Then
                Call ganaUsuario(Torneo_Dieciseisavos(i).usuario2)
            ElseIf Torneo_Dieciseisavos(i).usuario2 = 0 And Torneo_Dieciseisavos(i).usuario1 > 0 Then
                Call ganaUsuario(Torneo_Dieciseisavos(i).usuario1)
            ElseIf Torneo_Dieciseisavos(i).usuario1 > 0 And Torneo_Dieciseisavos(i).usuario2 > 0 Then
                If UserList(Torneo_Dieciseisavos(i).usuario1).flags.enTorneoAuto = False Then
                    Call ganaUsuario(Torneo_Dieciseisavos(i).usuario2)
                ElseIf UserList(Torneo_Dieciseisavos(i).usuario2).flags.enTorneoAuto = False Then
                    Call ganaUsuario(Torneo_Dieciseisavos(i).usuario1)
                Else
                    Call telepToAreaDuelo(Torneo_Dieciseisavos(i).usuario1, j)
                    UserList(Torneo_Dieciseisavos(i).usuario1).Counters.Torneo = 5
                    Call telepToAreaDuelo(Torneo_Dieciseisavos(i).usuario2, j)
                End If
            End If
        End If
        
        If j = 4 Then Exit For
    Next i
    
    Torneo_CuentaPelea = 4
    
End Sub

Public Sub comenzarOctavos()
    Torneo_RondaActual = eRonda.Ronda_Octavos
    Dim i As Byte
    Dim j As Byte
    j = 0

    MapInfo(Torneo_MAPATORNEO).Pk = False
    
    For i = 1 To 8
        If Torneo_Octavos(i).ganador = 0 Then
        
            j = j + 1
            
            If Torneo_Octavos(i).usuario1 = 0 And Torneo_Octavos(i).usuario2 > 0 Then
                Call ganaUsuario(Torneo_Octavos(i).usuario2)
            ElseIf Torneo_Octavos(i).usuario2 = 0 And Torneo_Octavos(i).usuario1 > 0 Then
                Call ganaUsuario(Torneo_Octavos(i).usuario1)
            ElseIf Torneo_Octavos(i).usuario1 > 0 And Torneo_Octavos(i).usuario2 > 0 Then
                If UserList(Torneo_Octavos(i).usuario1).flags.enTorneoAuto = False Then
                    Call ganaUsuario(Torneo_Octavos(i).usuario2)
                ElseIf UserList(Torneo_Octavos(i).usuario1).flags.enTorneoAuto = False Then
                    Call ganaUsuario(Torneo_Octavos(i).usuario1)
                Else
                    Call telepToAreaDuelo(Torneo_Octavos(i).usuario1, j)
                    UserList(Torneo_Octavos(i).usuario1).Counters.Torneo = 6
                    Call telepToAreaDuelo(Torneo_Octavos(i).usuario2, j)
                End If
            End If
        End If

        If j = 4 Then Exit For
    Next i
    
    Torneo_CuentaPelea = 4
    
End Sub

Public Sub comenzarCuartos()
    Torneo_RondaActual = eRonda.Ronda_Cuartos
    Dim i As Byte
    
    MapInfo(Torneo_MAPATORNEO).Pk = False
    
    For i = 1 To 4
        If Torneo_Cuartos(i).usuario1 = 0 And Torneo_Cuartos(i).usuario2 > 0 Then
            Call ganaUsuario(Torneo_Cuartos(i).usuario2)
        ElseIf Torneo_Cuartos(i).usuario2 = 0 And Torneo_Cuartos(i).usuario1 > 0 Then
            Call ganaUsuario(Torneo_Cuartos(i).usuario1)
        ElseIf Torneo_Cuartos(i).usuario1 > 0 And Torneo_Cuartos(i).usuario2 > 0 Then
            If UserList(Torneo_Cuartos(i).usuario1).flags.enTorneoAuto = False Then
                Call ganaUsuario(Torneo_Cuartos(i).usuario2)
            ElseIf UserList(Torneo_Cuartos(i).usuario2).flags.enTorneoAuto = False Then
                Call ganaUsuario(Torneo_Cuartos(i).usuario1)
            Else
                Call telepToAreaDuelo(Torneo_Cuartos(i).usuario1, i)
                UserList(Torneo_Cuartos(i).usuario1).Counters.Torneo = 7
                Call telepToAreaDuelo(Torneo_Cuartos(i).usuario2, i)
            End If
        End If
    Next i
    
    Torneo_CuentaPelea = 4
    
End Sub

Public Sub comenzarSemifinal()
    Torneo_RondaActual = eRonda.Ronda_Semi
    Dim i As Byte
    
    MapInfo(Torneo_MAPATORNEO).Pk = False
    
    For i = 1 To 2
        If Torneo_Semifinal(i).usuario1 = 0 And Torneo_Semifinal(i).usuario2 > 0 Then
            Call ganaUsuario(Torneo_Semifinal(i).usuario2)
        ElseIf Torneo_Semifinal(i).usuario2 = 0 And Torneo_Semifinal(i).usuario1 > 0 Then
            Call ganaUsuario(Torneo_Semifinal(i).usuario1)
        ElseIf Torneo_Semifinal(i).usuario1 > 0 And Torneo_Semifinal(i).usuario2 > 0 Then
            If UserList(Torneo_Semifinal(i).usuario1).flags.enTorneoAuto = False Then
                 Call ganaUsuario(Torneo_Semifinal(i).usuario2)
            ElseIf UserList(Torneo_Semifinal(i).usuario2).flags.enTorneoAuto = False Then
                 Call ganaUsuario(Torneo_Semifinal(i).usuario1)
            Else
                Call telepToAreaDuelo(Torneo_Semifinal(i).usuario1, i)
                UserList(Torneo_Semifinal(i).usuario1).Counters.Torneo = 8
                Call telepToAreaDuelo(Torneo_Semifinal(i).usuario2, i)
            End If
        End If
    Next i
    
    Torneo_CuentaPelea = 4
    
End Sub

Public Sub comenzarFinal()
    Torneo_RondaActual = eRonda.Ronda_Final
    
    MapInfo(Torneo_MAPATORNEO).Pk = False
        
    If Torneo_Final.usuario1 = 0 And Torneo_Final.usuario2 > 0 Then
        Call ganaUsuario(Torneo_Final.usuario2)
    ElseIf Torneo_Final.usuario2 = 0 And Torneo_Final.usuario1 > 0 Then
        Call ganaUsuario(Torneo_Final.usuario1)
    ElseIf Torneo_Final.usuario1 > 0 And Torneo_Final.usuario2 > 0 Then
        If UserList(Torneo_Final.usuario1).flags.enTorneoAuto = False Then
            Call ganaUsuario(Torneo_Final.usuario2)
        ElseIf UserList(Torneo_Final.usuario2).flags.enTorneoAuto = False Then
            Call ganaUsuario(Torneo_Final.usuario1)
        Else
            Call telepToAreaDuelo(Torneo_Final.usuario1, 1)
            Call telepToAreaDuelo(Torneo_Final.usuario2, 1)
        End If
    Else
        Call cerrarTorneo
    End If
    
    Torneo_CuentaPelea = 4
    
End Sub

Public Sub comenzarDeathmatch()
    Torneo_RondaActual = eRonda.Ronda_Deathmatch
    Dim i As Byte
    
    MapInfo(Torneo_MAPATORNEO).Pk = False
        
    For i = 1 To Torneo_CantidadInscriptos
       Call telepToAreaDuelo(Torneo_UsuariosInscriptos(i), 1)
    Next i
    
    Torneo_CuentaPelea = 4
    
End Sub

Public Sub comenzarDestruccion()
    Torneo_RondaActual = eRonda.Ronda_Destruccion
    Dim i As Byte
    Dim j As Byte

    MapInfo(Torneo_MAPATORNEO).Pk = False

    For i = 1 To Torneo_EQUIPOSDESTRUCCION
        For j = 1 To UBound(Torneo_DestruccionEquipos(i).usuarios)
            If Torneo_DestruccionEquipos(i).usuarios(j) > 0 And UserList(Torneo_DestruccionEquipos(i).usuarios(j)).flags.enTorneoAuto = True Then
                Call telepToAreaDuelo(Torneo_DestruccionEquipos(i).usuarios(j), Torneo_DestruccionEquipos(i).Numero)
            End If
        Next j
    Next i

    Call SendData(SendTarget.ToMap, 0, Torneo_MAPATORNEO, ServerPackages.dialogo & "TorneosAuto> El equipo que primero destruya la torre sera el vencedor!" & FONTTYPE_TORNEOAUTO)

End Sub

Public Sub ganaUsuario(ByVal UserIndex As Integer)
    Dim i As Byte
    
    UserList(UserIndex).Counters.Torneo = 0
    
    Select Case Torneo_RondaActual
    
        Case eRonda.Ronda_Dieciseisavos:
        
            For i = 1 To 16
            
                If Torneo_Dieciseisavos(i).usuario1 = UserIndex Then
                    Torneo_Dieciseisavos(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Octavos)
                    Call vuelveUlla(Torneo_Dieciseisavos(i).usuario2)
                    Exit Sub
                End If
                
                If Torneo_Dieciseisavos(i).usuario2 = UserIndex Then
                    Torneo_Dieciseisavos(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Octavos)
                    Call vuelveUlla(Torneo_Dieciseisavos(i).usuario1)
                    Exit Sub
                End If
                
            Next i
    
    
        Case eRonda.Ronda_Octavos:
        
            For i = 1 To 8
            
                If Torneo_Octavos(i).usuario1 = UserIndex Then
                    Torneo_Octavos(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Cuartos)
                    Call vuelveUlla(Torneo_Octavos(i).usuario2)
                    Exit Sub
                End If
                
                If Torneo_Octavos(i).usuario2 = UserIndex Then
                    Torneo_Octavos(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Cuartos)
                    Call vuelveUlla(Torneo_Octavos(i).usuario1)
                    Exit Sub
                End If
                
            Next i
    
        Case eRonda.Ronda_Cuartos:
        
            For i = 1 To 4
            
                If Torneo_Cuartos(i).usuario1 = UserIndex Then
                    Torneo_Cuartos(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Semi)
                    Call vuelveUlla(Torneo_Cuartos(i).usuario2)
                    Exit Sub
                End If
                
                If Torneo_Cuartos(i).usuario2 = UserIndex Then
                    Torneo_Cuartos(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Semi)
                    Call vuelveUlla(Torneo_Cuartos(i).usuario1)
                    Exit Sub
                End If
                
            Next i
            
            
        Case eRonda.Ronda_Semi
        
            For i = 1 To 2
            
                If Torneo_Semifinal(i).usuario1 = UserIndex Then
                    Torneo_Semifinal(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Final)
                    Call vuelveUlla(Torneo_Semifinal(i).usuario2)
                    Exit Sub
                End If
                
                If Torneo_Semifinal(i).usuario2 = UserIndex Then
                    Torneo_Semifinal(i).ganador = UserIndex
                    Call pasaDeRonda(UserIndex, eRonda.Ronda_Final)
                    Call vuelveUlla(Torneo_Semifinal(i).usuario1)
                    Exit Sub
                End If
                
            Next i
            
            
        Case eRonda.Ronda_Final
            
            If Torneo_Final.usuario1 = UserIndex Then
                Call vuelveUlla(Torneo_Final.usuario2)
                Torneo_Final.ganador = UserIndex
                Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto> " & UserList(UserIndex).Name & " ha ganado el Torneo Automático!" & FONTTYPE_TORNEOAUTO)
                Call saleSegundoTorneo(Torneo_Final.usuario2)
                Call ganaTorneo(UserIndex)
                Call inicializarTorneo
            End If
            
            If Torneo_Final.usuario2 = UserIndex Then
                Call vuelveUlla(Torneo_Final.usuario1)
                Torneo_Final.ganador = UserIndex
                Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto> " & UserList(UserIndex).Name & " ha ganado el Torneo Automático!" & FONTTYPE_TORNEOAUTO)
                Call saleSegundoTorneo(Torneo_Final.usuario1)
                Call ganaTorneo(UserIndex)
                Call inicializarTorneo
            End If
            
        End Select
        
        Torneo_Fixture = Torneo_Fixture & UserList(UserIndex).Name & ","
        
End Sub

Public Sub muerePareja(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
    Dim i As Byte
    
    Select Case Torneo_RondaActual
    
        Case eRonda.Ronda_Dieciseisavos:
        
            For i = 1 To 16
                
                'CHOTS | Muere justo el usuario1 del torneo
                If Torneo_Dieciseisavos(i).usuario1 = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Dieciseisavos(i).usuario2, eRonda.Ronda_Octavos)
                        Torneo_Dieciseisavos(i).ganador = Torneo_Dieciseisavos(i).usuario2
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If

                'CHOTS | Muere la pareja de usuario1
                If UserList(Torneo_Dieciseisavos(i).usuario1).torneoPareja = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Dieciseisavos(i).usuario2, eRonda.Ronda_Octavos)
                        Torneo_Dieciseisavos(i).ganador = Torneo_Dieciseisavos(i).usuario2
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If
                
                'CHOTS | Muere justo el usuario2 del torneo
                If Torneo_Dieciseisavos(i).usuario2 = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Dieciseisavos(i).usuario1, eRonda.Ronda_Octavos)
                        Torneo_Dieciseisavos(i).ganador = Torneo_Dieciseisavos(i).usuario1
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If

                'CHOTS | Muere la pareja de usuario2
                If UserList(Torneo_Dieciseisavos(i).usuario2).torneoPareja = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Dieciseisavos(i).usuario1, eRonda.Ronda_Octavos)
                        Torneo_Dieciseisavos(i).ganador = Torneo_Dieciseisavos(i).usuario1
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If
                
            Next i
    
    
        Case eRonda.Ronda_Octavos:
        
            For i = 1 To 8

                'CHOTS | Muere justo el usuario1 del torneo
                If Torneo_Octavos(i).usuario1 = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Octavos(i).usuario2, eRonda.Ronda_Cuartos)
                        Torneo_Octavos(i).ganador = Torneo_Octavos(i).usuario2
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If

                'CHOTS | Muere la pareja de usuario1
                If UserList(Torneo_Octavos(i).usuario1).torneoPareja = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Octavos(i).usuario2, eRonda.Ronda_Cuartos)
                        Torneo_Octavos(i).ganador = Torneo_Octavos(i).usuario2
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If
                
                'CHOTS | Muere justo el usuario2 del torneo
                If Torneo_Octavos(i).usuario2 = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Octavos(i).usuario1, eRonda.Ronda_Cuartos)
                        Torneo_Octavos(i).ganador = Torneo_Octavos(i).usuario1
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If

                'CHOTS | Muere la pareja de usuario2
                If UserList(Torneo_Octavos(i).usuario2).torneoPareja = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Octavos(i).usuario1, eRonda.Ronda_Cuartos)
                        Torneo_Octavos(i).ganador = Torneo_Octavos(i).usuario1
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If
                
            Next i
    
        Case eRonda.Ronda_Cuartos:
        
            For i = 1 To 4

                'CHOTS | Muere justo el usuario1 del torneo
                If Torneo_Cuartos(i).usuario1 = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Cuartos(i).usuario2, eRonda.Ronda_Semi)
                        Torneo_Cuartos(i).ganador = Torneo_Cuartos(i).usuario2
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If

                'CHOTS | Muere la pareja de usuario1
                If UserList(Torneo_Cuartos(i).usuario1).torneoPareja = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Cuartos(i).usuario2, eRonda.Ronda_Semi)
                        Torneo_Cuartos(i).ganador = Torneo_Cuartos(i).usuario2
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If
                
                'CHOTS | Muere justo el usuario2 del torneo
                If Torneo_Cuartos(i).usuario2 = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Cuartos(i).usuario1, eRonda.Ronda_Semi)
                        Torneo_Cuartos(i).ganador = Torneo_Cuartos(i).usuario1
                    End If
                    
                    Call UserDie(VictimIndex)
                    Exit Sub
                End If

                'CHOTS | Muere la pareja de usuario2
                If UserList(Torneo_Cuartos(i).usuario2).torneoPareja = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Cuartos(i).usuario1, eRonda.Ronda_Semi)
                        Torneo_Cuartos(i).ganador = Torneo_Cuartos(i).usuario1
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If
                
            Next i
            
            
        Case eRonda.Ronda_Semi
        
            For i = 1 To 2

                'CHOTS | Muere justo el usuario1 del torneo
                If Torneo_Semifinal(i).usuario1 = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Semifinal(i).usuario2, eRonda.Ronda_Final)
                        Torneo_Semifinal(i).ganador = Torneo_Semifinal(i).usuario2
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If

                'CHOTS | Muere la pareja de usuario1
                If UserList(Torneo_Semifinal(i).usuario1).torneoPareja = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Semifinal(i).usuario2, eRonda.Ronda_Final)
                        Torneo_Semifinal(i).ganador = Torneo_Semifinal(i).usuario2
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If
                
                'CHOTS | Muere justo el usuario2 del torneo
                If Torneo_Semifinal(i).usuario2 = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Semifinal(i).usuario1, eRonda.Ronda_Final)
                        Torneo_Semifinal(i).ganador = Torneo_Semifinal(i).usuario1
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If

                'CHOTS | Muere la pareja de usuario2
                If UserList(Torneo_Semifinal(i).usuario2).torneoPareja = VictimIndex Then
                    ' La pareja esta muerta, pasa de ronda el otro usuario
                    If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                        Call pasaDeRonda(Torneo_Semifinal(i).usuario1, eRonda.Ronda_Final)
                        Torneo_Semifinal(i).ganador = Torneo_Semifinal(i).usuario1
                        Call vuelveUlla(VictimIndex)
                    Else
                        Call UserDie(VictimIndex)
                        UserList(VictimIndex).flags.enDueloTorneoAuto = False
                    End If
                    Exit Sub
                End If
                
            Next i
            
            
        Case eRonda.Ronda_Final

            'CHOTS | Muere justo el usuario1 del torneo
            If Torneo_Final.usuario1 = VictimIndex Then
                ' La pareja esta muerta, pasa de ronda el otro usuario
                If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                    Call vuelveUlla(VictimIndex)
                    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto> La pareja de " & UserList(Torneo_Final.usuario2).Name & " y " & UserList(UserList(Torneo_Final.usuario2).torneoPareja).Name & " ha ganado el Torneo Automático!" & FONTTYPE_TORNEOAUTO)
                    Call saleSegundoTorneo(VictimIndex)
                    Torneo_Final.ganador = Torneo_Final.usuario2
                    Call ganaTorneo(Torneo_Final.usuario2)
                    Call inicializarTorneo
                Else
                    Call UserDie(VictimIndex)
                    UserList(VictimIndex).flags.enDueloTorneoAuto = False
                End If
                Exit Sub
            End If

            'CHOTS | Muere la pareja de usuario1
            If UserList(Torneo_Final.usuario1).torneoPareja = VictimIndex Then
                ' La pareja esta muerta, pasa de ronda el otro usuario
                If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                    Call vuelveUlla(VictimIndex)
                    Call saleSegundoTorneo(VictimIndex)
                    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto> La pareja de " & UserList(Torneo_Final.usuario2).Name & " y " & UserList(UserList(Torneo_Final.usuario2).torneoPareja).Name & " ha ganado el Torneo Automático!" & FONTTYPE_TORNEOAUTO)
                    Torneo_Final.ganador = Torneo_Final.usuario2
                    Call ganaTorneo(Torneo_Final.usuario2)
                    Call inicializarTorneo
                Else
                    Call UserDie(VictimIndex)
                    UserList(VictimIndex).flags.enDueloTorneoAuto = False
                End If
                Exit Sub
            End If
            
            'CHOTS | Muere justo el usuario2 del torneo
            If Torneo_Final.usuario2 = VictimIndex Then
                ' La pareja esta muerta, pasa de ronda el otro usuario
                If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                    Call vuelveUlla(VictimIndex)
                    Call saleSegundoTorneo(VictimIndex)
                    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto> La pareja de " & UserList(Torneo_Final.usuario1).Name & " y " & UserList(UserList(Torneo_Final.usuario1).torneoPareja).Name & " ha ganado el Torneo Automático!" & FONTTYPE_TORNEOAUTO)
                    Torneo_Final.ganador = Torneo_Final.usuario1
                    Call ganaTorneo(Torneo_Final.usuario1)
                    Call inicializarTorneo
                Else
                    Call UserDie(VictimIndex)
                    UserList(VictimIndex).flags.enDueloTorneoAuto = False
                End If
                Exit Sub
            End If

            'CHOTS | Muere la pareja de usuario2
            If UserList(Torneo_Final.usuario2).torneoPareja = VictimIndex Then
                ' La pareja esta muerta, pasa de ronda el otro usuario
                If UserList(UserList(VictimIndex).torneoPareja).flags.enDueloTorneoAuto = False Then
                    Call vuelveUlla(VictimIndex)
                    Call saleSegundoTorneo(VictimIndex)
                    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto> La pareja de " & UserList(Torneo_Final.usuario1).Name & " y " & UserList(UserList(Torneo_Final.usuario1).torneoPareja).Name & " ha ganado el Torneo Automático!" & FONTTYPE_TORNEOAUTO)
                    Torneo_Final.ganador = Torneo_Final.usuario1
                    Call ganaTorneo(Torneo_Final.usuario1)
                    Call inicializarTorneo
                Else
                    Call UserDie(VictimIndex)
                    UserList(VictimIndex).flags.enDueloTorneoAuto = False
                End If
                Exit Sub
            End If
            
        End Select
End Sub

Public Sub muereDeathmatch(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
    Dim i As Byte
    Dim cantidadDeVivos As Byte

    UserList(VictimIndex).showName = True
    Call vuelveUlla(VictimIndex)
    
    cantidadDeVivos = 0
    For i = 1 To Torneo_CantidadInscriptos
        If UserList(Torneo_UsuariosInscriptos(i)).flags.enTorneoAuto = True Then
            cantidadDeVivos = cantidadDeVivos + 1
        End If
    Next i

    If cantidadDeVivos = 1 Then
        UserList(AttackerIndex).showName = True
        Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto> " & UserList(AttackerIndex).Name & " ha ganado el Torneo Automático!" & FONTTYPE_TORNEOAUTO)
        Call saleSegundoTorneo(VictimIndex)
        Call ganaTorneo(AttackerIndex)
        Call inicializarTorneo
    End If
End Sub

Public Sub pasaDeRonda(ByVal UserIndex As Integer, ByVal ronda As eRonda)
    Dim i As Byte
    
    Select Case ronda

        Case eRonda.Ronda_Octavos:
            
            For i = 1 To 8
            
                If Torneo_Octavos(i).usuario1 = 0 Then
                    Torneo_Octavos(i).usuario1 = UserIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda!" & FONTTYPE_TORNEOAUTO)
                    If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Then
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a Octavos de final.", False)
                        Call telepToAreaDescanso(UserIndex)
                    ElseIf Torneo_Tipo = eTipoTorneo.t2vs2 Then
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).torneoPareja, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda!" & FONTTYPE_TORNEOAUTO)
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & " Pasaron a Octavos de final.", False)
                        Call telepToAreaDescanso(UserIndex)
                    End If
                    Exit Sub
                End If
                
                If Torneo_Octavos(i).usuario2 = 0 Then
                    Torneo_Octavos(i).usuario2 = UserIndex
                    
                    If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda! Tu rival será: " & UserList(Torneo_Octavos(i).usuario1).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, Torneo_Octavos(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será: " & UserList(UserIndex).Name & FONTTYPE_TORNEOAUTO)
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a Octavos de final. Juega vs " & UserList(Torneo_Octavos(i).usuario1).Name, False)
                        Call telepToAreaDescanso(UserIndex)
                    ElseIf Torneo_Tipo = eTipoTorneo.t2vs2 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda! Tu rival será la pareja de: " & UserList(Torneo_Octavos(i).usuario1).Name & " y " & UserList(UserList(Torneo_Octavos(i).usuario1).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).torneoPareja, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda! Tu rival será la pareja de: " & UserList(Torneo_Octavos(i).usuario1).Name & " y " & UserList(UserList(Torneo_Octavos(i).usuario1).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, Torneo_Octavos(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será la pareja de: " & UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, UserList(Torneo_Octavos(i).usuario1).torneoPareja, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será la pareja de: " & UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & " Pasaron a Octavos de final. Juegan vs " & UserList(Torneo_Octavos(i).usuario1).Name & " y " & UserList(UserList(Torneo_Octavos(i).usuario1).torneoPareja).Name, False)
                        Call telepToAreaDescanso(UserIndex)
                    End If

                    If i = 8 Then
                        Call setearCuenta(10, eRonda.Ronda_Octavos)
                    ElseIf (i = 6 Or i = 4 Or i = 2) Then
                        Call comenzarDieciseisavos
                    End If
                    Exit Sub
                End If
            
            Next i
            
    
        Case eRonda.Ronda_Cuartos:
            
            For i = 1 To 4
            
                If Torneo_Cuartos(i).usuario1 = 0 Then
                    Torneo_Cuartos(i).usuario1 = UserIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda!" & FONTTYPE_TORNEOAUTO)
                    If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Then
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a Cuartos de final.", False)
                        Call telepToAreaDescanso(UserIndex)
                    ElseIf Torneo_Tipo = eTipoTorneo.t2vs2 Then
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).torneoPareja, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda!" & FONTTYPE_TORNEOAUTO)
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & " Pasaron a Cuartos de final.", False)
                        Call telepToAreaDescanso(UserIndex)
                    End If
                    Exit Sub
                End If
                
                If Torneo_Cuartos(i).usuario2 = 0 Then
                    Torneo_Cuartos(i).usuario2 = UserIndex
                    
                    If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda! Tu rival será: " & UserList(Torneo_Cuartos(i).usuario1).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, Torneo_Cuartos(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será: " & UserList(UserIndex).Name & FONTTYPE_TORNEOAUTO)
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a Cuartos de final. Juega vs " & UserList(Torneo_Cuartos(i).usuario1).Name, False)
                        Call telepToAreaDescanso(UserIndex)
                    ElseIf Torneo_Tipo = eTipoTorneo.t2vs2 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda! Tu rival será la pareja de: " & UserList(Torneo_Cuartos(i).usuario1).Name & " y " & UserList(UserList(Torneo_Cuartos(i).usuario1).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).torneoPareja, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda! Tu rival será la pareja de: " & UserList(Torneo_Cuartos(i).usuario1).Name & " y " & UserList(UserList(Torneo_Cuartos(i).usuario1).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, Torneo_Cuartos(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será la pareja de: " & UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, UserList(Torneo_Cuartos(i).usuario1).torneoPareja, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será la pareja de: " & UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & " Pasaron a Cuartos de final. Juegan vs " & UserList(Torneo_Cuartos(i).usuario1).Name & " y " & UserList(UserList(Torneo_Cuartos(i).usuario1).torneoPareja).Name, False)
                        Call telepToAreaDescanso(UserIndex)
                    End If

                    If i = 4 Then
                        Call setearCuenta(10, eRonda.Ronda_Cuartos)
                    ElseIf i = 2 Then
                        Call comenzarOctavos
                    End If
                    Exit Sub
                End If
            
            Next i
    
        Case eRonda.Ronda_Semi:
            
            For i = 1 To 2
            
                If Torneo_Semifinal(i).usuario1 = 0 Then
                    Torneo_Semifinal(i).usuario1 = UserIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda!" & FONTTYPE_TORNEOAUTO)
                    If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Then
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a Semifinal.", False)
                        Call telepToAreaDescanso(UserIndex)
                    ElseIf Torneo_Tipo = eTipoTorneo.t2vs2 Then
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).torneoPareja, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda!" & FONTTYPE_TORNEOAUTO)
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & " Pasaron a Semifinal.", False)
                        Call telepToAreaDescanso(UserIndex)
                    End If
                    Exit Sub
                End If
                
                If Torneo_Semifinal(i).usuario2 = 0 Then
                    Torneo_Semifinal(i).usuario2 = UserIndex
                    If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda! Tu rival será: " & UserList(Torneo_Semifinal(i).usuario1).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, Torneo_Semifinal(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será: " & UserList(UserIndex).Name & FONTTYPE_TORNEOAUTO)
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a Semifinal. Juega vs " & UserList(Torneo_Semifinal(i).usuario1).Name, False)
                        Call telepToAreaDescanso(UserIndex)
                    ElseIf Torneo_Tipo = eTipoTorneo.t2vs2 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda! Tu rival será la pareja de: " & UserList(Torneo_Semifinal(i).usuario1).Name & " y " & UserList(UserList(Torneo_Semifinal(i).usuario1).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).torneoPareja, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda! Tu rival será la pareja de: " & UserList(Torneo_Semifinal(i).usuario1).Name & " y " & UserList(UserList(Torneo_Semifinal(i).usuario1).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, Torneo_Semifinal(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será la pareja de: " & UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, UserList(Torneo_Semifinal(i).usuario1).torneoPareja, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será la pareja de: " & UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & " Pasaron a la Semifinal. Juegan vs " & UserList(Torneo_Semifinal(i).usuario1).Name & " y " & UserList(UserList(Torneo_Semifinal(i).usuario1).torneoPareja).Name, False)
                        Call telepToAreaDescanso(UserIndex)
                    End If

                    If i = 2 Then Call setearCuenta(10, eRonda.Ronda_Semi)
                    Exit Sub
                End If
            
            Next i
            
        Case eRonda.Ronda_Final:
            
                If Torneo_Final.usuario1 = 0 Then
                    Torneo_Final.usuario1 = UserIndex
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado a la final del torneo!" & FONTTYPE_TORNEOAUTO)
                    If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Then
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a la Final.", False)
                        Call telepToAreaDescanso(UserIndex)
                    ElseIf Torneo_Tipo = eTipoTorneo.t2vs2 Then
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).torneoPareja, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado de ronda!" & FONTTYPE_TORNEOAUTO)
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & " Pasaron a la final.", False)
                        Call telepToAreaDescanso(UserIndex)
                    End If
                    Exit Sub
                End If
                
                If Torneo_Final.usuario2 = 0 Then
                    Torneo_Final.usuario2 = UserIndex
                    If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado a la final del torneo! Tu rival será: " & UserList(Torneo_Final.usuario1).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, Torneo_Final.usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será: " & UserList(UserIndex).Name & FONTTYPE_TORNEOAUTO)
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " Pasó a la Final. Juega vs " & UserList(Torneo_Final.usuario1).Name, False)
                        Call telepToAreaDescanso(UserIndex)
                    ElseIf Torneo_Tipo = eTipoTorneo.t2vs2 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado a la final del torneo! Tu rival será la pareja de: " & UserList(Torneo_Final.usuario1).Name & " y " & UserList(UserList(Torneo_Final.usuario1).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).torneoPareja, 0, ServerPackages.dialogo & "TorneosAuto> Felicitaciones, has avanzado a la final del torneo! Tu rival será la pareja de: " & UserList(Torneo_Final.usuario1).Name & " y " & UserList(UserList(Torneo_Final.usuario1).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, Torneo_Final.usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será la pareja de: " & UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call SendData(SendTarget.ToIndex, UserList(Torneo_Final.usuario1).torneoPareja, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival será la pareja de: " & UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & FONTTYPE_TORNEOAUTO)
                        Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & " Pasaron a la Final. Juegan vs " & UserList(Torneo_Final.usuario1).Name & " y " & UserList(UserList(Torneo_Final.usuario1).torneoPareja).Name, False)
                        Call telepToAreaDescanso(UserIndex)
                    End If

                    Call setearCuenta(15, eRonda.Ronda_Final)
                    Exit Sub
                End If
            
    End Select
                
        
End Sub

Public Sub vuelveUlla(ByVal UserIndex As Integer)
    On Error GoTo chotserror
    If UserIndex = 0 Or UserList(UserIndex).ConnID = -1 Or UserList(UserIndex).ConnIDValida = False Or UserList(UserIndex).flags.enTorneoAuto = False Then Exit Sub
    Dim Pos As WorldPos
    Dim nPos As WorldPos
    Pos.Map = Torneo_MAPAMUERTE
    Pos.X = 58
    Pos.Y = 45
    Call ClosestLegalPos(Pos, nPos)
    Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
    Call UserDie(UserIndex)
    UserList(UserIndex).flags.enTorneoAuto = False
    UserList(UserIndex).flags.enDueloTorneoAuto = False
    UserList(UserIndex).Counters.Torneo = 0

    'CHOTS | Si es 2vs2 llevamos a su pareja tambien
    If Torneo_Tipo = eTipoTorneo.t2vs2 Then
        If UserList(UserIndex).torneoPareja > 0 Then
            If UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                Call ClosestLegalPos(Pos, nPos)
                Call WarpUserChar(UserList(UserIndex).torneoPareja, nPos.Map, nPos.X, nPos.Y, True)
                Call UserDie(UserList(UserIndex).torneoPareja)
                UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = False
                UserList(UserList(UserIndex).torneoPareja).flags.enDueloTorneoAuto = False
                UserList(UserList(UserIndex).torneoPareja).Counters.Torneo = 0

                'CHOTS | Reseteamos parejas
                UserList(UserList(UserIndex).torneoPareja).torneoPareja = 0
                UserList(UserIndex).torneoPareja = 0
            End If
        End If
    End If
    Exit Sub
    
chotserror:
    Call LogError("Error en Vuelveulla " & Err.number & " " & Err.Description)
End Sub

Public Sub ganaTorneo(ByVal UserIndex As Integer)
    On Error GoTo chotserror
    Dim Pos As WorldPos
    Dim nPos As WorldPos
    Pos.Map = Torneo_MAPAMUERTE
    Pos.X = 58
    Pos.Y = 46

    If UserIndex > 0 Then
        If UserList(UserIndex).flags.enTorneoAuto = True Then
            If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.Deathmatch Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Or Torneo_Tipo = eTipoTorneo.Destruccion Then
                Call ClosestLegalPos(Pos, nPos)
                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                UserList(UserIndex).Stats.TorneosAuto(Torneo_Tipo) = UserList(UserIndex).Stats.TorneosAuto(Torneo_Tipo) + 1
                UserList(UserIndex).flags.enTorneoAuto = False
                UserList(UserIndex).flags.enDueloTorneoAuto = False
                Call darPremioTorneo(UserIndex)
                Call ActualizarRankingTorneos(UserIndex, Torneo_Tipo) 'CHOTS | Ranking de Torneos Automaticos
                Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " ha ganado el torneo.", False)
            ElseIf Torneo_Tipo = eTipoTorneo.t2vs2 Then
                Call ClosestLegalPos(Pos, nPos)
                Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, True)
                UserList(UserIndex).Stats.TorneosAuto(Torneo_Tipo) = UserList(UserIndex).Stats.TorneosAuto(Torneo_Tipo) + 1
                UserList(UserIndex).flags.enTorneoAuto = False
                UserList(UserIndex).flags.enDueloTorneoAuto = False
                If UserList(UserIndex).flags.Muerto = 1 Then Resucitar (UserIndex)
                
                Call darPremioTorneo(UserIndex)
                Call ActualizarRankingTorneos(UserIndex, Torneo_Tipo) 'CHOTS | Ranking de Torneos Automaticos

                If UserList(UserIndex).torneoPareja > 0 Then
                    Call ClosestLegalPos(Pos, nPos)
                    Call WarpUserChar(UserList(UserIndex).torneoPareja, nPos.Map, nPos.X, nPos.Y, True)
                    UserList(UserList(UserIndex).torneoPareja).Stats.TorneosAuto(Torneo_Tipo) = UserList(UserList(UserIndex).torneoPareja).Stats.TorneosAuto(Torneo_Tipo) + 1
                    UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = False
                    UserList(UserList(UserIndex).torneoPareja).flags.enDueloTorneoAuto = False
                    If UserList(UserList(UserIndex).torneoPareja).flags.Muerto = 1 Then Resucitar (UserList(UserIndex).torneoPareja)
                    Call darPremioTorneo(UserList(UserIndex).torneoPareja)
                    Call ActualizarRankingTorneos(UserList(UserIndex).torneoPareja, Torneo_Tipo) 'CHOTS | Ranking de Torneos Automaticos
                    Call LogGM("TORNEOAUTO", UserList(UserIndex).Name & " y " & UserList(UserList(UserIndex).torneoPareja).Name & " han ganado el torneo.", False)
                End If
            End If
        End If
    End If
    Exit Sub
chotserror:
    Call LogError("Error en GanaTorneo " & Err.number & " " & Err.Description)
End Sub

Public Sub inicializarTorneo()
Dim i as Integer
Torneo_HAYTORNEO = False
MinutosParaTorneo = 0
Torneo_Tipo = 0
For i = 1 to LastUser
    Userlist(i).flags.enTorneoAuto = False
    Userlist(i).flags.enDueloTorneoAuto = False
    Userlist(i).torneoPareja = 0
Next i
End Sub

Public Sub saleSegundoTorneo(ByVal UserIndex As Integer)
    On Error GoTo chotserror
    If UserIndex > 0 Then
        If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.Deathmatch Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Then
            Call darPremioSegundoTorneo(UserIndex)
        ElseIf Torneo_Tipo = eTipoTorneo.t2vs2 Then
            Call darPremioSegundoTorneo(UserIndex)
            If UserList(UserIndex).torneoPareja > 0 Then
                Call darPremioSegundoTorneo(UserList(UserIndex).torneoPareja)
            End If
        End If
    End If
    Exit Sub
chotserror:
    Call LogError("Error en saleSegundoTorneo " & Err.number & " " & Err.Description)
End Sub

Public Sub setearCuenta(ByVal segundos As Byte, ByRef ronda As eRonda)
    Dim razon As String

    Torneo_RondaActual = ronda
    
    Select Case ronda
        Case eRonda.Ronda_Dieciseisavos:
            razon = "Los diesiseisavos de final"
        Case eRonda.Ronda_Octavos:
            razon = "Los octavos de final"
        Case eRonda.Ronda_Cuartos:
            razon = "Los cuartos de final"
        Case eRonda.Ronda_Semi
            razon = "La semifinal"
        Case eRonda.Ronda_Final
            razon = "La Final"
        Case eRonda.Ronda_Deathmatch
            razon = "El Deathmatch"
        Case eRonda.Ronda_Destruccion
            razon = "El evento de Destruccion"
        Case Else
            razon = "El siguiente evento"
    End Select
            
    Call SendData(SendTarget.ToMap, 0, Torneo_MAPAMUERTE, ServerPackages.dialogo & "En " & segundos & " segundos comenzará " & razon & FONTTYPE_TORNEOAUTO)
    Torneo_CR.razon = razon
    Torneo_CR.segundos = segundos
    Torneo_CR.next = ronda

End Sub

Public Sub finalizarCuenta()
    Select Case Torneo_CR.next
        Case eRonda.Ronda_Dieciseisavos:
            Call comenzarDieciseisavos
            Exit Sub
            
        Case eRonda.Ronda_Octavos:
            Call comenzarOctavos
            Exit Sub
            
        Case eRonda.Ronda_Cuartos:
            Call comenzarCuartos
            Exit Sub
            
        Case eRonda.Ronda_Semi:
            Call comenzarSemifinal
            Exit Sub
            
        Case eRonda.Ronda_Final:
            Call comenzarFinal
            Exit Sub

        Case eRonda.Ronda_Deathmatch:
            Call comenzarDeathmatch
            Exit Sub

        Case eRonda.Ronda_Destruccion:
            Call comenzarDestruccion
            Exit Sub
            
    End Select
        
End Sub

Public Sub irseTorneo(ByVal UserIndex As Integer)
    Dim i As Integer
    
    Select Case Torneo_RondaActual
    
        Case eRonda.Ronda_Dieciseisavos:
            For i = 1 To 16
                If Torneo_Dieciseisavos(i).ganador <> UserIndex Then
                    If Torneo_Dieciseisavos(i).usuario1 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Dieciseisavos(i).usuario2, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Dieciseisavos(i).usuario2)
                        If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                            Call vuelveUlla(UserList(UserIndex).torneoPareja)
                        End If
                        Exit Sub
                    End If
                    
                    If Torneo_Dieciseisavos(i).usuario2 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Dieciseisavos(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Dieciseisavos(i).usuario1)
                        If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                            Call vuelveUlla(UserList(UserIndex).torneoPareja)
                        End If
                        Exit Sub
                    End If
                Else
                    For j = 1 To 8
                        If Torneo_Octavos(j).usuario1 = UserIndex Then
                            Torneo_Octavos(j).usuario1 = 0
                            If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                                Call vuelveUlla(UserList(UserIndex).torneoPareja)
                            End If
                            Exit Sub
                        End If
                        
                        If Torneo_Octavos(j).usuario2 = UserIndex Then
                            Torneo_Octavos(j).usuario2 = 0
                            If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                                Call vuelveUlla(UserList(UserIndex).torneoPareja)
                            End If
                            Exit Sub
                        End If
                    Next j
                End If
            Next i
            
        Case eRonda.Ronda_Octavos:
            For i = 1 To 8
                If Torneo_Octavos(i).ganador <> UserIndex Then
                    If Torneo_Octavos(i).usuario1 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Octavos(i).usuario2, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Octavos(i).usuario2)
                        If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                            Call vuelveUlla(UserList(UserIndex).torneoPareja)
                        End If
                        Exit Sub
                    End If
                    
                    If Torneo_Octavos(i).usuario2 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Octavos(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Octavos(i).usuario1)
                        If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                            Call vuelveUlla(UserList(UserIndex).torneoPareja)
                        End If
                        Exit Sub
                    End If
                Else
                    For j = 1 To 4
                        If Torneo_Cuartos(j).usuario1 = UserIndex Then
                            Torneo_Cuartos(j).usuario1 = 0
                            If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                                Call vuelveUlla(UserList(UserIndex).torneoPareja)
                            End If
                            Exit Sub
                        End If
                        
                        If Torneo_Cuartos(i).usuario2 = UserIndex Then
                            Torneo_Cuartos(j).usuario2 = 0
                            If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                                Call vuelveUlla(UserList(UserIndex).torneoPareja)
                            End If
                            Exit Sub
                        End If
                    Next j
                End If
            Next i
            
        Case eRonda.Ronda_Cuartos:
            For i = 1 To 4
                If Torneo_Cuartos(i).ganador <> UserIndex Then
                    If Torneo_Cuartos(i).usuario1 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Cuartos(i).usuario2, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Cuartos(i).usuario2)
                        If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                            Call vuelveUlla(UserList(UserIndex).torneoPareja)
                        End If
                        Exit Sub
                    End If
                    
                    If Torneo_Cuartos(i).usuario2 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Cuartos(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Cuartos(i).usuario1)
                        If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                            Call vuelveUlla(UserList(UserIndex).torneoPareja)
                        End If
                        Exit Sub
                    End If
                Else
                    For j = 1 To 2
                        If Torneo_Semifinal(j).usuario1 = UserIndex Then
                            Torneo_Semifinal(j).usuario1 = 0
                            If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                                Call vuelveUlla(UserList(UserIndex).torneoPareja)
                            End If
                            Exit Sub
                        End If
                        
                        If Torneo_Semifinal(i).usuario2 = UserIndex Then
                            Torneo_Semifinal(j).usuario2 = 0
                            If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                                Call vuelveUlla(UserList(UserIndex).torneoPareja)
                            End If
                            Exit Sub
                        End If
                    Next j
                End If
            Next i
            
        Case eRonda.Ronda_Semi:
            For i = 1 To 2
                If Torneo_Semifinal(i).ganador <> UserIndex Then
                    If Torneo_Semifinal(i).usuario1 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Semifinal(i).usuario2, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Semifinal(i).usuario2)
                        If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                            Call vuelveUlla(UserList(UserIndex).torneoPareja)
                        End If
                        Exit Sub
                    End If
                    
                    If Torneo_Semifinal(i).usuario2 = UserIndex Then
                        Call SendData(SendTarget.ToIndex, Torneo_Semifinal(i).usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                        Call ganaUsuario(Torneo_Semifinal(i).usuario1)
                        If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                            Call vuelveUlla(UserList(UserIndex).torneoPareja)
                        End If
                        Exit Sub
                    End If
                Else
                    If Torneo_Final.usuario1 = UserIndex Then
                        Torneo_Final.usuario1 = 0
                        If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                            Call vuelveUlla(UserList(UserIndex).torneoPareja)
                        End If
                        Exit Sub
                    End If
                    
                    If Torneo_Final.usuario2 = UserIndex Then
                        Torneo_Final.usuario2 = 0
                        If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                            Call vuelveUlla(UserList(UserIndex).torneoPareja)
                        End If
                        Exit Sub
                    End If
                End If
            Next i
            
        Case eRonda.Ronda_Final:
            If Torneo_Final.usuario1 = UserIndex Then
                Call SendData(SendTarget.ToIndex, Torneo_Final.usuario2, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                Call ganaUsuario(Torneo_Final.usuario2)
                If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                    Call vuelveUlla(UserList(UserIndex).torneoPareja)
                End If
                Exit Sub
            End If
            
            If Torneo_Final.usuario2 = UserIndex Then
                Call SendData(SendTarget.ToIndex, Torneo_Final.usuario1, 0, ServerPackages.dialogo & "TorneosAuto> Tu rival ha abandonado el torneo!" & FONTTYPE_TORNEOAUTO)
                Call ganaUsuario(Torneo_Final.usuario1)
                If Torneo_Tipo = eTipoTorneo.t2vs2 And UserList(UserIndex).torneoPareja > 0 And UserList(UserList(UserIndex).torneoPareja).flags.enTorneoAuto = True Then
                    Call vuelveUlla(UserList(UserIndex).torneoPareja)
                End If
                Exit Sub
            End If
            
        Case Else:
            For i = 1 To Torneo_CantidadInscriptos
            
                If Torneo_UsuariosInscriptos(i) = UserIndex Then
                    Torneo_UsuariosInscriptos(i) = 0

                    If Torneo_Tipo = eTipoTorneo.Deathmatch Then
                        UserList(i).showName = True
                    Else
                        'CHOTS | Deslogeo antes de que empiece lo sacamos y movemos el array si no es el ultimo
                        If i <> Torneo_CantidadInscriptos Then
                            Torneo_UsuariosInscriptos(i) = Torneo_UsuariosInscriptos(Torneo_CantidadInscriptos)
                            Torneo_UsuariosInscriptos(Torneo_CantidadInscriptos) = 0
                        End If

                        Torneo_CantidadInscriptos = Torneo_CantidadInscriptos - 1
                    End If

                    Exit Sub
                End If
                
            Next i
    
    End Select

    'CHOTS | Si llego hasta aca es porque el que deslogeo es una pareja, pero por las dudas corroboramos
    If Torneo_Tipo = eTipoTorneo.t2vs2 Then
        If UserList(UserIndex).torneoPareja > 0 Then
            UserList(UserList(UserIndex).torneoPareja).torneoPareja = 0
        End If
    End If
    
End Sub

Public Sub cerrarTorneo()
    Dim i As Integer
    Dim Pos As WorldPos
    Dim nPos As WorldPos
    
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto> El torneo ha sido cancelado." & FONTTYPE_TORNEOAUTO)
    
    For i = 1 To LastUser
        If UserList(i).flags.enTorneoAuto = True Then
            Pos.Map = 1
            Pos.X = 58
            Pos.Y = 45
            Call ClosestLegalPos(Pos, nPos)
            Call WarpUserChar(i, nPos.Map, nPos.X, nPos.Y, True)
            UserList(i).flags.enTorneoAuto = False
        End If
    Next i

    Torneo_HAYTORNEO = False
    MinutosParaTorneo = 0
    Torneo_Tipo = 0

End Sub

Public Sub terminarDuelo(ByVal UserIndex As Integer)
Dim i As Integer
    
    Select Case Torneo_RondaActual
    
        Case eRonda.Ronda_Dieciseisavos:
            For i = 1 To 16
            
                If Torneo_Dieciseisavos(i).usuario1 = UserIndex Then
                    Call SendData(SendTarget.ToIndex, Torneo_Dieciseisavos(i).usuario1, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_Dieciseisavos(i).usuario2, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    If val(RandomNumber(1, 2)) = 1 Then
                        Call ganaUsuario(Torneo_Dieciseisavos(i).usuario2)
                    Else
                        Call ganaUsuario(Torneo_Dieciseisavos(i).usuario1)
                    End If
                    Exit Sub
                End If

            Next i
            
        Case eRonda.Ronda_Octavos:
            For i = 1 To 8
            
                If Torneo_Octavos(i).usuario1 = UserIndex Then
                    Call SendData(SendTarget.ToIndex, Torneo_Octavos(i).usuario1, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_Octavos(i).usuario2, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    If val(RandomNumber(1, 2)) = 1 Then
                        Call ganaUsuario(Torneo_Octavos(i).usuario2)
                    Else
                        Call ganaUsuario(Torneo_Octavos(i).usuario1)
                    End If
                    Exit Sub
                End If
            Next i
            
        Case eRonda.Ronda_Cuartos:
            For i = 1 To 4
            
                If Torneo_Cuartos(i).usuario1 = UserIndex Then
                    Call SendData(SendTarget.ToIndex, Torneo_Cuartos(i).usuario1, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_Cuartos(i).usuario2, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    
                    If val(RandomNumber(1, 2)) = 1 Then
                        Call ganaUsuario(Torneo_Cuartos(i).usuario2)
                    Else
                        Call ganaUsuario(Torneo_Cuartos(i).usuario1)
                    End If
                    
                    Exit Sub
                End If

            Next i
            
        Case eRonda.Ronda_Semi:
            For i = 1 To 2
            
                If Torneo_Semifinal(i).usuario1 = UserIndex Then
                    Call SendData(SendTarget.ToIndex, Torneo_Semifinal(i).usuario1, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_Semifinal(i).usuario2, 0, ServerPackages.dialogo & "Se ha llegado al tiempo límite, se decidirá por azar." & FONTTYPE_TORNEOAUTO)
                    
                    If val(RandomNumber(1, 2)) = 1 Then
                        Call ganaUsuario(Torneo_Semifinal(i).usuario2)
                    Else
                        Call ganaUsuario(Torneo_Semifinal(i).usuario1)
                    End If
                    
                    Exit Sub
                End If
                
            Next i
    
    End Select
End Sub

Public Sub rearmarTorneo()
On Error GoTo CHOTSERR
Dim Pos As WorldPos
Dim nPos As WorldPos

If Torneo_Tipo = eTipoTorneo.t1vs1 Or Torneo_Tipo = eTipoTorneo.t2vs2 Or Torneo_Tipo = eTipoTorneo.Plantes Or Torneo_Tipo = eTipoTorneo.Aim Or Torneo_Tipo = eTipoTorneo.Destruccion Then
    If Torneo_CantidadInscriptos >= (Torneo_Cupo / 2) Then

        Call SendData(SendTarget.ToMap, 0, Torneo_MAPAMUERTE, ServerPackages.dialogo & "El torneo se ha reorganizado para " & Torneo_Cupo / 2 & " usuarios. Los sobrantes serán enviados a Ullathorpe" & FONTTYPE_TORNEOAUTO)

        'CHOTS | Los envía a Ulla a los sobrantes
        If Torneo_CantidadInscriptos > (Torneo_Cupo / 2) Then
            For i = ((Torneo_Cupo / 2) + 1) To Torneo_Cupo
                If Torneo_UsuariosInscriptos(i) > 0 Then
                    Pos.Map = 1
                    Pos.X = 58
                    Pos.Y = 45
                    Call ClosestLegalPos(Pos, nPos)
                    Call WarpUserChar(Torneo_UsuariosInscriptos(i), nPos.Map, nPos.X, nPos.Y, True)
                    UserList(Torneo_UsuariosInscriptos(i)).flags.enTorneoAuto = False
                    Torneo_UsuariosInscriptos(i) = 0
                End If
            Next i
        End If
        
        Torneo_Cupo = Torneo_Cupo / 2
        Torneo_CantidadInscriptos = Torneo_Cupo
        
        Call armarFixture

        Call LogGM("TORNEOAUTO", "El torneo se reorganizo para " & Torneo_Cupo & " participantes.", False)
    Else
        Call LogGM("TORNEOAUTO", "El torneo se cancelo por falta de participantes.", False)
        Call cerrarTorneo
    End If
ElseIf Torneo_Tipo = eTipoTorneo.Deathmatch Then
    If Torneo_CantidadInscriptos < 2 Then
        Call LogGM("TORNEOAUTO", "El torneo se cancelo por falta de participantes.", False)
        Call cerrarTorneo
    Else
        Torneo_Cupo = Torneo_CantidadInscriptos
        Call SendData(SendTarget.ToMap, 0, Torneo_MAPAMUERTE, ServerPackages.dialogo & "El torneo se ha reorganizado para " & Torneo_CantidadInscriptos & " usuarios." & FONTTYPE_TORNEOAUTO)
        Call LogGM("TORNEOAUTO", "El torneo se reorganizo para " & Torneo_CantidadInscriptos & " participantes.", False)
        Call armarFixture
    End If
End If
Exit Sub

CHOTSERR:
    Call LogError("Error en rearmarTorneo " & Err.number & " " & Err.Description)

End Sub

Public Function minutosProxTorneo() As String
    On Error GoTo CHOTSERR
    Dim minutelis As Long

    minutelis = 10 - MinutosParaTorneo
    
    minutosProxTorneo = Trim$(str$(minutelis))
    Exit Function
    
CHOTSERR:
    Call LogError("Error en Proxtorneos " & Err.number & " " & Err.Description)
    
End Function

Private Sub darPremioTorneo(ByVal UserIndex As Integer)

    If puedeDarPremioTorneo() Then
        Dim MiObj As Obj

        MiObj.Amount = 1
        MiObj.ObjIndex = TROFEOORO
        If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has ganado el torneo! Felicitaciones, aquí tienes tu premio." & FONTTYPE_TORNEOAUTO)
    End If

End Sub

Private Sub darPremioSegundoTorneo(ByVal UserIndex As Integer)

    If puedeDarPremioTorneo() Then
        Dim MiObj As Obj
        MiObj.Amount = 1
        MiObj.ObjIndex = TROFEOPLATA
        If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has obtenido el segundo lugar, aquí tienes tu premio." & FONTTYPE_TORNEOAUTO)
    End If

End Sub

Private Function puedeDarPremioTorneo() As Boolean
    puedeDarPremioTorneo = False

    If Torneo_Tipo = eTipoTorneo.t2vs2 Then
        If Torneo_CantidadInscriptos >= 4 Then puedeDarPremioTorneo = True
    Else
        If Torneo_CantidadInscriptos >= 8 Then puedeDarPremioTorneo = True
    End If
End Function

Public Function votarTorneo(ByVal VotanteIndex As Integer, ByVal Numero As Integer) As Boolean
    'Se podria usar collections aca pero mucha paja. Mejor malo conocido
    votarTorneo = False
    Dim i As Integer
    Dim UserName As String
    UserName = UCase$(UserList(VotanteIndex).Name)

    If val(Numero) < 1 Or val(Numero) > Torneo_TIPOTORNEOS Then
        votarTorneo = False
        Exit Function
    End If
    
    'CHOTS | Primero checkeamos que no haya votado
    For i = 1 To 255
        If UCase$(Torneo_Votantes(i)) = UserName Then
            votarTorneo = False
            Exit Function
        End If
    Next i

    'CHOTS | Ahora sumamos el voto
    For i = 1 To 255
        If UCase$(Torneo_Votantes(i)) = vbNullString Then
            Torneo_Votantes(i) = UserName
            votarTorneo = True

            If val(Numero) = 1 Then
                Torneo_Votos_1 = Torneo_Votos_1 + 1
            ElseIf val(Numero) = 2 Then
                Torneo_Votos_2 = Torneo_Votos_2 + 1
            ElseIf val(Numero) = 3 Then
                Torneo_Votos_3 = Torneo_Votos_3 + 1
            ElseIf val(Numero) = 4 Then
                Torneo_Votos_4 = Torneo_Votos_4 + 1
            ElseIf val(Numero) = 5 Then
                Torneo_Votos_5 = Torneo_Votos_5 + 1
            ElseIf val(Numero) = 6 Then
                Torneo_Votos_6 = Torneo_Votos_6 + 1
            End If
            
            Call LogGM("TORNEOAUTO", UserName & " votó " & Numero & ".", False)

            Exit Function
        End If
    Next i

End Function

Public Sub contarVotosTorneo()

    Call LogGM("TORNEOAUTO", "Votacion finalizada: 1(" & Torneo_Votos_1 & "), 2(" & Torneo_Votos_2 & "), 3(" & Torneo_Votos_3 & "), 4(" & Torneo_Votos_4 & "), 5(" & Torneo_Votos_5 & "), 6(" & Torneo_Votos_6 & ").", False)

    If Torneo_Votos_1 = 0 And Torneo_Votos_2 = 0 And Torneo_Votos_3 = 0 And Torneo_Votos_4 = 0 And Torneo_Votos_5 = 0 And Torneo_Votos_6 = 0 Then
        Torneo_Tipo = eTipoTorneo.Deathmatch
        Exit Sub
    End If

    If Torneo_Votos_1 >= Torneo_Votos_2 And Torneo_Votos_1 >= Torneo_Votos_3 And Torneo_Votos_1 >= Torneo_Votos_4 And Torneo_Votos_1 >= Torneo_Votos_5 And Torneo_Votos_1 >= Torneo_Votos_6 Then
        Torneo_Tipo = eTipoTorneo.t1vs1
        Exit Sub
    End If

    If Torneo_Votos_2 >= Torneo_Votos_1 And Torneo_Votos_2 >= Torneo_Votos_3 And Torneo_Votos_2 >= Torneo_Votos_4 And Torneo_Votos_2 >= Torneo_Votos_5 And Torneo_Votos_2 >= Torneo_Votos_6 Then
        Torneo_Tipo = eTipoTorneo.t2vs2
        Exit Sub
    End If

    If Torneo_Votos_3 >= Torneo_Votos_1 And Torneo_Votos_3 >= Torneo_Votos_2 And Torneo_Votos_3 >= Torneo_Votos_4 And Torneo_Votos_3 >= Torneo_Votos_5 And Torneo_Votos_3 >= Torneo_Votos_6 Then
        Torneo_Tipo = eTipoTorneo.Deathmatch
        Exit Sub
    End If

    If Torneo_Votos_4 >= Torneo_Votos_1 And Torneo_Votos_4 >= Torneo_Votos_2 And Torneo_Votos_4 >= Torneo_Votos_3 And Torneo_Votos_4 >= Torneo_Votos_5 And Torneo_Votos_4 >= Torneo_Votos_6 Then
        Torneo_Tipo = eTipoTorneo.Plantes
        Exit Sub
    End If

    If Torneo_Votos_5 >= Torneo_Votos_1 And Torneo_Votos_5 >= Torneo_Votos_2 And Torneo_Votos_5 >= Torneo_Votos_3 And Torneo_Votos_5 >= Torneo_Votos_4 And Torneo_Votos_5 >= Torneo_Votos_6 Then
        Torneo_Tipo = eTipoTorneo.Aim
        Exit Sub
    End If

    If Torneo_Votos_6 >= Torneo_Votos_1 And Torneo_Votos_6 >= Torneo_Votos_2 And Torneo_Votos_6 >= Torneo_Votos_3 And Torneo_Votos_6 >= Torneo_Votos_4 And Torneo_Votos_6 >= Torneo_Votos_5 Then
        Torneo_Tipo = eTipoTorneo.Destruccion
        Exit Sub
    End If

End Sub

Public Sub reinicializarVotacionesTorneo()
    Call LogGM("TORNEOAUTO", "Se ha abierto una votacion para un torneo Auto.", False)
    Torneo_Votos_1 = 0
    Torneo_Votos_2 = 0
    Torneo_Votos_3 = 0
    Torneo_Votos_4 = 0
    Torneo_Votos_5 = 0
    Torneo_Votos_6 = 0
    ReDim Torneo_Votantes(1 To 255)
End Sub

Public Function getTipoTorneoString() As String
    getTipoTorneoString = "Tipo no definido"

    If Torneo_Tipo = eTipoTorneo.t1vs1 Then
        getTipoTorneoString = "1vs1"
    ElseIf Torneo_Tipo = eTipoTorneo.t2vs2 Then
        getTipoTorneoString = "2vs2"
    ElseIf Torneo_Tipo = eTipoTorneo.Deathmatch Then
        getTipoTorneoString = "Deathmatch"
    ElseIf Torneo_Tipo = eTipoTorneo.Plantes Then
        getTipoTorneoString = "Plantes"
    ElseIf Torneo_Tipo = eTipoTorneo.Aim Then
        getTipoTorneoString = "Al Aim"
    ElseIf Torneo_Tipo = eTipoTorneo.Destruccion Then
        getTipoTorneoString = "Destruccion"
    End If
End Function

Public Sub armarParejasTorneo()
    On Error GoTo chotserror
    Dim i As Integer
    Dim j As Integer
    Dim Torneo_NuevosInscriptos() As Integer
    Dim nuevoSize As Integer
    nuevoSize = Torneo_CantidadInscriptos / 2
    ReDim Torneo_NuevosInscriptos(1 To nuevoSize) As Integer
    Dim k As Integer

    k = 1
    For i = 1 To Torneo_CantidadInscriptos
        
        'CHOTS | Primero le buscamos una pareja de distinta clase
        If UserList(Torneo_UsuariosInscriptos(i)).torneoPareja = 0 Then
            For j = Torneo_CantidadInscriptos To 1 Step -1
                If UserList(Torneo_UsuariosInscriptos(j)).torneoPareja = 0 And UCase$(UserList(Torneo_UsuariosInscriptos(i)).Clase) <> UCase$(UserList(Torneo_UsuariosInscriptos(j)).Clase) Then
                    UserList(Torneo_UsuariosInscriptos(i)).torneoPareja = Torneo_UsuariosInscriptos(j)
                    UserList(Torneo_UsuariosInscriptos(j)).torneoPareja = Torneo_UsuariosInscriptos(i)

                    Torneo_NuevosInscriptos(k) = Torneo_UsuariosInscriptos(j)
                    'CHOTS | Les avisamos quien es su pareja
                    Call SendData(SendTarget.ToIndex, Torneo_UsuariosInscriptos(i), 0, ServerPackages.dialogo & "TorneosAuto> Tu pareja para este torneo será: " & UserList(Torneo_UsuariosInscriptos(j)).Name & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_UsuariosInscriptos(j), 0, ServerPackages.dialogo & "TorneosAuto> Tu pareja para este torneo será: " & UserList(Torneo_UsuariosInscriptos(i)).Name & FONTTYPE_TORNEOAUTO)
                    k = k + 1
                    Exit For
                End If
            Next j
        End If

        'CHOTS | Si no le conseguimos, va a seguir sin pareja asi que buscamos cualquier otro user
        If UserList(Torneo_UsuariosInscriptos(i)).torneoPareja = 0 Then
            For j = Torneo_CantidadInscriptos To 1 Step -1
                If UserList(Torneo_UsuariosInscriptos(j)).torneoPareja = 0 Then
                    UserList(Torneo_UsuariosInscriptos(i)).torneoPareja = Torneo_UsuariosInscriptos(j)
                    UserList(Torneo_UsuariosInscriptos(j)).torneoPareja = Torneo_UsuariosInscriptos(i)

                    Torneo_NuevosInscriptos(k) = Torneo_UsuariosInscriptos(j)
                    'CHOTS | Les avisamos quien es su pareja
                    Call SendData(SendTarget.ToIndex, Torneo_UsuariosInscriptos(i), 0, ServerPackages.dialogo & "TorneosAuto> Tu pareja para este torneo será: " & UserList(Torneo_UsuariosInscriptos(j)).Name & FONTTYPE_TORNEOAUTO)
                    Call SendData(SendTarget.ToIndex, Torneo_UsuariosInscriptos(j), 0, ServerPackages.dialogo & "TorneosAuto> Tu pareja para este torneo será: " & UserList(Torneo_UsuariosInscriptos(i)).Name & FONTTYPE_TORNEOAUTO)
                    k = k + 1
                    Exit For
                End If
            Next j
        End If
        
    Next i

    Torneo_UsuariosInscriptos = Torneo_NuevosInscriptos
    Torneo_Cupo = nuevoSize
    Torneo_CantidadInscriptos = nuevoSize
    Exit Sub
chotserror:
    Call LogError("Error en armarParejasTorneo " & Err.number & " " & Err.Description)
End Sub

Public Sub armarEquiposDestruccion()
On Error GoTo chotserror
    Dim integrantesPorEquipo As Integer
    Dim i As Integer
    Dim j As Integer
    Dim nPos As WorldPos
    integrantesPorEquipo = Torneo_CantidadInscriptos / 4
    
    ReDim Torneo_DestruccionEquipos(1 To Torneo_EQUIPOSDESTRUCCION) As tEquipoDestruccion

    For i = 1 To Torneo_EQUIPOSDESTRUCCION
        Torneo_DestruccionEquipos(i).Numero = i
        ReDim Torneo_DestruccionEquipos(i).usuarios(1 To integrantesPorEquipo)
    Next i

    j = 1
    k = 1
    For i = 1 To Torneo_CantidadInscriptos

        Torneo_DestruccionEquipos(j).usuarios(k) = Torneo_UsuariosInscriptos(i)
        Call SendData(SendTarget.ToIndex, Torneo_UsuariosInscriptos(i), 0, ServerPackages.dialogo & "TorneosAuto> Has sido asignado al equipo " & j & FONTTYPE_TORNEOAUTO)

        j = j + 1

        If j > Torneo_EQUIPOSDESTRUCCION Then
            j = 1
            k = k + 1
        End If
    Next i

    'CHOTS | Spawneamos los NPCs
    For i = 1 To Torneo_EQUIPOSDESTRUCCION
        nPos.Map = Torneo_MAPATORNEO
        nPos.X = RandomNumber(Torneo_AreasDuelo(i).MinX, Torneo_AreasDuelo(i).MaxX)
        nPos.Y = RandomNumber(Torneo_AreasDuelo(i).MinY, Torneo_AreasDuelo(i).MaxY)
        Torneo_DestruccionEquipos(i).NpcIndex = SpawnNpc(Torneo_NPCDESTRUCCION, nPos, False, False)
    Next i

    Exit Sub
chotserror:
    Call LogError("Error en armarEquiposDestruccion " & Err.number & " " & Err.Description)
End Sub

'CHOTS | En Deathmatch los users no pueden hablar
Public Function isUserLocked(ByVal UserIndex As Integer) As Boolean
    isUserLocked = (UserList(UserIndex).showName = False And UserList(UserIndex).flags.enDueloTorneoAuto And Torneo_Tipo = eTipoTorneo.Deathmatch)
End Function


Public Sub CheckChangeBodyTorre(ByVal NpcIndex As Integer)
    Dim nuevoBody As Integer
    Dim oldBody As Integer

    oldBody = Npclist(NpcIndex).char.Body

    If Npclist(NpcIndex).Stats.MinHP < 10000 Then
        nuevoBody = 261
    ElseIf Npclist(NpcIndex).Stats.MinHP < 20000 Then
        nuevoBody = 260
    Else
        nuevoBody = 259
    End If

    If nuevoBody <> oldBody Then
        Call ChangeNPCChar(SendTarget.ToMap, 0, Npclist(NpcIndex).Pos.Map, NpcIndex, nuevoBody, Npclist(NpcIndex).char.Head, Npclist(NpcIndex).char.Heading)
    End If
End Sub

Public Sub muereNpcTorre(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    On Error GoTo chotserror
    Dim i As Integer
    Dim j As Integer
    Dim equipoGanador As Byte

    equipoGanador = 0

    For i = 1 To Torneo_EQUIPOSDESTRUCCION
        If Torneo_DestruccionEquipos(i).NpcIndex = NpcIndex Then
            equipoGanador = i
            Exit For
        End If
    Next i
    
    Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "TorneosAuto> El equipo " & i & " ha ganado el Torneo Automático!" & FONTTYPE_TORNEOAUTO)

    For i = 1 To Torneo_EQUIPOSDESTRUCCION
        For j = 1 To UBound(Torneo_DestruccionEquipos(i).usuarios)
            If i = equipoGanador Then
                Call ganaTorneo(Torneo_DestruccionEquipos(i).usuarios(j))
            Else
                Call vuelveUlla(Torneo_DestruccionEquipos(i).usuarios(j))
            End If
        Next j
    Next i

    Call inicializarTorneo
    Call QuitarNpcsSalaTorneo
    Exit Sub
chotserror:
    Call LogError("Error en MuereNpcTorre " & Err.number & " " & Err.Description)
End Sub

Public Sub QuitarNpcsSalaTorneo()
    Dim Y As Integer
    Dim X As Integer

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If MapData(Torneo_MAPATORNEO, X, Y).NpcIndex > 0 Then
                If Npclist(MapData(Torneo_MAPATORNEO, X, Y).NpcIndex).Numero > 500 Then Call QuitarNPC(MapData(Torneo_MAPATORNEO, X, Y).NpcIndex)
            End If
        Next X
    Next Y
End Sub
