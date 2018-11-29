Attribute VB_Name = "Ranking"
'Modulo de Ranking
'Creado por Juan Andrés Dalmasso(CHOTS)
'Para Lapsus AO
'CHOTS_AO@HOTMAIL.COM
'Modificado para Lapsus 3 por CHOTS
'17/10/2011
'Modificado para TwistAO 2018 por CHOTS
'01/08/2018

Public Ranking_Trofeos_Nick As String
Public Ranking_Trofeos_Cant As Integer

Public Type tRanking_Matados
    nombre As String
    Ciudadanos As Integer
    Criminales As Integer
End Type

Public Type tRanking_Torneos
    nombre As String
    Cantidad As Integer
End Type

Public Type tRanking_Duelos
    nombre As String
    Ganados As Integer
    Perdidos As Integer
End Type

Public Ranking_Torneos() As tRanking_Torneos

Public Ranking_Matados() As tRanking_Matados

Public Ranking_Duelos() As tRanking_Duelos

Sub GuardarRanking()
'CHOTS | Guarda los datos
Dim n As Integer
Dim i As Byte
n = FreeFile
Open (IniPath & "RANKING.INI") For Output As #n
Print #n, "[TROFEOS]"
Print #n, "Nombre=" & Ranking_Trofeos_Nick
Print #n, "Cantidad=" & Ranking_Trofeos_Cant
Print #n, ""

Print #n, "[MATADOS]"
For i = 1 To 10
    Print #n, "Nombre" & i & "=" & Ranking_Matados(i).nombre
    Print #n, "Ciudadanos" & i & "=" & Ranking_Matados(i).Ciudadanos
    Print #n, "Criminales" & i & "=" & Ranking_Matados(i).Criminales
Next i
Print #n, ""

Print #n, "[TORNEOS]"
For i = 1 To Torneo_TIPOTORNEOS
    Print #n, "Nombre" & i & "=" & Ranking_Torneos(i).nombre
    Print #n, "Cantidad" & i & "=" & Ranking_Torneos(i).Cantidad
Next i
Print #n, ""

Print #n, "[DUELOS]"
For i = 1 To 10
    Print #n, "Nombre" & i & "=" & Ranking_Duelos(i).nombre
    Print #n, "Ganados" & i & "=" & Ranking_Duelos(i).Ganados
    Print #n, "Perdidos" & i & "=" & Ranking_Duelos(i).Perdidos
Next i
Print #n, ""

Close #n

End Sub

Sub CargarRanking()
Dim i As Byte
'CHOTS | Carga los datos
Ranking_Trofeos_Nick = GetVar(IniPath & "RANKING.INI", "TROFEOS", "Nombre")
Ranking_Trofeos_Cant = val(GetVar(IniPath & "RANKING.INI", "TROFEOS", "Cantidad"))

ReDim Ranking_Matados(1 To 10) As tRanking_Matados
For i = 1 To 10
    Dim usuario As tRanking_Matados
    usuario.nombre = GetVar(IniPath & "RANKING.INI", "MATADOS", "Nombre" & i)
    usuario.Ciudadanos = GetVar(IniPath & "RANKING.INI", "MATADOS", "Ciudadanos" & i)
    usuario.Criminales = GetVar(IniPath & "RANKING.INI", "MATADOS", "Criminales" & i)
    Ranking_Matados(i) = usuario
Next i

ReDim Ranking_Torneos(1 To Torneo_TIPOTORNEOS) As tRanking_Torneos
For i = 1 To Torneo_TIPOTORNEOS
    Dim usuario1 As tRanking_Torneos
    usuario1.nombre = GetVar(IniPath & "RANKING.INI", "TORNEOS", "Nombre" & i)
    usuario1.Cantidad = GetVar(IniPath & "RANKING.INI", "TORNEOS", "Cantidad" & i)
    Ranking_Torneos(i) = usuario1
Next i

ReDim Ranking_Duelos(1 To 10) As tRanking_Duelos
For i = 1 To 10
    Dim usuario2 As tRanking_Duelos
    usuario2.nombre = GetVar(IniPath & "RANKING.INI", "DUELOS", "Nombre" & i)
    usuario2.Ganados = GetVar(IniPath & "RANKING.INI", "DUELOS", "Ganados" & i)
    usuario2.Perdidos = GetVar(IniPath & "RANKING.INI", "DUELOS", "Perdidos" & i)
    Ranking_Duelos(i) = usuario2
Next i

End Sub

Sub ActualizarRanking(ByVal UserIndex As Integer, ByVal Ranking As Byte)
'CHOTS | Tipos de Ranking
'CHOTS | 1: Trofeos
'CHOTS | 2: Frags
'CHOTS | 3: Matados
'CHOTS | 4: Duelos

If UserList(UserIndex).flags.Privilegios <> PlayerType.User Then Exit Sub

Select Case Ranking
    Case 1
        If UserList(UserIndex).Stats.TrofOro > Ranking_Trofeos_Cant Then
            Ranking_Trofeos_Nick = UserList(UserIndex).Name
            Ranking_Trofeos_Cant = UserList(UserIndex).Stats.TrofOro
        End If
        Exit Sub
        
    Case 3
        Call ActualizarRankingMatados(UserIndex)
        Exit Sub

    Case 4
        Call ActualizarRankingDuelos(UserIndex)
        Exit Sub
        
End Select

End Sub

Sub ActualizarRankingTorneos(ByVal UserIndex As Integer, ByVal TipoTorneo As Byte)
    If UserList(UserIndex).Stats.TorneosAuto(TipoTorneo) > Ranking_Torneos(TipoTorneo).Cantidad Then
        Ranking_Torneos(TipoTorneo).nombre = UserList(UserIndex).Name
        Ranking_Torneos(TipoTorneo).Cantidad = UserList(UserIndex).Stats.TorneosAuto(TipoTorneo)
    End If
End Sub


Public Sub ActualizarRankingMatados(ByVal UserIndex As Integer)
    Dim usuario As tRanking_Matados
    Dim posicion As Byte
    Dim totalMatados As Long
    totalMatados = UserList(UserIndex).Faccion.CiudadanosMatados + UserList(UserIndex).Faccion.CriminalesMatados

    If Not estaEnRankingMatados(UserList(UserIndex).Name) Then

        If totalMatados > (Ranking_Matados(10).Ciudadanos + Ranking_Matados(10).Criminales) Then
            usuario.nombre = UserList(UserIndex).Name
            usuario.Ciudadanos = UserList(UserIndex).Faccion.CiudadanosMatados
            usuario.Criminales = UserList(UserIndex).Faccion.CriminalesMatados
            Ranking_Matados(10) = usuario
            Call ordenarRankingMatados
        End If

    Else
        posicion = getPosRankingMatados(UserList(UserIndex).Name)
        Ranking_Matados(posicion).Ciudadanos = UserList(UserIndex).Faccion.CiudadanosMatados
        Ranking_Matados(posicion).Criminales = UserList(UserIndex).Faccion.CriminalesMatados
        Call ordenarRankingMatados
    End If
    
End Sub

Public Sub ActualizarRankingDuelos(ByVal UserIndex As Integer)
    Dim usuario As tRanking_Duelos
    Dim posicion As Byte
    Dim totalPuntos As Integer
    totalPuntos = UserList(UserIndex).Stats.DuelosGanados - UserList(UserIndex).Stats.DuelosPerdidos

    If Not estaEnRankingDuelos(UserList(UserIndex).Name) Then

        If totalPuntos > (Ranking_Duelos(10).Ganados - Ranking_Duelos(10).Perdidos) Then
            usuario.nombre = UserList(UserIndex).Name
            usuario.Ganados = UserList(UserIndex).Stats.DuelosGanados
            usuario.Perdidos = UserList(UserIndex).Stats.DuelosPerdidos
            Ranking_Duelos(10) = usuario
            Call ordenarRankingDuelos
        End If
    Else
        posicion = getPosRankingDuelos(UserList(UserIndex).Name)
        Ranking_Duelos(posicion).Ganados = UserList(UserIndex).Stats.DuelosGanados
        Ranking_Duelos(posicion).Perdidos = UserList(UserIndex).Stats.DuelosPerdidos
        Call ordenarRankingDuelos
    End If
    
End Sub

Private Function estaEnRankingMatados(ByVal nick As String) As Boolean
Dim i As Byte
estaEnRankingMatados = False

For i = 1 To 10
    If UCase$(Ranking_Matados(i).nombre) = UCase$(nick) Then
        estaEnRankingMatados = True
        Exit Function
    End If
Next i
End Function

Private Function getPosRankingMatados(ByVal nick As String) As Byte
Dim i As Byte
getPosRankingMatados = 0
For i = 1 To 10
    If UCase$(Ranking_Matados(i).nombre) = UCase$(nick) Then
        getPosRankingMatados = i
        Exit Function
    End If
Next i
End Function

Private Sub ordenarRankingMatados()
Dim i As Byte
Dim j As Byte
Dim aux As tRanking_Matados

For i = 1 To 10
    For j = (i + 1) To 10
        If (Ranking_Matados(i).Criminales + Ranking_Matados(i).Ciudadanos) < (Ranking_Matados(j).Criminales + Ranking_Matados(j).Ciudadanos) Then
            aux = Ranking_Matados(i)
            Ranking_Matados(i) = Ranking_Matados(j)
            Ranking_Matados(j) = aux
        End If
    Next j
Next i
            
End Sub

Private Function estaEnRankingDuelos(ByVal nick As String) As Boolean
Dim i As Byte
estaEnRankingDuelos = False

For i = 1 To 10
    If UCase$(Ranking_Duelos(i).nombre) = UCase$(nick) Then
        estaEnRankingDuelos = True
        Exit Function
    End If
Next i
End Function

Private Function getPosRankingDuelos(ByVal nick As String) As Byte
Dim i As Byte
getPosRankingDuelos = 0
For i = 1 To 10
    If UCase$(Ranking_Duelos(i).nombre) = UCase$(nick) Then
        getPosRankingDuelos = i
        Exit Function
    End If
Next i
End Function

Private Sub ordenarRankingDuelos()
Dim i As Byte
Dim j As Byte
Dim aux As tRanking_Duelos

For i = 1 To 10
    For j = (i + 1) To 10
        If (Ranking_Duelos(i).Ganados - Ranking_Duelos(i).Perdidos) < (Ranking_Duelos(j).Ganados - Ranking_Duelos(j).Perdidos) Then
            aux = Ranking_Duelos(i)
            Ranking_Duelos(i) = Ranking_Duelos(j)
            Ranking_Duelos(j) = aux
        End If
    Next j
Next i
            
End Sub

'CHOTS | Sistema de Ranking
Sub ActualizarWebUsuarios(Optional ByVal number As Integer = -1)
On Error GoTo chotserror
Dim baseUrl As String
baseUrl = "http://www.twistao.com/u1p2d3a4te.php?token=" & SecurityParameters.webToken & "&param="
Dim enviar As String

enviar = "1@" & IIf(number >= 0, number, NumUsers) & "@" & recordusuarios

frmMain.InetUsers.Execute baseUrl & enviar, "GET"
Exit Sub
chotserror:
    Call LogError("Error en ActualizarWebUsuarios " & Err.number & " " & Err.Description)
End Sub

Sub ActualizarWeb()
On Error GoTo chotserror
'CHOTS | Envía la web

Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Actualizando Web..." & FONTTYPE_SERVER)

Dim i, j, aux As Integer
Dim guerra1, guerra2, guerra3, guerra4, guerra5, guerra6, guerra7, guerra8, guerra9, guerra10 As String
Dim guerrasGanadas1, guerrasGanadas2, guerrasGanadas3, guerrasGanadas4, guerrasGanadas5, guerrasGanadas6, guerrasGanadas7, guerrasGanadas8, guerrasGanadas9, guerrasGanadas10 As Integer
Dim guerrasPerdidas1, guerrasPerdidas2, guerrasPerdidas3, guerrasPerdidas4, guerrasPerdidas5, guerrasPerdidas6, guerrasPerdidas7, guerrasPerdidas8, guerrasPerdidas9, guerrasPerdidas10 As Integer
Dim enviar As String
Dim baseUrl As String
baseUrl = "http://www.twistao.com/u1p2d3a4te.php?token=" & SecurityParameters.webToken & "&param="

If CANTIDADDECLANES < 10 Then

    guerra1 = "Lapsus Corp"
    guerra2 = "Lapsus Corp"
    guerra3 = "Lapsus Corp"
    guerra4 = "Lapsus Corp"
    guerra5 = "Lapsus Corp"
    guerra6 = "Lapsus Corp"
    guerra7 = "Lapsus Corp"
    guerra8 = "Lapsus Corp"
    guerra9 = "Lapsus Corp"
    guerra10 = "Lapsus Corp"

    guerrasGanadas1 = 0
    guerrasGanadas2 = 0
    guerrasGanadas3 = 0
    guerrasGanadas4 = 0
    guerrasGanadas5 = 0
    guerrasGanadas6 = 0
    guerrasGanadas7 = 0
    guerrasGanadas8 = 0
    guerrasGanadas9 = 0
    guerrasGanadas10 = 0

    guerrasPerdidas1 = 0
    guerrasPerdidas2 = 0
    guerrasPerdidas3 = 0
    guerrasPerdidas4 = 0
    guerrasPerdidas5 = 0
    guerrasPerdidas6 = 0
    guerrasPerdidas7 = 0
    guerrasPerdidas8 = 0
    guerrasPerdidas9 = 0
    guerrasPerdidas10 = 0

Else

    'CHOTS | Ranking de clanes por guerras

    ReDim vecClan(1 To CANTIDADDECLANES) As Long
    For i = 1 To CANTIDADDECLANES
        vecClan(i) = i
    Next i

    For i = 1 To (CANTIDADDECLANES - 1)
        For j = (i + 1) To CANTIDADDECLANES
            If (Guilds(vecClan(i)).GetGuerrasGanadas() - Guilds(vecClan(i)).GetGuerrasPerdidas()) < (Guilds(vecClan(j)).GetGuerrasGanadas() - Guilds(vecClan(j)).GetGuerrasPerdidas()) Then
                aux = vecClan(i)
                vecClan(i) = vecClan(j)
                vecClan(j) = aux
            End If
        Next j
    Next i

    guerra1 = Guilds(vecClan(1)).GuildName
    guerra2 = Guilds(vecClan(2)).GuildName
    guerra3 = Guilds(vecClan(3)).GuildName
    guerra4 = Guilds(vecClan(4)).GuildName
    guerra5 = Guilds(vecClan(5)).GuildName
    guerra6 = Guilds(vecClan(6)).GuildName
    guerra7 = Guilds(vecClan(7)).GuildName
    guerra8 = Guilds(vecClan(8)).GuildName
    guerra9 = Guilds(vecClan(9)).GuildName
    guerra10 = Guilds(vecClan(10)).GuildName

    guerrasGanadas1 = Guilds(vecClan(1)).GetGuerrasGanadas()
    guerrasGanadas2 = Guilds(vecClan(2)).GetGuerrasGanadas()
    guerrasGanadas3 = Guilds(vecClan(3)).GetGuerrasGanadas()
    guerrasGanadas4 = Guilds(vecClan(4)).GetGuerrasGanadas()
    guerrasGanadas5 = Guilds(vecClan(5)).GetGuerrasGanadas()
    guerrasGanadas6 = Guilds(vecClan(6)).GetGuerrasGanadas()
    guerrasGanadas7 = Guilds(vecClan(7)).GetGuerrasGanadas()
    guerrasGanadas8 = Guilds(vecClan(8)).GetGuerrasGanadas()
    guerrasGanadas9 = Guilds(vecClan(9)).GetGuerrasGanadas()
    guerrasGanadas10 = Guilds(vecClan(10)).GetGuerrasGanadas()

    guerrasPerdidas1 = Guilds(vecClan(1)).GetGuerrasPerdidas()
    guerrasPerdidas2 = Guilds(vecClan(2)).GetGuerrasPerdidas()
    guerrasPerdidas3 = Guilds(vecClan(3)).GetGuerrasPerdidas()
    guerrasPerdidas4 = Guilds(vecClan(4)).GetGuerrasPerdidas()
    guerrasPerdidas5 = Guilds(vecClan(5)).GetGuerrasPerdidas()
    guerrasPerdidas6 = Guilds(vecClan(6)).GetGuerrasPerdidas()
    guerrasPerdidas7 = Guilds(vecClan(7)).GetGuerrasPerdidas()
    guerrasPerdidas8 = Guilds(vecClan(8)).GetGuerrasPerdidas()
    guerrasPerdidas9 = Guilds(vecClan(9)).GetGuerrasPerdidas()
    guerrasPerdidas10 = Guilds(vecClan(10)).GetGuerrasPerdidas()
    
End If

'CHOTS | Ranking Torneos
enviar = "2"
For i = 1 To Torneo_TIPOTORNEOS
    enviar = enviar & "@" & Ranking_Torneos(i).nombre & "@" & Ranking_Torneos(i).Cantidad
Next i
If frmMain.InetRanking.StillExecuting Then
    Call frmMain.InetRanking.Cancel
End If
frmMain.InetRanking.Execute baseUrl & enviar, "GET"

'CHOTS | Ranking Users. Por que carajo no hice un FOR aca? Too late darling
enviar = "3@" & Ranking_Matados(1).nombre & "@" & Ranking_Matados(1).Ciudadanos & "@" & Ranking_Matados(1).Criminales & "@" & Ranking_Matados(2).nombre & "@" & Ranking_Matados(2).Ciudadanos & "@" & Ranking_Matados(2).Criminales & "@" & Ranking_Matados(3).nombre & "@" & Ranking_Matados(3).Ciudadanos & "@" & Ranking_Matados(3).Criminales & "@" & Ranking_Matados(4).nombre & "@" & Ranking_Matados(4).Ciudadanos & "@" & Ranking_Matados(4).Criminales & "@" & Ranking_Matados(5).nombre & "@" & Ranking_Matados(5).Ciudadanos & "@" & Ranking_Matados(5).Criminales & "@" & Ranking_Matados(6).nombre & "@" & Ranking_Matados(6).Ciudadanos & "@" & Ranking_Matados(6).Criminales & "@" & Ranking_Matados(7).nombre & "@" & Ranking_Matados(7).Ciudadanos & "@" & Ranking_Matados(7).Criminales & "@" & Ranking_Matados(8).nombre & "@" & Ranking_Matados(8).Ciudadanos & "@" & Ranking_Matados(8).Criminales & "@" & Ranking_Matados(9).nombre & "@" & Ranking_Matados(9).Ciudadanos & "@" & Ranking_Matados(9).Criminales & "@" & _
Ranking_Matados(10).nombre & "@" & Ranking_Matados(10).Ciudadanos & "@" & Ranking_Matados(10).Criminales
If frmMain.InetRankingUsers.StillExecuting Then
    Call frmMain.InetRankingUsers.Cancel
End If
frmMain.InetRankingUsers.Execute baseUrl & enviar, "GET"

'CHOTS | Ranking Duelos. Por que carajo no hice un FOR aca? Too late darling
enviar = "5@" & Ranking_Duelos(1).nombre & "@" & Ranking_Duelos(1).Ganados & "@" & Ranking_Duelos(1).Perdidos & "@" & Ranking_Duelos(2).nombre & "@" & Ranking_Duelos(2).Ganados & "@" & Ranking_Duelos(2).Perdidos & "@" & Ranking_Duelos(3).nombre & "@" & Ranking_Duelos(3).Ganados & "@" & Ranking_Duelos(3).Perdidos & "@" & Ranking_Duelos(4).nombre & "@" & Ranking_Duelos(4).Ganados & "@" & Ranking_Duelos(4).Perdidos & "@" & Ranking_Duelos(5).nombre & "@" & Ranking_Duelos(5).Ganados & "@" & Ranking_Duelos(5).Perdidos & "@" & Ranking_Duelos(6).nombre & "@" & Ranking_Duelos(6).Ganados & "@" & Ranking_Duelos(6).Perdidos & "@" & Ranking_Duelos(7).nombre & "@" & Ranking_Duelos(7).Ganados & "@" & Ranking_Duelos(7).Perdidos & "@" & Ranking_Duelos(8).nombre & "@" & Ranking_Duelos(8).Ganados & "@" & Ranking_Duelos(8).Perdidos & "@" & Ranking_Duelos(9).nombre & "@" & Ranking_Duelos(9).Ganados & "@" & Ranking_Duelos(9).Perdidos & "@" & _
Ranking_Duelos(10).nombre & "@" & Ranking_Duelos(10).Ganados & "@" & Ranking_Duelos(10).Perdidos
If frmMain.InetClanes.StillExecuting Then
    Call frmMain.InetClanes.Cancel
End If
frmMain.InetClanes.Execute baseUrl & enviar, "GET"

'CHOTS | Ranking Guerras
enviar = "4@" & guerra1 & "@" & guerrasGanadas1 & "@" & guerrasPerdidas1 & "@" & guerra2 & "@" & guerrasGanadas2 & "@" & guerrasPerdidas2 & "@" & guerra3 & "@" & guerrasGanadas3 & "@" & guerrasPerdidas3 & "@" & guerra4 & "@" & guerrasGanadas4 & "@" & guerrasPerdidas4 & "@" & guerra5 & "@" & guerrasGanadas5 & "@" & guerrasPerdidas5 & "@" & guerra6 & "@" & guerrasGanadas6 & "@" & guerrasPerdidas6 & "@" & guerra7 & "@" & guerrasGanadas7 & "@" & guerrasPerdidas7 & "@" & guerra8 & "@" & guerrasGanadas8 & "@" & guerrasPerdidas8 & "@" & guerra9 & "@" & guerrasGanadas9 & "@" & guerrasPerdidas9 & "@" & guerra10 & "@" & guerrasGanadas10 & "@" & guerrasPerdidas10
If frmMain.InetGuerras.StillExecuting Then
    Call frmMain.InetGuerras.Cancel
End If
frmMain.InetGuerras.Execute baseUrl & enviar, "GET"

Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Web Actualizada..." & FONTTYPE_SERVER)
Exit Sub
chotserror:
    Call LogError("Error en ActualizarWeb " & Err.number & " " & Err.Description)
End Sub
