Attribute VB_Name = "Admin"
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type
Public Apuestas As tAPuestas

Public npcs As Long
Public DebugSocket As Boolean

Public Horas As Long
Public Dias As Long
Public MinsRunning As Long

Public ReiniciarServer As Long

Public tInicioServer As Long
Public EstadisticasWeb As New clsEstadisticasIPC

Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public Intervalohambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloParalizadoNpc As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloDroga As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloUserPuedeAtacar As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]
Public IntervaloUserPuedeUsar As Long
Public IntervaloFlechasCazadores As Long

Public MinutosWs As Long
Public MinutosParaWs As Long
Public MinutosParaTorneo As Long
Public MinutosGrabar As Long
Public MinutosParaGrabar As Long
Public Puerto As Integer

Public MAXPASOS As Long

Public BootDelBackUp As Byte
Public DeNoche As Boolean

Public IpList As New Collection
Public ClientsCommandsQueue As Byte

Public Type TCPESStats
    BytesEnviados As Double
    BytesRecibidos As Double
    BytesEnviadosXSEG As Long
    BytesRecibidosXSEG As Long
    BytesEnviadosXSEGMax As Long
    BytesRecibidosXSEGMax As Long
    BytesEnviadosXSEGCuando As Date
    BytesRecibidosXSEGCuando As Date
End Type

Public TCPESStats As TCPESStats

'Public ResetThread As New clsThreading

Function VersionOK(ByVal Ver As String) As Boolean
VersionOK = (Ver = ULTIMAVERSION)
End Function

Public Function VersionesActuales(ByVal v1 As Integer, ByVal v2 As Integer, ByVal v3 As Integer, ByVal v4 As Integer, ByVal v5 As Integer, ByVal v6 As Integer, ByVal v7 As Integer) As Boolean
Dim rv As Boolean
Dim i As Integer
Dim f As String

rv = val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "GRAFICOS")) = v1
rv = rv And val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "WAVS")) = v2
rv = rv And val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "MIDIS")) = v3
rv = rv And val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "INIT")) = v4
rv = rv And val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "MAPAS")) = v5
rv = rv And val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "AOEXE")) = v6
rv = rv And val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "EXTRAS")) = v7
VersionesActuales = rv

End Function


Public Function ValidarLoginMSG(ByVal n As Integer) As Integer
On Error Resume Next
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(n)
AuxInteger2 = SDM(n)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function


Sub ReSpawnOrigPosNpcs()
On Error Resume Next

Dim i As Integer
Dim MiNPC As Npc
   
For i = 1 To LastNPC
   'OJO
   If Npclist(i).flags.NPCActive Then
        
        If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
        End If
        
        'tildada por sugerencia de yind
        'If Npclist(i).Contadores.TiempoExistencia > 0 Then
        '        Call MuereNpc(i, 0)
        'End If
   End If
   
Next i

End Sub

Sub WorldSave()
On Error Resume Next

Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> Iniciando WorldSave" & FONTTYPE_SERVER)

#If SeguridadAlkon Then
    Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
#End If

Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales

Call SendData(SendTarget.ToAll, 0, 0, ServerPackages.dialogo & "Servidor> WorldSave ha conclu�do" & FONTTYPE_SERVER)

End Sub

Public Sub PurgarPenas()
Dim i As Integer
For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
    
        If UserList(i).Counters.Pena > 0 Then
                
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call SendData(SendTarget.ToIndex, i, 0, ServerPackages.dialogo & "Has sido liberado!" & FONTTYPE_INFO)
                End If
                
        End If
        
    End If
Next i
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = "")
        
        UserList(UserIndex).Counters.Pena = Minutos
       
        
        Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
        
        If GmName = "" Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & "Has sido encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, ServerPackages.dialogo & GmName & " te ha encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & FONTTYPE_INFO)
        End If
        
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
On Error Resume Next
If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
    Kill CharPath & UCase$(UserName) & ".chr"
End If
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean

BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1)

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean

PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)

End Function

Public Function UnBan(ByVal Name As String) As Boolean
'Unban the character
Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")

'Remove it from the banned people database
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")
End Function
Public Function MD5okey(ByVal MD5Formateado As String, ByVal UserIndex As Integer) As Boolean
On Error GoTo errhandler
If Len(MD5Formateado) <> 32 Then
    MD5okey = False
    Exit Function
End If
Dim Codigo As Long
Codigo = (((((CLng(UserList(UserIndex).RandomCode) + 235) * 14) / 2) * 4) - 211)

If MD5String(CStr(Codigo)) = MD5Formateado Then
    MD5okey = True
Else
    MD5okey = False
End If
Exit Function
errhandler:
Debug.Print Err.Description & " " & Err.number
End Function

Public Function MD5ok(ByVal MD5Formateado As String) As Boolean
Dim i As Integer

If MD5ClientesActivado = 1 Then
    For i = 0 To UBound(MD5s)
        If (MD5Formateado = MD5s(i)) Then
            MD5ok = True
            Exit Function
        End If
    Next i
    MD5ok = False
Else
    MD5ok = True
End If

End Function

Public Sub MD5sCarga()
Dim LoopC As Integer

MD5ClientesActivado = val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))

If MD5ClientesActivado = 1 Then
    ReDim MD5s(val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))
    For LoopC = 0 To UBound(MD5s)
        MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
    Next LoopC
End If

End Sub

Public Sub BanIpAgrega(ByVal ip As String)
BanIps.Add ip

Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
Dim Dale As Boolean
Dim LoopC As Long

Dale = True
LoopC = 1
Do While LoopC <= BanIps.Count And Dale
    Dale = (BanIps.item(LoopC) <> ip)
    LoopC = LoopC + 1
Loop

If Dale Then
    BanIpBuscar = 0
Else
    BanIpBuscar = LoopC - 1
End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean

On Error Resume Next

Dim n As Long

n = BanIpBuscar(ip)
If n > 0 Then
    BanIps.Remove n
    BanIpGuardar
    BanIpQuita = True
Else
    BanIpQuita = False
End If

End Function

Public Sub BanIpGuardar()
Dim ArchivoBanIp As String
Dim ArchN As Long
Dim LoopC As Long

ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

ArchN = FreeFile()
Open ArchivoBanIp For Output As #ArchN

For LoopC = 1 To BanIps.Count
    Print #ArchN, BanIps.item(LoopC)
Next LoopC

Close #ArchN

End Sub

Public Sub BanIpCargar()
Dim ArchN As Long
Dim Tmp As String
Dim ArchivoBanIp As String

ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

Do While BanIps.Count > 0
    BanIps.Remove 1
Loop

ArchN = FreeFile()
Open ArchivoBanIp For Input As #ArchN

Do While Not EOF(ArchN)
    Line Input #ArchN, Tmp
    BanIps.Add Tmp
Loop

Close #ArchN

End Sub

Public Sub PalabrasInvalidasCargar()
Dim ArchN As Long
Dim Tmp As String
Dim ArchivoPalabrasInvalidas As String

ArchivoPalabrasInvalidas = App.Path & "\Dat\PalabrasInvalidas.dat"

Do While PalabrasInvalidas.Count > 0
    PalabrasInvalidas.Remove 1
Loop

ArchN = FreeFile()
Open ArchivoPalabrasInvalidas For Input As #ArchN

Do While Not EOF(ArchN)
    Line Input #ArchN, Tmp
    PalabrasInvalidas.Add Tmp
Loop

Close #ArchN

End Sub


Public Function UserDarPrivilegioLevel(ByVal Name As String) As Long
If EsDios(Name) Then
    UserDarPrivilegioLevel = 3
ElseIf EsSemiDios(Name) Then
    UserDarPrivilegioLevel = 2
ElseIf EsConsejero(Name) Then
    UserDarPrivilegioLevel = 1
Else
    UserDarPrivilegioLevel = 0
End If
End Function
