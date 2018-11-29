Attribute VB_Name = "Invocaciones"
'Modulo de Invocaciones
'Programado por Juan Andrés Dalmasso (CHOTS)
'CHOTS_AO@HOTMAIL.COM
'Para Lapsus AO 2.0
'05/09/2010
'Reprogramado por CHOTS
'Para TwistAO
'06/08/2018

'CHOTS | Mapa donde está la invocación
Public Const INVOCACION_MAPA As Byte = 46

'CHOTS | Mapa donde se tienen q parar para que salga el chobi
Public Const INVOCACION_X1 As Byte = 48
Public Const INVOCACION_Y1 As Byte = 48
Public Const INVOCACION_X2 As Byte = 48
Public Const INVOCACION_Y2 As Byte = 52
Public Const INVOCACION_X3 As Byte = 52
Public Const INVOCACION_Y3 As Byte = 48
Public Const INVOCACION_X4 As Byte = 52
Public Const INVOCACION_Y4 As Byte = 52

'CHOTS | Datos del chobi
Public Const INVOCACION_NPC As Integer = 606
Public Const INVOCACION_RESPAWNX As Byte = 50
Public Const INVOCACION_RESPWANY As Byte = 50
Public Const AYUDANTE_NPC As Integer = 608
Public Const AYUDANTE_RESPAWNX As Byte = 50
Public Const AYUDANTE_RESPAWNY As Byte = 44

Public INVOCACION_INVOCADO As Boolean

Public Sub InvocarInvocacion()
    Dim Pos As WorldPos
    Pos.Map = INVOCACION_MAPA
    Pos.X = INVOCACION_RESPAWNX
    Pos.Y = INVOCACION_RESPWANY
    Call SpawnNpc(INVOCACION_NPC, Pos, True, False)
    
    'CHOTS | Invocamos al ayudante
    Pos.Map = INVOCACION_MAPA
    Pos.X = AYUDANTE_RESPAWNX
    Pos.Y = AYUDANTE_RESPAWNY
    Call SpawnNpc(AYUDANTE_NPC, Pos, True, False)
    
    Call SendData(SendTarget.ToAll, 0, 0, "Z78")
    INVOCACION_INVOCADO = True
End Sub

Public Sub MuereInvocacion(ByVal Npc As Integer)
    Call QuitarNPC(Npc)
    INVOCACION_INVOCADO = False
End Sub

Public Function PuedeInvocar() As Boolean

PuedeInvocar = False

If INVOCACION_INVOCADO Then Exit Function

If (MapData(INVOCACION_MAPA, INVOCACION_X1, INVOCACION_Y1).UserIndex <> 0) And (MapData(INVOCACION_MAPA, INVOCACION_X2, INVOCACION_Y2).UserIndex <> 0) And (MapData(INVOCACION_MAPA, INVOCACION_X3, INVOCACION_Y3).UserIndex <> 0) And (MapData(INVOCACION_MAPA, INVOCACION_X4, INVOCACION_Y4).UserIndex <> 0) Then
    If UserList(MapData(INVOCACION_MAPA, INVOCACION_X1, INVOCACION_Y1).UserIndex).flags.Muerto = 0 And UserList(MapData(INVOCACION_MAPA, INVOCACION_X2, INVOCACION_Y2).UserIndex).flags.Muerto = 0 And UserList(MapData(INVOCACION_MAPA, INVOCACION_X3, INVOCACION_Y3).UserIndex).flags.Muerto = 0 And UserList(MapData(INVOCACION_MAPA, INVOCACION_X4, INVOCACION_Y4).UserIndex).flags.Muerto = 0 Then
        'CHOTS | No se puede invocar si el ayudante esta vivo
        Dim Y As Integer
        Dim X As Integer
        For Y = YMinMapSize To YMaxMapSize
            For X = XMinMapSize To XMaxMapSize
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(INVOCACION_MAPA, X, Y).NpcIndex > 0 And Npclist(MapData(INVOCACION_MAPA, X, Y).NpcIndex).Numero = AYUDANTE_NPC Then Exit Function
                End If
            Next X
        Next Y
        
        PuedeInvocar = True
        Exit Function
    End If
End If

End Function
