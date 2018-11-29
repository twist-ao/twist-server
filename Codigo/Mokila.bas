Attribute VB_Name = "Mokila"
'TwistAO
'Copyright (C) 2018 Dalmasso, Juan Andres
'
'Modulo de Mokila
'Ideado por Jadree (Adrian Schvartz) y CHOTS (Juan Andres Dalmasso)
'Programado por CHOTS (Juan Andres Dalmasso)
'Desde Wellington, New Zealand
'05/09/2018

Public Const ITEM_GEMA_AIRE As Integer = 604
Public Const ITEM_GEMA_AGUA As Integer = 604
Public Const ITEM_GEMA_FUEGO As Integer = 604
Public Const ITEM_GEMA_TIERRA As Integer = 604

Public Const MOKILA_MAPA as Byte = 72
Public Const MOKILA_AIRE_X As Byte = 48
Public Const MOKILA_AIRE_Y As Byte = 48
Public Const MOKILA_AGUA_X As Byte = 48
Public Const MOKILA_AGUA_Y As Byte = 52
Public Const MOKILA_FUEGO_X As Byte = 52
Public Const MOKILA_FUEGO_Y As Byte = 48
Public Const MOKILA_TIERRA_X As Byte = 52
Public Const MOKILA_TIERRA_Y As Byte = 52

Public Sub TimerMinutosMokila()
    On Error GoTo chotserror
    Dim cantidadGemas as Byte
    Dim tieneGemaAire, tieneGemaAgua, tieneGemaFuego, tieneGemaTierra as Boolean

    cantidadGemas = 0

    'CHOTS | Gema del Aire
    If MapData(MOKILA_MAPA, MOKILA_AIRE_X, MOKILA_AIRE_Y).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(i, X, Y).OBJInfo.ObjIndex).OBJType = otGema Then
            cantidadGemas = cantidadGemas + 1
            If MapData(i, X, Y).OBJInfo.ObjIndex = ITEM_GEMA_AIRE Then
                tieneGemaAire = True
            End If
        End If
    End If

    'CHOTS | Gema del Agua
    If MapData(MOKILA_MAPA, MOKILA_AGUA_X, MOKILA_AGUA_Y).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(i, X, Y).OBJInfo.ObjIndex).OBJType = otGema Then
            cantidadGemas = cantidadGemas + 1
            If MapData(i, X, Y).OBJInfo.ObjIndex = ITEM_GEMA_AGUA Then
                tieneGemaAgua = True
            End If
        End If
    End If

    'CHOTS | Gema del Fuego
    If MapData(MOKILA_MAPA, MOKILA_FUEGO_X, MOKILA_FUEGO_Y).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(i, X, Y).OBJInfo.ObjIndex).OBJType = otGema Then
            cantidadGemas = cantidadGemas + 1
            If MapData(i, X, Y).OBJInfo.ObjIndex = ITEM_GEMA_FUEGO Then
                tieneGemaFuego = True
            End If
        End If
    End If

    'CHOTS | Gema de Tierra
    If MapData(MOKILA_MAPA, MOKILA_TIERRA_X, MOKILA_TIERRA_Y).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(i, X, Y).OBJInfo.ObjIndex).OBJType = otGema Then
            cantidadGemas = cantidadGemas + 1
            If MapData(i, X, Y).OBJInfo.ObjIndex = ITEM_GEMA_TIERRA Then
                tieneGemaTierra = True
            End If
        End If
    End If

    If cantidadGemas = 4 Then
        If tieneGemaAire And tieneGemaAgua And tieneGemaFuego And tieneGemaTierra Then
            'Call InvocarMokila
        Else
            'Call EjecutarUsersMokila
        End If
    End If

Exit Sub
chotserror:
    Call LogError("Error en TimerMinutosMokila " & Err.number & " " & Err.Description)
End Sub

