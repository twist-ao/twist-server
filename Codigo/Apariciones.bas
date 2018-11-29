Attribute VB_Name = "Apariciones"
'TwistAO
'Copyright (C) 2018 Dalmasso, Juan Andres
'
'Modulo de Apariciones por dia
'Programado por CHOTS (Juan Andres Dalmasso)
'Desde Wellington, New Zealand
'30/08/2018

Public Const NPC_APARICION As Integer = 604
Public Const APARICION_MAPA As Byte = 55
Public Const APARICION_X As Byte = 41
Public Const APARICION_Y As Byte = 32
Public Const APARICION_HORA As Byte = 19
Public APARICION_APARECIDO As Boolean
Public APARICION_APARECERA As Boolean

Public Sub ApareceAparicion()
    Dim Pos As WorldPos
    Pos.Map = APARICION_MAPA
    Pos.X = APARICION_X
    Pos.Y = APARICION_Y
    Call SpawnNpc(NPC_APARICION, Pos, True, False)
    
    Call SendData(SendTarget.ToAll, 0, 0, "Z99")
    APARICION_APARECIDO = True
    Call LogGM("APARICIONES", "Aparece una Aparicion", False)
End Sub

Public Sub TimerMinutosAparicion()
    On Error GoTo chotserror

    Dim CurrentHour As Byte
    CurrentHour = val(DatePart("h", Now))

    If CurrentHour = APARICION_HORA And Not APARICION_APARECIDO And APARICION_APARECERA Then
        If RandomNumber(1, 3) = 1 Then
            Call ApareceAparicion
        End If
    End If

    If CurrentHour = (APARICION_HORA - 1) And Not APARICION_APARECERA Then
        APARICION_APARECERA = True
    End If

Exit Sub
chotserror:
    Call LogError("Error en TimerMinutosAparicion " & Err.number & " " & Err.Description)
End Sub

Public Sub MuereAparicion(ByVal npc As Integer)
    Call QuitarNPC(npc)
    APARICION_APARECERA = False
    APARICION_APARECIDO = False
End Sub

Public Sub AparicionTiraItems(ByRef npc As npc)
    On Error Resume Next

    Dim CurrentDay As Byte
    CurrentDay = val(DatePart("w", Now)) - 1

    If npc.Invent.NroItems > 0 And CurrentDay > 0 Then
        Dim MiObj As Obj

        If npc.Invent.Object(CurrentDay).ObjIndex > 0 Then
            If RandomNumber(1, 100) <= npc.Invent.Object(CurrentDay).ProbTirar Then
                Call LogGM("APARICIONES", "Aparicion tira item: " & npc.Invent.Object(CurrentDay).ObjIndex, False)
                MiObj.Amount = npc.Invent.Object(CurrentDay).Amount
                MiObj.ObjIndex = npc.Invent.Object(CurrentDay).ObjIndex
                Call TirarItemAlPiso(npc.Pos, MiObj)
            End If
        End If
    End If
End Sub
