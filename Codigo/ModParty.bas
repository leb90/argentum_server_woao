Attribute VB_Name = "ModParty"

Sub creaParty(ByVal Userindex As Integer, privada As Byte)

    On Error GoTo errhandler

    Dim n As Integer
    Dim encontrado As Boolean

    If UserList(Userindex).flags.party = False Then
        If numPartys >= MAXPARTYS Then
            Call SendData(ToIndex, Userindex, 0, "DD9A")
            'pluto:6.5
            Call LogParty(UserList(Userindex).Name & ": Intenta crear con Max")
            'Call SendData(ToIndex, UserIndex, 0, "||No puedes crear partys en este momento." & FONTTYPENAMES.FONTTYPE_INFO)
        Else
            encontrado = False
            n = 0

            Do While (n <= MAXPARTYS And encontrado <> True)
                n = n + 1

                If partylist(n).lider = 0 Then
                    encontrado = True

                End If

            Loop

            If encontrado = True Then
                UserList(Userindex).flags.partyNum = n
                UserList(Userindex).flags.party = True
                numPartys = numPartys + 1
                partylist(n).lider = Userindex
                partylist(n).expAc = 0
                partylist(n).reparto = 1
                partylist(n).privada = privada
                partylist(n).numMiembros = 1
                partylist(n).miembros(1).ID = Userindex
                partylist(n).miembros(1).privi = 100
                Call SendData(ToIndex, Userindex, 0, "DD10")
                'pluto:6.5
                Call LogParty(UserList(Userindex).Name & ": Crea N�: " & n & " Numpartys: " & numPartys)

                'Call SendData(ToIndex, UserIndex, 0, "||Has creado una party!" & FONTTYPENAMES.FONTTYPE_INFO)
            End If

        End If

    Else
        Call SendData(ToIndex, Userindex, 0, "DD11")
        'pluto:6.5
        Call LogParty(UserList(Userindex).Name & ": Intenta crear perteneciendo a otra.")

        'Call SendData(ToIndex, UserIndex, 0, "||Ya perteneces a una party!" & FONTTYPENAMES.FONTTYPE_INFO)
    End If

    Exit Sub
errhandler:
    Call LogError("Error en CreaPArty Nom:" & UserList(Userindex).Name & " UI:" & Userindex & " PRIV:" & privada & _
                  " N: " & Err.number & " D: " & Err.Description)
    '    Call LogError("Error en creaParty")

End Sub

Sub quitParty(ByVal Userindex As Integer)

    On Error GoTo errhandler

    Dim lpp As Byte
    Dim miembro As Integer
    Dim partyid As Integer

    If UserList(Userindex).flags.party = False Then
        Call SendData(ToIndex, Userindex, 0, "DD8A")
        'pluto:6.5
        Call LogParty(UserList(Userindex).Name & ": Intenta Cerrar party sin estar en party")
        'Call SendData(ToIndex, UserIndex, 0, "||No estas en ninguna party" & FONTTYPENAMES.FONTTYPE_INFO)
        Exit Sub

    End If

    If Userindex = partylist(UserList(Userindex).flags.partyNum).lider And UserList(Userindex).flags.party = True Then
        partyid = UserList(Userindex).flags.partyNum
        'pluto:6.5
        Call LogParty(UserList(Userindex).Name & ": Finaliza party")
        lpp = MAXMIEMBROS

        Do While lpp > 0

            If partylist(partyid).miembros(lpp).ID <> 0 Then
                miembro = partylist(partyid).miembros(lpp).ID
                Call quitUserParty(miembro)
                partylist(partyid).miembros(lpp).ID = 0

                'UserList(miembro).flags.party = False
                'Call SendData(ToIndex, miembro, 0, "||Party finalizada!" & FONTTYPENAMES.FONTTYPE_INFO)
            End If

            lpp = lpp - 1
        Loop
        Call SendData(toParty, Userindex, 0, "DD12")
        'Call SendData(toParty, miembro, 0, "||Party finalizada!" & FONTTYPENAMES.FONTTYPE_INFO)
        UserList(Userindex).flags.party = False
        partylist(partyid).lider = 0
        partylist(partyid).expAc = 0
        partylist(partyid).numMiembros = 0
        numPartys = numPartys - 1
        'pluto:6.5
        Call LogParty("NumeroPartys: " & numPartys)
    Else
        Call SendData(ToIndex, Userindex, 0, "DD13")
        'pluto:6.5
        Call LogParty(UserList(Userindex).Name & ": Intenta Finalizar party")

        'Call SendData(ToIndex, UserIndex, 0, "||Debes ser el lider de la party para poder finalizarla." & FONTTYPENAMES.FONTTYPE_INFO)
    End If

    Exit Sub
errhandler:
    'Call LogError("Error en quitaPArty")
    Call LogError("Error en quitaParty Nom:" & UserList(Userindex).Name & " UI:" & Userindex & " PID:" & partyid & _
                  " N: " & Err.number & " D: " & Err.Description)

End Sub

Sub addUserParty(ByVal Userindex As Integer, PartyIndex As Integer)

    On Error GoTo errhandler

    Dim n As Integer

    'pluto:6.5
    Call LogParty(UserList(Userindex).Name & ": a�adir a party n� " & PartyIndex)

    If partylist(PartyIndex).numMiembros >= MAXMIEMBROS Then
        Call SendData(ToIndex, partylist(PartyIndex).lider, 0, "||Party llena" & "�" & FontTypeNames.FONTTYPE_INFO)
        'pluto:6.5
        Call LogParty(UserList(Userindex).Name & ": no se a�ade por party llena.")
        Exit Sub
    Else

    End If

    'pluto:6.7-----------------
    For n = 1 To MaxUsers

        If UserList(n).flags.invitado = UserList(Userindex).Name Then UserList(n).flags.invitado = ""
    Next
    '------------------------
    n = 0

    Do While (n < MAXMIEMBROS)
        n = n + 1

        If partylist(PartyIndex).miembros(n).ID = 0 And UserList(Userindex).flags.party = False Then
            partylist(PartyIndex).miembros(n).ID = Userindex
            partylist(PartyIndex).miembros(n).privi = 0
            partylist(PartyIndex).numMiembros = partylist(PartyIndex).numMiembros + 1
            UserList(Userindex).flags.partyNum = PartyIndex
            UserList(Userindex).flags.party = True
            UserList(Userindex).flags.invitado = ""
            'pluto:6.5
            Call LogParty(UserList(Userindex).Name & ": a�adido ha Party n� " & PartyIndex & " en pos " & n)
            Call LogParty("Party n� " & PartyIndex & " Miembros: " & partylist(PartyIndex).numMiembros)

            Call sendMiembrosParty(Userindex)
            'pluto:6.3--------
            Call SendData(toParty, Userindex, 0, "DD14" & UserList(Userindex).Name)
            'Dim fp As Byte
            'Dim mie As String
            'For fp = 1 To MAXMIEMBROS
            'mie = mie + UserList(partylist(PartyIndex).miembros(fp).ID).Char.CharIndex & ","
            'Next
            'Call SendData(ToIndex, UserIndex, 0, "O6" & PartyIndex & "," & mie)
            '-----------------

            'Call SendData(toParty, partylist(partyindex).lider, 0, "||" & UserList(UserIndex).Name & " se ha incorporado a la party." & FONTTYPENAMES.FONTTYPE_INFO)
            'a�adir user ponemos reparto proporcional
            partylist(PartyIndex).reparto = 1

            If partylist(PartyIndex).reparto = 1 Then
                Call BalanceaPrivisLVL(PartyIndex)
            ElseIf partylist(PartyIndex).reparto = 2 Then
                Call SendData(ToIndex, Userindex, 0, "DD15")
                'Call SendData(ToIndex, partylist(partyindex).lider, 0, "||Modifica los privilegios para el nuevo usuario" & FONTTYPENAMES.FONTTYPE_INFO)
            ElseIf partylist(PartyIndex).reparto = 3 Then
                Call BalanceaPrivisMiembros(PartyIndex)

            End If

            Call sendPriviParty(Userindex)

        End If

        If partylist(PartyIndex).Solicitudes(n) = Userindex Then
            partylist(PartyIndex).Solicitudes(n) = 0
            partylist(PartyIndex).numSolicitudes = partylist(PartyIndex).numSolicitudes - 1
            'pluto:6.5
            Call LogParty(UserList(Userindex).Name & ": Borrado de solicitudes")
            Call LogParty("N� Party: " & PartyIndex & " Solicitudes: " & partylist(PartyIndex).numSolicitudes)

        End If

    Loop

    Exit Sub
errhandler:
    Call LogError("Error en addUserParty Nom:" & UserList(Userindex).Name & " UI:" & Userindex & " PID:" & PartyIndex _
                  & " N: " & Err.number & " D: " & Err.Description)

    'Call LogError("Error en addUserPArty")
End Sub

Sub addSoliParty(ByVal Userindex As Integer, PartyIndex As Integer)

    On Error GoTo errhandler

    Dim n As Integer
    Dim encontrado As Boolean
    'pluto:6.5
    Call LogParty(UserList(Userindex).Name & ": a�adir solicitud a la party n� " & PartyIndex)

    If UserList(Userindex).flags.party = False Then
        encontrado = False

        For n = 1 To MAXMIEMBROS

            If partylist(PartyIndex).Solicitudes(n) = Userindex Then
                Call SendData(ToIndex, Userindex, 0, "DD26")
                'pluto:6.5
                Call LogParty(UserList(Userindex).Name & ": no a�adida pq ya env�o antes.")

                Exit Sub

            End If

        Next
        n = 0

        If partylist(PartyIndex).numSolicitudes >= MAXMIEMBROS Then
            Call SendData(ToIndex, Userindex, 0, "DD16")
            'pluto:6.5
            Call LogParty(UserList(Userindex).Name & ": no a�adida por cola llena.")
            'Call SendData(ToIndex, UserIndex, 0, "||Cola de solicitudes llena, no puedes unirte en este momento." & FONTTYPENAMES.FONTTYPE_INFO)
        Else

            Do While (n < MAXMIEMBROS And encontrado <> True)
                n = n + 1

                If partylist(PartyIndex).Solicitudes(n) = 0 Then
                    encontrado = True

                End If

            Loop

            If encontrado = True Then
                ' UserList(UserIndex).flags.partyNum = PartyIndex
                partylist(PartyIndex).Solicitudes(n) = Userindex
                partylist(PartyIndex).numSolicitudes = partylist(PartyIndex).numSolicitudes + 1
                Call SendData(ToIndex, partylist(PartyIndex).lider, 0, "DD17" & UserList(Userindex).Name)
                'Call SendData(ToIndex, partylist(partyindex).lider, 0, "||" & UserList(UserIndex).Name & " solicita entrar en la party ." & FONTTYPENAMES.FONTTYPE_INFO)
                Call SendData(ToIndex, Userindex, 0, "DD18" & UserList(partylist(PartyIndex).lider).Name)
                'Call SendData(ToIndex, UserIndex, 0, "||Solicitud enviada a la party de " + UserList(partylist(partyindex).lider).Name + " ." & FONTTYPENAMES.FONTTYPE_INFO)
                'pluto:6.5
                Call LogParty(UserList(Userindex).Name & ": solicitud a�adida a la party n� " & PartyIndex & _
                              " en pos " & n)
                Call LogParty("Total solicitudes: " & partylist(PartyIndex).numSolicitudes)

            End If

        End If

    Else
        Call SendData(ToIndex, Userindex, 0, "DD11")
        'pluto:6.5
        Call LogParty(UserList(Userindex).Name & ": no a�adida pq ya pertenece a una party.")

        'Call SendData(ToIndex, UserIndex, 0, "||Ya perteneces a una party." & FONTTYPENAMES.FONTTYPE_INFO)
    End If

    Exit Sub
errhandler:
    Call LogError("Error en addSoliParty Nom:" & UserList(Userindex).Name & " UI:" & Userindex & " PID:" & PartyIndex _
                  & " N: " & Err.number & " D: " & Err.Description)

    'Call LogError("Error en addUserPArty")
End Sub

Sub quitSoliParty(ByVal Userindex As Integer, PartyIndex As Integer)

    On Error GoTo errhandler

    Dim n As Integer
    Dim encontrado As Boolean
    'pluto:6.5
    Call LogParty(UserList(Userindex).Name & ": quitar solicitud a la party n� " & PartyIndex)

    encontrado = False
    n = 0

    Do While (n < MAXMIEMBROS And encontrado <> True)
        n = n + 1

        If partylist(PartyIndex).Solicitudes(n) = Userindex Then
            encontrado = True

        End If

    Loop

    If encontrado = True Then
        partylist(PartyIndex).Solicitudes(n) = 0
        partylist(PartyIndex).numSolicitudes = partylist(PartyIndex).numSolicitudes - 1
        UserList(Userindex).flags.partyNum = 0
        UserList(Userindex).flags.party = False
        'pluto:6.5
        Call LogParty(UserList(Userindex).Name & ": solicitud quitada en pos " & n)
        Call LogParty("Party: " & PartyIndex & " Solicitudes: " & partylist(PartyIndex).numSolicitudes)

    Else
        Call LogParty(UserList(Userindex).Name & ": error quitar user no encontrado en party: " & PartyIndex)

        GoTo errhandler

    End If

    Exit Sub
errhandler:
    Call LogError("Error en quitSoliParty Nom:" & UserList(Userindex).Name & " UI:" & Userindex & " PID:" & _
                  PartyIndex & " N: " & Err.number & " D: " & Err.Description)
    'Call LogError("Error en quitSoliPArty")

End Sub

Sub quitUserParty(ByVal Userindex As Integer)

    On Error GoTo errhandler

    Dim n As Integer
    Dim encontrado As Boolean
    Dim PartyIndex As Integer

    'pluto:6.5
    Call LogParty(UserList(Userindex).Name & ": vamos a quitarlo de party")

    If Userindex = 0 Then Exit Sub
    If UserList(Userindex).flags.party = True Then
        If esLider(Userindex) = True And partylist(UserList(Userindex).flags.partyNum).numMiembros > 1 Then
            'pluto:6.5
            Call LogParty(UserList(Userindex).Name & ": no se quita pq es lider")

            Exit Sub

        End If

        PartyIndex = UserList(Userindex).flags.partyNum
        encontrado = False
        'pluto:6.5
        Call LogParty(UserList(Userindex).Name & ": est� en la party " & PartyIndex)

        'n = 1
        'Do While (n < MAXMIEMBROS)
        For n = 1 To MAXMIEMBROS

            If partylist(PartyIndex).miembros(n).ID = Userindex Then
                'Debug.Print UserList(UserIndex).Name
                partylist(PartyIndex).miembros(n).ID = 0
                partylist(PartyIndex).miembros(n).privi = 0
                'partylist(PartyIndex).numMiembros = partylist(PartyIndex).numMiembros - 1

                Call SendData(ToIndex, Userindex, 0, "DD19" & partylist(UserList(Userindex).flags.partyNum).expAc)
                Call SendData(ToIndex, Userindex, 0, "DD20")
                'Call SendData(ToIndex, UserIndex, 0, "||Has abandonado la party!" & FONTTYPENAMES.FONTTYPE_INFO)
                Call SendData(ToIndex, Userindex, 0, "W10,")
                'Call SendData(ToIndex, UserIndex, 0, "||Has ganado un total de " & partylist(UserList(UserIndex).flags.partyNum).expAc & " puntos de experiencia" & FONTTYPENAMES.FONTTYPE_INFO)
                'Call sendMiembrosParty(partylist(UserList(UserIndex).flags.partyNum).lider)
                'pluto:6.3---------
                Call SendData(toParty, Userindex, 0, "O5" & UserList(Userindex).Char.CharIndex)

                '-----------------
                'pluto:6.3 ponemos esto detras de enviar miembrosparty
                partylist(PartyIndex).numMiembros = partylist(PartyIndex).numMiembros - 1
                'pluto:6.5
                Call LogParty(UserList(Userindex).Name & ": quitado en pos " & n)
                Call LogParty("Miembros Party: " & partylist(PartyIndex).numMiembros)

                UserList(Userindex).flags.partyNum = 0
                UserList(Userindex).flags.party = False
                UserList(Userindex).flags.invitado = ""

                'pluto:6.7 a�ade reparto 2
                If partylist(PartyIndex).reparto = 1 Or partylist(PartyIndex).reparto = 2 Then
                    Call BalanceaPrivisLVL(PartyIndex)
                ElseIf partylist(PartyIndex).reparto = 3 Then
                    Call BalanceaPrivisMiembros(PartyIndex)

                End If

                'pluto:6.7-----------
                Call sendMiembrosParty(partylist(PartyIndex).lider)
                Call sendPriviParty(partylist(PartyIndex).lider)
                '----------------------

            End If

            'n = n + 1
        Next
    Else
        Call SendData(ToIndex, Userindex, 0, "DD8A")
        'pluto:6.5
        Call LogParty(UserList(Userindex).Name & " no est� en party")

        'Call SendData(ToIndex, UserIndex, 0, "||No estas en ninguna party" & FONTTYPENAMES.FONTTYPE_INFO)
    End If

    Exit Sub
errhandler:
    Call LogError("Error en quitUserParty Nom:" & UserList(Userindex).Name & " UI:" & Userindex & " PID:" & _
                  PartyIndex & "n: " & n & " N: " & Err.number & " D: " & Err.Description)

    'Call LogError("Error en quitUserPArty")
End Sub

Sub InvitaParty(indexAnfitrion As Integer, indexInvitado As Integer)

    On Error GoTo errhandler

    'pluto:6.5
    Call LogParty(UserList(indexAnfitrion).Name & " invita a " & UserList(indexInvitado).Name)

    If UserList(indexInvitado).flags.party = True Then
        Call SendData(ToIndex, indexAnfitrion, 0, "DD21" & UserList(indexInvitado).Name)
        'Call SendData(ToIndex, indexAnfitrion, 0, "||No puedes invitar a " & UserList(indexInvitado).Name & ", ya esta en una party." & FONTTYPENAMES.FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, indexAnfitrion, 0, "DD22" & UserList(indexInvitado).Name)
        'Call SendData(ToIndex, indexAnfitrion, 0, "||Has invitado a " & UserList(indexInvitado).Name & " a la party." & FONTTYPENAMES.FONTTYPE_INFO)
        Call SendData(ToIndex, indexInvitado, 0, "DD23" & UserList(indexAnfitrion).Name)
        'Call SendData(ToIndex, indexInvitado, 0, "||" & UserList(indexAnfitrion).Name & " te ha invitado a crear una party. Escribe /unirme para unirte" & FONTTYPENAMES.FONTTYPE_INFO)
        UserList(indexInvitado).flags.invitado = UserList(indexAnfitrion).Name
        'pluto:6.7
        UserList(indexAnfitrion).flags.invitado = ""

    End If

    Exit Sub
errhandler:
    Call LogError("Error en InvitaParty Anfitrion:" & UserList(indexAnfitrion).Name & " AnfiID:" & indexAnfitrion & _
                  " Invitado:" & UserList(indexInvitado).Name & " InviID:" & indexInvitado & " N: " & Err.number & " D: " & _
                  Err.Description)

End Sub

Function totalexpParty(ByVal PartyIndex As Integer) As Long

    On Error GoTo errhandler

    Dim n As Integer
    Dim total As Double
    total = 0

    For n = 1 To MAXMIEMBROS

        If partylist(PartyIndex).miembros(n).ID <> 0 Then
            total = total + UserList(partylist(PartyIndex).miembros(n).ID).Stats.ELV

        End If

    Next
    totalexpParty = total
    Exit Function
errhandler:
    Call LogError("Error en totalexpParty PID " & PartyIndex & " N: " & Err.number & " D: " & Err.Description)

End Function

Sub PartyReparteExp(NpcIndex As npc, Userindex As Integer)

    On Error GoTo errhandler

    Dim n As Byte
    Dim expIndi As Double
    Dim b As Long
    Dim aa As Integer
    Dim oo As Integer

    partylist(UserList(Userindex).flags.partyNum).expAc = partylist(UserList(Userindex).flags.partyNum).expAc + _
                                                          NpcIndex.GiveEXP

    For n = 1 To MAXMIEMBROS
        oo = partylist(UserList(Userindex).flags.partyNum).miembros(n).ID

        If oo <> 0 Then
            If UserList(oo).Pos.Map = NpcIndex.Pos.Map Then
                If UserList(oo).flags.Muerto = 0 Then
                    If UserList(oo).Bebe = 0 Then

                        b = partylist(UserList(Userindex).flags.partyNum).miembros(n).privi
                        expIndi = (NpcIndex.GiveEXP / 100) * b

                        If ServerPrimario = 1 Then
                            If UserList(oo).Remort > 0 Then
                                expIndi = expIndi * 1
                            Else
                                expIndi = expIndi * 1

                            End If

                        Else    'secundario Pluto:6.5

                            If UserList(Userindex).Remort > 0 Then
                                expIndi = expIndi * 1
                            Else
                                expIndi = expIndi * 1

                            End If

                        End If    'primario =1

                        If UserList(oo).flags.Montura > 0 Then



                            aa = Int(expIndi / 4000)    ' ORIGINAL: Int(expIndi / 1000)
                        Else
                            aa = 0

                        End If

                
                                        'IRON AO: Aumento de Experiencia
                        'Alas Angelicales y Rojas
                        If UserList(oo).Invent.AlaEqpObjIndex = 1375 Then
                            expIndi = Int(expIndi * 1.05)

                        End If

                        'Alas Vampirescas
                        If UserList(oo).Invent.AlaEqpObjIndex = 1376 Then
                            expIndi = Int(expIndi * 1.1)

                            '
                        End If

                        'Alas azules
                        If UserList(oo).Invent.AlaEqpObjIndex = 1377 Then
                            expIndi = Int(expIndi * 1.15)

                        End If

                        'Alas War
                        If UserList(oo).Invent.AlaEqpObjIndex = 1378 Then
                            expIndi = Int(expIndi * 1.2)

                        End If


                        'El user tiene montura (hay que repartir exp con ella)
                        If UserList(oo).flags.Montura > 0 And UserList(oo).flags.ClaseMontura > 0 Then

                            'a�ade topelevel
                            If PMascotas(UserList(oo).flags.ClaseMontura).TopeLevel > UserList(oo).Montura.Nivel( _
                               UserList(oo).flags.ClaseMontura) Then

                                'Comprobamos que no este bugueada
                                If UserList(oo).Montura.Elu(UserList(oo).flags.ClaseMontura) = 0 Then
                                    Call SendData(ToGM, 0, 0, "|| Matanpc Mascota Bugueada: " & UserList(oo).Name & _
                                                              "�" & FontTypeNames.FONTTYPE_COMERCIO)
                                    Call LogCasino("BUG MataNpcMASCOTAparty Serie: " & UserList(oo).Serie & " IP: " & _
                                                   UserList(oo).ip & " Nom: " & UserList(oo).Name)

                                End If

                                '----------------
                                'Le metemos la exp a la montura
                                Call AddtoVar(UserList(oo).Montura.exp(UserList(oo).flags.ClaseMontura), Int(expIndi _
                                                                                                             / 4000), MAXEXP)
                                Call CheckMonturaLevel(oo)

                            End If

                        End If    'topelevel

                        'expIndi = (NpcIndex.GiveEXP * (partylist(UserList(UserIndex).flags.partyNum).miembros(n).privi) / 100)
                        'expIndi = NpcIndex.GiveEXP * (partylist(UserList(UserIndex).flags.partyNum).miembros(n).privi) / 100)
                        'expIndi = expIndi \ 100

                        'pluto:6.3 AUMENTO EXP----------
                        'If ServerPrimario = 1 Then
                        'If UserList(oo).Remort > 0 Then
                        'expIndi = expIndi * 2
                        'GoTo Nomire
                        'End If

                        ' Select Case UserList(oo).Stats.ELV
                        ' Case Is < 50
                        ' expIndi = expIndi * 10
                        'Case 50 To 60
                        'expIndi = expIndi * 5
                        'Case Is > 60
                        'expIndi = expIndi * 3
                        'End Select
                        'Else 'secundario Pluto:6.5

                        ' If UserList(oo).Remort > 0 Then
                        ' expIndi = expIndi * 1
                        'GoTo Nomire
                        'End If

                        '       Select Case UserList(oo).Stats.ELV
                        '      Case Is < 30
                        '     expIndi = expIndi * 10
                        '    Case 30 To 40
                        '   expIndi = expIndi * 5
                        '  Case 41 To 50
                        ' expIndi = expIndi * 3
                        'Case Is > 50
                        'expIndi = expIndi * 2
                        'End Select

                        'End If 'primario =1
                        'If UserList(oo).Remort > 0 Then expIndi = expIndi * 2
                        'Debug.Print UserList(oo).Name
                        '-----------------------------
Nomire:
                        Call AddtoVar(UserList(oo).Stats.exp, expIndi, MAXEXP)
                        Call SendData(ToIndex, oo, 0, "V6" & expIndi & "," & aa)
                        Call CheckUserLevel(oo)
                        Call senduserstatsbox(oo)

                    End If

                End If

            End If

        End If

    Next
    Exit Sub
errhandler:
    Call LogError("Error en PartyReparteExp Nom:" & UserList(Userindex).Name & " UI: " & Userindex & " NPCID: " & _
                  NpcIndex.Name & " N: " & Err.number & " D: " & Err.Description)

End Sub

Function partyid(ByVal liderName As String) As Integer

    On Error GoTo errhandler

    Dim n As Integer
    Dim encontrado As Boolean
    encontrado = False
    n = 0

    Do While (n < numPartys And encontrado <> True)
        n = n + 1

        If UCase$(UserList(partylist(n).lider).Name) = UCase$(liderName) Then
            encontrado = True

        End If

    Loop

    If encontrado = True Then
        partyid = n

    End If

    Exit Function
errhandler:
    Call LogError("Error en partyid NomLider:" & liderName & " N: " & Err.number & " D: " & Err.Description)

End Function

Function esLider(ByVal Userindex As Integer) As Boolean

    On Error GoTo errhandler

    esLider = False

    If UserList(Userindex).flags.party = True Then
        If Userindex = partylist(UserList(Userindex).flags.partyNum).lider Then
            esLider = True

        End If

    End If

    Exit Function
errhandler:
    Call LogError("Error en esLider Nom:" & UserList(Userindex).Name & " UI:" & Userindex & " N: " & Err.number & _
                  " D: " & Err.Description)

End Function

Sub sendExpParty(exp As Long, Userindex As Integer)

    On Error GoTo errhandler

    Dim n As Integer

    For n = 1 To partylist(UserList(Userindex).flags.partyNum).numMiembros

        If partylist(UserList(Userindex).flags.partyNum).miembros(n).ID <> 0 Then
            If UserList(partylist(UserList(Userindex).flags.partyNum).miembros(n).ID).flags.Muerto = 0 Then
                Call SendData(ToIndex, partylist(UserList(Userindex).flags.partyNum).miembros(n).ID, 0, "V6" & (exp * _
                                                                                                                UserList(partylist(UserList(Userindex).flags.partyNum).miembros(n).ID).Stats.ELV) / _
                                                                                                                totalexpParty(UserList(Userindex).flags.partyNum) & ",")

            End If

        End If

    Next
    Exit Sub
errhandler:
    Call LogError("Error en sendExpParty Nom:" & UserList(Userindex).Name & " UI:" & Userindex & " exp:" & exp & _
                  " N: " & Err.number & " D: " & Err.Description)

End Sub

Sub sendMiembrosParty(Userindex As Integer)

    On Error GoTo errhandler

    Dim miempar$
    Dim npar As Byte

    If UserList(Userindex).flags.party = False Then Exit Sub
    npar = 1
    miempar$ = partylist(UserList(Userindex).flags.partyNum).numMiembros & ", "

    Do While (npar <= MAXMIEMBROS)

        If partylist(UserList(Userindex).flags.partyNum).miembros(npar).ID <> 0 Then
            'pluto:6.3-------
            Call SendData(ToIndex, partylist(UserList(Userindex).flags.partyNum).miembros(npar).ID, 0, "O4" & _
                                                                                                       UserList(Userindex).flags.partyNum)
            '-----------------
            miempar$ = miempar$ & UserList(partylist(UserList(Userindex).flags.partyNum).miembros(npar).ID).Name & _
                       "," & UserList(partylist(UserList(Userindex).flags.partyNum).miembros(npar).ID).Char.CharIndex & _
                       ","

        End If

        npar = npar + 1
    Loop
    npar = 1

    Do While (npar <= MAXMIEMBROS)

        If partylist(UserList(Userindex).flags.partyNum).miembros(npar).ID <> 0 Then
            Call SendData(ToIndex, partylist(UserList(Userindex).flags.partyNum).miembros(npar).ID, 0, "W1" & miempar$)

        End If

        npar = npar + 1
    Loop
    Exit Sub
errhandler:
    Call LogError("Error en sendMiembrosParty Nom:" & UserList(Userindex).Name & " UI:" & Userindex & " N: " & _
                  Err.number & " D: " & Err.Description)

End Sub

Sub sendPriviParty(Userindex As Integer)

    On Error GoTo errhandler

    Dim miempar$
    Dim npar As Byte

    If UserList(Userindex).flags.party = False Then Exit Sub
    npar = 1
    miempar$ = partylist(UserList(Userindex).flags.partyNum).numMiembros & ", "

    Do While (npar <= MAXMIEMBROS)

        If partylist(UserList(Userindex).flags.partyNum).miembros(npar).ID <> 0 Then
            miempar$ = miempar$ & partylist(UserList(Userindex).flags.partyNum).miembros(npar).privi & ", "

        End If

        npar = npar + 1
    Loop
    Call SendData(toParty, Userindex, 0, "W3" & miempar$)
    Exit Sub
errhandler:
    Call LogError("Error en sendPriviParty Nom:" & UserList(Userindex).Name & " UI:" & Userindex & " N: " & _
                  Err.number & " D: " & Err.Description)

End Sub

Sub sendSolicitudesParty(Userindex As Integer)

    On Error GoTo errhandler

    If esLider(Userindex) = True Then
        Dim miempar2$
        Dim npar As Byte
        npar = 1
        miempar2$ = partylist(UserList(Userindex).flags.partyNum).numSolicitudes & ", "

        Do While (npar <= MAXMIEMBROS)

            If partylist(UserList(Userindex).flags.partyNum).Solicitudes(npar) <> 0 Then
                miempar2$ = miempar2$ & UserList(partylist(UserList(Userindex).flags.partyNum).Solicitudes( _
                                                 npar)).Name & ", "

            End If

            npar = npar + 1
        Loop
        Call SendData(ToIndex, Userindex, 0, "W2" & miempar2$)

    End If

    Exit Sub
errhandler:
    Call LogError("Error en sendSolicitudesParty Nom:" & UserList(Userindex).Name & " UI:" & Userindex & " N: " & _
                  Err.number & " D: " & Err.Description)

End Sub

Sub resetParty(PartyIndex As Integer)

'completar
    On Error GoTo errhandler

    'pluto:6.5
    Call LogParty("Party: " & PartyIndex & " reseteada.")

    partylist(PartyIndex).expAc = 0
    partylist(PartyIndex).lider = 0
    partylist(PartyIndex).numMiembros = 0
    partylist(PartyIndex).numSolicitudes = 0
    partylist(PartyIndex).reparto = 1
    Dim lpp2 As Integer

    For lpp2 = 1 To MAXMIEMBROS
        partylist(partyidex).miembros(lpp2).ID = 0
        partylist(partyidex).miembros(lpp2).privi = 0
        partylist(partyidex).Solicitudes(lpp2) = 0
    Next
    numPartys = numPartys - 1
    'pluto:6.5
    Call LogParty("Numero de Partys: " & numPartys)
    Exit Sub
errhandler:
    Call LogError("Error en resetparty PID:" & PartyIndex & " N: " & Err.number & " D: " & Err.Description)

End Sub

Sub BalanceaPrivisLVL(PartyIndex As Integer)

    On Error GoTo errhandler

    Dim n As Integer

    For n = 1 To MAXMIEMBROS    'partylist(PartyIndex).numMiembros

        If partylist(PartyIndex).miembros(n).ID <> 0 Then
            partylist(PartyIndex).miembros(n).privi = (UserList(partylist(PartyIndex).miembros(n).ID).Stats.ELV * _
                                                       100) \ totalexpParty(PartyIndex)

        End If

    Next
    Exit Sub
errhandler:
    Call LogError("Error en BalanceaPrivisLVL PID:" & PartyIndex & " N: " & Err.number & " D: " & Err.Description)

End Sub

Sub BalanceaPrivisMiembros(PartyIndex As Integer)

    On Error GoTo errhandler

    Dim n As Integer

    For n = 1 To MAXMIEMBROS    'partylist(PartyIndex).numMiembros

        If partylist(PartyIndex).miembros(n).ID <> 0 Then
            partylist(PartyIndex).miembros(n).privi = 100 \ partylist(PartyIndex).numMiembros

        End If

    Next
    Exit Sub
errhandler:
    Call LogError("Error en BalanceaPrivisMiembros PID:" & PartyIndex & " N: " & Err.number & " D: " & Err.Description)

End Sub

