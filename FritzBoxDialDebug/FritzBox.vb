Imports System.Text
Imports System.Xml
Imports System.Threading
Imports System.ComponentModel
Public Class FritzBox
    Friend sDefaultSID As String = "0000000000000000"
    Private c_ini As InI
    Private c_Crypt As Rijndael
    Private c_hf As Helfer
    Private form_Config As formCfg

    Private FBFehler As ErrObject
    Private FBEncoding As System.Text.Encoding = Encoding.UTF8
    Private sSID As String = sDefaultSID ' Startwert: UNg�ltige SID
    Private threadTelefon As Thread
    Private sDateiPfad As String
    Private bRausschreiben As Boolean = False
    Private sFBAddr As String

    Public Sub New(ByVal IniPath As String, _
                   ByVal iniKlasse As InI, _
                   ByVal HelferKlasse As Helfer, _
                   ByVal CryptKlasse As Rijndael, _
                   ByRef TelEinlesen As Boolean, _
                   Optional ByVal frmconfig As formCfg = Nothing)

        Dim EncodeingFritzBox As String

        ' Zuweisen der an die Klasse �bergebenen Parameter an die internen Variablen, damit sie in der Klasse global verf�gbar sind
        sDateiPfad = IniPath
        c_ini = iniKlasse
        c_hf = HelferKlasse
        c_hf.KeyChange(sDateiPfad)
        c_Crypt = CryptKlasse

        If Not frmconfig Is Nothing Then
            form_Config = frmconfig
            bRausschreiben = True
            setline("Konfigurationsmen� erhalten")
        End If

        sFBAddr = c_ini.Read(sDateiPfad, "Optionen", "TBFBAdr", "fritz.box")

        EncodeingFritzBox = c_ini.Read(sDateiPfad, "Optionen", "EncodeingFritzBox", "-1")
        If EncodeingFritzBox = "-1" Then
            Dim R�ckgabe As String
            R�ckgabe = c_hf.httpRead("http://" & sFBAddr, FBEncoding, FBFehler)
            If FBFehler Is Nothing Then
                FBEncoding = c_hf.GetEncoding(c_hf.StringEntnehmen(R�ckgabe, "charset=", """>"))
                c_ini.Write(sDateiPfad, "Optionen", "EncodeingFritzBox", FBEncoding.HeaderName)
            Else
                c_hf.LogFile("FBError (FritzBox.New): " & Err.Number & " - " & Err.Description & " - " & "http://" & sFBAddr)
            End If
        Else
            FBEncoding = c_hf.GetEncoding(EncodeingFritzBox)
        End If
        If c_ini.Read(sDateiPfad, "Telefone", "Anzahl", "-1") = "-1" And TelEinlesen Then
            c_hf.LogFile("Telefone, Anzahl nicht vorhanden. Starte Einleseroutine in STA-Thread.")
            threadTelefon = New Thread(AddressOf FritzBoxDaten)
            With threadTelefon
                .SetApartmentState(ApartmentState.STA)
                .IsBackground = True
                .Start()
            End With
        End If

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#Region "Login & Logout"
    Public Function FBLogin(ByRef Fw550 As Boolean, Optional ByVal InpupBenutzer As String = vbNullString, Optional ByVal InpupPasswort As String = "-1") As String
        Dim sLink As String
        Dim slogin_xml As String

        sLink = "http://" & sFBAddr & "/login_sid.lua?sid=" & sSID
        slogin_xml = c_hf.httpRead(sLink, FBEncoding, FBFehler)

        If FBFehler Is Nothing Then
            If InStr(slogin_xml, "FRITZ!Box Anmeldung", CompareMethod.Text) = 0 And Not Len(slogin_xml) = 0 Then

                If Not InpupPasswort = "-1" Then
                    c_ini.Write(sDateiPfad, "Optionen", "TBPasswort", c_Crypt.EncryptString128Bit(InpupPasswort, "Fritz!Box Script"))
                    c_ini.Write(sDateiPfad, "Optionen", "TBBenutzer", InpupBenutzer)
                    SaveSetting("FritzBox", "Optionen", "Zugang", "Fritz!Box Script")
                    c_hf.KeyChange(sDateiPfad)
                End If

                Dim sBlockTime As String
                Dim sChallenge As String
                Dim sFBBenutzer As String = c_ini.Read(sDateiPfad, "Optionen", "TBBenutzer", vbNullString)
                Dim sFBPasswort As String = c_ini.Read(sDateiPfad, "Optionen", "TBPasswort", vbNullString)
                Dim sFormData As String
                Dim sResponse As String
                Dim sSIDResponse As String
                Dim sZugang As String = GetSetting("FritzBox", "Optionen", "Zugang", "-1")
                Dim XMLDocLogin As New XmlDocument()

                '<SessionInfo>
                '   <SID>ff88e4d39354992f</SID>
                '   <Challenge>ab7190d6</Challenge>
                '   <BlockTime>128</BlockTime>
                '   <Rights>
                '       <Name>BoxAdmin</Name>
                '       <Access>2</Access>
                '       <Name>Phone</Name>
                '       </Access>2</Access>
                '       <Name>NAS></Name>
                '       <Access>2</Access>
                '   </Rights>
                '</SessionInfo> 

                '<?xml version="1.0" encoding="utf-8"?>
                '<SessionInfo>
                '   <iswriteaccess>0</iswriteaccess>
                '   <SID>0000000000000000</SID>
                '   <Challenge>dbef619d</Challenge>
                '</SessionInfo>


                With XMLDocLogin
                    .LoadXml(slogin_xml)

                    If .Item("SessionInfo").Item("SID").InnerText() = sDefaultSID Then
                        sChallenge = .Item("SessionInfo").Item("Challenge").InnerText()
                        With c_Crypt
                            sSIDResponse = String.Concat(sChallenge, "-", .getMd5Hash(String.Concat(sChallenge, "-", .DecryptString128Bit(sFBPasswort, sZugang)), Encoding.Unicode))
                        End With
                        If .InnerXml.Contains("Rights") Then
                            'If .Item("SessionInfo").InnerXml.Contains("Rights") Then
                            sBlockTime = .Item("SessionInfo").Item("BlockTime").InnerText()
                            If sBlockTime = "0" Then
                                sLink = "http://" & sFBAddr & "/login_sid.lua?username=" & sFBBenutzer & "&response=" & sSIDResponse
                                sResponse = c_hf.httpRead(sLink, FBEncoding, FBFehler)
                                If FBFehler Is Nothing Then
                                    Fw550 = True
                                Else
                                    c_hf.LogFile("FBError (FBLogin): " & Err.Number & " - " & Err.Description & " - " & sLink)
                                End If
                            Else
                                c_hf.FBDB_MsgBox("Die Fritz!Box l�sst keinen weiteren Anmeldeversuch in den n�chsten " & sBlockTime & "Sekunden zu.  Versuchen Sie es sp�ter erneut.", MsgBoxStyle.Critical, "FBLogin")
                                c_hf.LogFile("Der anzumendende Nutzer verf�gt �ber keine ausreichende Berechtigung.")
                                Return sDefaultSID
                            End If
                        Else
                            If CBool(.Item("SessionInfo").Item("iswriteaccess").InnerText) Then
                                c_hf.LogFile("Die Fritz!Box ben�tigt kein Passwort. Das AddIn wird nicht funktionieren.")
                                Return .Item("SessionInfo").Item("SID").InnerText()
                            End If

                            sLink = "http://" & sFBAddr & "/cgi-bin/webcm"
                            sFormData = "getpage=../html/login_sid.xml&login:command/response=" + sSIDResponse
                            sResponse = c_hf.httpWrite(sLink, sFormData, FBEncoding)

                            Fw550 = False
                        End If

                        .LoadXml(sResponse)

                        sSID = .Item("SessionInfo").Item("SID").InnerText()

                        If Not sSID = sDefaultSID Then
                            If Fw550 Then
                                If Not c_hf.IsOneOf("BoxAdmin", Split(.SelectSingleNode("//Rights").InnerText, "2")) Then
                                    c_hf.LogFile("Es fehlt die Berechtigung f�r den Zugriff auf die Fritz!Box. Benutzer: " & sFBBenutzer)
                                    FBLogout(sSID)
                                    sSID = sDefaultSID
                                End If
                                c_ini.Write(sDateiPfad, "Optionen", sFBBenutzer, CStr(IIf(sSID = sDefaultSID, 0, 2)))
                            End If
                        Else
                            c_hf.LogFile("Die Anmeldedaten sind falsch." & sSID)
                        End If

                    ElseIf .Item("SessionInfo").Item("SID").InnerText() = sSID Then
                        c_hf.LogFile("Eine g�ltige SessionID ist bereits vorhanden: " & sSID)
                    End If
                End With
                XMLDocLogin = Nothing
            End If
        Else
            c_hf.LogFile("FBError (FBLogin): " & Err.Number & " - " & Err.Description & " - " & sLink)
        End If
        Return sSID
    End Function

    Public Function FBLogout(ByRef sSID As String) As Boolean
        ' Die Komplement�rfunktion zu FBLogin. Beendet die Session, indem ein Logout durchgef�hrt wird.

        Dim sLink As String
        Dim Response As String
        Dim xml As New XmlDocument()

        sLink = "http://" & sFBAddr & "/login_sid.lua?sid=" & sSID
        Response = c_hf.httpRead(sLink, FBEncoding, FBFehler)
        If FBFehler Is Nothing Then
            With xml
                .LoadXml(Response)
                If .InnerXml.Contains("Rights") Then
                    sLink = "http://" & sFBAddr & "/home/home.lua?sid=" & sSID & "&logout=1"
                Else
                    sLink = "http://" & sFBAddr & "/logout.lua?sid=" & sSID
                End If
            End With
            xml = Nothing
            Response = c_hf.httpRead(sLink, FBEncoding, FBFehler)
            c_hf.KeyChange(sDateiPfad)
            If FBFehler Is Nothing Then
                If Not InStr(Response, "Sie haben sich erfolgreich von der FRITZ!Box abgemeldet.", CompareMethod.Text) = 0 Or _
                    Not InStr(Response, "Sie haben sich erfolgreich von der Benutzeroberfl�che Ihrer FRITZ!Box abgemeldet.", CompareMethod.Text) = 0 Then
                    c_hf.LogFile("Logout erfolgreich")
                    sSID = sDefaultSID
                    Return True
                Else
                    c_hf.LogFile("Logout eventuell NICHT erfolgreich!")
                    sSID = sDefaultSID
                    Return False
                End If
            Else
                c_hf.LogFile("FBError (FBLogout): " & Err.Number & " - " & Err.Description & " - " & sLink)
            End If
        Else
            c_hf.LogFile("FBError (FBLogout): " & Err.Number & " - " & Err.Description & " - " & sLink)
        End If
        Return False
    End Function
#End Region

#Region "Telefonnummern, Telefonnamen"
    Friend Overloads Sub FritzBoxDaten()
        Dim FW550 As Boolean = True
        Dim sLink As String
        Dim tempstring As String
        Dim tempstring_code As String

        If bRausschreiben Then setline("Fritz!Box Adresse: " & sFBAddr)

        FBLogin(FW550)
        If Not sSID = sDefaultSID Then
            sLink = "http://" & sFBAddr & "/fon_num/fon_num_list.lua?sid=" & sSID
            If bRausschreiben Then
                setline("Fritz!Box SessionID: " & sSID)
                setline("Fritz!Box Firmware  5.50: " & FW550.ToString)
            End If

            If bRausschreiben Then
                If form_Config.CBTelefonDatei.Checked Then sLink = form_Config.TBTelefonDatei.Text
                setline("Fritz!Box Telefon Quelldatei: " & sLink)
            End If

            tempstring = c_hf.httpRead(sLink, FBEncoding, FBFehler)
            If FBFehler Is Nothing Then
                If InStr(tempstring, "FRITZ!Box Anmeldung", CompareMethod.Text) = 0 Then
                    tempstring = Replace(tempstring, Chr(34), "'", , , CompareMethod.Text)   ' " in ' umwandeln 
                    tempstring = Replace(tempstring, Chr(13), "", , , CompareMethod.Text)
                    If InStr(tempstring, "Luacgi not readable") = 0 Then
                        tempstring_code = c_hf.StringEntnehmen(tempstring, "<code>", "</code>")

                        If Not tempstring_code = "-1" Then
                            tempstring = tempstring_code
                        Else
                            tempstring = c_hf.StringEntnehmen(tempstring, "<pre>", "</pre>")
                        End If
                        If Not tempstring = "-1" Then
                            FritzBoxDatenN(tempstring)
                            FBLogout(sSID)
                        Else
                            c_hf.FBDB_MsgBox("Fehler bei dem Herunterladen der Telefone: Telefonieseite kann nicht gelesen werden.", MsgBoxStyle.Critical, "FritzBoxDaten #3")
                        End If
                    Else
                        FritzBoxDatenA()
                    End If

                Else
                    c_hf.FBDB_MsgBox("Fehler bei dem Herunterladen der Telefone: Die Anmeldedaten sind falsch oder es fehlt die Berechtigung f�r diesen Bereich.", MsgBoxStyle.Critical, "FritzBoxDaten #1")
                End If
            Else
                c_hf.LogFile("FBError (FritzBoxDaten): " & Err.Number & " - " & Err.Description & " - " & sLink)
            End If
        Else
            c_hf.FBDB_MsgBox("Fehler bei dem Herunterladen der Telefone: Die Anmeldedaten sind falsch oder es fehlt die Berechtigung f�r diesen Bereich.", MsgBoxStyle.Critical, "FritzBoxDaten #2")
        End If
    End Sub

    Private Sub FritzBoxDatenA()
        If bRausschreiben Then setline("Fritz!Box Telefone Auslesen gestartet. (alt)")

        Dim Vorwahl As String = c_ini.Read(sDateiPfad, "Optionen", "TBVorwahl", "")  ' In den Einstellungen eingegebene Vorwahl
        Dim TelName As String                 ' Gefundener Telefonname
        Dim TelNr As String                 ' Dazugeh�rige Telefonnummer
        Dim SIPID As String = "-1"
        Dim pos(6) As Integer                   ' Positionsmarker
        Dim Anzahl As Integer = 0
        Dim AnzahlISDN As Integer = 0
        Dim ID As Integer
        Dim PortName() As String = Split("readFon123;readNTHotDialList;readDect1;readFonControl;readVoipExt;readTam;readFaxMail", ";", , CompareMethod.Text)
        Dim EndPortName = Split("return list;return list;return list;return list;return Result;return list;return list", ";", , CompareMethod.Text)
        Dim Section As String
        Dim TelefonString() As String
        Dim j As Integer = 0
        Dim SIP(20) As String
        Dim TAM(10) As String
        Dim MSN(10) As String
        Dim EingerichteteTelefone As String = vbNullString
        Dim DialPort As String
        Dim POTS As String
        Dim Mobil As String
        Dim AllIn As String
        Dim outgoing As String
        Dim tempstring As String
        'Alten Einstellungen l�schen
        Dim sLink As String
        'MSNs emitteln

        sLink = "http://" & sFBAddr & "/cgi-bin/webcm?sid=" & sSID & "&getpage=../html/de/menus/menu2.html&var:lang=de&var:menu=fon&var:pagename=fondevices"
        If bRausschreiben Then
            If form_Config.CBTelefonDatei.Checked Then
                sLink = form_Config.TBTelefonDatei.Text
            End If
        End If
        If bRausschreiben Then setline("Fritz!Box Telefon Quelldatei: " & sLink)

        tempstring = c_hf.httpRead(sLink, FBEncoding, FBFehler)
        If FBFehler Is Nothing Then

            If Not InStr(tempstring, "FRITZ!Box Anmeldung", CompareMethod.Text) = 0 Then
                c_hf.FBDB_MsgBox("Fehler bei dem Herunterladen der Telefone. Anmeldung fehlerhaft o.A.!", MsgBoxStyle.Critical, "FritzBoxDaten_FWbelow5_50")
                Exit Sub
            End If
            c_ini.Write(sDateiPfad, "Telefone", vbNullString, "")
            tempstring = Replace(tempstring, Chr(34), "'", , , CompareMethod.Text)   ' " in ' umwandeln

            FBLogout(sSID)
            If bRausschreiben Then setline("Fritz!Box Logout. SID " & sSID & " ung�ltig")

            pos(0) = 1
            For i = 0 To 9
                pos(0) = InStr(pos(0), tempstring, "nrs.msn.push('", CompareMethod.Text) + 14
                If Not pos(0) = 14 Then
                    pos(1) = InStr(pos(0), tempstring, "'", CompareMethod.Text)
                    TelNr = Mid(tempstring, pos(0), pos(1) - pos(0))
                    If Not TelNr = "" Then
                        TelNr = c_hf.OrtsVorwahlEntfernen(TelNr, Vorwahl)
                        MSN(i) = TelNr
                        j = i

                        If bRausschreiben Then
                            setline("MSN-telefonnummer (MSN) gefunden: MSN" & CStr(i) & ", " & TelNr)
                        Else
                            c_ini.Write(sDateiPfad, "Telefone", "MSN" & CStr(i), TelNr)
                        End If

                    End If
                End If
            Next
            ReDim Preserve MSN(j)
            'Internetnummern ermitteln
            j = 0
            For i = 0 To 19
                pos(0) = InStr(pos(0), tempstring, "nrs.sip.push('", CompareMethod.Text) + 14
                If Not pos(0) = 14 Then
                    pos(1) = InStr(pos(0), tempstring, "'", CompareMethod.Text)
                    TelNr = Mid(tempstring, pos(0), pos(1) - pos(0))
                    If Not TelNr = "" Then
                        TelNr = c_hf.OrtsVorwahlEntfernen(TelNr, Vorwahl)
                        SIP(i) = TelNr

                        SIPID = CStr(i)
                        j = i
                        If bRausschreiben Then
                            setline("Internettelefonnummer (SIP) gefunden: SIP" & CStr(i) & ", " & TelNr)
                        Else
                            c_ini.Write(sDateiPfad, "Telefone", "SIP" & CStr(i), TelNr)
                        End If

                    End If
                End If
            Next
            ReDim Preserve SIP(j)
            j = 0

            If bRausschreiben Then
                setline("Letzte SIP: " & SIPID)
            Else
                c_ini.Write(sDateiPfad, "Telefone", "SIPID", SIPID)
            End If

            'TAM Nr ermitteln
            For i = 0 To 9
                pos(0) = InStr(pos(0), tempstring, "nrs.tam.push('", CompareMethod.Text) + 14
                If Not pos(0) = 14 Then
                    pos(1) = InStr(pos(0), tempstring, "'", CompareMethod.Text)
                    TelNr = Mid(tempstring, pos(0), pos(1) - pos(0))
                    If Not TelNr = "" Then
                        TelNr = c_hf.OrtsVorwahlEntfernen(TelNr, Vorwahl)
                        TAM(i) = TelNr

                        If bRausschreiben Then
                            setline("Anrufbeantworternummer (TAM) gefunden: TAM" & CStr(i) & ", " & TelNr)
                        Else
                            c_ini.Write(sDateiPfad, "Telefone", "TAM" & CStr(i), TelNr)
                        End If

                        j = i
                    End If
                End If
            Next
            ReDim Preserve TAM(j)
            'Festnetznummer ermitteln
            pos(0) = InStr(1, tempstring, "telcfg:settings/MSN/POTS", CompareMethod.Text)
            pos(1) = InStr(pos(0), tempstring, "value='", CompareMethod.Text) + 7
            pos(2) = InStr(pos(1), tempstring, "' id", CompareMethod.Text)
            POTS = Mid(tempstring, pos(1), pos(2) - pos(1))
            POTS = c_hf.OrtsVorwahlEntfernen(POTS, Vorwahl)
            If Not POTS = vbNullString Then

                If bRausschreiben Then
                    setline("Plain old telephone service (POTS) gefunden: POTS, " & POTS)
                Else
                    c_ini.Write(sDateiPfad, "Telefone", "POTS", POTS)
                End If

            End If


            'Mobilnummer ermitteln
            pos(0) = InStr(1, tempstring, "function readFonNumbers() {", CompareMethod.Text)
            pos(1) = InStr(pos(0), tempstring, "nrs.mobil = '", CompareMethod.Text) + Len("nrs.mobil = '")
            pos(2) = InStr(pos(1), tempstring, "';", CompareMethod.Text)
            Mobil = CStr(IIf(pos(1) = Len("nrs.mobil = '"), vbNullString, Mid(tempstring, pos(1), pos(2) - pos(1))))
            If Not Mobil = vbNullString Then

                If bRausschreiben Then
                    setline("Mobilnummer (Mobil) gefunden: Mobil, " & Mobil)
                Else
                    c_ini.Write(sDateiPfad, "Telefone", "Mobil", Mobil)
                End If

            End If

            Dim FAX(0) As String
            AllIn = AlleNummern(MSN, SIP, TAM, FAX, POTS, Mobil)

            'Telefone ermitteln
            pos(0) = 1
            If CBool(c_ini.Read(sDateiPfad, "Optionen", "CBAuslesen", "True")) Then
                For i = 0 To UBound(PortName)
                    pos(0) = InStr(pos(0), tempstring, PortName(i), CompareMethod.Text)
                    pos(1) = InStr(pos(0), tempstring, EndPortName(i), CompareMethod.Text) + Len(EndPortName(i))
                    If pos(1) = Len(EndPortName(i)) Then
                        ' Die JavaFunktion "readVoipExt" f�r die IPTelefone endet ab der Firmware *80 auf "return Result;". (fr�her auf "return list;")
                        pos(1) = InStr(pos(0), tempstring, "return list;", CompareMethod.Text) + Len("return list;")
                    End If
                    Section = Mid(tempstring, pos(0), pos(1) - pos(0))
                    TelefonString = Split(Section, "});", , CompareMethod.Text)

                    For Each Telefon In TelefonString
                        If InStr(Telefon, "return list") = 0 And InStr(Telefon, "Isdn-Default") = 0 Then
                            pos(0) = InStr(Telefon, "name: ", CompareMethod.Text) + Len("name: ")
                            pos(1) = InStr(pos(0), Telefon, ",", CompareMethod.Text)
                            If Not pos(0) = 6 Or Not pos(1) = 0 Then
                                TelName = Mid(Telefon, pos(0), pos(1) - pos(0))
                                If TelName = "fonName" Then
                                    pos(0) = InStr(Telefon, "fonName = '", CompareMethod.Text) + Len("fonName = '")
                                    pos(1) = InStr(pos(0), Telefon, "'", CompareMethod.Text)
                                    TelName = Mid(Telefon, pos(0), pos(1) - pos(0))
                                Else
                                    TelName = Replace(TelName, "'", "", , , CompareMethod.Text)
                                End If
                                pos(2) = InStr(pos(1), Telefon, "number: ", CompareMethod.Text) + Len("number: ")
                                pos(3) = InStr(pos(2), Telefon, Chr(10), CompareMethod.Text)
                                TelNr = Replace(Trim(Mid(Telefon, pos(2), pos(3) - pos(2))), "'", "", , , CompareMethod.Text)
                                TelNr = Replace(TelNr, Chr(10), "", , , CompareMethod.Text)
                                TelNr = Replace(TelNr, Chr(13), "", , , CompareMethod.Text)
                                If Right(TelNr, 1) = "," Then TelNr = Left(TelNr, Len(TelNr) - 1) ' F�r die Firmware *85
                                If Right(TelNr, 1) = "#" Then TelNr = Left(TelNr, Len(TelNr) - 1) ' F�r die Firmware *85
                                If Left(TelNr, 3) = "SIP" Then TelNr = SIP(CInt(Mid(TelNr, 4, 1)))
                                If Not Trim(TelName) = "" And Not Trim(TelNr) = "" Then
                                    Select Case i
                                        Case 0 ' FON 1-3
                                            pos(2) = InStr(pos(1), Telefon, "allin: ('", CompareMethod.Text) + Len("allin: ('")
                                            pos(3) = InStr(pos(2), Telefon, "')", CompareMethod.Text)
                                            If Mid(Telefon, pos(2), pos(3) - pos(2)) = "1'=='1" Then
                                                TelNr = AllIn
                                            Else
                                                TelNr = c_hf.OrtsVorwahlEntfernen(TelNr, Vorwahl)
                                            End If
                                            pos(4) = InStr(Telefon, "n = parseInt('", CompareMethod.Text) + Len("n = parseInt('")
                                            pos(5) = InStr(pos(4), Telefon, "'", CompareMethod.Text)
                                            ID = CInt(Mid(Telefon, pos(4), pos(5) - pos(4))) + 1
                                            pos(2) = InStr(pos(1), Telefon, "outgoing: '", CompareMethod.Text) + Len("outgoing: '")
                                            pos(3) = InStr(pos(2), Telefon, "'", CompareMethod.Text)
                                            outgoing = Mid(Telefon, pos(2), pos(3) - pos(2))
                                            If Strings.Right(outgoing, 1) = "#" Then outgoing = Strings.Left(outgoing, Len(outgoing) - 1) ' F�r die Firmware *85
                                            If Left(outgoing, 3) = "SIP" Then outgoing = SIP(CInt(Mid(outgoing, 4, 1)))
                                            EingerichteteTelefone = String.Concat(EingerichteteTelefone, CStr(ID), ";")

                                            If bRausschreiben Then
                                                setline("Analogtelefon gefunden: FON" & CStr(ID) & ", " & outgoing & ", " & TelNr & ", " & TelName)
                                            Else
                                                c_ini.Write(sDateiPfad, "Telefone", CStr(ID), outgoing & ";" & TelNr & ";" & TelName)
                                            End If

                                            Anzahl += 1
                                        Case 1 ' S0-Port
                                            pos(2) = InStr(Telefon, "partyNo = '", CompareMethod.Text) + Len("partyNo = '")
                                            pos(3) = InStr(pos(2), Telefon, "'", CompareMethod.Text)
                                            If Not pos(2) = pos(3) Then
                                                AnzahlISDN += 1
                                                pos(4) = InStr(pos(1), Telefon, "allin: ('", CompareMethod.Text) + Len("allin: ('")
                                                pos(5) = InStr(pos(2), Telefon, "')", CompareMethod.Text)
                                                If Mid(Telefon, pos(4), pos(5) - pos(4)) = "true" Then
                                                    TelNr = AllIn
                                                Else
                                                    TelNr = Trim(Mid(Telefon, pos(2), pos(3) - pos(2)))
                                                    TelNr = c_hf.OrtsVorwahlEntfernen(TelNr, Vorwahl)
                                                End If
                                                pos(4) = InStr(Telefon, "n = parseInt('", CompareMethod.Text) + Len("n = parseInt('")
                                                pos(5) = InStr(pos(4), Telefon, "'", CompareMethod.Text)
                                                ID = CInt(Mid(Telefon, pos(4), pos(5) - pos(4)))
                                                pos(2) = InStr(pos(1), Telefon, "outgoing: '", CompareMethod.Text) + Len("outgoing: '")
                                                pos(3) = InStr(pos(2), Telefon, "'", CompareMethod.Text)
                                                outgoing = Mid(Telefon, pos(2), pos(3) - pos(2))
                                                If Strings.Right(outgoing, 1) = "#" Then outgoing = Strings.Left(outgoing, Len(outgoing) - 1) ' F�r die Firmware *85
                                                If Left(outgoing, 3) = "SIP" Then outgoing = SIP(CInt(Mid(outgoing, 4, 1)))
                                                DialPort = "5" & ID
                                                EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")

                                                If bRausschreiben Then
                                                    setline("S0-Telefon gefunden: " & DialPort & ", " & ", " & TelNr & ", " & TelName)
                                                Else
                                                    c_ini.Write(sDateiPfad, "Telefone", DialPort, outgoing & ";" & TelNr & ";" & TelName)
                                                End If

                                            End If
                                        Case 2 ' DECT Fritz!Fon 7150
                                            Anzahl += 1
                                            pos(2) = InStr(Telefon, "n = parseInt('", CompareMethod.Text) + Len("n = parseInt('")
                                            pos(3) = InStr(pos(2), Telefon, "'", CompareMethod.Text)
                                            ID = CInt(Trim(Mid(Telefon, pos(2), pos(3) - pos(2))))
                                            TelNr = c_hf.OrtsVorwahlEntfernen(TelNr, Vorwahl)
                                            DialPort = "6" & ID
                                            EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")

                                            If bRausschreiben Then
                                                setline("DECT Fritz!Fon 7150 gefunden: " & DialPort & ", " & ", " & TelNr & ", " & TelName)
                                            Else
                                                c_ini.Write(sDateiPfad, "Telefone", DialPort, TelNr & ";" & TelName)
                                            End If

                                        Case 3 ' DECT
                                            Dim isUnpersonalizedMini() As String
                                            Dim tempTelNr As String
                                            pos(2) = InStr(Telefon, "isUnpersonalizedMini = '", CompareMethod.Text) + Len("isUnpersonalizedMini = '")
                                            pos(3) = InStr(pos(2), Telefon, "';", CompareMethod.Text)
                                            isUnpersonalizedMini = Split(Mid(Telefon, pos(2), pos(3) - pos(2)), "' == '", , CompareMethod.Text)
                                            If Not isUnpersonalizedMini(0) = isUnpersonalizedMini(1) Then
                                                Anzahl += 1
                                                pos(2) = InStr(Telefon, "intern: isUnpersonalizedMini ? '' : '**", CompareMethod.Text) + Len("intern: isUnpersonalizedMini ? '' : '**") + 2
                                                pos(3) = InStr(pos(2), Telefon, "'", CompareMethod.Text)
                                                ID = CInt(Trim(Mid(Telefon, pos(2), pos(3) - pos(2))))
                                                pos(2) = InStr(pos(1), Telefon, "allin: ('", CompareMethod.Text) + Len("allin: ('")
                                                pos(3) = InStr(pos(2), Telefon, "')", CompareMethod.Text)
                                                If Mid(Telefon, pos(2), pos(3) - pos(2)) = "1'=='1" Then
                                                    TelNr = AllIn
                                                Else
                                                    pos(2) = InStr(Telefon, "num = '", CompareMethod.Text) + Len("num = '")
                                                    TelNr = vbNullString
                                                    If Not pos(2) = 7 Then
                                                        Do
                                                            pos(3) = InStr(pos(2), Telefon, "'", CompareMethod.Text)
                                                            tempTelNr = Mid(Telefon, pos(2), pos(3) - pos(2))
                                                            TelNr = c_hf.OrtsVorwahlEntfernen(TelNr, Vorwahl)
                                                            TelNr += CStr(IIf(Right(TelNr, 1) = "#", vbNullString, tempTelNr & "_"))
                                                            pos(2) = InStr(pos(3), Telefon, "num = '", CompareMethod.Text) + Len("num = '")
                                                        Loop Until pos(2) = 7
                                                        TelNr = Left(TelNr, Len(TelNr) - 1)
                                                    Else
                                                        pos(2) = InStr(TelNr, ":", CompareMethod.Text) + 2
                                                        TelNr = Trim(Mid(TelNr, pos(2)))
                                                        TelNr = c_hf.OrtsVorwahlEntfernen(TelNr, Vorwahl)
                                                    End If
                                                End If
                                                pos(2) = InStr(pos(1), Telefon, "outgoing: isUnpersonalizedMini ? '' : '", CompareMethod.Text) + Len("outgoing: isUnpersonalizedMini ? '' : '")
                                                pos(3) = InStr(pos(2), Telefon, "'", CompareMethod.Text)
                                                outgoing = Mid(Telefon, pos(2), pos(3) - pos(2))
                                                If Left(outgoing, 3) = "SIP" Then outgoing = SIP(CInt(Mid(outgoing, 4, 1)))
                                                DialPort = "6" & ID
                                                EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")

                                                If bRausschreiben Then
                                                    setline("DECT-Telefon gefunden: " & DialPort & ", " & outgoing & ", " & TelNr & ", " & TelName)
                                                Else
                                                    c_ini.Write(sDateiPfad, "Telefone", DialPort, outgoing & ";" & TelNr & ";" & TelName)
                                                End If

                                            End If
                                        Case 4 ' IP-Telefone
                                            If Not Trim(TelName) = "TelCfg[Index].Name" Then
                                                pos(4) = InStr(Telefon, "n = parseInt('", CompareMethod.Text) + Len("n = parseInt('")
                                                pos(5) = InStr(pos(4), Telefon, "'", CompareMethod.Text)
                                                ID = CInt(Mid(Telefon, pos(4), pos(5) - pos(4)))
                                                Anzahl += 1
                                                DialPort = "2" & ID
                                                EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")

                                                If bRausschreiben Then
                                                    setline("IP-Telefon gefunden: " & DialPort & ", " & TelNr & ", " & TelName)
                                                Else
                                                    c_ini.Write(sDateiPfad, "Telefone", DialPort, TelNr & ";" & TelName)
                                                End If
                                            Else
                                                Dim LANTelefone() As String = Split(Telefon, "in_nums = [];", , CompareMethod.Text)
                                                Dim InNums As String = vbNullString
                                                Dim NetInfo As String
                                                Dim NetInfoPush As String = vbNullString
                                                pos(0) = InStr(LANTelefone(LANTelefone.Length - 1), "NetInfo.push(parseInt('", CompareMethod.Text)
                                                If Not pos(0) = 0 Then
                                                    NetInfo = Mid(LANTelefone(LANTelefone.Length - 1), pos(0))
                                                    pos(0) = 1
                                                    Do
                                                        pos(1) = InStr(pos(0), NetInfo, "', 10));", CompareMethod.Text) + Len("', 10));")
                                                        NetInfoPush = Mid(NetInfo, pos(0) + Len("NetInfo.push(parseInt('"), 3) & CStr(IIf(Not NetInfoPush = vbNullString, ";" & NetInfoPush, vbNullString))
                                                        pos(0) = InStr(pos(1), NetInfo, "NetInfo.push(parseInt('", CompareMethod.Text)
                                                    Loop Until pos(0) = 0
                                                End If
                                                For Each LANTelefon In LANTelefone
                                                    If Not InStr(LANTelefon, "TelCfg.push( { Enabled : '", vbTextCompare) = 0 Then
                                                        Dim tempTelNr As String
                                                        pos(2) = InStr(LANTelefon, "num = '", CompareMethod.Text) + Len("num = '")
                                                        TelNr = vbNullString
                                                        If Not pos(2) = 7 Then
                                                            InNums = vbNullString
                                                            Do
                                                                pos(3) = InStr(pos(2), LANTelefon, "'", CompareMethod.Text)
                                                                tempTelNr = Mid(LANTelefon, pos(2), pos(3) - pos(2))
                                                                TelNr = c_hf.OrtsVorwahlEntfernen(tempTelNr, Vorwahl)
                                                                InNums += CStr(IIf(Strings.Right(TelNr, 1) = "#", vbNullString, TelNr & "_"))
                                                                pos(2) = InStr(pos(3), LANTelefon, "num = '", CompareMethod.Text) + Len("num = '")
                                                            Loop Until pos(2) = 7
                                                            InNums = Left(InNums, Len(InNums) - 1)
                                                        End If

                                                        pos(0) = InStr(LANTelefon, "Name : '", CompareMethod.Text) + Len("Name : '")
                                                        pos(1) = InStr(pos(0), LANTelefon, "'", CompareMethod.Text)
                                                        TelName = Mid(LANTelefon, pos(0), pos(1) - pos(0))
                                                        If Not TelName = vbNullString Then
                                                            pos(2) = InStr(pos(1), Telefon, "AllIn: ('", CompareMethod.Text) + Len("AllIn: ('")
                                                            pos(3) = InStr(pos(2), Telefon, "')", CompareMethod.Text)
                                                            If Mid(Telefon, pos(2), pos(3) - pos(2)) = "1' == '1" Then
                                                                TelNr = AllIn
                                                            Else
                                                                If Not InStr(LANTelefon, "InNums : in_nums", CompareMethod.Text) = 0 Then
                                                                    TelNr = InNums
                                                                Else
                                                                    pos(2) = InStr(pos(1), LANTelefon, "Number0 : '", CompareMethod.Text) + Len("Number0 : '")
                                                                    pos(3) = InStr(pos(2), LANTelefon, "'", CompareMethod.Text)
                                                                    TelNr = c_hf.OrtsVorwahlEntfernen(Mid(LANTelefon, pos(2), pos(3) - pos(2)), Vorwahl)
                                                                End If
                                                            End If
                                                            pos(4) = InStr(LANTelefon, "g_txtIpPhone + ' 62", CompareMethod.Text) + Len("g_txtIpPhone + ' 62")
                                                            ID = CInt(Mid(LANTelefon, pos(4), 1))
                                                            If NetInfoPush = vbNullString Then
                                                                If Not InStr(LANTelefon, "TelCfg.push( { Enabled : '1',", CompareMethod.Text) = 0 Then
                                                                    DialPort = "2" & ID
                                                                    EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")
                                                                    Anzahl += 1
                                                                    If bRausschreiben Then
                                                                        setline("IP-Telefon gefunden: " & DialPort & ", " & TelNr & ", " & TelName)
                                                                    Else
                                                                        c_ini.Write(sDateiPfad, "Telefone", DialPort, ";" & TelNr & ";" & TelName)
                                                                    End If

                                                                End If
                                                            Else
                                                                If c_hf.IsOneOf("62" & ID, Split(NetInfoPush, ";", , CompareMethod.Text)) Then
                                                                    DialPort = "2" & ID
                                                                    EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")
                                                                    Anzahl += 1
                                                                    If bRausschreiben Then
                                                                        setline("IP-Telefon gefunden: " & DialPort & ", " & TelNr & ", " & TelName)
                                                                    Else
                                                                        c_ini.Write(sDateiPfad, "Telefone", DialPort, ";" & TelNr & ";" & TelName)
                                                                    End If

                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        Case 5 ' Anrufbeantworter
                                            Dim tamMsnBits As Integer
                                            TelNr = vbNullString
                                            pos(2) = InStr(Telefon, "tamDisplay = '", CompareMethod.Text) + Len("tamDisplay = '")
                                            pos(3) = InStr(pos(2), Telefon, "'", CompareMethod.Text)
                                            If Mid(Telefon, pos(2), pos(3) - pos(2)) = "1" Then
                                                pos(4) = InStr(Telefon, "n = parseInt('", CompareMethod.Text) + Len("n = parseInt('")
                                                pos(5) = InStr(pos(4), Telefon, "'", CompareMethod.Text)
                                                ID = CInt(Mid(Telefon, pos(4), pos(5) - pos(4)))
                                                pos(4) = InStr(Telefon, "var tamMsnBits = parseInt('", CompareMethod.Text) + Len("var tamMsnBits = parseInt('")
                                                pos(5) = InStr(pos(4), Telefon, "'", CompareMethod.Text)
                                                tamMsnBits = CInt(Mid(Telefon, pos(4), pos(5) - pos(4)))
                                                For j = 0 To TAM.Length - 1
                                                    If Not TAM(j) Is Nothing Then
                                                        If (tamMsnBits And (1 << j)) > 0 Then ' Aus AVM Quellcode Funktion isBitSet �bernommen 
                                                            TelNr += TAM(j) & "_"
                                                        End If
                                                    End If
                                                Next
                                                If Not TelNr = vbNullString Then
                                                    TelNr = Left(TelNr, Len(TelNr) - 1)
                                                    DialPort = "60" & ID
                                                    EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")

                                                    If bRausschreiben Then
                                                        setline("Anrufbeantworter gefunden: " & DialPort & ", " & ", " & TelNr & ", " & TelName)
                                                    Else
                                                        c_ini.Write(sDateiPfad, "Telefone", DialPort, ";" & TelNr & ";" & TelName)
                                                    End If

                                                    Anzahl += 1
                                                End If
                                            End If
                                        Case 6 ' integrierter Faxempfang
                                            Dim FAXMSN(9) As String
                                            TelNr = vbNullString
                                            pos(2) = InStr(Telefon, "var isActive = '", CompareMethod.Text) + Len("var isActive = '")
                                            pos(3) = InStr(pos(2), Telefon, "'", CompareMethod.Text)
                                            If Not pos(2) = pos(3) Then
                                                If CInt(Mid(Telefon, pos(2), pos(3) - pos(2))) > 0 Then
                                                    TelName = "Faxempfang"
                                                    If InStr(Telefon, "allin: true", CompareMethod.Text) = 0 Then
                                                        pos(2) = InStr(Telefon, "var faxMsn = '", CompareMethod.Text) + Len("var faxMsn = '")
                                                        pos(3) = InStr(pos(2), Telefon, "'", CompareMethod.Text)
                                                        If Not pos(2) = Len("var faxMsn = '") Then
                                                            TelNr = Mid(Telefon, pos(2), pos(3) - pos(2))
                                                        Else
                                                            pos(3) = 1
                                                            For j = 0 To 9
                                                                pos(2) = InStr(pos(3), Telefon, "msn = '", CompareMethod.Text) + Len("msn = '")
                                                                pos(3) = InStr(pos(2), Telefon, "'", CompareMethod.Text)
                                                                FAXMSN(j) = Mid(Telefon, pos(2), pos(3) - pos(2))
                                                            Next
                                                            pos(2) = InStr(Telefon, "number: faxMsns[", CompareMethod.Text) + Len("number: faxMsns[")
                                                            pos(3) = InStr(pos(2), Telefon, "]", CompareMethod.Text)
                                                            TelNr = FAXMSN(CInt(Mid(Telefon, pos(2), pos(3) - pos(2))))
                                                        End If
                                                    Else
                                                        TelNr = AllIn
                                                    End If
                                                    DialPort = "5"
                                                    EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")

                                                    If bRausschreiben Then
                                                        setline("Die integrierte Faxfunktion ist eingeschaltet: " & DialPort & ", " & TelNr & "," & TelName)
                                                    Else
                                                        c_ini.Write(sDateiPfad, "Telefone", DialPort, ";" & TelNr & ";" & TelName)
                                                    End If

                                                    Anzahl += 1
                                                End If
                                            End If
                                    End Select
                                End If
                            End If
                        End If
                    Next
                Next
            End If

            If Not AnzahlISDN = 0 Then
                If bRausschreiben Then
                    setline("S0-Basis hinzugef�gt.")
                Else
                    c_ini.Write(sDateiPfad, "Telefone", "50", ";;ISDN-Basis")
                    EingerichteteTelefone = String.Concat(EingerichteteTelefone, "50", ";")
                End If

            End If

            EingerichteteTelefone = Strings.Left(EingerichteteTelefone, Strings.Len(EingerichteteTelefone) - 1)

            If bRausschreiben Then
                setline("Anzahl Telefone: " & Anzahl)
                setline("Anzahl ISDN: " & AnzahlISDN)
                setline("Gesamtanzahl: " & Anzahl + AnzahlISDN)
            Else
                c_ini.Write(sDateiPfad, "Telefone", "EingerichteteTelefone", EingerichteteTelefone)
                c_ini.Write(sDateiPfad, "Telefone", "Anzahl", CStr(Anzahl + AnzahlISDN))
            End If

        Else
            c_hf.LogFile("FBError (FritzBoxDatenA): " & Err.Number & " - " & Err.Description & " - " & sLink)
        End If

    End Sub ' (FritzBoxDaten f�r �ltere Firmware)

    Private Sub FritzBoxDatenN(ByVal Code As String)
        If bRausschreiben Then setline("Fritz!Box Telefone Auslesen gestartet (Neu).")

        Dim Vorwahl As String = c_ini.Read(sDateiPfad, "Optionen", "TBVorwahl", "")                 ' In den Einstellungen eingegebene Vorwahl
        Dim Landesvorwahl As String
        Dim TelName As String                 ' Gefundener Telefonname
        Dim TelNr As String                 ' Dazugeh�rige Telefonnummer
        Dim SIPID As String = "-1"
        Dim pos(1) As Integer
        Dim i As Integer                   ' Laufvariable
        Dim j As Integer
        Dim k As Integer
        Dim TelAnzahl As Integer                   ' Anzahl der gefundenen Telefone
        Dim SIP(20) As String
        Dim TAM(10) As String
        Dim MSNPort(2, 9) As String
        Dim MSN(9) As String
        Dim FAX(9) As String
        Dim Mobil As String
        Dim POTS As String
        Dim allin As String
        Dim AnzahlFON123 As Integer = 0
        Dim AnzahlISDN As Integer = 0
        Dim AnzahlDECT As Integer = 0
        Dim AnzahlLANWLAN As Integer = 0
        Dim AnzahlTAM As Integer = 0
        Dim AnzahlFAX As Integer = 0
        Dim DialPort As String
        Dim outgoing As String
        Dim tmpstrTelefone As String
        Dim tmpstrUser() As String
        Dim Node As String
        Dim tmpTelNr As String
        Dim Port As String
        Dim EingerichteteTelefone As String = vbNullString
        Dim EingerichteteFax = vbNullString

        If Not bRausschreiben Then c_ini.Write(sDateiPfad, "Telefone", vbNullString, "")
        'SIP Nummern
        With c_hf
            For Each SIPi In Split(.StringEntnehmen(Code, "['sip:settings/sip/list(" & .StringEntnehmen(Code, "['sip:settings/sip/list(", ")'] = {") & ")'] = {", "}" & Chr(10) & "  },"), " },", , CompareMethod.Text)
                If .StringEntnehmen(SIPi, "['activated'] = '", "'") = "1" Then
                    TelNr = .OrtsVorwahlEntfernen(.StringEntnehmen(SIPi, "['displayname'] = '", "'"), Vorwahl)
                    Node = UCase(.StringEntnehmen(SIPi, "['_node'] = '", "'"))
                    SIPID = .StringEntnehmen(SIPi, "['ID'] = '", "'")
                    SIP(CInt(SIPID)) = TelNr
                    If bRausschreiben Then
                        setline("Internettelefonnummer (SIP) gefunden: " & Node & ", " & TelNr)
                    Else
                        c_ini.Write(sDateiPfad, "Telefone", Node, TelNr)
                    End If
                End If
            Next

            SIP = (From x In SIP Where Not x Like "" Select x).ToArray
            If bRausschreiben Then
                setline("Letzte SIP: " & SIPID)
            Else
                c_ini.Write(sDateiPfad, "Telefone", "SIPID", SIPID)
            End If
            For i = 0 To 9
                TelNr = .StringEntnehmen(Code, "['telcfg:settings/MSN/MSN" & i & "'] = '", "'")
                If Not TelNr = "-1" Then
                    If Not Len(TelNr) = 0 Then
                        TelNr = .OrtsVorwahlEntfernen(TelNr, Vorwahl)
                        MSN(i) = TelNr
                        If bRausschreiben Then
                            setline("MSN-telefonnummer (MSN) gefunden: MSN" & CStr(i) & ", " & TelNr)
                        Else
                            c_ini.Write(sDateiPfad, "Telefone", "MSN" & CStr(i), TelNr)
                        End If
                    End If
                End If
            Next

            For i = 0 To 2
                If Not .StringEntnehmen(Code, "['telcfg:settings/MSN/Port" & i & "/Name'] = '", "'") = "-1" Then
                    For j = 0 To 9
                        TelNr = .StringEntnehmen(Code, "['telcfg:settings/MSN/Port" & i & "/MSN" & j & "'] = '", "'")
                        If Not TelNr = "-1" Then
                            If Not Len(TelNr) = 0 Then
                                If Strings.Left(TelNr, 3) = "SIP" Then
                                    TelNr = SIP(CInt(Mid(TelNr, 4, 1)))
                                Else
                                    TelNr = .OrtsVorwahlEntfernen(TelNr, Vorwahl)
                                End If

                                If Not .IsOneOf(TelNr, MSN) Then
                                    For k = 0 To 9
                                        If MSN(k) = "" Then
                                            MSN(k) = TelNr
                                            If bRausschreiben Then
                                                setline("MSN-telefonnummer (MSN) gefunden: MSN" & CStr(k) & ", " & TelNr)
                                            Else
                                                c_ini.Write(sDateiPfad, "Telefone", "MSN" & CStr(k), TelNr)
                                            End If
                                            Exit For
                                        End If
                                    Next
                                End If
                                MSNPort(i, j) = TelNr
                            End If
                        End If
                    Next
                End If
            Next
            'Dim res = From x In tmp Select x Distinct 'Doppelte entfernen
            MSN = (From x In MSN Select x Distinct).ToArray 'Doppelte entfernen
            MSN = (From x In MSN Where Not x Like "" Select x).ToArray

            For i = 0 To 9
                TelNr = .StringEntnehmen(Code, "['tam:settings/MSN" & i & "'] = '", "'")
                If Not TelNr = "-1" Then
                    If Not Len(TelNr) = 0 Then
                        If Strings.Left(TelNr, 3) = "SIP" Then
                            TelNr = SIP(CInt(Mid(TelNr, 4, 1)))
                        ElseIf Strings.Left(TelNr, 3) = "SIP" Then
                            TelNr = .OrtsVorwahlEntfernen(TelNr, Vorwahl)
                        End If

                        If bRausschreiben Then
                            setline("Anrufbeantworternummer (TAM) gefunden: TAM" & CStr(i) & ", " & TelNr)
                        Else
                            c_ini.Write(sDateiPfad, "Telefone", "TAM" & CStr(i), TelNr)
                        End If

                        TAM(i) = TelNr
                    End If
                End If
            Next
            'TAM = (From x In TAM Where Not x Like "" Select x).ToArray

            For i = 0 To 9
                TelNr = .StringEntnehmen(Code, "['telcfg:settings/FaxMSN" & i & "'] = '", "'")
                If Not TelNr = "-1" Then
                    If Not Len(TelNr) = 0 Then
                        If Strings.Left(TelNr, 3) = "SIP" Then
                            TelNr = SIP(CInt(Mid(TelNr, 4, 1)))
                        Else
                            TelNr = .OrtsVorwahlEntfernen(TelNr, Vorwahl)
                        End If

                        If bRausschreiben Then
                            setline("Faxnummer (FAX) gefunden: FAX" & CStr(i) & ", " & TelNr)
                        Else
                            c_ini.Write(sDateiPfad, "Telefone", "FAX" & CStr(i), TelNr)
                        End If

                        FAX(i) = TelNr
                    End If
                End If
            Next
            FAX = (From x In FAX Where Not x Like "" Select x).ToArray

            POTS = .StringEntnehmen(Code, "['telcfg:settings/MSN/POTS'] = '", "'")
            If Not POTS = "-1" Then
                If Strings.Left(POTS, 3) = "SIP" Then
                    POTS = SIP(CInt(Mid(POTS, 4, 1)))
                Else
                    POTS = .OrtsVorwahlEntfernen(POTS, Vorwahl)
                End If

                If bRausschreiben Then
                    setline("Plain old telephone service (POTS) gefunden: " & POTS)
                Else
                    c_ini.Write(sDateiPfad, "Telefone", "POTS", POTS)
                End If

            End If


            Mobil = .StringEntnehmen(Code, "['telcfg:settings/Mobile/MSN'] = '", "'")
            If Not Mobil = "-1" Then
                If Strings.Left(Mobil, 3) = "SIP" Then
                    Mobil = SIP(CInt(Mid(Mobil, 4, 1)))
                Else
                    Mobil = .OrtsVorwahlEntfernen(Mobil, Vorwahl)
                End If

                If bRausschreiben Then
                    setline("Mobilnummer (Mobil) gefunden: " & Mobil)
                Else
                    c_ini.Write(sDateiPfad, "Telefone", "Mobil", Mobil)
                End If

            End If

            allin = AlleNummern(MSN, SIP, TAM, FAX, POTS, Mobil)

            TelAnzahl = 0
            pos(0) = 1

            'FON
            For Each Telefon In Split(.StringEntnehmen(Code, "['telcfg:settings/MSN/Port/list(" & .StringEntnehmen(Code, "['telcfg:settings/MSN/Port/list(", ")'] = {") & ")'] = {", "}" & Chr(10) & "  },"), " },", , CompareMethod.Text)
                TelName = .StringEntnehmen(Telefon, "['Name'] = '", "'")
                If Not (TelName = "-1" Or TelName = vbNullString) Then
                    TelNr = vbNullString
                    Port = Right(.StringEntnehmen(Telefon, "['_node'] = '", "'"), 1)

                    Dim tmparray(9) As String
                    For i = 0 To 9
                        tmpTelNr = MSNPort(CInt(Port), i)
                        If Not tmpTelNr = "" Then
                            tmparray(i) = MSNPort(CInt(Port), i)
                        Else
                            Exit For
                        End If
                    Next
                    tmparray = (From x In tmparray Where Not x Like "" Select x).ToArray
                    If tmparray.Length = 0 Then
                        ReDim tmparray(9)
                        For i = 0 To 9
                            tmpTelNr = MSN(i)
                            If Not tmpTelNr = "" Then
                                tmparray(i) = MSN(i)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                    outgoing = tmparray(0)
                    TelNr = String.Join("_", tmparray)
                    DialPort = CStr(CInt(Port) + 1)
                    AnzahlFON123 += 1


                    If bRausschreiben Then
                        setline("Analogtelefon gefunden: FON" & DialPort & ", " & outgoing & ", " & TelNr & ", " & TelName)
                    Else
                        c_ini.Write(sDateiPfad, "Telefone", DialPort, outgoing & ";" & TelNr & ";" & TelName)
                    End If

                    EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")
                    If .StringEntnehmen(Telefon, "['Fax'] = '", "'") = "1" Then
                        EingerichteteFax = String.Concat(EingerichteteFax, DialPort, ";")
                        If bRausschreiben Then setline("Analogtelefon FON" & DialPort & " ist ein FAX.")
                    End If

                End If
            Next

            ' DECT

            tmpstrTelefone = .StringEntnehmen(Code, "['telcfg:settings/Foncontrol/User/list(" & .StringEntnehmen(Code, "['telcfg:settings/Foncontrol/User/list(", ")'] = {") & ")'] = {", "}" & Chr(10) & "  },")

            For Each DectTelefon In Split(tmpstrTelefone, "] = {", , CompareMethod.Text)

                DialPort = .StringEntnehmen(DectTelefon, "['Intern'] = '", "'")
                If Not (DialPort = "-1" Or DialPort = vbNullString) Then
                    TelNr = vbNullString
                    DialPort = "6" & Strings.Right(DialPort, 1)

                    TelName = .StringEntnehmen(DectTelefon, "['Name'] = '", "'")

                    Node = .StringEntnehmen(DectTelefon, "['_node'] = '", "'")

                    If .StringEntnehmen(Code, "['telcfg:settings/Foncontrol/" & Node & "/RingOnAllMSNs'] = '", "',") = "1" Then
                        TelNr = allin
                    Else
                        tmpstrUser = Split(.StringEntnehmen(Code, "['telcfg:settings/Foncontrol/" & Node & "/MSN/list(Number)'] = {", "}" & Chr(10) & "  },"), "['Number'] = '", , CompareMethod.Text)

                        tmpstrUser(0) = vbNullString
                        For l As Integer = 1 To tmpstrUser.Length - 1
                            tmpstrUser(l) = Strings.Left(tmpstrUser(l), InStr(tmpstrUser(l), "'", CompareMethod.Text) - 1)
                        Next
                        ' Etwas unsch�ner Code
                        Dim res2 = From x In tmpstrUser Where Not x Like "" Select x ' Leere entfernen
                        For Each Nr In res2
                            TelNr = TelNr & "_" & .OrtsVorwahlEntfernen(Nr, Vorwahl)
                        Next
                        TelNr = Mid(TelNr, 2) 'Strings.Left(TelNr, Len(TelNr) - 1)
                    End If
                    ' Etwas unsch�ner Code
                    outgoing = Split(TelNr, "_", , CompareMethod.Text)(0)
                    AnzahlDECT += 1
                    EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")

                    If bRausschreiben Then
                        setline("DECT-Telefon gefunden: " & DialPort & ", " & outgoing & ", " & TelNr & ", " & TelName)
                    Else
                        c_ini.Write(sDateiPfad, "Telefone", DialPort, outgoing & ";" & TelNr & ";" & TelName)
                    End If

                End If
            Next


            'IP-Telefone
            tmpstrUser = Split(.StringEntnehmen(Code, "['telcfg:settings/VoipExtension/list(" & .StringEntnehmen(Code, "['telcfg:settings/VoipExtension/list(", ")'] = {") & ")'] = {", "}" & Chr(10) & "  },"), " },", , CompareMethod.Text)
            For Each Telefon In tmpstrUser
                If .StringEntnehmen(Telefon, "['enabled'] = '", "'") = "1" Then
                    TelName = .StringEntnehmen(Telefon, "['Name'] = '", "'")
                    TelNr = vbNullString
                    Port = .StringEntnehmen(Telefon, "['_node'] = '", "'")
                    For j = 0 To 9
                        tmpTelNr = .StringEntnehmen(Code, "['telcfg:settings/" & Port & "/Number" & j & "'] = '", "'")
                        If Not tmpTelNr = "-1" Then
                            If Not Len(tmpTelNr) = 0 Then
                                If Strings.Left(tmpTelNr, 3) = "SIP" Then
                                    tmpTelNr = SIP(CInt(Mid(tmpTelNr, 4, 1)))
                                Else
                                    tmpTelNr = .OrtsVorwahlEntfernen(tmpTelNr, Vorwahl)
                                End If
                                TelNr = tmpTelNr & "_" & TelNr
                            End If
                        End If
                    Next
                    If Not TelNr = vbNullString Then
                        TelNr = Strings.Left(TelNr, Len(TelNr) - 1)
                    End If

                    DialPort = "2" & Strings.Right(Port, 1)
                    AnzahlLANWLAN += 1
                    EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")

                    If bRausschreiben Then
                        setline("IP-Telefon gefunden: " & DialPort & ", " & TelNr & ", " & TelName)
                    Else
                        c_ini.Write(sDateiPfad, "Telefone", DialPort, ";" & TelNr & ";" & TelName)
                    End If

                End If
            Next

            Dim S0Typ As String
            ' S0-Port
            For i = 1 To 8
                TelName = .StringEntnehmen(Code, "['telcfg:settings/NTHotDialList/Name" & i & "'] = '", "'")
                If Not TelName = "-1" Then
                    If Not TelName = vbNullString Then
                        TelNr = .StringEntnehmen(Code, "['telcfg:settings/NTHotDialList/Number" & i & "'] = '", "'")
                        If Not TelNr = "-1" Then
                            DialPort = "5" & i

                            If bRausschreiben Then
                                setline("S0-Telefon gefunden: " & DialPort & ", " & ", " & TelNr & ", " & TelName)
                            Else
                                c_ini.Write(sDateiPfad, "Telefone", DialPort, TelNr & ";" & TelNr & ";" & TelName)
                            End If

                            S0Typ = .StringEntnehmen(Code, "['telcfg:settings/NTHotDialList/Type" & i & "'] = '", "'")
                            If Not TelNr = "-1" Then
                                Select Case S0Typ
                                    Case "Fax"
                                        EingerichteteFax = String.Concat(EingerichteteFax, DialPort, ";")
                                        If bRausschreiben Then setline("S0-telefon " & DialPort & " ist ein FAX.")
                                        'Case "Isdn"
                                        'Case "Fon"
                                        'Case Else
                                End Select

                            End If
                            AnzahlISDN += 1
                            EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")

                        End If
                    End If
                End If
            Next
            If Not AnzahlISDN = 0 Then
                DialPort = "50"
                EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")

                If bRausschreiben Then
                    setline("S0-Basis hinzugef�gt.")
                Else
                    c_ini.Write(sDateiPfad, "Telefone", DialPort, ";;ISDN-Basis")
                End If

            End If

            ' TAM
            tmpstrUser = Split(.StringEntnehmen(Code, "['tam:settings/TAM/list(" & .StringEntnehmen(Code, "['tam:settings/TAM/list(", ")'] = {") & ")'] = {", "}" & Chr(10) & "  },"), " },", , CompareMethod.Text)
            For Each Anrufbeantworter In tmpstrUser
                If .StringEntnehmen(Anrufbeantworter, "['Active'] = '", "'") = "1" Then
                    TelName = .StringEntnehmen(Anrufbeantworter, "['Name'] = '", "'")
                    Port = .StringEntnehmen(Anrufbeantworter, "['_node'] = '", "'")
                    TelNr = TAM(CInt(Strings.Right(Port, 1)))
                    AnzahlTAM += 1
                    DialPort = "60" & Strings.Right(Port, 1)
                    EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")

                    If bRausschreiben Then
                        setline("Anrufbeantworter gefunden: " & DialPort & ", " & ", " & TelNr & ", " & TelName)
                    Else
                        c_ini.Write(sDateiPfad, "Telefone", DialPort, ";" & TelNr & ";" & TelName)
                    End If

                End If
            Next


            ' integrierter Faxempfang

            DialPort = .StringEntnehmen(Code, "['telcfg:settings/FaxMailActive'] = '", "'")
            If DialPort = "1" Then
                TelNr = ""
                DialPort = "5"
                AnzahlFAX += 1
                EingerichteteTelefone = String.Concat(EingerichteteTelefone, DialPort, ";")

                EingerichteteFax = String.Concat(EingerichteteFax, DialPort, ";")

                If bRausschreiben Then
                    setline("Die integrierte Faxfunktion ist eingeschaltet: " & DialPort & ", " & TelNr & "," & "Faxempfang")
                Else
                    c_ini.Write(sDateiPfad, "Telefone", DialPort, ";" & TelNr & ";" & "Faxempfang")
                End If

            End If

            If Not EingerichteteFax Is Nothing Then
                EingerichteteFax = Strings.Left(EingerichteteFax, Strings.Len(EingerichteteFax) - 1)
                If Not bRausschreiben Then c_ini.Write(sDateiPfad, "Telefone", "EingerichteteFax", EingerichteteFax)
            End If

            Landesvorwahl = .StringEntnehmen(Code, "['country'] = '", "'")
            If Len(Landesvorwahl) > 2 Then
                c_ini.Write(sDateiPfad, "Optionen", "TBLandesVW", "0" & Landesvorwahl)
            End If

            EingerichteteTelefone = Strings.Left(EingerichteteTelefone, Strings.Len(EingerichteteTelefone) - 1)
            TelAnzahl = AnzahlDECT + AnzahlFAX + AnzahlFON123 + AnzahlISDN + AnzahlLANWLAN + AnzahlTAM
            If bRausschreiben Then
                setline("Anzahl FON: " & AnzahlFON123)
                setline("Anzahl DECT: " & AnzahlDECT)
                setline("Anzahl ISDN: " & AnzahlISDN)
                setline("Anzahl LANWLAN: " & AnzahlLANWLAN)
                setline("Anzahl TAM: " & AnzahlTAM)
                setline("Anzahl FAX: " & AnzahlFAX)
                setline("Gesamtanzahl: " & TelAnzahl)
            Else
                c_ini.Write(sDateiPfad, "Telefone", "EingerichteteTelefone", EingerichteteTelefone)
                c_ini.Write(sDateiPfad, "Telefone", "Anzahl", CStr(TelAnzahl))
            End If
        End With


    End Sub

    Function AlleNummern(ByVal MSN() As String, ByVal SIP() As String, ByVal TAM() As String, ByVal FAX() As String, ByVal POTS As String, ByVal Mobil As String) As String
        AlleNummern = vbNullString
        Dim max(MSN.Length + SIP.Length + TAM.Length + FAX.Length) As String
        Dim tmp() As String = Split(Strings.Join(MSN, ";") & ";" & Strings.Join(SIP, ";") & ";" & Strings.Join(TAM, ";") & ";" & Strings.Join(FAX, ";") & ";" & POTS & ";" & Mobil, ";", , CompareMethod.Text)
        Dim res = From x In tmp Select x Distinct 'Doppelte entfernen
        Dim res2 = From x In res Where Not x Like "" Select x ' Leere entfernen
        For Each Nr In res2
            AlleNummern = Nr & "_" & AlleNummern
        Next
        AlleNummern = Strings.Left(AlleNummern, Len(AlleNummern) - 1)
    End Function
#End Region

#Region "W�hlen"
    Friend Function SendDialRequestToBox(ByVal DialCode As String, ByVal DialPort As String, HangUp As Boolean) As String
        ' �bertr�gt die zum Verbindungsaufbau notwendigen Daten per WinHttp an die FritzBox
        ' Parameter:  dialCode (string):    zu w�hlende Nummer
        '             fonanschluss (long):  Welcher Anschluss wird verwendet?
        '             HangUp (bool):        Soll Verbindung abgebrochen werden
        ' R�ckgabewert (String):            Antworttext (Status)
        '
        Dim formdata As String             ' an die FritzBox zu sendende Daten
        Dim Response As String             ' Antwort der FritzBox
        Dim Link As String
        '
        SendDialRequestToBox = "Fehler!" & vbCrLf & "Entwickler kontaktieren."            ' Antwortstring
        If Not sSID = sDefaultSID And Len(sSID) = Len(sDefaultSID) Then
            Link = "http://" & sFBAddr & "/cgi-bin/webcm"
            formdata = "sid=" & sSID & "&getpage=&telcfg:settings/UseClickToDial=1&telcfg:settings/DialPort=" & DialPort & "&telcfg:command/" & CStr(IIf(HangUp, "Hangup", "Dial=" & DialCode))
            Response = c_hf.httpWrite(Link, formdata, FBEncoding)

            If Response = vbNullString Then
                SendDialRequestToBox = CStr(IIf(HangUp, "Verbindungsaufbau" & vbCrLf & "wurde abgebrochen!", "W�hle " & DialCode & vbCrLf & "Jetzt abheben!"))
            Else
                SendDialRequestToBox = "Fehler!" & vbCrLf & "Logfile beachten!"
                c_hf.LogFile("SendDialRequestToBox: Response: " & Response)
            End If
        Else
            c_hf.FBDB_MsgBox("Fehler bei dem Login. SessionID: " & sSID & "!", MsgBoxStyle.Critical, "sendDialRequestToBox")
        End If
    End Function
#End Region

#Region "Journalimort"

    Public Function DownloadAnrListe() As String
        Dim sLink(1) As String
        Dim ReturnString As String = vbNullString

        sSID = FBLogin(True)
        If Not sSID = sDefaultSID Then
            sLink(0) = "http://" & sFBAddr & "/fon_num/foncalls_list.lua?sid=" & sSID
            sLink(1) = "http://" & sFBAddr & "/fon_num/foncalls_list.lua?sid=" & sSID & "&csv="

            ReturnString = c_hf.httpRead(sLink(0), FBEncoding, FBFehler)
            If FBFehler Is Nothing Then
                If Not InStr(ReturnString, "Luacgi not readable", CompareMethod.Text) = 0 Then
                    sLink(0) = "http://" & sFBAddr & "/cgi-bin/webcm?sid=" & sSID & "&getpage=../html/de/menus/menu2.html&var:lang=de&var:menu=fon&var:pagename=foncalls"
                    c_hf.httpRead(sLink(0), FBEncoding, FBFehler)
                    sLink(1) = "http://" & sFBAddr & "/cgi-bin/webcm?sid=" & sSID & "&getpage=../html/de/FRITZ!Box_Anrufliste.csv"
                End If
                ReturnString = c_hf.httpRead(sLink(1), FBEncoding, FBFehler)
            Else
                c_hf.LogFile("FBError (DownloadAnrListe): " & Err.Number & " - " & Err.Description & " - " & sLink(0))
            End If
        Else
            c_hf.FBDB_MsgBox("Der Login in die Fritz!Box ist fehlgeschlagen" & vbCrLf & vbCrLf & "Die Anmeldedaten sind falsch oder es fehlt die Berechtigung f�r diesen Bereich.", MsgBoxStyle.Critical, "DownloadAnrListe_DoWork")
            c_hf.LogFile("Die Anmeldedaten sind falsch oder es fehlt die Berechtigung f�r diesen Bereich. (DownloadAnrListe_DoWork)")
        End If
        Return ReturnString
    End Function

#End Region

#Region "SetLine in Config"
    Private Sub setline(ByVal Zeile As String)
        If bRausschreiben Then form_Config.AddLine(Zeile)
    End Sub
#End Region

    Friend Function GetFBAddr() As String
        Return sFBAddr
    End Function

End Class
