﻿Imports System.Drawing
Imports System.Collections.Generic

Public Class Wählclient
    Implements IDisposable

    Private C_DP As DataProvider
    Private frm_Wählbox As formWählbox
    Private C_hf As Helfer
    Private C_KF As Contacts
    Private C_GUI As GraphicalUserInterface
    Private C_OlI As OutlookInterface
    Private C_FBox As FritzBox
    Private C_Phoner As PhonerInterface
    Private C_XML As XML
    Friend _listFormWählbox As New List(Of formWählbox)


    Public Sub New(ByVal DataProviderKlasse As DataProvider, _
                   ByVal HelferKlasse As Helfer, _
                   ByVal KontaktKlasse As Contacts, _
                   ByVal InterfaceKlasse As GraphicalUserInterface, _
                   ByVal OutlInter As OutlookInterface, _
                   ByVal FritzBoxKlasse As FritzBox, _
                   ByVal PhonerKlasse As PhonerInterface, _
                   ByVal XMLKlasse As XML)
        C_hf = HelferKlasse
        C_KF = KontaktKlasse
        C_GUI = InterfaceKlasse
        C_DP = DataProviderKlasse
        C_OlI = OutlInter
        C_FBox = FritzBoxKlasse
        C_Phoner = PhonerKlasse
        C_XML = XMLKlasse
    End Sub

#Region "Alles was mit dem Wählen zu tun hat"
    Friend Sub WählboxStart(ByVal olAuswahl As Outlook.Selection)
        ' wird durch das Symbol 'Wählen' in der 'FritzBox'-Symbolleiste ausgeführt
        Dim olNamespace As Outlook.NameSpace
        Dim aktKontakt As Outlook.ContactItem       ' aktuell ausgewählter Kontakt
        Dim pos1 As Integer
        Dim pos2 As Integer
        Dim vCard As String

        olNamespace = C_OlI.OutlookApplication.GetNamespace("MAPI")
        ' Ist überhaupt etwas ausgewählt?
        If (olAuswahl.Count = 1) Then
            If TypeOf olAuswahl.Item(1) Is Outlook.MailItem Then
                ' Es wurde eine Mail ausgewählt
                ' Den zur Email-Adresse gehörigen Kontakt suchen
                Dim aktMail As Outlook.MailItem = CType(olAuswahl.Item(1), Outlook.MailItem)

                If Not aktMail.SenderEmailAddress = C_DP.P_Def_StringEmpty Then
                    aktKontakt = C_KF.KontaktSuche(KontaktID:=C_DP.P_Def_StringEmpty, _
                                                   StoreID:=C_DP.P_Def_StringEmpty, _
                                                   alleOrdner:=C_DP.P_CBKHO, _
                                                   TelNr:=C_DP.P_Def_StringEmpty, _
                                                   EMailAdresse:=aktMail.SenderEmailAddress)
                    If aktKontakt IsNot Nothing Then
                        Wählbox(aktKontakt, C_DP.P_Def_StringEmpty, C_DP.P_Def_StringEmpty, False)
                    Else
                        C_hf.FBDB_MsgBox("Es ist kein Kontakt mit der E-Mail-Adresse " & aktMail.SenderEmailAddress & " vorhanden!", MsgBoxStyle.Information, "WählboxStart")
                    End If
                    C_hf.NAR(aktKontakt)
                End If
                With C_hf
                    .NAR(olNamespace)
                    .NAR(aktMail)
                End With
                aktKontakt = Nothing
                olNamespace = Nothing
                aktMail = Nothing
            ElseIf TypeOf olAuswahl.Item(1) Is Outlook.ContactItem Then
                ' Es wurde gleich ein Kontakt gewählt!
                ' Nun direkt den Wähldialog für den Kontakt anzeigen.
                aktKontakt = CType(olAuswahl.Item(1), Outlook.ContactItem)
                Wählbox(aktKontakt, C_DP.P_Def_StringEmpty, C_DP.P_Def_StringEmpty, False)
                C_hf.NAR(aktKontakt)
                aktKontakt = Nothing
            ElseIf TypeOf olAuswahl.Item(1) Is Outlook.AppointmentItem Then
#If Not OVer = 15 Then
                Dim oAppItem As Outlook.AppointmentItem = CType(olAuswahl.Item(1), Outlook.AppointmentItem)
                Dim oAppLink As Outlook.Link
                Dim oAppThing As Object
                For Each oAppLink In oAppItem.Links
                    oAppThing = oAppLink.Item
                    If TypeOf oAppThing Is Outlook.ContactItem Then 'Nur, wenn der Link auf einen Kontakt zeigt....
                        Wählbox(CType(oAppThing, Outlook.ContactItem), C_DP.P_Def_StringEmpty, C_DP.P_Def_StringEmpty, False)
                        Exit For
                    End If
                Next 'oAppLink
                C_hf.NAR(oAppItem)
                oAppItem = Nothing
#End If
            ElseIf TypeOf olAuswahl.Item(1) Is Outlook.JournalItem Then
                ' Es wurde ein Journaleintrag gewählt!
                Dim aktJournal As Outlook.JournalItem = CType(olAuswahl.Item(1), Outlook.JournalItem)
                If InStr(aktJournal.Body, C_DP.P_Def_StringUnknown) = 0 _
                    And Not InStr(aktJournal.Categories, "FritzBox Anrufmonitor") = 0 Then
#If Not OVer = 15 Then
                    If Not aktJournal.Links.Count = 0 Then
                        Dim oAppLink As Outlook.Link
                        Dim oAppThing As Object
                        For Each oAppLink In aktJournal.Links
                            oAppThing = oAppLink.Item
                            If TypeOf oAppThing Is Outlook.ContactItem Then
                                Wählbox(CType(oAppThing, Outlook.ContactItem), Mid(aktJournal.Body, 11, InStr(1, aktJournal.Body, vbNewLine) - 11), C_DP.P_Def_StringEmpty, False)
                                Exit For
                            End If
                            C_hf.NAR(oAppThing)
                        Next 'oAppLink
                    Else
#End If
                        pos1 = InStr(1, aktJournal.Body, C_DP.P_Def_Begin_vCard, CompareMethod.Text)
                        pos2 = InStr(1, aktJournal.Body, C_DP.P_Def_End_vCard, CompareMethod.Text)
                        If Not pos1 = 0 And Not pos2 = 0 Then
                            pos2 = pos2 + Len(C_DP.P_Def_End_vCard)
                            vCard = Mid(aktJournal.Body, pos1, pos2 - pos1)
                        Else
                            vCard = C_DP.P_Def_StringEmpty
                        End If

                        Wählbox(Nothing, Mid(aktJournal.Body, 11, InStr(1, aktJournal.Body, vbNewLine) - 11), vCard, False)
#If Not OVer = 15 Then
                    End If
#End If

                End If
            Else
                C_hf.FBDB_MsgBox("Es muss entweder ein Kontakt, eine eingegangene E-Mail-Adresse oder ein Journaleintrag ausgewählt sein!", MsgBoxStyle.Exclamation, "WählboxStart")
            End If
        Else
            C_hf.FBDB_MsgBox("Es muss entweder ein Kontakt, eine eingegangene E-Mail-Adresse oder ein Journaleintrag ausgewählt sein!", MsgBoxStyle.Exclamation, "WählboxStart")
        End If
        olAuswahl = Nothing
        Exit Sub
    End Sub ' (WählboxStart)

    Friend Sub Wählbox(ByVal oContact As Outlook.ContactItem, ByVal TelNr As String, ByVal vCard As String, ByVal Direktwahl As Boolean)
        ' macht alle Eintragungen in 'formWählbox'
        ' aus FritzBoxDial übernommen und überarbeitet
        ' Parameter:  oContact (ContactItem): Kontaktdaten des Anzurufenden
        '             TelNr (String):         Telefonnummer des Anzurufenden
        Dim alleTelNr(14) As String ' alle im Kontakt enthaltenen Telefonnummern
        Dim alleNrTypen(14) As String ' die Bezeichnungen der Telefonnummern
        Dim i, iTelNr As Integer    ' Zählvariablen
        Dim ImgPath As String = C_DP.P_Def_StringEmpty   ' Position innerhalb eines Strings
        Dim row(2) As String

        frm_Wählbox = New formWählbox(Direktwahl, C_DP, C_hf, C_GUI, C_FBox, C_Phoner, C_KF, Me, C_XML)
        _listFormWählbox.Add(frm_Wählbox)

        If oContact Is Nothing Then
            frm_Wählbox.Tag = C_DP.P_Def_ErrorMinusOne_String & ";" & vCard ' C_DP.P_Def_ErrorMinusOne
        Else
            frm_Wählbox.Tag = oContact.EntryID & ";" & CType(oContact.Parent, Outlook.MAPIFolder).StoreID
        End If

        iTelNr = 1 'Index Zeile im DataGrid des Formulars
        ' Ist der Kontakt nicht vorhanden (z.B. Rückruf)?
        If oContact Is Nothing Then
            If Not Direktwahl Then
                ' Ortsvorwahl vor die Nummer setzen, falls eine Rufnummer nicht mit "0" beginnt und nicht mit "11"
                ' (Rufnummern die mit "11" beginnen sind Notrufnummern oder andere Sondernummern)
                If Not Left(C_hf.nurZiffern(TelNr), 1) = "0" And Not Left(C_hf.nurZiffern(TelNr), 2) = "11" Then TelNr = C_DP.P_TBVorwahl & TelNr

                frm_Wählbox.Text = "Anruf: " & CStr(IIf(vCard = C_DP.P_Def_StringEmpty Or vCard = C_DP.P_Def_ErrorMinusTwo_String, TelNr, ReadFNfromVCard(vCard)))
                ' Liste füllen
                row(0) = CStr(iTelNr) 'Index Zeile im DataGrid
                row(2) = TelNr
                frm_Wählbox.ListTel.Rows.Add(row)
            Else
                frm_Wählbox.Text = "Anruf: Direktwahl"
                frm_Wählbox.ContactImage.Visible = False
                frm_Wählbox.TelNrBox.Text = TelNr
            End If
        Else
            ' Welche Telefonnummerntypen sollen angezeigt werden?
            With oContact
                ' Fenstertitel setzen
                frm_Wählbox.Text = Replace("Anruf: " & .FullName & " (" & .CompanyName & ")", " ()", "", , , CompareMethod.Text)
                ' Die einzelnen Telefonnummern in ein Array packen
                ' Die deutsche Bezeichnung der Nummerntypen
                ' für Anzeigezwecke auch in ein Array packen.
                alleTelNr(1) = .AssistantTelephoneNumber : alleNrTypen(1) = "Assistent"
                alleTelNr(2) = .BusinessTelephoneNumber : alleNrTypen(2) = "Geschäftlich"
                alleTelNr(3) = .Business2TelephoneNumber : alleNrTypen(3) = "Geschäftlich2"
                alleTelNr(4) = .CallbackTelephoneNumber : alleNrTypen(4) = "Rückmeldung"
                alleTelNr(5) = .CarTelephoneNumber : alleNrTypen(5) = "Auto"
                alleTelNr(6) = .CompanyMainTelephoneNumber : alleNrTypen(6) = "Firma"
                alleTelNr(7) = .HomeTelephoneNumber : alleNrTypen(7) = "Privat"
                alleTelNr(8) = .Home2TelephoneNumber : alleNrTypen(8) = "Privat2"
                alleTelNr(9) = .ISDNNumber : alleNrTypen(9) = "ISDN"
                alleTelNr(10) = .MobileTelephoneNumber : alleNrTypen(10) = "Mobiltelefon"
                alleTelNr(11) = .OtherTelephoneNumber : alleNrTypen(11) = "Weitere"
                alleTelNr(12) = .PagerNumber : alleNrTypen(12) = "Pager"
                alleTelNr(13) = .PrimaryTelephoneNumber : alleNrTypen(13) = "Haupttelefon"
                alleTelNr(14) = .RadioTelephoneNumber : alleNrTypen(14) = "Funkruf"

                ' Kontaktbild anzeigen
                ImgPath = C_KF.KontaktBild(oContact)
                If Not ImgPath = C_DP.P_Def_StringEmpty Then
                    Dim orgImage As Image
                    Using fs As New IO.FileStream(ImgPath, IO.FileMode.Open)
                        orgImage = Image.FromStream(fs)
                    End Using
                    C_KF.DelKontaktBild(ImgPath)
                    With frm_Wählbox.ContactImage
                        Dim Bildgröße As New Size(.Width, CInt((.Width * orgImage.Size.Height) / orgImage.Size.Width))
                        Dim showImage As Image = New Bitmap(Bildgröße.Width, Bildgröße.Height)
                        Using g As Graphics = Graphics.FromImage(showImage)
                            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                            g.DrawImage(orgImage, 0, 0, Bildgröße.Width, Bildgröße.Height)
                            'g.Dispose()
                        End Using
                        .Image = showImage
                    End With
                Else
                    frm_Wählbox.ContactImage.Visible = False
                End If
            End With
            ' Liste füllen
            For i = LBound(alleTelNr) + 1 To UBound(alleTelNr)
                If Not alleTelNr(i) = C_DP.P_Def_StringEmpty Then
                    ' Wenn die Telefonnummer nicht leer ist, dann in die Liste hinzufügen
                    row(0) = CStr(iTelNr) 'Index wird eins hochgezählt
                    'Ortsvorwahl vor die Nummer setzen, falls eine Rufnummer nicht mit "0" beginnt und nicht mit "11"
                    '(Rufnummern die mit "11" beginnen sind Notrufnummern oder andere Sondernummern)
                    If Not Left(C_hf.nurZiffern(alleTelNr(i)), 1) = "0" And Not Left(C_hf.nurZiffern(alleTelNr(i)), 2) = "11" Then _
                        alleTelNr(i) = C_DP.P_TBVorwahl & alleTelNr(i)
                    If C_hf.nurZiffern(alleTelNr(i)) = C_hf.nurZiffern(TelNr) Then
                        row(1) = alleNrTypen(i) & " *"
                    Else
                        row(1) = alleNrTypen(i)
                    End If
                    row(2) = alleTelNr(i)
                    iTelNr += 1
                    frm_Wählbox.ListTel.Rows.Add(row)
                End If
            Next
            'VIP
            frm_Wählbox.BVIP.Enabled = True
            frm_Wählbox.BVIP.Checked = C_GUI.IsVIP(oContact)
        End If
        With frm_Wählbox
            If Not .ListTel.RowCount = 0 Or Direktwahl Then

                .Show()
                .BringToFront()
                If Direktwahl Then
                    .TelNrBox.Focus()
                    .AcceptButton = .ButtonWeiter
                Else
                    .Focus()
                End If

            Else
                C_hf.FBDB_MsgBox("Der Kontakt hat keine Telefonnummern.", MsgBoxStyle.Exclamation, "Wählbox")
            End If
        End With
        ' Wähldialog anzeigen

    End Sub '(Wählbox)

    Sub OnActionListen(ByVal index As String)
        Dim oContact As Outlook.ContactItem
        Dim Telefonat As String() = Split(index, ";", , CompareMethod.Text) ' ####List;ID
        ' KontaktID, StoreID, TelNr ermitteln
        Dim KontaktID As String
        Dim StoreID As String
        Dim TelNr As String
        'Dim Anrufer As String
        Dim vCard As String
        Dim ListNodeNames As New ArrayList
        Dim ListNodeValues As New ArrayList
        Dim xPathTeile As New ArrayList

        ' TelNr
        ListNodeNames.Add("TelNr")
        ListNodeValues.Add(C_DP.P_Def_ErrorMinusOne_String)

        ' Anrufer
        ListNodeNames.Add("Anrufer")
        ListNodeValues.Add(C_DP.P_Def_ErrorMinusOne_String)

        ' StoreID
        ListNodeNames.Add("StoreID")
        ListNodeValues.Add(C_DP.P_Def_ErrorMinusOne_String)

        ' KontaktID
        ListNodeNames.Add("KontaktID")
        ListNodeValues.Add(C_DP.P_Def_ErrorMinusOne_String & ";")

        ' vCard
        ListNodeNames.Add("vCard")
        ListNodeValues.Add(C_DP.P_Def_ErrorMinusOne_String & ";")

        With xPathTeile
            .Add(Telefonat(0))
            .Add("Eintrag")
        End With
        C_XML.ReadXMLNode(C_DP.XMLDoc, xPathTeile, ListNodeNames, ListNodeValues, "ID", Telefonat(1))

        'Anrufer = CStr(ListNodeValues.Item(ListNodeNames.IndexOf("Anrufer")))
        TelNr = CStr(ListNodeValues.Item(ListNodeNames.IndexOf("TelNr")))
        KontaktID = CStr(ListNodeValues.Item(ListNodeNames.IndexOf("KontaktID")))
        StoreID = CStr(ListNodeValues.Item(ListNodeNames.IndexOf("StoreID")))
        vCard = CStr(ListNodeValues.Item(ListNodeNames.IndexOf("vCard")))

        If Not StoreID = C_DP.P_Def_ErrorMinusOne_String Then
            'If Not KontaktID = C_DP.P_Def_ErrorMinusOne And Not StoreID = C_DP.P_Def_ErrorMinusOne Then
            oContact = C_KF.GetOutlookKontakt(KontaktID, StoreID)
            If oContact Is Nothing Then
                Select Case Telefonat(0)
                    Case C_DP.P_Def_NameListVIP
                        If C_hf.FBDB_MsgBox("Der zuwählende Kontakt wurde nicht gefunden. Er wurde möglicherweise gelöscht oder verschoben. Soll der zugehörige VIP-Eintrag entfernt werden?", MsgBoxStyle.YesNo, "OnActionAnrListen") = MsgBoxResult.Yes Then
                            C_GUI.RemoveVIP(KontaktID, StoreID)
                        End If
                    Case Else
                        C_hf.FBDB_MsgBox("Der zuwählende Kontakt wurde nicht gefunden. Er wurde möglicherweise gelöscht oder verschoben.", MsgBoxStyle.Critical, "OnActionAnrListen")
                End Select
            End If
        Else
            oContact = Nothing
        End If

        Wählbox(oContact, TelNr, vCard, False) '.TooltipText = TelNr. - .Caption = evtl. vorh. Name.
    End Sub

    ''' <summary>
    '''  Wird durch formAnrMon Button Rückruf (für das direkte Rückrufen des letzten Anrufers) ausgelöst.
    ''' </summary>
    ''' <param name="Telefonat">Hinterlegtes Telefonat</param>
    ''' <remarks></remarks>
    Friend Sub Rueckruf(ByVal Telefonat As C_Telefonat)
        With Telefonat
            Wählbox(.olContact, .TelNr, .vCard, False)
        End With
    End Sub

    Public Sub WählenAusInspector()
        'Mit diesem Makro ist es möglich direkt aus einem geöffneten Kontakt oder Journaleintrag zu wählen. ähnlich wählboxstart

        Dim olAuswahl As Outlook.Inspector ' das aktuelle Inspector-Fenster (Kontakt oder Journal)
        Dim TelNr As String    ' Telefonnummer des zu Suchenden
        Dim vCard As String
        Dim pos1 As Integer
        Dim pos2 As Integer
        Dim Absender As String
        Dim olContact As Outlook.ContactItem

        olAuswahl = ThisAddIn.P_oApp.ActiveInspector

        If TypeOf olAuswahl.CurrentItem Is Outlook.ContactItem Then ' ist aktuelles Fenster ein Kontakt?
            olContact = CType(olAuswahl.CurrentItem, Outlook.ContactItem)
            Wählbox(olContact, C_DP.P_Def_StringEmpty, C_DP.P_Def_StringEmpty, False)
            C_hf.NAR(olContact) : olContact = Nothing
        ElseIf TypeOf olAuswahl.CurrentItem Is Outlook.JournalItem Then ' ist aktuelles Fenster ein Journal?
            Dim olJournal As Outlook.JournalItem = CType(olAuswahl.CurrentItem, Outlook.JournalItem)
            If Not InStr(olJournal.Categories, "FritzBox Anrufmonitor") = 0 Then
                ' wurde der Eintrag vom Anrufmonitor angelegt?
                ' TelNr aus dem .Body entnehmen
                TelNr = Mid(olJournal.Body, 11, InStr(1, olJournal.Body, vbNewLine) - 11)
                If Not TelNr = C_DP.P_Def_StringUnknown Then
#If Not OVer = 15 Then
                    If Not olJournal.Links.Count = 0 Then 'KontaktID des darangehangenen Kontaktes ermitteln
                        Dim olLink As Outlook.Link = Nothing
                        For Each olLink In olJournal.Links
                            If TypeOf olLink.Item Is Outlook.ContactItem Then
                                olContact = CType(olLink.Item, Outlook.ContactItem)
                                Wählbox(olContact, C_DP.P_Def_StringEmpty, C_DP.P_Def_StringEmpty, False)
                                C_hf.NAR(olContact) ' : olContact = Nothing
                                Exit Sub
                            End If
                        Next
                        C_hf.NAR(olLink) : olLink = Nothing
                    Else ' Wenn in dem Journal kein Link hinterlegt ist, suche nach einer vCard im Body des Journaleintrags.
#End If
                        pos1 = InStr(1, olJournal.Body, C_DP.P_Def_Begin_vCard, CompareMethod.Text)
                        pos2 = InStr(1, olJournal.Body, C_DP.P_Def_End_vCard, CompareMethod.Text)
                        If Not pos1 = 0 And Not pos2 = 0 Then
                            pos2 = pos2 + Len(C_DP.P_Def_End_vCard)
                            vCard = Mid(olJournal.Body, pos1, pos2 - pos1)
                        Else
                            vCard = C_DP.P_Def_StringEmpty
                        End If

                        If Not TelNr Is C_DP.P_Def_StringEmpty Then Wählbox(Nothing, TelNr, vCard, False)
#If Not OVer = 15 Then
                    End If
#End If
                End If
            End If
        ElseIf TypeOf olAuswahl.CurrentItem Is Outlook.MailItem Then ' ist aktuelles Fenster ein Mail?
            Dim oContact As Outlook.ContactItem
            Dim olMail As Outlook.MailItem = CType(olAuswahl.CurrentItem, Outlook.MailItem)
            Absender = olMail.SenderEmailAddress
            ' Nun den zur Email-Adresse gehörigen Kontakt suchen
            If Not Absender = C_DP.P_Def_StringEmpty Then
                oContact = C_KF.KontaktSuche(C_DP.P_Def_StringEmpty, Absender, C_DP.P_Def_StringEmpty, C_DP.P_Def_StringEmpty, C_DP.P_CBKHO)
                If oContact IsNot Nothing Then
                    Wählbox(oContact, C_DP.P_Def_StringEmpty, C_DP.P_Def_StringEmpty, False)
                Else
                    C_hf.FBDB_MsgBox("Es ist kein Kontakt mit der E-Mail-Adresse " & Absender & " vorhanden!", MsgBoxStyle.Exclamation, "WählenAusKontakt")
                End If
            End If
        End If

    End Sub '(WählenAusKontakt)
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' So ermitteln Sie überflüssige Aufrufe

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                frm_Wählbox.Dispose()
                ' TODO: Verwalteten Zustand löschen (verwaltete Objekte).
            End If

            ' TODO: Nicht verwaltete Ressourcen (nicht verwaltete Objekte) freigeben und Finalize() unten überschreiben.
            ' TODO: Große Felder auf NULL festlegen.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: Finalize() nur überschreiben, wenn Dispose(ByVal disposing As Boolean) oben über Code zum Freigeben von nicht verwalteten Ressourcen verfügt.
    'Protected Overrides Sub Finalize()
    '    ' Ändern Sie diesen Code nicht. Fügen Sie oben in Dispose(ByVal disposing As Boolean) Bereinigungscode ein.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' Dieser Code wird von Visual Basic hinzugefügt, um das Dispose-Muster richtig zu implementieren.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Ändern Sie diesen Code nicht. Fügen Sie oben in Dispose(disposing As Boolean) Bereinigungscode ein.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
