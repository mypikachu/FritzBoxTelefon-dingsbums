 Friend SUB UploadAddressbook(BYVAL BookID AS STRING, BYVAL Adressbuch AS STRING)
  ' To do: Mehrere Telefonbucher sind möglich. Zugriff prüfen.
  ' http://www.ip-phone-forum.de/showthread.php?t=226605
  ' sPhonebookId = "1" ' 255 = Intern ' 256 = Clip Info ' 0 = Haupttelefonbuch
  ' sPhonebookExportName = "Test" ' muss mindestens ein Zeichen enthalten ab ID 1
  '
  DIM row AS STRING
  DIM cmd AS STRING
  DIM ReturnValue AS STRING
  DIM XMLFBAddressbuch AS XmlDocument

  IF SID = C_DP.P_Def_SessionID THEN FBLogin(True)
  IF NOT SID = C_DP.P_Def_SessionID AND LEN(SID) = LEN(C_DP.P_Def_SessionID) THEN

    row = "---" & 12345 + RND() * 16777216
    cmd = row & vbCrLf & "Content-Disposition: form-data; name=""sid""" & vbCrLf & vbCrLf & SID & vbCrLf _
    & row & vbCrLf & "Content-Disposition: form-data; name=""PhonebookId""" & vbCrLf & vbCrLf & BookID & vbCrLf _
    & row & vbCrLf & "Content-Disposition: form-data; name=""PhonebookImportFile""" & vbCrLf & vbCrLf & "@" + Adressbuch + ";type=text/xml" & vbCrLf _
    & row & "--" & vbCrLf

    WITH C_hf
      ReturnValue = .httpPOST(P_Link_FB_ExportAddressbook, cmd, FBEncoding)
      IF ReturnValue.StartsWith("<?xml") THEN
        XMLFBAddressbuch = NEW XmlDocument()
        TRY
          XMLFBAddressbuch.LoadXml(ReturnValue)
        CATCH ex AS Exception
          .LogFile(C_DP.P_Fehler_Export_Addressbuch)
        END TRY
      END IF
    END WITH

  ELSE

    C_hf.FBDB_MsgBox(C_DP.P_FritzBox_Dial_Error3(SID), MsgBoxStyle.Critical, "UploadAddressbook")

  END IF

 END SUB
