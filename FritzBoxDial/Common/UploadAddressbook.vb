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
 
'
'
FUNCTION HTTPTransferRtMpfd(sMode AS STRING, sLink AS STRING, sFormdata AS STRING, OPT sRow AS STRING) AS STRING
' Link: Link zur Webpage
' Mode "Get" oder "Post"
' Microsoft WinHTTP Services, version 5.1
' PowerBasic PB/WIN 9
'
 LOCAL v1 AS VARIANT, v2 AS VARIANT, v3 AS VARIANT, v4 AS VARIANT
 LOCAL nL1 AS LONG, nL2 AS LONG
 LOCAL sHost AS STRING
'
 DIM objXMLHTTP AS DISPATCH
'
 SET objXMLHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
'
 sHost = MID$(sLink, LEN("://") + INSTR(sLink, "://") , INSTR(LEN("://") + INSTR(sLink, "://"),sLink, "/") - (LEN("://") + INSTR(sLink, "://")))
'
 v1 = sMode ' "GET " oder "POST "
 v2 = sLink ' "http://192.168.2.1/html/top_start_passwort.htm"
 v3 = 0
'
 OBJECT CALL objXMLHTTP.OPEN(v1, v2, v3)
'
' v1 = "Content-Type"
' v2 = "application/x-www-form-urlencoded"
' OBJECT CALL objXMLHTTP.setRequestHeader(v1, v2)
'
 IF TALLY(sHost, ".") = 3 THEN
' MSGBOX sHost
  v1 = "HOST" : v2 = sHost
  OBJECT CALL objXMLHTTP.setRequestHeader (v1, v2)
 END IF
 v1 = "Connection" : v2 = "Keep-Alive"
 OBJECT CALL objXMLHTTP.setRequestHeader(v1, v2)
 IF sRow = "" THEN
  v1 = "Content-Type" : v2 = "application/x-www-form-urlencoded"
 ELSE
  v1 = "Content-Type" : v2 = "multipart/form-data; boundary=" + sRow
 END IF
 OBJECT CALL objXMLHTTP.setRequestHeader(v1, v2)
 v1 = "Content-Length" : v2 = LEN(sFormdata)
 OBJECT CALL objXMLHTTP.setRequestHeader(v1, v2)
'
 sRow = ""
'
 v1 = sFormdata ' ""
 OBJECT CALL objXMLHTTP.SEND(v1)
'
 OBJECT GET objXMLHTTP.STATUS TO v2
 OBJECT GET objXMLHTTP.readyState TO v3
 '
 nL1 = VARIANT#(v3) : nL2 = VARIANT#(v2)
'
 IF nL1 = 0 AND nL2 = 200 THEN ' "WinHttp.WinHttpRequest.5.1"
'
  OBJECT GET objXMLHTTP.getAllResponseHeaders TO v1 : sHost = VARIANT$(v1)
  IF sHost <> "" THEN sRow = sHost
'  MSGBOX format$(VARIANTVT(v1)) + $crlf + sHost
'
  OBJECT GET objXMLHTTP.responseText TO v1
'  MSGBOX FORMAT$(VARIANTVT(v1)) + $CRLF + sHost
'
  FUNCTION = VARIANT$(v1)
'
 END IF
'
 SET objXMLHTTP = NOTHING
'
END FUNCTION
'
'        
