<%
		dim Mailer		'Holds third party mail component
		dim RsConsultant	'Holds Consultants Contact information
		dim strKontaktnavn 	'as string
		dim strKontaktepost	'as string
		dim strKontaktDirTlf	'as string


		set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		strNavn = objCons.DataValues("Fornavn") + " " + objCons.DataValues("Etternavn")
		set RsConsultant = objCons.GetContact()
		if not RsConsultant.EOF and not RsConsultant.BOF then
			strKontaktnavn		= RsConsultant("Fornavn") + " " + RsConsultant("Etternavn")
			strKontaktepost		= RsConsultant("epost")
			if (len(trim(RsConsultant("dirtlf").value))>5) then
				strKontaktDirTlf = RsConsultant("dirtlf")
			else
				strKontaktDirTlf = Application("TlfReception")
			end if

		end if
		if strKontaktepost = "" or strKontaktnavn = "" then
			strKvittering = "<p class='warning'>Kommentar kunne ikke sendes til konsulenten.<br>Konsulentens kontaktperson har ikke e-post adresse.</p>"
		else
			if IsObject(Mailer) then
				Mailer.FromName    = strKontaktnavn
				Mailer.FromAddress = strKontaktepost
				Mailer.RemoteHost  = Application("XtraMailServer")
				Mailer.AddRecipient objCons.DataValues("Fornavn") & " " & objCons.DataValues("Etternavn"), objCons.DataValues("Epost")
				Mailer.Subject     = "Ang. vurdering av endret CV"
				strTekst = "Hei " & strNavn & "!" & vbCrLf & vbCrLf
				strTekst = strTekst & strKomm & vbCrLf & vbCrLf
				strTekst = strTekst & "Mvh" & vbCrLf & strKontaktnavn & vbCrLf & "Xtra " & vbCrLf & "Tlf " & strKontaktDirtlf
				Mailer.BodyText    = strTekst
				if Mailer.SendMail then
					strKvittering = "<p class='warning'>Kommentar er sendt til konsulenten.</p>"
				else
					strKvittering = "<p class='warning'>Kommentar kunne ikke sendes til konsulenten.</p>"
				end if
			else
				strKvittering = "<p class='warning'>Kommentar kunne ikke sendes til konsulenten.</p>"
			end if

		end if
		vis_kvittering = true
%>

