<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<%
profil = Session("Profil")

If  Mid(profil,2,1) > 1 Then 
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionTimeOut = Session("xtra_ConnectionTimeout")
	Conn.CommandTimeOut = Session("xtra_CommandTimeout")
	Conn.Open Session("xtra_ConnectionString"), Session("xtra_RuntimeUserName"), Session("xtra_RuntimePassword")

	if Request.QueryString("FirmaID") <> "" Then
	   Set rsFirma  = Conn.Execute("Select FirmaID, Firma, OrgNr, AnsvmedID, BransjeID, StatusID, KategoriID, Kreditgrense, telefon, fax, epost, hjemmeside, kredittopplysning from Suspect where FirmaID = " &Request.QueryString("FirmaID"))

	   strFirmaID		= rsFirma("FirmaID")
	   strFirm			= rsFirma("Firma") 
	   strOrgNo			= rsFirma("OrgNr")
	   lBransjeID		= rsFirma("BransjeID")
	   strTelefon		= rsFirma("Telefon")
	   strFax			= rsFirma("Fax") 
	   strHjemmeside	= rsFirma("Hjemmeside") 

	   ' Set heading
	   strHeading="Endre Suspect  " & rsFirma("Firma")

	    ' Close and release recordset
	    rsFirma.Close
	    Set rsFirma = Nothing
	   
	   Set rsAdresse = Conn.Execute("Select adrId, adressetype, adresse , Postnr, Poststed from suspect_ADRESSE where adresseRelID = " & Request.QueryString("FirmaID") & " and adressetype=1" )

	   ' Set adress values
	   strAdrID  = rsAdresse("AdrID")
	   strAdress = rsAdresse("Adresse")
	   strPostnr   = rsadresse("Postnr")
	   strPostSted = rsAdresse("Poststed")

	    ' Close and release recordset
	    rsAdresse.Close
	    Set rsAdresse = Nothing

	   Set rsKontakt = Conn.Execute("Select KontaktId, Etternavn, Fornavn, Telefon, Fax, Epost, Notat from SUSPECT_KONTAKT where FirmaID = " & Request.QueryString("FirmaID")  )

	   ' Set kontakt values
	   strKontaktID  = rsKontakt("KontaktId")
	   strEtternavn = rsKontakt("Etternavn")
	   strFornavn = rsKontakt("Fornavn")
	   strKtelefon = rsKontakt("Telefon")
	   strKFax = rsKontakt("Fax")
	   strKEPost = rsKontakt("EPost")
	   strNotat = rsKontakt("Notat")

	    ' Close and release recordset
	    rsKontakt.Close
	    Set rsKontakt = Nothing
	Else
	   strHeading = "Ny Suspect"
	End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <meta http-equiv="Content-Script-Type" content="text/javascript">
    <meta name="Developer" content="Electric Farm ASA">
	<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
	<title><%=strHeading %></title>
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
</head>
<script language="javaScript" type="text/javascript">
//###############globale variabler################################
var i=0;

function focused(f){
	i=f.substring(3,6);
	i=parseInt(i);
}

//#############lager felles variabler#############################
function shortKey(e) {			
	var keyChar = String.fromCharCode(event.keyCode);
	var modKey  = event.ctrlKey;
	var modKey2 = event.shiftKey;

	//#############linker i submeny######################################
	if (modKey && modKey2 && keyChar=="S"){	
		location=("kundesoek.asp");
	}
	if (modKey && modKey2 && keyChar=="Y"){	
		location=("kundeny.asp");
	}
	if (modKey && modKey2 && keyChar=="U"){	
		location=("suspectny.asp");
	}
	if (modKey && modKey2 && keyChar=="C"){	
		location=("suspectsoek.asp");
	}
	//########## menyer i toppframe#####################################
	if (modKey && modKey2 && keyChar=="H"){	
		parent.frames[1].location=("hotlistSub.asp");
	}
	if (modKey && modKey2 && keyChar=="K"){	
		parent.frames[1].location=("kundeSub.asp");
	}
	if (modKey && modKey2 && keyChar=="V"){	
		parent.frames[1].location=("vikarSub.asp");
	}
	if (modKey && modKey2 && keyChar=="O"){	
		parent.frames[1].location=("oppdragSub.asp");
	}
	if (modKey && modKey2 && keyChar=="R"){	
		parent.frames[1].location=("rapportSub.asp");
	}
	if (modKey && modKey2 && keyChar=="A"){	
		parent.frames[1].location=("adminSub.asp");
	}
	//#############scrolling på linkene#################################
	if (modKey && modKey2 && keyChar=="N"){	
		i=i+1;
		lnk='lnk'+i;
		if (document.all(lnk)){
			document.all(lnk).focus();
		}else{
			i=i-1;
			lnk='lnk'+i;
			document.all(lnk).focus();
		}
	}
	if (modKey && modKey2 && keyChar=="L"){	
		if (i>0){
			i=i-1;
			lnk='lnk'+i;
		}else{
			i=0;
			lnk='lnk'+i;
		}
		document.all(lnk).focus();
	}
	//############# klikk på søkeknapp/lagreknapp #################################
	if (modKey && modKey2 && keyChar=="Q"){	
		document.all('lnk19').click();
	}
}
//########### her catches eventen som trigger shortcut'en###########
document.onkeydown = shortKey;

//###############kode for å sette fokus ved onLoad()################
function fokus(){
	if(document.all('lnk0')){
		document.all('lnk0').focus();
	}
}
</script>
</head>

<body onLoad="fokus()">
	<div class="pageContainer" id="pageContainer">
<form ACTION="suspectdb.asp" METHOD="POST">
<input NAME="tbxFirmaID" TYPE="HIDDEN" VALUE="<%=strFirmaID%>">
<input NAME="tbxAdrID" TYPE="HIDDEN" VALUE="<%=strAdrID%>">
<input NAME="dbxAdrType" TYPE="HIDDEN" Value="1"> 
<input NAME="tbxKontaktID" TYPE="HIDDEN" VALUE="<%=strKontaktID%>">

<table cellpadding='0' cellspacing='0'>
<tr>
<td COLSPAN="8"><strong><u>Kontakt</u></strong></td>
<tr>
<td><font color="green">Navn:</font></td>
<td><input NAME="tbxFirm" ID="lnk0" onFocus="focused(this.id)" TYPE="TEXT" SIZE="30" MAXLENGTH="50" Value="<%=strFirm %>"></td>
<td>Org.No:</td>
<td><input NAME="tbxOrgNo" ID="lnk1" onFocus="focused(this.id)" TYPE="TEXT" SIZE="10" MAXLENGTH="11" Value="<%=strOrgNo %>"></td>
<tr>
<td><font color="green">Besøksadresse:</font></td>
<td><input NAME="tbxAdresse" ID="lnk2" onFocus="focused(this.id)" TYPE="TEXT" SIZE="30" MAXLENGTH="50" Value="<%=strAdress %>"></td>
<td>Postnr/Sted:</td>
<td NOWRAP>
<input NAME="tbxPostnr" ID="lnk3" onFocus="focused(this.id)" TYPE="TEXT" SIZE="5" MAXLENGTH="5" Value="<%=strPostnr%>">
<input NAME="tbxPoststed" ID="lnk4" onFocus="focused(this.id)" TYPE="TEXT" SIZE="20" MAXLENGTH="50" Value="<%=strPoststed%>"></td>

<tr>
<td>Hjemmeside:</td>
<td><input NAME="tbxHjemmeside" ID="lnk5" onFocus="focused(this.id)" TYPE="TEXT" SIZE="30" MAXLENGTH="50" VALUE="<%=strHjemmeside %>"></td>
<td>Bransje:</td>
<td><select NAME="dbxBransje" ID="lnk6" onFocus="focused(this.id)">
    <option VALUE="0">
<% 
	Set rsBransje = Conn.Execute("Select bransjeid, bransje from h_firma_bransje")
	Do Until rsBransje.EOF
	   If rsBransje("BransjeID") = lBransjeID Then
	      strValueSelected = rsBransje("BransjeID") & " SELECTED"
	   Else
	      strValueSelected = rsBransje("BransjeID")
	   End If  
	%>
	    <option VALUE="<%=strValueSelected %>"><%=rsBransje("Bransje") %>
	<% 
		rsBransje.MoveNext
	Loop

	' Close and release recordset
	rsBransje.Close
	Set rsBransje = Nothing

%>
</select></td>
<tr>
<td>Telefon:</td>
<td><input NAME="tbxTelefon" ID="lnk7" onFocus="focused(this.id)" TYPE="TEXT" SIZE="15" MAXLENGTH="20" VALUE="<%=strTelefon %>"></td>
<td>Fax:</td>
<td><input NAME="tbxFax" ID="lnk8" onFocus="focused(this.id)" TYPE="TEXT" SIZE="15" MAXLENGTH="20" VALUE="<%=strFax %>"></td>
<tr>
<td COLSPAN="3"><hl> </td>
<tr>
<td COLSPAN="8"><strong><u>Kontaktperson</u></strong></td>
<tr>
<td>Fornavn:
<td><input NAME="tbxFornavn" Size="20" ID="lnk9" onFocus="focused(this.id)" TYPE="TEXT" Value="<% =strForNavn %>">
<td>Etternavn:
<td><input NAME="tbxEtternavn" Size="20" ID="lnk10" onFocus="focused(this.id)" TYPE="TEXT" Value="<% =strEtternavn %>">
<tr>
<td>Telefon:</td>
<td><input NAME="tbxKTelefon" ID="lnk11" onFocus="focused(this.id)" TYPE="TEXT" SIZE="15" MAXLENGTH="20" VALUE="<% =strKTelefon %>"></td>
<td>Fax:</td>
<td><input NAME="tbxKFax" ID="lnk12" onFocus="focused(this.id)" TYPE="TEXT" SIZE="15" MAXLENGTH="20" VALUE="<% =strKFax %>"></td>
<tr>
<td>E-Post:</td>
<td><input NAME="tbxKEPost" ID="lnk13" onFocus="focused(this.id)" TYPE="TEXT" SIZE="30" MAXLENGTH="50" VALUE="<% =strKEPost %>"></td>
<tr>
<td COLSPAN="8"><strong><u>Til Xtra </u></strong></td>
<tr>
<td>Melding:</td>
<td COLSPAN="5"><textarea NAME="tbxMemo" ID="lnk14" onFocus="focused(this.id)" COLS="60" ROWS="3"><%=strNotat%></textarea></td>
<tr>
<td></td>
<td COLSPAN="4">Ønsker å bli kontaktet pr :
      <input NAME="rbnKontaktkode" ID="lnk15" onFocus="focused(this.id)" TYPE="RADIO" VALUE="1" <%=strEPost%>>E-Post
      <input NAME="rbnKontaktkode" ID="lnk16" onFocus="focused(this.id)" TYPE="RADIO" VALUE="2" <%=strTelefon%>>Telefon
      <input NAME="rbnKontaktkode" ID="lnk17" onFocus="focused(this.id)" TYPE="RADIO" VALUE="3" <%=strFax%>>Fax
      <input NAME="rbnKontaktkode" ID="lnk18" onFocus="focused(this.id)" TYPE="RADIO" VALUE="3" <%=strBrev%>>Brev
</td>

</table>

<table cellpadding='0' cellspacing='0'>
	<% 
	If Request.QueryString("FirmaID")="" Then
	   Response.write "<td><INPUT NAME=pbDataAction ID=lnk19 onFocus="&"focused(this.id)"&" TYPE=SUBMIT  VALUE=  Lagre  >"
	   Response.write "<td><INPUT NAME=pbDataAction ID=lnk20 onFocus="&"focused(this.id)"&" TYPE=SUBMIT  VALUE=Nullstill>"
	   
	Else
	   
	   Response.write "<td><INPUT NAME=pbDataAction ID=lnk19 onFocus="&"focused(this.id)"&" TYPE=SUBMIT  VALUE=  Lagre  >"
	   Response.write "<td><INPUT NAME=pbDataAction ID=lnk20 onFocus="&"focused(this.id)"&" TYPE=SUBMIT  VALUE=  Slette >"
	   Response.write "<td><INPUT NAME=pbDataAction ID=lnk21 onFocus="&"focused(this.id)"&" TYPE=SUBMIT  VALUE=Nullstill>"
	End If
	%>
</table>
</form>

    </div>
</body>
</html>

<% 
	Else
		Response.Redirect("default.asp") 
	End If 
%>