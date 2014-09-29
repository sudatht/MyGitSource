<%@language = "VBScript"%>
<%option explicit%>
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%
	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	dim rsConsultant
	dim oCon
	dim strSQL
	dim strVikarID
	dim strFornavn
	dim strEtternavn
	
	if (len(trim(Request("vikarid"))) = 0) then
		AddErrorMessage("Ingen vikar spesifisert.")
		call RenderErrorMessage()
	else
		strVikarID = Request("vikarid")
	end if
	
	' Open database connection
	Set oCon = GetConnection(GetConnectionstring(XIS, ""))	

	strSQL = "SELECT fornavn, etternavn FROM vikar WHERE vikarid = " & strVikarID
		
	set rsConsultant = GetFirehoseRS(strSQL, oCon)
	if (HasRows(rsConsultant) = false) then
		set rsConsultant = nothing
		AddErrorMessage("Vikar ikke funnet.")
		call RenderErrorMessage()
	else
		strFornavn = rsConsultant.fields("fornavn").value
		strEtternavn = rsConsultant.fields("etternavn").value
	end if
	rsConsultant.close
	set rsConsultant = nothing		
	CloseConnection(oCon)
	set oCon = nothing
%>
<HTML>
	<HEAD>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<TITLE>Velg fil du &oslash;nsker &aring; laste opp</TITLE>
	</HEAD>
	<script language="javaScript" type="text/javascript">
	function validateFile()
	{
		if (document.all.Ufile.value.length==0)
		{
			alert("Du må velge fil å laste opp!");
			return false;
		}
		return true;
	}
	</script>	
	<body>
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead">
				<h1>Last opp fil</h1>
			</div>
			<div class="content content2">
				<p>Velg fil du &oslash;nsker &aring; laste opp for vikar <strong><%= strFornavn & " " & strEtternavn%></strong>.</p>
				<form enctype="multipart/form-data" method="post" action="formresp.asp?vikarid=<%=Request("vikarid")%>" onSubmit="return(validateFile());">
					<input type="file" id="Ufile" name="Ufile">&nbsp;<input type="submit" Value="Last opp">
				</form>
			</div>
		</div>
	</body>
</html>