<%@ LANGUAGE="VBSCRIPT" %>
<%Option explicit%>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="includes\SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.inc"-->
<!--#INCLUDE FILE="includes\Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes\xis.rights.inc"-->
<!--#INCLUDE FILE="includes\Xis.HTML.Error.inc"-->
<%

	If (HasUserRight(ACCESS_CONSULTANT, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if

	Dim LngVikarId			'Holds consultant id  
	Dim lKompetanseID		'Holds qualification/jobwish id  
	Dim intTypeID			'Holds either 3 (qualification) or 4 (jobwish)
	Dim Conn
	Dim RsVikar				'as adodb.recordset, holds consultant information
	Dim rsRank
	Dim rsExperience
	Dim intExperience
	Dim rsEducation
	Dim intEducation
	Dim rsKompetanse		'as adodb.recordset, holds qualification information
	Dim rsLevel				'as adodb.recordset, holds qualification level information
	Dim StrSQL				'holds qualification/jobwish SQL
	Dim StrHeading			'holds heading of page
	Dim StrTitle			'holds Title of page
	Dim strVikarName		'holds name of consultant
	Dim strKTittel			'Title of jobwish/kompetanse
	Dim strKommentar 
	Dim strSelected
	Dim Sel
	Dim lntTittelID    
	Dim lntLevelID    
	Dim lntRangering

	dim blnShowHotList
	'Hotlist menu items variables
	dim strAddToHotlistLink
	dim strHotlistType

	'Consultant menu variables
	dim strClass
	dim strJSEvents

	' Move parameters to local variables
	LngVikarId      = clng(Request.Querystring("VikarID"))
	lKompetanseID	= Request.Querystring("KompetanseID")
	intTypeID       = CLng( Request.Querystring("TypeID") )

	'Verify VikarID
	If LngVikarId = "" Then
	   AddErrorMessage("Feil: VikarId mangler!")
		call RenderErrorMessage()
	End If

	' Get a database connection
	Set Conn = GetClientConnection(GetConnectionstring(XIS, ""))	

	' Get vikar name
	strSQL = "SELECT Etternavn, Fornavn FROM Vikar WHERE VikarID = " & LngVikarId 
	Set rsVikar = GetFirehoseRS(strSQL, Conn)
	strVikarName = rsVikar("Fornavn") & " " & rsVikar("Etternavn")
	'Release recordset
	rsVikar.close
	Set rsVikar = Nothing

	' Build heading
	if intTypeID = 3 then
		strHeading = "Produktkompetanse for <a href=""vikarVis.asp?VikarID=" & LngVikarid & """>" & strVikarName & "</a>"
		StrTitle = "Kompetanse"
	elseif intTypeID = 4 then
		strHeading = "Fagkompetanse for <a href=""vikarVis.asp?VikarID=" & LngVikarid & """>" & strVikarName & "</a>"
		StrTitle = "Fagkompetanse"
	end if

	'Retrieve selected kompetanse / jobwish from DB
	strSQL = "SELECT [VK].[K_LevelID], [VK].[K_TittelID], [VK].[Rangering],  " & _
	"[VK].[K_TypeID], [VK].[Kommentar] , [KT].[kTittel], " & _
	"[VK].[Relevant_Education], [VK].[Relevant_WorkExperience] " & _	
	"FROM [VIKAR_KOMPETANSE] AS [VK] " & _
	"INNER JOIN [H_Komp_Tittel] AS [KT] ON [VK].[K_TittelID] = [KT].[K_TittelID] " & _
	"WHERE [KompetanseID] = " & lKompetanseID
	
	Set rsKompetanse = GetFirehoseRS(strSQL, Conn)
   
   ' No records found ?
	If (HasRows(rsKompetanse) = false) Then 
		Set rsKompetanse = Nothing
		CloseConnection(Conn)
		set Conn = nothing
		AddErrorMessage("Feil: Feil i Parameter! Kompetanse/fagkompetanse eksisterer ikke!")
		call RenderErrorMessage()	
	End If

   ' move selected kompetanse/jobwish to variables
   lntTittelID		= rsKompetanse("K_TittelID")
   lntLevelID		= rsKompetanse("K_LevelID")
   lntRangering		= rsKompetanse("Rangering")
   strKommentar		= rsKompetanse("Kommentar")
   strKTittel		= rsKompetanse("KTittel")
   intEducation		= rsKompetanse("Relevant_Education")
   intExperience	= rsKompetanse("Relevant_WorkExperience")   
   
   ' Close and release recordset
   rsKompetanse.close
   Set rsKompetanse = Nothing

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"
    "http://www.w3.org/TR/REC-html40/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <meta http-equiv="Content-Script-Type" content="text/javascript">
    <meta name="Developer" content="Electric Farm ASA">
	<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
	<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
	<script type="text/javascript" src="/xtra/js/contentMenu.js"></script>
	<script type="text/javascript" src="/xtra/js/fontSizer.js"></script>	
	<title><%=StrTitle%></title>
</head>
<body>
	<div class="pageContainer" id="pageContainer">
		<form action="kompetanseDB.asp" method="post" name="frmKompetanse">
			<input TYPE="HIDDEN" NAME="tbxVikarID" VALUE="<%=LngVikarId%>">
			<input TYPE="HIDDEN" NAME="tbxKompetanseID" VALUE="<%=lKompetanseID%>">
			<input TYPE="HIDDEN" NAME="tbxTypeID" VALUE="<%=intTypeID%>">
			<input TYPE="HIDDEN" NAME="tbxAction" VALUE="lagre">
			<div class="contentHead1">
				<h1><%=strHeading%></h1>
				<div class="contentMenu">
					<table cellpadding="0" cellspacing="0" width="96%">
						<tr>
							<td>
								<%
								strClass = "menu"
								strJSEvents = ""
								%>
								<table cellpadding="0" cellspacing="2">
									<tr>
										<td class="<%=strClass%>" id="menu1" <%=strJSEvents%>>
											<a onClick="javascript:document.all.tbxAction.value='lagre';document.all.frmKompetanse.submit();" href="#" title="Lagre kompetanse">
											<img src="/xtra/images/icon_save.gif" width="18" height="15" alt="" align="absmiddle">Lagre</a>
										</td>
										<td class="<%=strClass%>" id="menu2" <%=strJSEvents%>>
											<a onClick="javascript:document.all.tbxAction.value='slette';document.all.frmKompetanse.submit();" href="#" title="Slette kompetanse">										
											Slette</a>
										</td>
									</tr>
								</table>
							</td>
							<td class="right">
							<!--#include file="Includes/contentToolsMenu.asp"-->
							</td>
						</tr>
					</table>
				</div>
			</div>
			<div class="content">
				<div class="listing">
					<p><strong>Type: <%=strKTittel%></strong></p>
						<table>
							<%
							if IntTypeID = 3 then
								%>
								<tr>
									<th>Rangering:</th>
									<th>Kommentar:</th>
								</tr>
								<tr>
									<td>
										<select NAME="tbxRangering">
											<Option VALUE="0"></Option>
											<%
											' Get Kompetanse level
											strSQL = "SELECT DISTINCT K_RangeringID, K_rangering FROM H_KOMP_RANGERING"
											Set rsRank = GetFirehoseRS(strSQL, Conn)
												while (not rsRank.EOF)
													If rsRank.fields("K_RangeringID") = lntRangering Then
														sel = "SELECTED"
													Else
														sel = ""
													End If
													%>
													<option VALUE="<%=rsRank.fields("K_RangeringID")%>"<%=sel%>><%=rsRank.fields("K_rangering").value%>
													<% 
												rsRank.MoveNext
												Wend
											' Close and release recordset
												rsRank.close
												Set rsRank = Nothing
											%>
										</select>					
									</td>
									<%
								elseif IntTypeID = 4 then
									%>
									<tr>
										<th>Erfaring:</th>
										<th>Utdannelse:</th>
										<th>Kommentar:</th>
									</tr>
									<tr>
										<td>
											<select NAME="tbxExperience">
												<Option value="0"></Option>
												<%
												' Get all experience levels
												strSQL = "SELECT distinct [K_ErfaringsNivaaID], [K_Erfaring] FROM [H_KOMP_ERFARINGNIVAA]"
												Set rsExperience = GetFirehoseRS(strSQL, Conn )
												while (not rsExperience.EOF)
													If rsExperience.fields("K_ErfaringsNivaaID") = intExperience Then
														sel = "SELECTED"
													Else
														sel = ""
													End If
													%>
													<option VALUE="<%=rsExperience.fields("K_ErfaringsNivaaID")%>"<%=sel%>><%=rsExperience.fields("K_Erfaring").value%>
													<% 
												rsExperience.MoveNext
												Wend
												' Close and release recordset
												rsExperience.close
												Set rsExperience = Nothing 
												%>
											</select>
										</td>
										<td>
											<select name="tbxEducation">
												<option value="0"></option>
												<%
												' Get all experience levels
												strSQL = "SELECT distinct [K_UNivaaID], [k_nivaa] FROM [H_KOMP_UTDANNELSENIVA]"
												Set rsEducation = GetFirehoseRS(strSQL, Conn)
												while (not rsEducation.EOF)
													If rsEducation.fields("K_UNivaaID") = intEducation Then
														sel = "SELECTED"
													Else
														sel = ""
													End If
													%>
													<option VALUE="<%=rsEducation.fields("K_UNivaaID")%>"<%=sel%>><%=rsEducation.fields("k_nivaa").value%>
													<% 
													rsEducation.MoveNext
												Wend
												' Close and release recordset
												rsEducation.close
												Set rsEducation = Nothing
												%>
											</select>					
										</td>
										<%
									end if
									%>
									<td><input NAME="tbxKommentar" TYPE="TEXT" SIZE="30" VALUE="<%=strKommentar %>"></td>
								</tr>
						</table>
					</div>
				</div>
			</form>
		</div>
	</body>
</html>

