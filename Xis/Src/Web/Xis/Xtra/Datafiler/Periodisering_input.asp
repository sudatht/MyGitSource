<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>
<!--#INCLUDE FILE="../includes/SuperOffice.Constants.inc"-->
<!--#INCLUDE FILE="../includes/SuperOffice.Page.Navigation.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="../includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="../includes/xis.rights.inc"-->
<%
	If (HasUserRight(ACCESS_ADMIN, RIGHT_READ) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	dim Conn
	dim strSQL
	dim action
	dim title
	
	if(request("action")="as") then
		action = "../reports/AccruedReportByDepartment.aspx"
		title = "Periodisering  - AS"
	else
		action = "../reports/AccruedReportByDepartment.aspx" 
		title = "Periodisering - Vanlig"
	end if
	
	Set Conn = GetConnection(GetConnectionstring(XIS, ""))
%>
<html>
	<head>
		<title><%=title%></title>
		<link rel="stylesheet" href="/xtra/css/main.css" type="text/css" title="xtra intranett stylesheet">
		<link rel="stylesheet" href="/xtra/css/print.css" type="text/css" title="xtra intranett stylesheet" media="print">
		<script type="text/javascript" language="javascript" src="../js/javascript.js"></script>
	</head>
	<body onload="document.all.tildato.focus()">
		<div class="pageContainer" id="pageContainer">
			<div class="contentHead1">
				<h1><%=title%></h1>
			</div>
			<div class="content">
				<p>
					<FORM name="dato" ACTION="<%=action%>" method="post">
					
					<table>
					<tr>
					  <td>
						Tildato: <INPUT SIZE=6 NAME="tildato"  id="tildato" ONBLUR="dateCheck(this.form, this.name)">
						</td>
						<td>
						Avdelingskontor:
							<select NAME="dbxAvdeling" ID="lnk5">
								<option VALUE='0' selected>Alle Avdelinger</option>
								
								<%
								Dim rsAvdeling 'as adodb.recordset
								strSQL = "SELECT avdelingid, avdeling FROM avdeling ORDER BY avdelingid"
								Set rsAvdeling = GetFirehoseRS(strSQL, Conn)
								Do Until rsAvdeling.EOF
									Response.Write "<option value='" & rsAvdeling("avdelingID") & "'>" & rsAvdeling("avdeling") & "</option>"
									rsAvdeling.MoveNext
								Loop
								' Close and release recordset
								rsAvdeling.Close
								Set rsAvdeling = Nothing
								%>
							</select>
							
							</td>
					<td>	Ikke lønnet: <INPUT CHECKED class="CHECKBOX" id="checkbox1" name="ikke_loennet" type="checkbox"> </td>
            		<td>    Ikke fakturert: <INPUT CHECKED class="CHECKBOX" id="checkbox2" name="ikke_fakt" type="checkbox"></td>
							<td>	<INPUT TYPE=submit VALUE="Hent" id="view" onClick="displayProgressBar();"> </td>
                    
     </tr>
	 </table>
      </form>
	  
				
  </div>
		</div>
		<table  style="width:100%;height:80%;" border_width="0">
      <tr>
        <td style="width:100%;height:80%;vertical-align:middle;text-align:center"  align="center"> <input type="image" id="prgBar" name="prgBar" border="none" height="60px" width="60px" src="" style="display:none;border-width:0px;" /></td>
      </tr>
    </table>
	</body>
</html>
<%
CloseConnection(Conn)
set Conn = nothing
%>