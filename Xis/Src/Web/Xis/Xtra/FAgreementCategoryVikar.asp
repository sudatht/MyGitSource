<%@ LANGUAGE="VBSCRIPT" %>
<!--#INCLUDE FILE="includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Renderfunctions.inc"-->

<%

	dim vikar
	dim contactno
	dim contactname
	dim vikarid
	dim oppdragno
	dim faid
	dim cpaction
	
	vikar = Request("vikarName")
	vikarid = Request("VikarId")
	oppdragno = Request("OppdragNr")
	faid = Request("FaId")
	cpaction = Request("cpaction")

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
<link href="css/styles.css" rel="stylesheet" type="text/css" />

</head>

<body>


<!--START title box -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td class="des_main_title_box"><span class="txt_main_tiltle">Tilordne kategori</span> </td>
	<td width="136"><img src="images/logo.jpg" width="136" height="36" /></td>
  </tr>
</table>
<!--END title box -->

<!--START content box -->
<div class="des_content_box">
	<form Name="SaveVikarCategory" action="FAgreementCategoryDB.asp" METHOD="post" ID="Form2">
	<input name="OppdragID" type="HIDDEN" value="<%=oppdragno%>" id="Hidden2">

	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
	  <td width="150" height="20" class="txt_bold">AnsattNummer:<br /></td>
	  <td width="200"><%=vikarid%></td>
	  <td width="150" class="txt_bold">Navn:</td>
	  <td><%=vikar%></td>
	</tr>
	</table>
	
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
			<td width="150" height="20" class="txt_bold">Kategori:<br /></td>
			<td>
			<select name="dbxCategory" id="dbxCategory" class="form_list_menu" style="width: 225px">
		
				<%
				if isNull(faid) then
     					response.write GetActiveCategoryList(0)
     				Else
     					response.write GetActiveCategoryList(Clng(faid))
     				End If
				%>
			</select>
			</td>
			</tr>      

		    
		</table>
	
	<table>
		
		<INPUT TYPE="hidden" NAME="hdnCpAction" VALUE="<%=cpaction%>"  ID="hdn1">
		
		<INPUT TYPE=SUBMIT NAME="submit" VALUE="Lagre"  ID="Submit1">
		<INPUT TYPE=BUTTON NAME="btnCancel" VALUE="Avbryt" onClick="window.close()" >
	</table>
	</form>
</div>
<!--END content box -->


</body>
</html>