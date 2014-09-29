<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = 0%>

<!--#INCLUDE FILE="includes/Xis.DB.inc"-->
<!--#INCLUDE FILE="includes/Xis.DB.Utils.inc"-->
<!--#INCLUDE FILE="includes/Xis.HTML.Error.inc"-->
<!--#INCLUDE FILE="includes/xis.rights.inc"-->
<!--#INCLUDE FILE="includes/Xis.Settings.inc"-->
<%
	If (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = false) Then
		call Response.Redirect("/xtra/IngenTilgang.asp")
	end if
	
	dim cpaction
	cpaction = Request.form("hdnCpAction")
	
	Set ConnTrans = GetClientConnection(GetConnectionstring(XIS, ""))
	ConnTrans.BeginTrans
	
	
	strSQL = "EXECUTE [spECUpdateOppdragVikarCategory] " & Request.form("OppdragId") &_
		", " & Request.form("dbxCategory")

	If ExecuteCRUDSQL(strSQL, ConnTrans) = false then
		ConnTrans.RollbackTrans
		CloseConnection(ConnTrans)
		set ConnTrans = nothing
		AddErrorMessage("Feil Oppdrag Vikar.")
		call RenderErrorMessage()
	Else
	
	ConnTrans.CommitTrans
	CloseConnection(ConnTrans)
	set ConnTrans = nothing
%>	
	<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
	<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	<title></title>

	<script language="javaScript" type="text/javascript">


		window.opener.document.all.hdnJobAction.value = "<%=cpaction%>";
		window.opener.document.all.frmJobNew.submit();
		window.close();

	</script>
	</head>

	<body>
	</body>
	
	<% End if %>
