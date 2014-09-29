<%@ Language=JScript %>
<script runat=server language=Vbscript>
Function BinToHex(bytes)
	Dim bslen,i,b,l,h,chars,sb
	chars="0123456789ABCDEF"
	sb=""
	For i=1 To LenB(bytes)
		b=AscB(MidB(bytes,i,1))
		l=b Mod 16
		h=(b-l)/16
		sb = sb & Mid(chars,h+1,1) & Mid(chars,l+1,1)
	Next
	BinToHex=sb
End Function
</script>
<%

var type=String(Request.QueryString("type"));

if(type=="emptyhtml")
{
	Response.ContentType="text/html";
	Response.Write("<html><head></head><body></body></html>");
	Response.End();
}
else if(type=="serverip")
{
	Response.ContentType="text/plain";
	Response.Write(Request.ServerVariables("LOCAL_ADDR"));
	Response.End();
}
else if(type=="license")
{
	var thisfile=String(Request.ServerVariables("SCRIPT_NAME"));
	var licensefile=thisfile.replace(/[^\/]+\/[^\/]+$/,"")+"license/aspedit.lic";
	licensefile=String(Server.MapPath(licensefile));

	var stream=new ActiveXObject("ADODB.Stream");
	stream.Type=1;
	stream.Mode=3;
	stream.Open();
	stream.LoadFromFile(licensefile);
	var data=stream.Read(stream.Size);
	stream.Close();
	Response.Write(BinToHex(data));
	Response.End();
	
}

%>