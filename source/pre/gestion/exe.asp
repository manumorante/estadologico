<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="inc/secure.asp" -->
<!--#include file="inc/inc_conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>SQL</title>
</head>
<body>
<%
sql = ""& request.Form("sql")
if sql <> "" then
	on error resume next
	conn_.execute sql
	if err <> 0 then
		unerror = true : msgerror = err.description
	end if
	on error goto 0
end if

if unerror then
	Response.Write msgerror
end if
%>
<form id="f" name="f" method="post" action="exe.asp">
  <input name="sql" type="text" id="sql" value="<%=sql%>" size="120" />
  <input type="submit" name="Submit" value="Enviar" />
</form>
</body>
</html>
