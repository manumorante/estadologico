<%
	if msginfo = "" then
		msginfo = ""& request.QueryString("msginfo")
	end if

	if msginfo <> "" then
%>
		<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#DAF8D6">
			<tr>
			<td><font color="#006600"><strong>Info:</strong></font> <%=msginfo%></td>
			</tr>
</table>
	<%end if%>