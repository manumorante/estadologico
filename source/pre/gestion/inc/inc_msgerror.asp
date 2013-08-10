<%
	if msgerror = "" then
		msgerror = ""& request.QueryString("msgerror")
	end if

	if msgerror <> "" then
%>
		<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFDFDF">
			<tr>
			<td><font color="#990000"><strong>Error:</strong></font> <%=msgerror%></td>
			</tr>
		</table>
	<%end if%>