<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="inc/secure.asp" -->
<!--#include file="inc/inc_conn.asp" -->
<!--#include file="inc/inc_rutinas.asp" -->
<%
randomize
r= formatNumber(rnd(10)*1000,0)

dim re
dim ac
dim id_cliente
dim id
dim orden

ac = ""& request.QueryString("ac")
id_cliente = mNumero(""& request.QueryString("id_cliente"))

id = ""& request("id")
if ""&id <> "u" then
	id = mNumero(id)
end if
orden = ""& request.QueryString("orden")
key = ""& request.QueryString("key")


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="estilos.css" type="text/css" rel="stylesheet">
<title>Gesti&oacute;n</title>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="160" valign="top" bgcolor="#6375D6" class="fondoIzq"><table width="100%" border="0" cellspacing="0" cellpadding="10">
      <tr>
        <td><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="50" align="center" valign="top"><img src="arch/logo.jpg" width="163" height="34"></td>
            </tr>
          </table>
          <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="25" background="arch/f_titulo_menu_iz.gif">&nbsp;&nbsp;<b><font color="#316AC5">Opciones</font></b></td>
            </tr>
          </table>
          <table width="100%"  border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
            <tr>
              <td><table width="100%"  border="0" cellspacing="0" cellpadding="12">
                  <tr>
                    <td bgcolor="D6DFF7"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><a href="index.asp?ac=nuevo_servicio&r=<%=r%>" class="menu"><img src="arch/new.gif" alt="Nuevo" width="18" height="18" vspace="2" border="0" align="absmiddle"></a><a href="index.asp?ac=servicios&r=<%=r%>" class="menu"><img src="arch/ico_pro.gif" alt="Servicios" width="16" height="16" hspace="6" vspace="2" border="0" align="absmiddle">Servicios</a> </td>
                          </tr>
                        <tr>
                          <td><a href="index.asp?ac=nuevo_contacto&r=<%=r%>" class="menu"><img src="arch/new.gif" alt="Nuevo" width="18" height="18" vspace="2" border="0" align="absmiddle"></a><a href="index.asp?ac=contactos&r=<%=r%>" class="menu"><img src="arch/ico_usuarios.gif" alt="Contactos" width="18" height="18" hspace="5" vspace="2" border="0" align="absmiddle">Contactos</a></td>
                          </tr>
                        <tr>
                          <td><a href="index.asp?ac=nueva_empresa&r=<%=r%>" class="menu"><img src="arch/new.gif" alt="Nuevo" width="18" height="18" vspace="2" border="0" align="absmiddle"></a><a href="index.asp?r=<%=r%>" class="menu"><img src="arch/casa.gif" alt="Empresas" width="18" height="18" hspace="5" vspace="2" border="0" align="absmiddle">Empresas</a></td>
                          </tr>

                        <tr>
                          <td><a href="index.asp?ac=nueva_empresa&r=<%=r%>" class="menu"><img src="arch/new.gif" alt="Nuevo" width="18" height="18" vspace="2" border="0" align="absmiddle"></a><a href="index.asp?ac=tipos_de_servicios" title="Editar lista de tipos de servicios"><img src="arch/editar.gif" alt="Tipos" width="18" height="18" hspace="5" vspace="2" border="0" align="absmiddle">Tipos</a></td>
                          </tr>
                    </table></td>
                  </tr>
              </table></td>
            </tr>
          </table>
          <br>
          <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="25" background="arch/f_titulo_menu_iz.gif">&nbsp;&nbsp;<b><font color="#316AC5">Acciones</font></b></td>
            </tr>
          </table>
          <table width="100%"  border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
            <tr>
              <td><table width="100%"  border="0" cellspacing="0" cellpadding="12">
                  <tr>
                    <td bgcolor="D6DFF7"><form action="index.asp" method="get" name="buscar" id="buscar">
                      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><img src="arch/ico_usuarios.gif" alt="Buscar" width="18" height="18" hspace="5" vspace="2" align="absmiddle"><span class="suave">Buscar</span></td>
                        </tr>
                        <tr>
                          <td><select name="ac" id="ac">
                            <option value="servicios" selected>Servicios</option>
                            <option value="contactos">Contactos</option>
                            <option value="empresas">Empresas</option>
                          </select>
</td>
                        </tr>
                        <tr>
                          <td><input name="key" type="text" class="campo" id="key" value="<%=key%>">                          </td>
                        </tr>
                        <tr>
                          <td align="right"><input name="r" type="hidden" id="r" value="<%=r%>">
                            <input type="submit" value="Enviar"></td>
                        </tr>
                      </table>
                                        </form>
                    </td>
                  </tr>
              </table></td>
            </tr>
          </table></td>
      </tr>
    </table></td>
    <td valign="top" bgcolor="#FFFFFF">
      <table width="100%"  border="0" cellspacing="0" cellpadding="10">
        <tr>
          <td><%


Select case ac

case "servicio"

	' Modificar servicio
	if request.Form("fac") = "Modificar" then

		set re = mConsulta("SELECT * FROM SERVICIOS WHERE S_ID = "& id,conn_,2)
			id = re("S_ID")
			re("S_TIPO") = mNumero(request.Form("s_tipo"))
			re("S_TITULO") = ""& request.Form("s_titulo")
			re("S_DESCRIPCION") = ""& request.Form("s_descripcion")
			re("S_GASTOS") = mNumero(request.Form("s_gastos"))
			re("S_ADELANTO") = mNumero(request.Form("s_adelanto"))
			precio = mNumero(request.Form("s_precio"))
			pagado = mNumero(request.Form("s_pagado"))
			if precio >0 then
				por_pagado = FormatNumber((pagado / precio)*100,0)
			else
				por_pagado = 100
			end if
			re("S_PROGRESO") = mNumero(request.Form("s_progreso"))
			re("S_PRECIO") = precio
			re("S_PAGADO") = pagado
			re("S_POR_PAGADO") = por_pagado
			re("S_HORAS_ESTIMADAS") = mNumero(request.Form("s_horas_estimadas"))
			re("S_HORAS_DEDICADAS") = mNumero(request.Form("s_horas_dedicadas"))
			re("S_INICIO") = request.Form("dia_inicio") &"/"& request.Form("mes_inicio") &"/"& request.Form("ano_inicio")
			re("S_FIN") = request.Form("dia_fin") &"/"& request.Form("mes_fin") &"/"& request.Form("ano_fin")
			re("S_ESTADO") = ""& request.Form("s_estado")
		
		re.update
		mCierra(re)
		
		Response.Redirect("index.asp?ac=servicios")

	' Insertar Gasto
	' --------------------------------------------------------------------------------------------------------------
	elseif request.Form("fac") = "Gasto" then
	
		concepto = ""& request.Form("g_concepto")
		importe = mNumero(request.Form("g_importe"))

		dia = request.Form("dia_fecha")
		mes = request.Form("mes_fecha")
		ano = request.Form("ano_fecha")
		fecha = CDate(dia&"/"&mes&"/"&ano)
		
		if concepto = "" then
			unerror = true : msgerror = "No se ha indicado el concepto del gasto"
		end if
		
		if not unerror then
			if importe <= 0 then
				unerror = true : msgerror = "No se ha indicado un importe válido"
			end if
		end if
		
		if not unerror then
			set re = mConsulta("SELECT * FROM GASTOS",conn_,2)
			re.addNew()
				re("G_ID_SERVICIO") = id
				re("G_CONCEPTO") = concepto
				re("G_IMPORTE") = mMonedaBD(importe)
				re("G_FECHA") = fecha
			re.update()
			mCierra(re)
		end if
		
		Response.Redirect("index.asp?ac=servicio&id="& id &"&msgerror="& msgerror &"&r="& r)
	
	' Mostar servicio
	else

		if ""& id = "u" then
			sql = "SELECT * FROM SERVICIOS ORDER BY S_ID DESC"
		else
			sql = "SELECT * FROM SERVICIOS WHERE S_ID = "& id
		end if

		set re = mConsulta(sql,conn_,1)
		id = re("S_ID")

	%>
	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
		<tr>
		<td>
		<p><table width="100%"  border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td><b>Ampliar / Editar servicio</b></td>
                </tr>
                <tr>
                  <td><img src="arch/linea.gif" width="323" height="1"></td>
                </tr>
              </table>
			    </p>

			<!--#include file="inc/inc_msgerror.asp" -->
			<!--#include file="inc/inc_msginfo.asp" -->


				<form action="index.asp?ac=servicio" method="post" name="f" id="f">
				<input type="hidden" name="vuelta" value="<%=request.ServerVariables("HTTP_REFERER")%>">
				<input type="hidden" name="id" value="<%=id%>">
			      <table  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td valign="top"><table  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><fieldset>
                            <legend>Datos del servicio</legend>
                            <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                              <tr>
                                <td><table width="100%"  border="0" cellspacing="0" cellpadding="1">
                                    <tr>
                                      <td>Tipo</td>
                                    </tr>
                                    <tr>
                                      <td><select name="s_tipo" class="campo" id="s_tipo" onChange="changeTipo(this);">
                                          <%
										  set tipos = mConsulta("SELECT * FROM TIPOS_DE_SERVICIOS ORDER BY T_TITULO",conn_,2)
										  while not tipos.eof
										  %>
                                          <option value="<%=tipos("T_ID")%>" <%if re("S_TIPO") = tipos("T_ID") then%>selected<%end if%>><%=tipos("T_TITULO")%></option>
                                          <%
											tipos.movenext
											wend
											%>
                                        </select>                                      </td>
                                    </tr>
                                    <tr>
                                      <td>Titulo</td>
                                    </tr>
                                    <tr>
                                      <td><input name="s_titulo" type="text" class="campo" id="s_titulo" value="<%=re("S_TITULO")%>" size="30">
                                        (<%=id%>)</td>
                                    </tr>
                                    <tr>
                                      <td>Descripci&oacute;n</td>
                                    </tr>
                                    <tr>
                                      <td><textarea name="s_descripcion" cols="50" rows="6" wrap="VIRTUAL" class="area" id="s_descripcion"><%=re("S_DESCRIPCION")%></textarea></td>
                                    </tr>
                                  </table></td>
                              </tr>
                            </table>
                          </fieldset></td>
                        </tr>
                      </table>
                        </td>
                      <td width="10" valign="top">&nbsp;</td>
                      <td valign="top">
                          <table  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><fieldset>
                                <legend>Cliente</legend>
                                <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                                  <tr>
                                    <td>
									<%
									id_cliente = mNumero(request.QueryString("id_cliente"))
									empresa = ""& request.QueryString("empresa")
									if id_cliente >0 then
										%>
										<input name="s_id_empresa" type="hidden" id="s_id_empresa" value="<%=id_cliente%>">
										<input name="s_empresa_cliente" type="hidden" id="s_empresa_cliente" value="<%=empresa%>">
										<%=empresa%>
										<%
									else									
									
										set em = mConsulta("SELECT E_NOMBRE FROM EMPRESAS WHERE E_ID = "& re("S_ID_EMPRESA"),conn_,1)
									%>
									<table width="100%"  border="0" cellspacing="0" cellpadding="1">
                                        <tr>
                                          <td align="left"><a href="index.asp?ac=ampliar_empresa&id_cliente=<%=re("S_ID_EMPRESA")%>&r=<%=r%>"><%=em("E_NOMBRE")%></a></td>
                                        </tr>
                                    </table>
									
									<%end if%>									</td>
                                  </tr>
                                </table>
                              </fieldset></td>
                            </tr>
                          </table>
						  <br>
						  <table  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><fieldset>

							  <legend>Tiempo</legend>
                              <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                                  <tr>
                                    <td><table  border="0" cellspacing="0" cellpadding="1">
                                      <tr>
                                        <td align="right">Horas estimadas </td>
                                        <td>&nbsp;</td>
                                        <td>                                              <input name="s_horas_estimadas" type="text" class="campo" id="s_horas_estimadas" value="<%=re("S_HORAS_ESTIMADAS")%>" size="5"> 
                                          h                                        </td>
                                      </tr>
                                      <tr>
                                        <td align="right">Horas dedicadas </td>
                                        <td>&nbsp;</td>
                                        <td><input name="s_horas_dedicadas" type="text" class="campo" id="s_horas_dedicadas" value="<%=re("S_HORAS_DEDICADAS")%>" size="5">
    h </td>
                                      </tr>
                                    </table>
                                      <br>
                                      <table width="100%"  border="0" cellspacing="0" cellpadding="1">
                                        <tr>
                                          <td align="right">Inicio</td>
                                          <td>&nbsp;</td>
                                          <td><select name="dia_inicio" class="campo" id="dia_inicio">
										  <%for n=1 to 31%>
                                            <option value="<%=n%>"><%=n%></option>
											<%next%>
                                          </select>
                                            <select name="mes_inicio" class="campo" id="mes_inicio">
                                              <option value="01" selected>Enero</option>
                                              <option value="02">Febrero</option>
                                              <option value="03">Marzo</option>
                                              <option value="04">Abril</option>
                                              <option value="05">Mayo</option>
                                              <option value="06">Junio</option>
                                              <option value="07">Julio</option>
                                              <option value="08">Agosto</option>
                                              <option value="09">Septiembre</option>
                                              <option value="10">Octubre</option>
                                              <option value="11">Noviembre</option>
                                              <option value="12">Diciembre</option>
										    </select>
											  <select name="ano_inicio" class="campo" id="ano_inicio">
										  <%
										  ano = year(date)
										  i=0
										  for n=ano-2 to ano+2
										  i=i+1%>
                                            <option value="<%=n%>" <%if i=2 then Response.Write "selected" end if%>><%=n%></option>
											<%next%>
                                          </select>
                                          </select></td>
                                        </tr>
                                        <tr>
                                          <td align="right">Fin</td>
                                          <td>&nbsp;</td>
                                          <td><select name="dia_fin" class="campo" id="dia_fin">
                                            <%for n=1 to 31%>
                                            <option value="<%=n%>"><%=n%></option>
                                            <%next%>
                                          </select>
                                            <select name="mes_fin" class="campo" id="mes_fin">
                                              <option value="01" selected>Enero</option>
                                              <option value="02">Febrero</option>
                                              <option value="03">Marzo</option>
                                              <option value="04">Abril</option>
                                              <option value="05">Mayo</option>
                                              <option value="06">Junio</option>
                                              <option value="07">Julio</option>
                                              <option value="08">Agosto</option>
                                              <option value="09">Septiembre</option>
                                              <option value="10">Octubre</option>
                                              <option value="11">Noviembre</option>
                                              <option value="12">Diciembre</option>
                                            </select>
                                            <select name="ano_fin" class="campo" id="ano_fin">
                                              <%
										  ano = year(date)
										  i=0
										  for n=ano-2 to ano+2
										  i=i+1%>
                                              <option value="<%=n%>" <%if i=2 then Response.Write "selected" end if%>><%=n%></option>
                                              <%next%>
                                            </select></td>
                                        </tr>
                                    </table></td>
                                  </tr>
                                </table>
                              </fieldset></td>
                            </tr>
                          </table>
						  </td>
                    </tr>
                  </table>
			      <br>
<table  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td valign="top"><table  border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td><fieldset>
                              <legend>Gastos</legend>
                              <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                                <tr>
                                  <td><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#6376D6">
                                    <tr>
                                      <td height="18" valign="bottom">&nbsp;</td>
                                      <td colspan="3" valign="bottom"><strong><font color="#FFFFFF">Nuevo gasto </font></strong></td>
                                      <td valign="bottom">&nbsp;</td>
                                    </tr>
                                    <tr>
                                      <td height="18" valign="bottom">&nbsp;</td>
                                      <td colspan="3" valign="bottom"><input name="g_concepto" type="text" class="campo" id="g_concepto" style="width:100%;"></td>
                                      <td valign="bottom">&nbsp;</td>
                                    </tr>
                                    <tr>
                                      <td height="32">&nbsp;</td>
                                      <td><input name="g_importe" type="text" class="campo" id="g_importe" size="7"></td>
                                      <td><%=mWOFecha("", "fecha","")%></td>
                                      <td align="right"><input name="fac" type="submit" class="boton" id="fac" title="Insertar nuevo gasto" value="Gasto"></td>
                                      <td>&nbsp;</td>
                                    </tr>
                                  </table>
                                  <%set gastos = mConsulta("SELECT * FROM GASTOS WHERE G_ID_SERVICIO = "& id &" ORDER BY G_FECHA DESC",conn_,1)%>
                                      <table width="100%"  border="0" cellspacing="0" cellpadding="3">
                                        <%if gastos.eof then%>
                                        <tr>
                                          <td colspan="4" align="center"><strong><font color="#009900">No hay gastos</font></strong></td>
                                        </tr>
                                        <%else
										gastos_total = 0
										while not gastos.eof
										
										importe = gastos("G_IMPORTE")
										gastos_total = gastos_total + importe
										
										fila1 = "#EDF3FE"
										fila2 = "#FFFFFF"
										if fila = fila1 then
											fila = fila2
										else
											fila = fila1
										end if
										%>
                                        <tr bgcolor="<%=fila%>">
                                          <td align="left"><%=gastos("G_CONCEPTO")%></td>
                                          <td align="right"><span class="suave" title="Modificado <%=gastos("G_AUTO_FECHA")%>"><font size="1" face="Arial, Helvetica, sans-serif"><%=gastos("G_FECHA")%></font></span></td>
                                          <td align="right"><%=mEuros(importe)%>&euro;</td>
                                          <td align="right"><a href="#" onClick="if (confirm(&quot;&iquest;Borrar gasto '<%=gastos("G_CONCEPTO")%>'?&quot;)){
location.href='index.asp?ac=eliminar_gasto_servicio&amp;ids=<%=id%>&amp;idg=<%=gastos("G_ID")%>'
}"><img src="arch/x.gif" alt="Eliminar" width="18" height="18" border="0"></a></td>
                                        </tr>
                                        <%gastos.movenext
										wend%>
                                        <tr>
                                          <td align="right" bgcolor="#6376D6">&nbsp;</td>
                                          <td align="right" bgcolor="#6376D6">&nbsp;</td>
                                          <td align="right" bgcolor="#6376D6"><strong><font color="#FFFFFF"><%=mEuros(gastos_total)%>&euro;</font></strong></td>
                                          <td align="right" bgcolor="#6376D6">&nbsp;</td>
                                        </tr>
                                        <%
										end if%>
                                      </table>
                                    <% mCierra(gastos) %>
                                  </td>
                                </tr>
                              </table>
                            </fieldset></td>
                          </tr>
                      </table></td>
                      <td width="10" valign="top">&nbsp;</td>
                      <td valign="top"><table  border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td><fieldset>
                                  <legend>Precio</legend>
                                  <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                                    <tr>
                                      <td><table  border="0" cellspacing="0" cellpadding="1">
                                          <tr>
                                            <td align="right">Progreso</td>
                                            <td>&nbsp;</td>
                                            <td><input name="s_progreso" type="text" class="campo" id="s_progreso" value="<%=re("S_PROGRESO")%>" size="5">
                                            %</td>
                                            <td align="right">&nbsp;</td>
                                            <td align="right">&nbsp;</td>
                                          </tr>
                                          <tr>
                                            <td align="right"><strong>Precio</strong></td>
                                            <td>&nbsp;</td>
                                            <td><input name="s_precio" type="text" class="campo" id="s_precio" value="<%=re("S_PRECIO")%>" size="5">
                                              &euro; </td>
                                            <td align="right">&nbsp;</td>
                                            <td align="right">&nbsp;</td>
                                          </tr>
                                          <tr>
                                            <td align="right">Gastos</td>
                                            <td>&nbsp;</td>
                                            <td><input name="s_gastos" type="text" class="campo" id="s_gastos" value="<%=mEuros(gastos_total)%>" size="5" readonly>
                                              &euro; </td>
                                            <td align="right">&nbsp;</td>
                                            <td align="right">&nbsp;</td>
                                          </tr>
                                          <tr>
                                            <td align="right">Adelanto</td>
                                            <td>&nbsp;</td>
                                            <td><input name="s_adelanto" type="text" class="campo" id="s_adelanto" value="<%=re("S_ADELANTO")%>" size="5">
                                              &euro; </td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                          </tr>
                                          <tr>
                                            <td align="right">Pagado</td>
                                            <td>&nbsp;</td>
                                            <td><input name="s_pagado" type="text" class="campo" id="s_pagado" value="<%=re("S_PAGADO")%>" size="5">
                                              &euro; </td>
                                            <td>&nbsp;</td>
                                            <td><%
										  if re("S_PRECIO") > 0 then
										  por_pagado = (re("S_PAGADO")/re("S_PRECIO"))*100
										  else
										  por_pagado = 100
										  end if%>
                                              <%=por_pagado%>% </td>
                                          </tr>
                                      </table></td>
                                    </tr>
                                  </table>
                                </fieldset></td>
                              </tr>
                        </table></td>
                    </tr>
                  </table>
			      <br>
			      <table width="100%"  border="0" cellpadding="15" cellspacing="0" bgcolor="#6477D7">
                    <tr>
                      <td align="left"><input name="fac" type="submit" id="fac" value="Modificar"></td>
                    </tr>
                  </table>
			    </form>
			    </td>
			</tr>
			
             
            </table>
	<%
	end if


case "eliminar_gasto_servicio"
	
	ids = mNumero(request.QueryString("ids"))
	idg = mNumero(request.QueryString("idg"))
	if ids>0 and idg>0 then
		set re = mConsulta("DELETE * FROM GASTOS WHERE G_ID_SERVICIO = "& ids &" AND G_ID = "& idg, conn_, 3)
		mCierra(re)
	end if
	
	Response.Redirect("index.asp?ac=servicio&id="& ids &"&msginfo=Gasto eliminado correctamente")

case "eliminar_tipo_de_servicio"

	if id>0 then
		set re = mConsulta("DELETE * FROM TIPOS_DE_SERVICIOS WHERE T_ID = "& id &"",conn_,3)
		mCierra(re)
	end if
	
	Response.Redirect(request.ServerVariables("HTTP_REFERER"))

case "eliminar_contacto"

	if id>0 then
		set re = mConsulta("DELETE * FROM CONTACTOS WHERE C_ID = "& id &"",conn_,3)
		mCierra(re)
	end if
	Response.Redirect(request.ServerVariables("HTTP_REFERER"))

case "eliminar_empresa"

	if id>0 then
		set re = mConsulta("DELETE * FROM EMPRESAS WHERE E_ID = "& id &"",conn_,3)
		mCierra(re)
	end if
	Response.Redirect(request.ServerVariables("HTTP_REFERER"))

case "eliminar_servicio"

	if id>0 then
		set re = mConsulta("DELETE * FROM SERVICIOS WHERE S_ID = "& id &"",conn_,3)
		mCierra(re)
	end if
	Response.Redirect(request.ServerVariables("HTTP_REFERER"))

case "nuevo_servicio"

	if request.Form() <> "" then

		set re = mConsulta("SELECT * FROM SERVICIOS",conn_,2)
		re.addNew
		
			re("S_TIPO") = ""& request.Form("s_tipo")
			re("S_TITULO") = ""& request.Form("s_titulo")
			re("S_DESCRIPCION") = ""& request.Form("s_descripcion")
			re("S_PRECIO") = mNumero(request.Form("s_precio"))
			re("S_HORAS_ESTIMADAS") = mNumero(request.Form("s_horas_estimadas"))
			re("S_HORAS_DEDICADAS") = mNumero(request.Form("s_horas_dedicadas"))
			re("S_ID_EMPRESA") = mNumero(request.Form("s_id_empresa"))
			re("S_INICIO") = request.Form("dia_inicio") &"/"& request.Form("mes_inicio") &"/"& request.Form("ano_inicio")
			re("S_FIN") = request.Form("dia_fin") &"/"& request.Form("mes_fin") &"/"& request.Form("ano_fin")
			re("S_ESTADO") = ""& request.Form("s_estado")
		
		re.update
		mCierra(re)
		
		Response.Redirect("index.asp?ac=servicio&id=u")

	else
	%>
	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
		<tr>
		<td>
		<p><table width="100%"  border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td><b>Nuevo servicio</b></td>
                </tr>
                <tr>
                  <td><img src="arch/linea.gif" width="323" height="1"></td>
                </tr>
              </table>
			    </p>
			    <form action="index.asp?ac=nuevo_servicio" method="post" name="f" id="f">
				<input type="hidden" name="vuelta" value="<%=request.ServerVariables("HTTP_REFERER")%>">
			      <table  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td>
<p>
                          <table  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><fieldset>
                                <legend>Cliente</legend>
                                <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                                  <tr>
                                    <td>
									<%
									id_cliente = mNumero(request.QueryString("id_cliente"))
									empresa = ""& request.QueryString("empresa")
									if id_cliente >0 then
										%>
										<input name="s_id_empresa" type="hidden" id="s_id_empresa" value="<%=id_cliente%>">
										<input name="s_empresa_cliente" type="hidden" id="s_empresa_cliente" value="<%=empresa%>">
										<%=empresa%>
										<%
									else									
									
										set re = mConsulta("SELECT E_NOMBRE, E_ID FROM EMPRESAS ORDER BY E_NOMBRE",conn_,1)
									%>
									<table width="100%"  border="0" cellspacing="0" cellpadding="1">
                                        <tr>
                                          <td align="right">Empresa</td>
                                          <td>&nbsp;</td>
                                          <td>
										<%if not re.eof then%>
											<select name="s_id_empresa" class="campo">
												<%while not re.eof%>
													<option value="<%=re("E_ID")%>"><%=re("E_NOMBRE")%></option>
													<%re.movenext
												wend%>
											</select>                                            
                                            <%else%>
											<select name="s_id_empresa" disabled="disabled" class="campo">
												<option value="">Ninguna</option>
											</select>
										<%end if%>
										<a href="index.asp?ac=nueva_empresa" title="A&ntilde;adir una nueva empresa"><img src="arch/new.gif" width="18" height="18" border="0" align="absmiddle">Nueva</a></td>
                                        </tr>
                                    </table>
									
									<%end if%>
									
									</td>
                                  </tr>
                                </table>
                              </fieldset></td>
                            </tr>
                          </table>
						  </p>
					  <p>
					  <script language="javascript" type="text/javascript">
					  function changeTipo(c) {
//					  	f.s_titulo.value = c.nodetext
					  }
					  </script>
                          <table  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><fieldset>
                                <legend>Datos del servicio</legend>
                                <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                                  <tr>
                                    <td><table width="100%"  border="0" cellspacing="0" cellpadding="1">
                                        <tr>
                                          <td align="right">Tipo</td>
                                          <td>&nbsp;</td>
                                          <td>
										  <select name="s_tipo" class="campo" id="s_tipo" onChange="changeTipo(this);">
										  <%
										  set tipos = mConsulta("SELECT * FROM TIPOS_DE_SERVICIOS ORDER BY T_TITULO",conn_,2)
										  while not tipos.eof
										  %>
                                            <option value="<%=tipos("T_ID")%>"><%=tipos("T_TITULO")%></option>
											<%
											tipos.movenext
											wend
											%>
                                          </select>										  
										  <a href="index.asp?ac=tipos_de_servicios" title="Editar lista de tipos de servicios"><img src="arch/editar.gif" width="18" height="18" border="0" align="absmiddle">Editar</a></td>
                                        </tr>
                                        <tr>
                                          <td align="right">Titulo</td>
                                          <td>&nbsp;</td>
                                          <td><input name="s_titulo" type="text" class="campo" id="s_titulo" size="30"></td>
                                        </tr>
                                        <tr>
                                          <td align="right" valign="top">Descripci&oacute;n</td>
                                          <td>&nbsp;</td>
                                          <td><textarea name="s_descripcion" cols="25" rows="10" wrap="VIRTUAL" class="area" id="s_descripcion"></textarea></td>
                                        </tr>
                                    </table>
                                      <br>
                                      <table  border="0" cellspacing="0" cellpadding="1">
                                        <tr>
                                          <td align="right">Precio</td>
                                          <td>&nbsp;</td>
                                          <td><input name="s_precio" type="text" class="campo" id="s_precio" value="" size="5">
&euro; </td>
                                        </tr>
                                      </table></td>
                                  </tr>
                                </table>
                              </fieldset></td>
                            </tr>
                          </table>
						  </p>
<p>
                          <table  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><fieldset>

							  <legend>Tiempo</legend>
                              <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                                  <tr>
                                    <td><table  border="0" cellspacing="0" cellpadding="1">
                                      <tr>
                                        <td align="right">Horas estimadas </td>
                                        <td>&nbsp;</td>
                                        <td>                                              <input name="s_horas_estimadas" type="text" class="campo" id="s_horas_estimadas" value="" size="5"> 
                                          h
                                        </td>
                                      </tr>
                                      <tr>
                                        <td align="right">Horas dedicadas </td>
                                        <td>&nbsp;</td>
                                        <td><input name="s_horas_dedicadas" type="text" class="campo" id="s_horas_dedicadas" value="" size="5">
    h </td>
                                      </tr>
                                    </table>
                                      <br>
                                      <table width="100%"  border="0" cellspacing="0" cellpadding="1">
                                        <tr>
                                          <td align="right">Inicio</td>
                                          <td>&nbsp;</td>
                                          <td><select name="dia_inicio" class="campo" id="dia_inicio">
										  <%for n=1 to 31%>
                                            <option value="<%=n%>"><%=n%></option>
											<%next%>
                                          </select>
                                            <select name="mes_inicio" class="campo" id="mes_inicio">
                                              <option value="01" selected>Enero</option>
                                              <option value="02">Febrero</option>
                                              <option value="03">Marzo</option>
                                              <option value="04">Abril</option>
                                              <option value="05">Mayo</option>
                                              <option value="06">Junio</option>
                                              <option value="07">Julio</option>
                                              <option value="08">Agosto</option>
                                              <option value="09">Septiembre</option>
                                              <option value="10">Octubre</option>
                                              <option value="11">Noviembre</option>
                                              <option value="12">Diciembre</option>
										    </select>
											  <select name="ano_inicio" class="campo" id="ano_inicio">
										  <%
										  ano = year(date)
										  i=0
										  for n=ano-2 to ano+2
										  i=i+1%>
                                            <option value="<%=n%>" <%if i=2 then Response.Write "selected" end if%>><%=n%></option>
											<%next%>
                                          </select>
                                          </select></td>
                                        </tr>
                                        <tr>
                                          <td align="right">Fin</td>
                                          <td>&nbsp;</td>
                                          <td><select name="dia_fin" class="campo" id="dia_fin">
                                            <%for n=1 to 31%>
                                            <option value="<%=n%>"><%=n%></option>
                                            <%next%>
                                          </select>
                                            <select name="mes_fin" class="campo" id="mes_fin">
                                              <option value="01" selected>Enero</option>
                                              <option value="02">Febrero</option>
                                              <option value="03">Marzo</option>
                                              <option value="04">Abril</option>
                                              <option value="05">Mayo</option>
                                              <option value="06">Junio</option>
                                              <option value="07">Julio</option>
                                              <option value="08">Agosto</option>
                                              <option value="09">Septiembre</option>
                                              <option value="10">Octubre</option>
                                              <option value="11">Noviembre</option>
                                              <option value="12">Diciembre</option>
                                            </select>
                                            <select name="ano_fin" class="campo" id="ano_fin">
                                              <%
										  ano = year(date)
										  i=0
										  for n=ano-2 to ano+2
										  i=i+1%>
                                              <option value="<%=n%>" <%if i=2 then Response.Write "selected" end if%>><%=n%></option>
                                              <%next%>
                                            </select></td>
                                        </tr>
                                    </table></td>
                                  </tr>
                                </table>
                              </fieldset></td>
                            </tr>
                          </table>
						  </p>


                          <p>
                          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td align="right"><input type="submit" value="Enviar"></td>
                            </tr>
                        </table></td>
                    </tr>
                  </table>
			      </form>
			    </td>
			</tr>
			
             
            </table>
	<%
	end if

case "nuevo_contacto"

	if request.Form() <> "" then

		set re = mConsulta("SELECT * FROM CONTACTOS",conn_,2)
		re.addNew
		
			re("C_NOMBRE") = ""& request.Form("c_nombre")
			re("C_APELLIDOS") = ""& request.Form("c_apellidos")
			re("C_TELEFONO") = ""& request.Form("c_telefono")
			re("C_MOVIL") = ""& request.Form("c_movil")
			re("C_EMAIL") = ""& request.Form("c_email")
			re("C_NOTAS") = ""& request.Form("c_notas")
		
		re.update
		mCierra(re)
		
		Response.Redirect("index.asp?ac=contactos")

	else
	%>
	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
		<tr>
		<td>
		<p><table width="100%"  border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td><b>Nuevo contacto</b></td>
                </tr>
                <tr>
                  <td><img src="arch/linea.gif" width="323" height="1"></td>
                </tr>
              </table>
			    </p>
			    <form action="index.asp?ac=nuevo_contacto" method="post" name="f" id="f">
			      <table  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><p>
                          <table  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><fieldset>
                                <legend>Datos del contacto</legend>
                                <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                                  <tr>
                                    <td><table width="100%"  border="0" cellspacing="0" cellpadding="1">
                                        <tr>
                                          <td align="right">Nombre</td>
                                          <td>&nbsp;</td>
                                          <td><input name="c_nombre" type="text" class="campo" id="c_nombre"></td>
                                        </tr>
                                        <tr>
                                          <td align="right">Apellidos</td>
                                          <td>&nbsp;</td>
                                          <td><input name="c_apellidos" type="text" class="campo" id="c_apellidos"></td>
                                        </tr>
                                        <tr>
                                          <td align="right">M&oacute;vil</td>
                                          <td>&nbsp;</td>
                                          <td><input name="c_movil" type="text" class="campo" id="c_movil"></td>
                                        </tr>
                                        <tr>
                                          <td align="right">Tel&eacute;fono</td>
                                          <td>&nbsp;</td>
                                          <td><input name="c_telefono" type="text" class="campo" id="c_telefono"></td>
                                        </tr>
                                        <tr>
                                          <td align="right">E-mail</td>
                                          <td>&nbsp;</td>
                                          <td><input name="c_email" type="text" class="campo" id="c_email"></td>
                                        </tr>
                                    </table></td>
                                  </tr>
                                </table>
                              </fieldset></td>
                            </tr>
                          </table>
                          <p>
                          <table  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><fieldset>
                                <legend>Notas</legend>
                                <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                                  <tr>
                                    <td><textarea name="c_notas" cols="23" rows="5" wrap="VIRTUAL" class="area" id="c_notas"></textarea></td>
                                  </tr>
                                </table>
                              </fieldset></td>
                            </tr>
                          </table>
                          <p></p>
                          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td align="right"><input type="submit" value="Enviar"></td>
                            </tr>
                        </table></td>
                    </tr>
                  </table>
			      </form>
			    </td>
			</tr>
			
             
            </table>
	<%
	end if


case "nueva_empresa"

	if request.Form() <> "" then

		set re = mConsulta("SELECT * FROM EMPRESAS",conn_,2)
		re.addNew
		
			re("E_CONTACTOS") = ""& request.Form("e_contactos")
			re("E_NOMBRE") = ""& request.Form("e_nombre")
			re("E_EMAIL") = ""& request.Form("e_email")
			re("E_TELEFONO") = ""& request.Form("e_telefono")
			re("E_DIRECCION") = ""& request.Form("e_direccion")
			re("E_POBLACION") = ""& request.Form("e_poblacion")
			re("E_CP") = mNumero(request.Form("e_cp"))
			re("E_PROVINCIA") = ""& request.Form("e_provincia")
			re("E_PAIS") = ""& request.Form("e_pais")
			re("E_ACTIVIDAD") = ""& request.Form("e_actividad")
		
		re.update
		mCierra(re)
		
		Response.Redirect("index.asp?ac=empresas")

	else
	%>
	<table width="100%"  border="0" cellspacing="0" cellpadding="2">
		<tr>
		<td>
		<p><table width="100%"  border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td><b>Nueva empresa</b></td>
                </tr>
                <tr>
                  <td><img src="arch/linea.gif" width="323" height="1"></td>
                </tr>
              </table>
			    </p>
			    <form action="index.asp?ac=nueva_empresa" method="post" name="f" id="f">
			      <table  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td valign="top"><table  border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td><fieldset>
                              <legend>Contactos vinculados</legend>
                              <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                                <tr>
                                  <td><table width="100%"  border="0" cellspacing="0" cellpadding="1">
                                      <tr>
                                        <td><iframe frameborder="0" scrolling="yes" src="scan_contactos.asp" width="225px" height="400px"></iframe></td>
                                      </tr>
                                    </table>
                                      <a href="index.asp?ac=nuevo_contacto"><img src="arch/new.gif" width="18" height="18" border="0" align="absmiddle"> Nuevo</a></td>
                                </tr>
                              </table>
                            </fieldset></td>
                          </tr>
                      </table></td>
                      <td width="5" valign="top">&nbsp;</td>
                      <td valign="top">
                          <table  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td><fieldset>
                                <legend>Datos de la empresa</legend>
                                <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                                  <tr>
                                    <td><table width="100%"  border="0" cellspacing="0" cellpadding="1">
                                        <tr>
                                          <td align="right"><input name="e_contactos" type="hidden" id="e_contactos" value="">
                                          Nombre</td>
                                          <td>&nbsp;</td>
                                          <td><input name="e_nombre" type="text" class="campo" id="e_nombre"></td>
                                        </tr>
                                        <tr>
                                          <td align="right">Actividad</td>
                                          <td>&nbsp;</td>
                                          <td><input name="e_actividad" type="text" class="campo" id="e_actividad"></td>
                                        </tr>
                                        <tr>
                                          <td align="right">&nbsp;</td>
                                          <td>&nbsp;</td>
                                          <td>&nbsp;</td>
                                        </tr>
                                        <tr>
                                          <td align="right">E-mail</td>
                                          <td>&nbsp;</td>
                                          <td><input name="e_email" type="text" class="campo" id="e_email"></td>
                                        </tr>
                                        <tr>
                                          <td align="right">Tel&eacute;fono</td>
                                          <td>&nbsp;</td>
                                          <td><input name="e_telefono" type="text" class="campo" id="e_telefono"></td>
                                        </tr>
                                        <tr>
                                          <td align="right">Provincia</td>
                                          <td>&nbsp;</td>
                                          <td><input name="e_provincia" type="text" class="campo" id="e_provincia"></td>
                                        </tr>
                                        <tr>
                                          <td align="right">Poblaci&oacute;n</td>
                                          <td>&nbsp;</td>
                                          <td><input name="e_poblacion" type="text" class="campo" id="e_poblacion"></td>
                                        </tr>
                                        <tr>
                                          <td align="right">Direcci&oacute;n</td>
                                          <td>&nbsp;</td>
                                          <td><input name="e_direccion" type="text" class="campo" id="e_direccion"></td>
                                        </tr>
                                        <tr>
                                          <td align="right">C. postal </td>
                                          <td>&nbsp;</td>
                                          <td><input name="e_cp" type="text" class="campo" id="e_cp"></td>
                                        </tr>
                                    </table>                                    </td>
                                  </tr>
                                </table>
                              </fieldset></td>
                            </tr>
                          </table>
                          <br>
                          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td align="right"><input type="submit" value="Enviar"></td>
                            </tr>
                        </table></td>
                      </tr>
                  </table>
			      </form>
			    </td>
			</tr>
			
             
            </table>
	<%
	end if
	
case "servicios"

	sql = "SELECT * FROM EMPRESAS, SERVICIOS, TIPOS_DE_SERVICIOS WHERE (T_ID = S_TIPO) AND (S_ID_EMPRESA = E_ID)"

	if key <> "" then
		sql = sql & " AND (S_TITULO LIKE '%"& key &"%' OR S_DESCRIPCION LIKE '%"& key &"%')"
	end if

	if mNumero(request.QueryString("tipo")) >0 then
		sql = sql & " AND S_TIPO = "& mNumero(request.QueryString("tipo"))
	end if

	if mNumero(request.QueryString("empresa")) >0 then
		sql = sql & " AND S_ID_EMPRESA = "& mNumero(request.QueryString("empresa"))
	end if

	if orden <> "" then
		sql = sql & " ORDER BY S_"& orden &" ASC"
	else
		sql = sql & " ORDER BY S_POR_PAGADO ASC"
	end if
	
	sql = "SELECT * FROM EMPRESAS, SERVICIOS, TIPOS_DE_SERVICIOS"
	
	Response.Write "sql: "& sql
	set re = mConsulta(sql,conn_,1)
	%>
            <table width="100%"  border="0" cellspacing="0" cellpadding="2">
			<tr>
			  <td><table width="100%"  border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td><b>Servicios</b></td>
                </tr>
                <tr>
                  <td><img src="arch/linea.gif" width="100%" height="1"></td>
                </tr>
              </table>
			    </td>
			</tr>
			
            </table>
			
			

			<script language="javascript" type="text/javascript">
			function changeTipos(c) {
				location.href='index.asp?ac=servicios&tipo='+ c.value +'&empresa=<%=request.QueryString("empresa")%>'
			}

			function changeEmpresa(c) {
				location.href='index.asp?ac=servicios&tipo=<%=request.QueryString("tipo")%>&empresa='+ c.value
			}
			</script>
            <table width="100%"  border="0" cellspacing="0" cellpadding="2">
              <tr bgcolor="#6375D6">
                <td width="10">&nbsp;</td>
                <td><a href="index.asp?ac=servicios&orden=titulo&r=<%=r%>"><b><font color="#FFFFFF">T&iacute;tulo</font></b></a></td>
                <td><a href="index.asp?ac=servicios&orden=tipo&r=<%=r%>"><b><font color="#FFFFFF">Tipo</font></b></a></td>
                <td colspan="2"><a href="index.asp?ac=servicios&orden=id_empresa&r=<%=r%>"><b><font color="#FFFFFF">Cliente</font></b></a></td>
                <td align="right"><a href="index.asp?ac=servicios&orden=gastos&r=<%=r%>"><b><font color="#FFFFFF">Gastos</font></b></a></td>
                <td align="right"><a href="index.asp?ac=servicios&orden=adelanto&r=<%=r%>"><b><font color="#FFFFFF">Adelanto</font></b></a></td>
                <td align="center" bgcolor="#6375D6"><table width="70" border="0" cellpadding="1" cellspacing="0">
                  <tr>
                    <td width="50%" align="center"><table width="100%" border="0" cellpadding="2" cellspacing="0" bgcolor="#009900">
                      <tr>
                        <td align="center"><a href="index.asp?ac=servicios&orden=progreso&r=<%=r%>"><strong><font color="#FFFFFF">%</font></strong></a></td>
                      </tr>
                    </table>
                      </td>
                    <td width="50%" align="center"><table width="100%" border="0" cellpadding="2" cellspacing="0" bgcolor="#CCCC00">
                      <tr>
                        <td align="center"><a href="index.asp?ac=servicios&orden=por_pagado&r=<%=r%>"><strong><font color="#FFFFFF">&euro;</font></strong></a></td>
                      </tr>
                    </table>
                      </td>
                  </tr>
                </table></td>
                <td align="right"><a href="index.asp?ac=servicios&orden=precio&r=<%=r%>"><b><font color="#FFFFFF">Precio</font></b></a></td>
                <td align="right">&nbsp;</td>
                <td width="10" align="right">&nbsp;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>
				<%
				set tipos = mConsulta("SELECT * FROM TIPOS_DE_SERVICIOS ORDER BY T_TITULO",conn_,1)
				%>
				<select name="" class="campo" onChange="changeTipos(this);">
                  <option value="">Todos</option>
				<%while not tipos.eof%>
                  <option value="<%=tipos("T_ID")%>" <%if ""& request.QueryString("tipo") = ""& tipos("T_ID") then Response.Write "selected" end if%>><%=tipos("T_TITULO")%></option>
				  <%tipos.movenext
				  wend%>
                </select></td>
                <td colspan="2"><%
				set empresas = mConsulta("SELECT * FROM EMPRESAS ORDER BY E_NOMBRE",conn_,1)
				%>
                    <select name="" class="campo" onChange="changeEmpresa(this);">
                    <option value="">Todos</option>
                      <%while not empresas.eof%>
                      <option value="<%=empresas("E_ID")%>" <%if ""& request.QueryString("empresa") = ""& empresas("E_ID") then Response.Write "selected" end if%>><%=empresas("E_NOMBRE")%></option>
                      <%empresas.movenext
				  wend%>
                    </select></td>
                <td align="right">&nbsp;</td>
                <td align="right">&nbsp;</td>
                <td align="right">&nbsp;</td>
                <td align="right">&nbsp;</td>
                <td align="right">&nbsp;</td>
                <td align="right">&nbsp;</td>
              </tr>

              <%
				if not re.eof then
					total = 0
					totalG = 0
					while not re.eof
						precio = mNumero(re("S_PRECIO"))
						gastos = mNumero(re("S_GASTOS"))
						adelanto = mNumero(re("S_ADELANTO"))
						pagado = mNumero(re("S_PAGADO"))
						' Porcentaje adelantado
						if adelanto > 0 and precio > 0 then
							por_adelanto = int((adelanto/precio)*100)
						else
							por_adelanto = 0
						end if
						total = total + precio
						totalG = totalG + gastos
						fila1 = "#EDF3FE"
						fila2 = "#FFFFFF"
						if fila = fila1 then
							fila = fila2
						else
							fila = fila1
						end if%>
					  <tr bgcolor="<%=fila%>">
						<td>&nbsp;</td>
						<td><a href="index.asp?ac=servicio&id=<%=re("S_ID")%>"><img src="arch/lupa_mini.gif" width="14" height="14" border="0" align="absmiddle"> <%if re("S_TITULO") <> "" then Response.Write re("S_TITULO") else Response.Write re("T_TITULO") end if%></a></td>
						<td><%=re("T_TITULO")%></td>
						<td colspan="2"><a href="index.asp?ac=ampliar_empresa&id_cliente=<%=re("E_ID")%>&r=<%=R%>"><%=re("E_NOMBRE")%></a></td>
						<td align="right"><%if ""& session("usuario") = "1" then Response.Write mEuros(gastos) end if%></td>
						<td align="right"><%=mEuros(adelanto)%> <span class="suave"><%=por_adelanto%>%</span></td>
						<td align="right"><table border="0" cellspacing="0" cellpadding="1">
                          <tr>
                            <td><table width="65" height="7" border="0" cellpadding="0" cellspacing="0" bgcolor="#B1C3D9" title="Progreso <%'=re("S_PROGRESO")%>%">
                              <tr>
                                <td align="left" bgcolor="#DDF1DA">
								<%'if re("S_PROGRESO") > 0 then%>
								<table width="<%'=re("S_PROGRESO")%>%" height="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#009900">
                                      <tr>
                                        <td align="center"><img src="arch/spacer.gif" width="1" height="1"></td>
                                      </tr>
                                  </table>
								  <%'else%>
								<table  border="0" cellpadding="0" cellspacing="0">
                                      <tr>
                                        <td align="center"><img src="arch/spacer.gif" width="1" height="1"></td>
                                      </tr>
                                  </table>
								  <%'end if%></td>
                              </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td>
							<table width="65" height="7" border="0" cellpadding="0" cellspacing="0" bgcolor="#B1C3D9" title="Pagado <%'=re("S_POR_PAGADO")%>%">
                              <tr>
                                <td align="left" bgcolor="#F0EDCA">
								<%'if re("S_POR_PAGADO") >0 then%>
								<table width="<%'=re("S_POR_PAGADO")%>%" height="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCC00">
                                      <tr>
                                        <td align="center"><img src="arch/spacer.gif" width="1" height="1"></td>
                                      </tr>
                                  </table>
								  <%'else%>
								<table height="100%"  border="0" cellpadding="0" cellspacing="0">
                                      <tr>
                                        <td align="center"><img src="arch/spacer.gif" width="1" height="1"></td>
                                      </tr>
                                  </table>
								  <%'end if%>
								</td>
                              </tr>
                            </table></td>
                          </tr>
                        </table>
						  </td>
						<td align="right"><%=mEuros(precio)%></td>
						<td width="18" align="center"><a href="index.asp?ac=eliminar_servicio&id=<%'=re("S_ID")%>" onClick="if(!confirm('¿Eliminar <%'=re("S_TITULO")%>?')){return false;}"><img src="arch/x.gif" alt="Eliminar" width="18" height="18" border="0" align="absmiddle"></a></td>
						<td align="right">&nbsp;</td>
					  </tr>
		
					  <%re.movenext
				wend%>

			<%else%>
			<TR><TD colspan="11"><div align="center" class="msginfo">No hay ningún servicio<%if key<>"" then%> con la búsqueda <strong><%=key%></strong><%end if%></div></TD></TR>
			<%end if%>
              <tr>
                <td colspan="11"><img src="arch/linea.gif" width="100%" height="1"></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td align="right"><%if ""& session("usuario") = "1" then Response.Write mEuros(totalG) &" &euro;" end if%></td>
                <td align="right">&nbsp;</td>
                <td align="right">&nbsp;</td>
                <td align="right"><span title="<%=formatnumber(total*166.386,0)%>"><%=mEuros(total)%> </span></td>
                <td align="center">&nbsp;</td>
                <td align="right">&nbsp;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td align="right">&nbsp;</td>
                <td align="right">&nbsp;</td>
                <td align="right">&nbsp;</td>
                <td align="right"><%if ""& session("usuario") = "1" then%><span title="<%=formatnumber((total-totalg)*166.386,0)%> pts"><b><%=mEuros(total-totalg)%> </b></span>
                  <%end if%></td>
                <td align="center">&nbsp;</td>
                <td align="right">&nbsp;</td>
              </tr>
            </table>

			
            <%
	mCierra(re)

case "ampliar_contacto"

	set re = mConsulta("SELECT * FROM CONTACTOS WHERE C_ID = "& id, conn_, 1)
	if re.eof then
		unerror = true : msgerror = "No se ha encontrado el contacto solicitado."
	end if
	
	if not unerror then
		%>
            <table  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td colspan="2" class="nombre_empresa_ampliado"><%=re("C_NOMBRE")%>&nbsp;<%=re("C_APELLIDOS")%></td>
              </tr>
              <tr>
                <td colspan="2">&nbsp;</td>
              </tr>
              <tr>
                <td><b>Tel&eacute;fono</b></td>
                <td><b>E-mail</b></td>
              </tr>
              <tr>
                <td><%=re("C_TELEFONO")%></td>
                <td><%=re("C_EMAIL")%></td>
              </tr>
              <tr>
                <td colspan="2">&nbsp;</td>
              </tr>
              <tr>
                <td colspan="2"><b>Direcci&oacute;n</b></td>
              </tr>
              <tr>
                <td colspan="2"><%=re("C_DIRECCION")%><br>
				<%=re("C_CP")%>&nbsp;<%=re("C_POBLACION")%><br>
				<%=re("C_PROVINCIA")%><br>
				<%=re("C_PAIS")%></td>
              </tr>
            </table>
            <%
	end if


case "ampliar_empresa"

	set re = mConsulta("SELECT * FROM EMPRESAS WHERE E_ID = "& id_cliente, conn_, 1)
	if re.eof then
		unerror = true : msgerror = "No se ha encontrado el cliente solicitado."
	end if
	
	if not unerror then
		%>
            <span class="nombre_empresa_ampliado"><%=re("E_NOMBRE")%></span>
            <br>
            <br>
            <a href="index.asp?ac=nuevo_servicio&id_cliente=<%=re("E_ID")%>&empresa=<%=re("E_NOMBRE")%>"><img src="arch/new.gif" width="18" height="18" border="0" align="absmiddle"> Servicio nuevo</a> <a href="index.asp?ac=nuevo_contacto&id_empresa=<%=re("E_ID")%>&empresa=<%=re("E_NOMBRE")%>"><img src="arch/new.gif" width="18" height="18" border="0" align="absmiddle"> Contacto nuevo</a>
            <br>
            <br>            
            <table  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td valign="top">                  <fieldset>
                  <legend>Contacto </legend>
                    <table width="100%"  border="0" cellspacing="0" cellpadding="6">
                      <tr>
                        <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td><b>Tel&eacute;fono</b></td>
                          </tr>
                          <tr>
                            <td><%=re("E_TELEFONO")%></td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                          </tr>
                          <tr>
                            <td><b>Direcci&oacute;n</b></td>
                          </tr>
                          <tr>
                            <td><%=re("E_DIRECCION")%><br>
                                <%=re("E_CP")%>&nbsp;<%=re("E_POBLACION")%><br>
                                <%=re("E_PROVINCIA")%><br>
                                <%=re("E_PAIS")%></td>
                          </tr>
                        </table></td>
                      </tr>
                    </table>
                  </fieldset></td>
                <td width="5" valign="top">&nbsp;</td>
                <td valign="top">			<b>Contactos vinculados</b><br>			
                  <%
			' Si sólo hay un contacto
			if mCPalabras(re("E_CONTACTOS"),"|") <= 2 then

				set re = mConsulta("SELECT * FROM CONTACTOS WHERE C_ID = "& mNumero(replace(re("E_CONTACTOS"),"|","")) &"", conn_, 1)
				if not re.eof then%>					<a href="index.asp?ac=ampliar_contacto&id=<%=re("C_ID")%>"><%=re("C_NOMBRE")%>&nbsp;<%=re("C_APELLIDOS")%></a>					<%else%>
					No se ha encontrado el contacto vinculado.
			      <%end if

			else
				%>
			      Hay más de un contacto. No implementado.
                  <%
			end if%>
			</td>
              </tr>
            </table>
            <br>

			
			<%' Lista de servicios
			' ------------------------------------------
			set re = mConsulta("SELECT * FROM SERVICIOS, TIPOS_DE_SERVICIOS WHERE (T_ID = S_TIPO) AND S_ID_EMPRESA = "& id_cliente &" ORDER BY S_INICIO DESC",conn_,1)
			%>
					<table width="100%"  border="0" cellpadding="2" cellspacing="0">
					<tr>
					  <td><table width="100%"  border="0" cellpadding="2" cellspacing="0">
						<tr>
						  <td><b>Servicios</b></td>
						</tr>
						<tr>
						  <td><img src="arch/linea.gif" width="100%" height="1"></td>
						</tr>
					  </table>
					  </td>
					</tr>
					
			</table>
					
					<%if not re.eof then%>
					
					
					<table width="100%"  border="0" cellpadding="2" cellspacing="0">
					  <tr bgcolor="#6375D6">
						<td width="10">&nbsp;</td>
						<td><b><font color="#FFFFFF">T&iacute;tulo</font></b></td>
						<td><b><font color="#FFFFFF">Tipo</font></b></td>
						<td><b><font color="#FFFFFF">Descripci&oacute;n</font></b></td>
						<td align="right"><b><font color="#FFFFFF">Gastos</font></b></td>
						<td align="right"><b><font color="#FFFFFF">Precio</font></b></td>
						<td align="right">&nbsp;</td>
						<td width="10" align="right">&nbsp;</td>
					  </tr>
					  <%
					  total = 0
					  totalG = 0
					  while not re.eof
					  gastos = re("S_GASTOS")
					  precio = re("S_PRECIO")
					  total = total + precio
					  totalG = totalG + gastos
						fila1 = "#EDF3FE"
						fila2 = "#FFFFFF"
						if fila = fila1 then
							fila = fila2
						else
							fila = fila1
						end if%>
					  <tr bgcolor="<%=fila%>">
						<td>&nbsp;</td>
						<td><%=re("S_TITULO")%></td>
						<td><%=re("T_TITULO")%></td>
						<td><%=re("S_DESCRIPCION")%></td>
						<td align="right"><%if ""& session("usuario") = "1" then Response.Write mEuros(gastos) &" &euro;" end if%></td>
						<td align="right"><%=mEuros(precio)%>&euro;</td>
						<td width="18" align="center"><a href="index.asp?ac=eliminar_servicio&id=<%=re("S_ID")%>" onClick="if(!confirm('¿Eliminar <%=re("S_TITULO")%>?')){return false;}"><img src="arch/x.gif" alt="Eliminar" width="18" height="18" border="0" align="absmiddle"></a></td>
						<td align="right">&nbsp;</td>
					  </tr>

					  <%re.movenext
				wend%>
					  <tr>
					    <td colspan="8"><img src="arch/linea.gif" width="100%" height="1"></td>
				      </tr>
					  <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td align="right"><%if ""& session("usuario") = "1" then Response.Write mEuros(totalG) &" &euro;" end if%></td>
                        <td align="right"><span title="<%=formatnumber(total*166.386,0)%>"><%=mEuros(total)%> &euro;</span></td>
                        <td align="center">&nbsp;</td>
                        <td align="right">&nbsp;</td>
				      </tr>
					  <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td align="right">&nbsp;</td>
                        <td align="right"><%if ""& session("usuario") = "1" then%><span title="<%=formatnumber((total-totalg)*166.386,0)%> pts"><b><%=mEuros(total-totalg)%> &euro;</b></span><%end if%></td>
                        <td align="center">&nbsp;</td>
                        <td align="right">&nbsp;</td>
				      </tr>
			</table>
					<%else%>
					
					  <div align="center" class="msginfo">No hay ningún servicio</div>
					<%end if%>
					
					<%
			mCierra(re)

	end if

case "contactos"

select case orden
case "nombre"
	orden = " ORDER BY C_NOMBRE"
case "email"
	orden = " ORDER BY C_EMAIL"
case "movil"
	orden = " ORDER BY C_MOVIL"
case "telefono"
	orden = " ORDER BY C_TELEFONO"
end select
set re = mConsulta("SELECT * FROM CONTACTOS"& orden &"",conn_,1)
	%>
            <table  border="0" cellspacing="0" cellpadding="2">
			<tr><td><table width="100%"  border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td><b>Contactos</b></td>
              </tr>
              <tr>
                <td><img src="arch/linea.gif" width="323" height="1"></td>
              </tr>
            </table></td>
			</tr>
			</table>
			
			<%if not re.eof then %>
			<table width="100%"  border="0" cellspacing="0" cellpadding="2">
			  <tr bgcolor="#6375D6">
			    <td width="10">&nbsp;</td>
				<td><a href="index.asp?ac=<%=ac%>&orden=nombre"><b><font color="#FFFFFF">Nombre</font></b></a></td>
                <td><a href="index.asp?ac=<%=ac%>&orden=email"><b><font color="#FFFFFF">E-mail</font></b></a></td>
                <td><a href="index.asp?ac=<%=ac%>&orden=movil"><b><font color="#FFFFFF">M&oacute;vil</font></b></a></td>
                <td><a href="index.asp?ac=<%=ac%>&orden=telefono"><b><font color="#FFFFFF">Tel&eacute;fono</font></b></a></td>
                <td>&nbsp;</td>
                <td width="10">&nbsp;</td>
			  </tr>
			<%while not re.eof
				fila1 = "#EDF3FE"
				fila2 = "#FFFFFF"
				if fila = fila1 then
					fila = fila2
				else
					fila = fila1
				end if%>

				<tr bgcolor="<%=fila%>">
				  <td>&nbsp;</td>
				<td><a href="index.asp?ac=ampliar_contacto&id=<%=re("C_ID")%>&r=<%=r%>"><%=re("C_NOMBRE")%></a></td>
				<td><a href="mailto:<%=re("C_EMAIL")%>"><%=re("C_EMAIL")%></a></td>
				<td><%=re("C_MOVIL")%></td>
				<td><%=re("C_TELEFONO")%></td>
                <td width="18"><a href="index.asp?ac=eliminar_contacto&id=<%=re("C_ID")%>" onClick="if(!confirm('&iquest;Eliminar <%=re("C_NOMBRE")%>?')){return false;}"><img src="arch/x.gif" alt="Eliminar" width="18" height="18" border="0" align="absmiddle"></a></td>
                <td>&nbsp;</td>
              </tr>
              <%re.movenext
		wend%>
              
            </table>
			
			<%else%>
			
              <div align="center" class="msginfo">No hay ningún cliente</div>
			<%end if%>
            <%
	mCierra(re)

case "tipos_de_servicios"

if ""& request.Form("t_titulo") <> "" then

		set re = mConsulta("SELECT * FROM TIPOS_DE_SERVICIOS",conn_,2)
		re.addNew
		
			re("T_TITULO") = ""& request.Form("t_titulo")
			re("T_DESCRIPCION") = ""& request.Form("t_descripcion")
		
		re.update
		mCierra(re)
		
		Response.Redirect("index.asp?ac=tipos_de_servicios")

else

set re = mConsulta("SELECT * FROM TIPOS_DE_SERVICIOS",conn_,1)
	%>
            <table width="100%"  border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td><table width="100%"  border="0" cellspacing="0" cellpadding="2">
                    <tr>
                      <td><b>Tipos de servicios</b></td>
                    </tr>
                    <tr>
                      <td><img src="arch/linea.gif" width="323" height="1"></td>
                    </tr>
                </table></td>
              </tr>
            </table>
            <table  border="0" cellspacing="0" cellpadding="0">
              <tr valign="top">
                <td><form action="index.asp?ac=tipos_de_servicios" method="post" name="f" id="f">
                  <table  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><fieldset>
                        <legend>Datos del contacto</legend>
                        <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                          <tr>
                            <td><table width="100%"  border="0" cellspacing="0" cellpadding="1">
                                <tr>
                                  <td align="right">T&iacute;tulo</td>
                                  <td>&nbsp;</td>
                                  <td><input name="t_titulo" type="text" class="campo" id="t_titulo" size="23"></td>
                                </tr>
                            </table></td>
                          </tr>
                        </table>
                      </fieldset></td>
                    </tr>
                  </table>
                  <br>
                  <table  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><fieldset>
                        <legend>Descripci&oacute;n</legend>
                        <table width="100%"  border="0" cellspacing="0" cellpadding="10">
                          <tr>
                            <td><textarea name="t_descripcion" cols="23" rows="5" wrap="VIRTUAL" class="area" id="t_descripcion"></textarea></td>
                          </tr>
                        </table>
                      </fieldset></td>
                    </tr>
                  </table>
                  <br>
                  <div align="right"><input type="submit" value="Nuevo"></div>
                </form>
                </td>
                <td width="15">&nbsp;</td>
                <td><%if not re.eof then %>
				  <table width="100%"  border="0" cellpadding="1" cellspacing="0" bgcolor="#6375D6">
                    <tr>
                      <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
                        <tr>
                          <td><table width="100%"  border="0" cellspacing="0" cellpadding="2">
                              <tr bgcolor="#6375D6">
                                <td width="5">&nbsp;</td>
                                <td><b><font color="#FFFFFF">T&iacute;tulo</font></b></td>
                                <td>&nbsp;</td>
                                <td width="5">&nbsp;</td>
                              </tr>
                              <%while not re.eof
				fila1 = "#EDF3FE"
				fila2 = "#FFFFFF"
				if fila = fila1 then
					fila = fila2
				else
					fila = fila1
				end if%>
                              <tr bgcolor="<%=fila%>">
                                <td>&nbsp;</td>
                                <td><%=re("T_TITULO")%></td>
                                <td width="18"><a href="index.asp?ac=eliminar_tipo_de_servicio&id=<%=re("T_ID")%>" onClick="if(!confirm('&iquest;Eliminar <%=re("T_TITULO")%>?')){return false;}"><img src="arch/x.gif" alt="Eliminar" width="18" height="18" border="0" align="absmiddle"></a></td>
                                <td>&nbsp;</td>
                              </tr>
                              <%re.movenext
		wend%>
                          </table></td>
                        </tr>
                      </table></td>
                    </tr>
                  </table>
				  <%else%>
                  <div align="center" class="msginfo">No hay ning&uacute;n tipo</div>
                <%end if%></td>
              </tr>
            </table>
            <%
	mCierra(re)
end if ' if de insertar


case "empresas"

	set re = mConsulta("SELECT * FROM EMPRESAS ORDER BY E_NOMBRE",conn_,1)
	%>
            <table width="100%"  border="0" cellspacing="0" cellpadding="2">
			<tr><td><table width="100%"  border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td><b>Empresas</b></td>
              </tr>
              <tr>
                <td><img src="arch/linea.gif" width="323" height="1"></td>
              </tr>
            </table></td>
			</tr>
			</table>
			
			<%if not re.eof then %>
			<table width="100%"  border="0" cellspacing="0" cellpadding="2">
			  <tr bgcolor="#6375D6">
			    <td width="10">&nbsp;</td>
				<td><b><font color="#FFFFFF">Empresa</font></b></td>
                <td><b><font color="#FFFFFF">E-mail</font></b></td>
                <td><b><font color="#FFFFFF">Tel&eacute;fono</font></b></td>
                <td>&nbsp;</td>
                <td width="10">&nbsp;</td>
			  </tr>
			<%while not re.eof
				fila1 = "#EDF3FE"
				fila2 = "#FFFFFF"
				if fila = fila1 then
					fila = fila2
				else
					fila = fila1
				end if%>

				<tr bgcolor="<%=fila%>">
				  <td>&nbsp;</td>
				<td><a href="index.asp?ac=ampliar_empresa&id_cliente=<%=re("E_ID")%>&r=<%=r%>" title="<%=re("E_ACTIVIDAD")%> - <%=re("E_ID")%>"><%=re("E_NOMBRE")%></a></td>
				<td><a href="mailto:<%=re("E_EMAIL")%>"><%=re("E_EMAIL")%></a></td>
				<td><%=re("E_TELEFONO")%></td>
                <td width="18"><a href="index.asp?ac=eliminar_empresa&id=<%=re("E_ID")%>" onClick="if(!confirm('&iquest;Eliminar <%=re("E_NOMBRE")%>?')){return false;}"><img src="arch/x.gif" alt="Eliminar" width="18" height="18" border="0" align="absmiddle"></a></td>
                <td>&nbsp;</td>
              </tr>
              <%re.movenext
		wend%>
              
            </table>
			
			<%else%>
			
              <div align="center" class="msginfo">No hay ningún cliente</div>
			<%end if%>
            <%
	mCierra(re)

case else

	Response.Redirect("index.asp?ac=servicios&r="&r)
	
end select
%></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
