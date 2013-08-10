<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/datos/inc_config_gen.asp" -->
<!--#include virtual="/admin/inc_rutinas.asp" -->
<html> 
<head> 
<style type="text/css"> 
<!-- 
body, table,tr,td {
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:8pt;} 
--> 
</style> 
</head> 
<body bgcolor="#FFFFFF" text="#000000"> 

<% 
	dim idioma
	dim cualid
	dim ruta_xml_db_config
	dim ruta_db
	dim xmlObj
	dim conn_
	dim oConn
	dim objADOX

	idioma = ""&session("idioma")
	cualid = ""&session("cualid")
	ruta_xml_db_config = "/"& c_s &"datos/"& idioma &"/"& cualid &"/db_config.xml"
	
	if idioma = "" or cualid = "" then
		unerror = true : msgerror = "No está validado en el sistema."
	end if
	
	if not unerror then
		ruta_db = "/"& c_s &"datos/"& idioma &"/"& cualid &"/"& cualid &".mdb"
		conn_ = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source="& Server.MapPath(ruta_db)
		if not existe(Server.MapPath(ruta_db)) then
			unerror = true : msgerror = "La base de datos: <b>"& ruta_db &"</b> no existe."
		end if
	end if

	if not unerror then
		Set xmlObj = CreateObject("MSXML.DOMDocument")
		if not xmlObj.Load(server.MapPath(ruta_xml_db_config)) then
			unerror = true : msgerror = "No se ha podido cargar el XML de configuración: "&ruta_xml_db_config 
		end if
	end if

	if not unerror then
		Set tablas = xmlObj.selectnodes("/configuracion/tabla")
		if not typeOK(tablas) then
			unerror = true : msgerror = "El XML de configuración no es correcto. Consulte a su administrador."
		end if
	end if

	if not unerror then
		on error resume next
		Set oConn = Server.CreateObject("ADODB.Connection")
		Set objADOX = Server.CreateObject("ADOX.Catalog")
		objADOX.ActiveConnection = conn_
		oConn.Open conn_
		if err<>0 then
			unerror = true : msgerror = "No se ha logrado iniciar el objeto ADOX.Catalog.<br>"& err.description
		end if
		on error goto 0
	end if
	
	if not unerror then

		informe = ""

		for t=0 to tablas.length-1
			Set tabla = xmlObj.selectnodes("/configuracion/tabla").item(t)
			tablanueva = 0
			if not Existelatabla(tabla.getattribute("nombre")) then
				sql = "CREATE TABLE "&tabla.getattribute("nombre")&" "&tabla.getattribute("claveunica")
				oConn.Execute sql	
				tablanueva = 1
				informe = informe & "<b>Creando tabla. </b>"&tabla.getattribute("nombre")&"<br>"
			else
				'informe = informe & "<b>No creo la tabla. </b>"&tabla.getattribute("nombre")&" porque ya existe <br>"
			end if
			ejecutar = 0
			sql = "ALTER TABLE "& tabla.getattribute("nombre")&" "
			
			for n=0 to tabla.childnodes.length-1
				latabla = tabla.getattribute("nombre")
				lacolumna = tabla.childnodes(n).getattribute("nombre")
				tipo = tabla.childnodes(n).getattribute("tipo")
				if tablanueva=0 then
					if Existelacolumna(latabla,lacolumna) then
						tipocol = Tipocolumna(latabla,lacolumna)
						coTipo = correspondenciatipo(tipocol)
						if instr(tipo,"TEXT")>0 then
							if not strcomp(coTipo &"("&tamanocolumna(latabla,lacolumna)&")",tipo)=0 then
								informe = informe & "Columna <b>"& lacolumna &"</b> pasa de tipo <b>"& coTipo &"</b> a <b>TEXT</b><br>"
								sql = "ALTER TABLE "&latabla&" ALTER COLUMN "& lacolumna &" "& tipo
								oConn.Execute sql
							end if
							
						elseif instr(tipo,"INTEGER") then
							if not strcomp(coTipo,left(tipo,7))=0 then
								informe = informe & "Columna <b>"& lacolumna &"</b> pasa de tipo <b>"& coTipo &"</b> a <b>INTEGER</b><br>"
								sql = "ALTER TABLE "&latabla&" ALTER COLUMN "& lacolumna &" "& tipo
								oConn.Execute sql
							end if

						elseif instr(tipo,"DOUBLE") then
							if not strcomp(coTipo,left(tipo,6))=0 then
								informe = informe & "Columna <b>"& lacolumna &"</b> pasa de tipo <b>"& coTipo &"</b> a <b>DOUBLE</b><br>"
								sql = "ALTER TABLE "&latabla&" ALTER COLUMN "& lacolumna &" "& tipo
								oConn.Execute sql
							end if

						elseif instr(tipo,"MEMO")  then
							if not strcomp(coTipo,left(tipo,4))=0 then
								informe = informe & "Columna <b>"& lacolumna &"</b> pasa de tipo <b>"& coTipo &"</b> a <b>MEMO</b><br>"
								sql = "ALTER TABLE "&latabla&" ALTER COLUMN "& lacolumna &" "& tipo
								oConn.Execute sql
							end if

						elseif instr(tipo,"DATETIME") then
							if not strcomp(coTipo,left(tipo,8))=0 then
								informe = informe & "Columna <b>"& lacolumna &"</b> pasa de tipo <b>"& coTipo &"</b> a <b>DATETIME</b><br>"
								sql = "ALTER TABLE "& latabla &" ALTER COLUMN "& lacolumna &" "& tipo
								oConn.Execute sql
							end if
						end if
					else
						if ejecutar = 0 then
							sql = sql & " ADD COLUMN "& lacolumna &" "& tipo &" "
						else
							sql = sql & ", "& lacolumna &" "& tipo &" "
						end if
						informe = informe & "<b>Creando fila: </b>"& lacolumna &"<br>"
						ejecutar=1
					end if 
				end if
'				response.flush
			next
			
			if ejecutar=1 then
				on error resume next
				oConn.Execute sql
				if err<>0 then
					unerror = true
					msgerror = msgerror & "<br><b>Error en SQL</b><br>" & err.description & "<br><b>SQL: </b>"& sql
				end if
				on error goto 0
			end if
			tablanueva=0
			ejecutar=0
		next
		informe = informe & ""
		oCOnn.Close
		Set oCOnn = Nothing
	end if  ' not unerror
	
	if informe <> "" then
		Response.Write informe
	else
		Response.Write "<b>La base de datos está actualizada.</b>"
	end if
	
	if unerror then%>
		<br>
		<table width="100%"  border="0" cellspacing="3" cellpadding="4">
			<tr>
			<td><%=msgerror%></td>
			</tr>
		</table>
	<%end if

	Function correspondenciatipo (referencia)
		Select case referencia
		case "202"
			correspondenciatipo="TEXT"
		case "203"
			correspondenciatipo="MEMO"
		case "11"
			correspondenciatipo="YESNO"
		case "3"
			correspondenciatipo="INTEGER"
		case "5"
			correspondenciatipo="DOUBLE"
		case "7"
			correspondenciatipo="DATETIME"
		case else
			correspondenciatipo="nodetectado"
		end select
	End Function
	
	Function Existelacolumna(tabla, Byval ColName)
		For Each field in objADOX.Tables(tabla).Columns
			If LCase(ColName) = LCase(field.Name) Then
				'Already exists
				Existelacolumna = True
				Exit Function			
			End If
		Next
		Existelacolumna = False
	End Function

	Function tamanocolumna(tabla, Byval ColName)
		For Each field in objADOX.Tables(tabla).Columns
			If LCase(ColName) = LCase(field.Name) Then
				'Already exists
				tamanocolumna = field.DefinedSize
				Exit Function			
			End If
		Next
		tamanocolumna = False
	End Function

	Function Tipocolumna(tabla, Byval ColName)
		For Each field in objADOX.Tables(tabla).Columns 
			If LCase(ColName) = LCase(field.Name) Then
				'Already exists
				Tipocolumna = field.type
				tamano=field.DefinedSize
				Exit Function			
			End If
		Next
		Tipocolumna = False
	End Function

	Function Existelatabla(Byval TableName)
		For Each tableloop in objADOX.Tables
			If LCase(TableName) = LCase(tableloop.Name) Then
				'Already exists
				Existelatabla = True
				Exit Function
			End If
		Next
		Existelatabla = False
	End Function
%>
</body> 
</html> 
