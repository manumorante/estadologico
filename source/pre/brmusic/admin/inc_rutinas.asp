<%
		' ***   TRATAMIENTO DE CADENAS   ***
		' ------------------------------------------------------------------------------------------------------

		function quitarAcentos(cadena)
			if cadena <> "" then
				cadena = Replace(cadena," ","")
				cadena = Replace(cadena,"á","a")
				cadena = Replace(cadena,"é","e")
				cadena = Replace(cadena,"í","i")
				cadena = Replace(cadena,"ó","o")
				cadena = Replace(cadena,"ú","u")
				CADENA = REPLACE(CADENA,"Á","A")
				CADENA = REPLACE(CADENA,"É","E")
				CADENA = REPLACE(CADENA,"Í","I")
				CADENA = REPLACE(CADENA,"Ó","O")
				CADENA = REPLACE(CADENA,"Ú","U")
				quitarAcentos = cadena
			else
				quitaracentos = cadena
			end if
		end function

		function code(texto)
			dim t
			t = ""&texto
			if t<> "" then
				t = replace(t,"á","[a]")
				t = replace(t,"é","[e]")
				t = replace(t,"í","[i]")
				t = replace(t,"ó","[o]")
				t = replace(t,"ú","[u]")
				t = replace(t,"Á","[A]")
				t = replace(t,"É","[E]")
				t = replace(t,"Í","[I]")
				t = replace(t,"Ó","[O]")
				t = replace(t,"Ú","[U]")
				
				t = replace(t,"à","[a2]")
				t = replace(t,"è","[e2]")
				t = replace(t,"ì","[i2]")
				t = replace(t,"ò","[o2]")
				t = replace(t,"ù","[u2]")
				t = replace(t,"À","[A2]")
				t = replace(t,"È","[E2]")
				t = replace(t,"Ì","[I2]")
				t = replace(t,"Ò","[O2]")
				t = replace(t,"Ù","[U2]")
				
				t = replace(t,"ñ","[n]")
				t = replace(t,"Ñ","[N]")

				t = replace(t,"&","[as]")
				t = replace(t,"=","[ig]")
			
				' Cambio comillas dobles por simples
				t = replace(t,chr(34),"'")
			end if
		
			code = t
		end function
		
	
		function uncode(texto)
			dim t : t = ""&texto
			if t <> "" then
				t = replace(texto,"[a]","á")
				t = replace(t,"[e]","é")
				t = replace(t,"[i]","í")
				t = replace(t,"[o]","ó")
				t = replace(t,"[u]","ú")
				t = replace(t,"[A]","Á")
				t = replace(t,"[E]","É")
				t = replace(t,"[I]","Í")
				t = replace(t,"[O]","Ó")
				t = replace(t,"[U]","Ú")

				t = replace(t,"[a2]","à")
				t = replace(t,"[e2]","è")
				t = replace(t,"[i2]","ì")
				t = replace(t,"[o2]","ò")
				t = replace(t,"[u2]","ù")
				t = replace(t,"[A2]","À")
				t = replace(t,"[E2]","È")
				t = replace(t,"[I2]","Ì")
				t = replace(t,"[O2]","Ò")
				t = replace(t,"[U2]","Ù")
				
				t = replace(t,"[n]","ñ")
				t = replace(t,"[N]","Ñ")
				
				t = replace(t,"[as]","&")
				t = replace(t,"[ig]","=")
			end if
			uncode = t
		end function
		
		' Filtro que adapta los caracteres no válidos para XML
		function filtroXML(pCadena)
			dim cadena
			cadena = ""&pCadena
			
			cadena = replace(cadena,"€","E")
			cadena = replace(cadena,"&nbsp;"," ")
			cadena = replace(cadena,vbCrlf,"")
		
			filtroXML = cadena
		end function

		' Euros
		function euros(n)
			euros = FormatNumber(numero(n),2)
		end function

		' Devuelve true o false si es numero o no
		function esNumero(n)
			if ""&n <> "" and isNumeric(n) and len(n) > 0 then
				esNumero = true
			else
				esNumero = false
			end if
		end function
		
		' Convierte lo que le pasemos a un estado boleano
		function bool(n)
			bool = cbool(numero(n))
		end function

		' Convierte lo que le pasemos a número, en caso de nos ser un número válido devuelve 0
		function numero(n)
			if ""&n <> "" and isNumeric(n) then
				numero = 0+n
			else
				numero = 0
			end if			
		end function

		' Convierte lo que le pasemos a número decimal, en caso de nos ser un número válido devuelve 0
		function decimal(n)
			if ""&n <> "" then
				n = replace(replace(n,"'",","),".",",")
				if ""&n <> "" and isNumeric(n) then
					decimal = CDbl(n)
				end if
			else
				decimal = 0
			end if			
		end function

		' Devuelve un texto cortado a un maximo de caracteres
		function unpoco (txt,maximo)
			if len(txt) > maximo then
				while mid(txt,maximo,1) <> " "
					maximo = maximo - 1
				wend
				unpoco = left(txt,maximo) & " ..."
			else
				unpoco = txt
			end if
		end function


		' Filtro para caracteres especiales que XML no acepta. => los paso a equivalencias HTML
		function filtroHtml(pfiltroHtml)
			dim t : t = ""&pfiltroHtml
			if t <> "" then
				t = replace(t,"€","&euro;")
				t = replace(t,"ª","&ordf;")
				t = replace(t,"º","&ordm;")
				t = replace(t,"Ç","&Ccedil;")
				t = replace(t,"ç","&ccedil;")
			end if
			filtroHtml = t
		end function

		' le pasamos una cadena separado por comas y un valor. Nos devuelce True si el valor se encuentra entre alguna de las comas
		function estaEnCadena (texto,este)
			dim partes
			partes = Split(""&texto, ",")
			for each a in partes
				if trim(este) = trim(a) then
					estaEnCadena = true
					exit function
					
				end if
			next
			estaEnCadena = false
		end function

		' Obtener la extension de un archivo
		function getExtension(t)
			dim n : t = Ucase(""&t)
			for n=0 to len(t)
				c = inStrRev(t,".")
				getExtension = mid(t,c+1,len(t)-c)
			next
		end function

		' Obtener la última palabra de una cadena: algo/otracosa/ => "otracosa"
		Function nombreArchivo(valor)
			dim n
			if ""& valor <> "" then
				for n=0 to len(valor)-1
					if Mid(valor,len(valor)-n,1)="\" or Mid(valor,len(valor)-n,1)="/" then
						n=len(valor)+1
					else
						nombreArchivo=Mid(valor,len(valor)-n,1)&nombreArchivo
					end if
				next
			else
				nombreArchivo = ""
			end if
		end Function
		
		' Cuenta el numero total de veces que encuentra la palabra que le indicamos en una cadena que le indicamos
		function cuentaPalabras(cadena, palabra)
			dim pos, conta
			pos = 1
			conta = 0
			While inStr(pos, cadena, palabra) > 0
				conta = conta + 1
				pos = inStr(pos, cadena, palabra) + Len(palabra)
			Wend   
			cuentaPalabras = conta    
		End Function

		' Devuelve texto listo para verse en navegadores HTML.
		function escribeHtml(ptxt)
			dim txt : txt = ""&ptxt
			if txt = "" then
				escribeHtml = ptxt
			else

				txt = replace(txt,"á","&aacute;")
				txt = replace(txt,"é","&eacute;")
				txt = replace(txt,"í","&iacute;")
				txt = replace(txt,"ó","&oacute;")
				txt = replace(txt,"ú","&uacute;")

				txt = replace(txt,"Á","&Aacute;")
				txt = replace(txt,"É","&Eacute;")
				txt = replace(txt,"Í","&Iacute;")
				txt = replace(txt,"Ó","&Oacute;")
				txt = replace(txt,"Ú","&Uacute;")

				txt = replace(txt,chr(34),"&quot;")
				txt = replace(txt,"·","&middot;")
				txt = replace(txt,"¿","&iquest;")
				txt = replace(txt,"¡","&iexcl;")
				txt = replace(txt,"º","&ordm;")
				txt = replace(txt,"ª","&ordf;")
				txt = replace(txt,"ç","&ccedil;")
				txt = replace(txt,"Ç","&Ccedil;")
				txt = replace(txt,"ñ","&ntilde;")
				txt = replace(txt,"Ñ","&Ntilde;")

				txt = replace(txt,vbCrlf,"<br>")

				escribeHtml = txt

			end if
		end function

		' Escribe
		function escribe ( txt )
			txt = replace(txt,"+"," ")
			txt = replace(txt,"%E1","á")
			txt = replace(txt,"%E9","é")
			txt = replace(txt,"%ED","í")
			txt = replace(txt,"%F3","ó")
			txt = replace(txt,"%FA","ú")
			txt = replace(txt,"%C1","Á")
			txt = replace(txt,"%C9","É")
			txt = replace(txt,"%CD","Í")
			txt = replace(txt,"%D3","Ó")
			txt = replace(txt,"%DA","Ú")
		'	'--
			txt = replace(txt,"%28","(")
			txt = replace(txt,"%29",")")
			txt = replace(txt,"%2C",",")
			txt = replace(txt,"%3B",";")
			txt = replace(txt,"%3A",":")
		'	' --
			txt = replace(txt,"%BA","º")
			txt = replace(txt,"%AA","ª")
			txt = replace(txt,"%BF","¿")
			txt = replace(txt,"%3F","?")
			txt = replace(txt,"%F1","ñ")
			txt = replace(txt,"%5B","[")
			txt = replace(txt,"%5D","]")
			txt = replace(txt,"%26","&")
			txt = replace(txt,"%22","&quot;")
			txt = replace(txt,"%27","'")
			txt = replace(txt,"%2F","/")
			txt = replace(txt,"%5C","\")
			txt = replace(txt,"%80","€")
			txt = replace(txt,"%2B","+")
			txt = replace(txt,"%25","%")
			txt = replace(txt,"%3D","=")
		
			txt = replace(txt,"%3C","<")
			txt = replace(txt,"%3E",">")
			txt = replace(txt,"%21","!")
			txt = replace(txt,"%A1","¡")
			txt = replace(txt,"%23","#")
			txt = replace(txt,"%E4","ä")
			txt = replace(txt,"%EB","ë")
			txt = replace(txt,"%EF","ï")
			txt = replace(txt,"%F6","ö")
			txt = replace(txt,"%FC","ü")
			txt = replace(txt,"%A8","¨")
			txt = replace(txt,"%C4","Ä")
			txt = replace(txt,"%CB","Ë")
			txt = replace(txt,"%CF","Ï")
			txt = replace(txt,"%D6","Ö")
			txt = replace(txt,"%DC","Ü")
			txt = replace(txt,"%E2","â")
			txt = replace(txt,"%EA","ê")
			txt = replace(txt,"%EE","î")
			txt = replace(txt,"%F4","ô")
			txt = replace(txt,"%FB","û")
			txt = replace(txt,"%C2","Â")
			txt = replace(txt,"%CA","Ê")
			txt = replace(txt,"%CE","Î")
			txt = replace(txt,"%D4","Ô")
			txt = replace(txt,"%DB","Û")
			txt = replace(txt,"%E0","à")
			txt = replace(txt,"%E8","è")
			txt = replace(txt,"%EC","ì")
			txt = replace(txt,"%F2","ò")
			txt = replace(txt,"%F9","ù")
			txt = replace(txt,"%C0","À")
			txt = replace(txt,"%C8","È")
			txt = replace(txt,"%CC","Ì")
			txt = replace(txt,"%D2","Ò")
			txt = replace(txt,"%D9","Ù")
			txt = replace(txt,"%7C","|")
			txt = replace(txt,"%AC","¬")
			txt = replace(txt,"%0D%0A",vbCrLf)
			escribe = (txt)
		End Function

		'Pasa un texto a UTF
		Function utf (dato)
			valor = dato
			valor = Replace(dato,"á","%C3%A1")
			valor = Replace(valor,"é","%C3%A9")
			valor = Replace(valor,"í","%C3%AD")
			valor = Replace(valor,"ó","%C3%B3")
			valor = Replace(valor,"ú","%C3%BA")
			valor = Replace(valor,"Á","%C3%81")
			valor = Replace(valor,"É","%C3%89")
			valor = Replace(valor,"Í","%C3%A9")
			valor = Replace(valor,"Ó","%C3%8D")
			valor = Replace(valor,"Ú","%C3%9A")
			utf = valor
		end Function
		
		' ***   MANEJO DE ARCHIVOS Y DIRECTORIOS EN EL SERVIDOR   *** -----------------------------------------------
		
		' Abrir un archivo en formato texto plano
		function abrirTXT(archivo)
			if existe(archivo) then
				dim fso
				set fso = Server.CreateObject("Scripting.FileSystemObject")
				set abrirTXT = fso.OpenTextFile(archivo,1,true)
				set fso = nothing
			end if
		end function

		' Leer una linea especifia de un archivo de texto (le pasamos el objeto abierto)
		function leerLinea(objTxt,linea)
			dim contenido, lineasLeidas, paraDeLeer
			paraDeLeer = false
			contenido = ""
			if typeOK(objTxt) and ""&linea <> "" then
				lineasLeidas = 0
				Do while Not objTxt.AtEndoFStream and not paraDeLeer
					contenido = objTxt.Readline
					lineasLeidas = lineasLeidas + 1
					if ""&linea <> "" then
						if lineasLeidas = linea then
							paraDeLeer = true
						end if
					end if
				Loop
			end if
			leerLinea = contenido
		end function
		
		' Crear un archivo de texto plano pasandole el contenido
		function crearTXT(archivo, texto)
			if ""&archivo<>"" then
				dim fso, txt
				set fso = Server.CreateObject("Scripting.FileSystemObject")
				on error resume next
				set txt = fso.CreateTextFile (archivo,true)
'				txt.WriteLine(texto) ' Sólo una línea
				txt.Write(texto)
				txt.Close

				if err=0 then
					set crearTXT = txt
				else
					set crearTXT = nothing
				end if
				set txt = nothing
				set fso = nothing
				on error goto 0
			end if
		end function
		
		' Crear un carpeta nueva
		function nuevaCarpeta(ruta,sobrescribir)
			dim fso, carpeta
			if ""&ruta <> "" then
				set fso = Server.CreateObject("Scripting.FileSystemObject")
				if fso.FolderExists(ruta) then
					if sobrescribir then
						set carpeta = fso.getFolder(ruta)
						carpeta.delete(true)
						fso.CreateFolder(ruta)
						nuevaCarpeta = true
					else
						nuevaCarpeta = false
					end if
				else
					fso.CreateFolder(ruta)
					nuevaCarpeta = true
				end if
			end if
		end function

		' Abrir un archivo en formato texto plano
		function abrirTXT_2()
			Set MiFSO=Server.CreateObject("Scripting.FileSystemObject")
			Set fichero_leer=MiFSO.OpenTextFile(Server.MapPath(".")&"/"&fichero,1,true)
			contenido=""
			Do while Not fichero_leer.AtEndoFStream and not Instr(Replace(Trim(Lcase(contenido))," ",""),"<body>")>0
				contenido = fichero_leer.Readline
				'response.write( fichero_leer.AtEndoFStream&"-"&Instr(Replace(Trim(Lcase(contenido))," ",""),"<body>")&"-"&contenido&"<br>")
			Loop
		end function

		' Devuelve true o false según exista o no el archivo que le pasamos
		Function existe(fichero)
			dim fso
			set fso = Server.CreateObject("Scripting.FileSystemObject")
			existe = fso.FileExists(fichero)
			set fso = Nothing
		end Function

		' Borrar un archivo del servidor
		Function borrarArchivo (archivo)
			Dim fso
			set fso = Server.CreateObject("Scripting.FileSystemObject")
			if fso.FileExists(archivo) then
				fso.DeleteFile archivo, true
				borrarArchivo = true
			else
				borrarArchivo = false
			end if
			set fso = nothing
		end function

		' Mete el ancho y alto de la imagen en las variable que le paso
		function imgWH(sFileName, width, height, tipo)
			dim strGIF, strType
			strImageType = "(desconocido)"
			imgWH = False
			strType = ""&GetBytes(sFileName, 0, 3)
			if strType = "GIF" then
				strImageType = "GIF"
				Width = lngConvert(GetBytes(sFileName, 7, 2))
				Height = lngConvert(GetBytes(sFileName, 9, 2))
				imgWH = True
			else
				strBuff = GetBytes(sFileName, 0, -1) ' Get all bytes from file
				lngSize = len(strBuff)
				flgFound = 0
				strTarget = chr(255) & chr(216) & chr(255)
				flgFound = instr(strBuff, strTarget)
				if flgFound = 0 then
					exit function
				end if
				strImageType = "JPG"
				lngPos = flgFound + 2
				ExitLoop = false
				do while ExitLoop = False and lngPos < lngSize
					do while asc(mid(strBuff, lngPos, 1)) = 255 and lngPos < lngSize
						lngPos = lngPos + 1
					loop
					if asc(mid(strBuff, lngPos, 1)) < 192 or asc(mid(strBuff, lngPos, 1)) > 195 then
						lngMarkerSize = lngConvert2(mid(strBuff, lngPos + 1, 2))
						lngPos = lngPos + lngMarkerSize + 1
					else
						ExitLoop = True
					end if
				loop
				if ExitLoop = False then
					Width = -1
					Height = -1
				else
					Height = lngConvert2(mid(strBuff, lngPos + 4, 2))
					Width = lngConvert2(mid(strBuff, lngPos + 6, 2))
					imgWH = True
				end if
			end if
			tipo = strImageType
		end function
				
		function GetBytes(sFileName, offset, bytes)
			Dim objFSO, objFTemp, objTextStream, lngSize
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			Set objFTemp = objFSO.GetFile(sFileName)
			lngSize = objFTemp.Size		
			set objFTemp = nothing
			fsoForReading = 1
			Set objTextStream = objFSO.OpenTextFile(sFileName, fsoForReading)
			if offset > 0 then
				strBuff = objTextStream.Read(offset - 1)
			end if
			if bytes = -1 then ' Get All!
				GetBytes = objTextStream.Read(lngSize) 'ReadAll
			else
				GetBytes = objTextStream.Read(bytes)
			end if
			objTextStream.Close
			set objTextStream = nothing
			set objFSO = nothing
		end function
		
		function lngConvert(strTemp)
			lngConvert = clng(asc(left(strTemp, 1)) + ((asc(right(strTemp, 1)) * 256)))
		end function
		
		function lngConvert2(strTemp)
			lngConvert2 = clng(asc(right(strTemp, 1)) + ((asc(left(strTemp, 1)) * 256)))
		end function

		
		'   ***   CALCULOS NUMÉRICOS   ***
		'--------------------------------------------------------------------------------------------

		' Devuelve en parametro mayor
		function getMayor (n1,n2)
			if ""&n1 <> "" and ""&n2 <> "" and isNumeric(n1) and isNumeric(n2) then
				if cint(n1) > cint(n2) then
					getMayor = n1
				elseif cint(n2) > cint(n1) then
					getMayor = n2
				else
					getMayor = n1
				end if
			else
				getMayor = 0
			end if
		end function

		' Devuelve la fecha actual formateado
		function pintafecha()
			dim fecha
			fecha = FormatDateTime(Now, vbLongDate)
			pintafecha = Ucase(Left(fecha,1)) & Right(fecha,len(fecha)-1)
		end function
		
		' Comprobamos que el parametro no sea ni Nothing ni Empty ni Null
		function typeOK(param_ob)
			dim tname
			tname = lcase(typeName(param_ob))
			if tname <> "nothing" and tname <> "empty" and tname <> "null" and tname <> "string" and tname <> "array" and tname <> "integer" then
				typeOK = true
			else
				typeOK = false
			end if
		end function
		
		' *** RELACIONADOS CON CONSULTA SQL ***
		
		' Formatea la consulta SQL para pintarla en una alerta JAVA SCRIPT
		function pintaSqlJS(psql)
			dim sqlt
			sqlt = ""&psql
			sqlt = replace(sqlt,"SELECT","SELECT\n")
			sqlt = replace(sqlt,",",",\n")
			sqlt = replace(sqlt,"AND","\nAND")
			sqlt = replace(sqlt,"FROM","\n\nFROM\n")
			sqlt = replace(sqlt,"WHERE","\n\nWHERE\n")
			sqlt = replace(sqlt,"ORDER BY","\n\nORDER BY\n")
			pintaSqlJS = sqlt
		end function
		
		' Efectua un execute sobre una SQL (depende de una conn_)
		function exeSql(sql,conn)
		
			if ""& sql = "" then
				exeSql = "Sentencia SQL vacia."
			else
				if ""&typeName(conn_activa) = "Connection" then
					on error resume next
					conn_activa.Execute sql
				elseif ""& conn <> "" then
					dim oConn
					set oConn = server.CreateObject("ADODB.Connection")
					oConn.Open conn
					oConn.Execute sql
					oConn.Close
					set oConn = nothing
				else
					unerror = true : msgerror = "No hay una conexión activa o la cadena de conexíon esta vacia o es incorrecta."
					exit function
				end if
	
				if err=0 then
					exeSql = ""
				else
					exeSql = "(exeSql): " & err.description & "<br><b>conn</b>: " & conn & "<br><b>sql</b>: " & sql
				end if
				on error goto 0
				
			end if

		end function


%>