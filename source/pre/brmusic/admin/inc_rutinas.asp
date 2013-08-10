<%
		' ***   TRATAMIENTO DE CADENAS   ***
		' ------------------------------------------------------------------------------------------------------

		function quitarAcentos(cadena)
			if cadena <> "" then
				cadena = Replace(cadena," ","")
				cadena = Replace(cadena,"�","a")
				cadena = Replace(cadena,"�","e")
				cadena = Replace(cadena,"�","i")
				cadena = Replace(cadena,"�","o")
				cadena = Replace(cadena,"�","u")
				CADENA = REPLACE(CADENA,"�","A")
				CADENA = REPLACE(CADENA,"�","E")
				CADENA = REPLACE(CADENA,"�","I")
				CADENA = REPLACE(CADENA,"�","O")
				CADENA = REPLACE(CADENA,"�","U")
				quitarAcentos = cadena
			else
				quitaracentos = cadena
			end if
		end function

		function code(texto)
			dim t
			t = ""&texto
			if t<> "" then
				t = replace(t,"�","[a]")
				t = replace(t,"�","[e]")
				t = replace(t,"�","[i]")
				t = replace(t,"�","[o]")
				t = replace(t,"�","[u]")
				t = replace(t,"�","[A]")
				t = replace(t,"�","[E]")
				t = replace(t,"�","[I]")
				t = replace(t,"�","[O]")
				t = replace(t,"�","[U]")
				
				t = replace(t,"�","[a2]")
				t = replace(t,"�","[e2]")
				t = replace(t,"�","[i2]")
				t = replace(t,"�","[o2]")
				t = replace(t,"�","[u2]")
				t = replace(t,"�","[A2]")
				t = replace(t,"�","[E2]")
				t = replace(t,"�","[I2]")
				t = replace(t,"�","[O2]")
				t = replace(t,"�","[U2]")
				
				t = replace(t,"�","[n]")
				t = replace(t,"�","[N]")

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
				t = replace(texto,"[a]","�")
				t = replace(t,"[e]","�")
				t = replace(t,"[i]","�")
				t = replace(t,"[o]","�")
				t = replace(t,"[u]","�")
				t = replace(t,"[A]","�")
				t = replace(t,"[E]","�")
				t = replace(t,"[I]","�")
				t = replace(t,"[O]","�")
				t = replace(t,"[U]","�")

				t = replace(t,"[a2]","�")
				t = replace(t,"[e2]","�")
				t = replace(t,"[i2]","�")
				t = replace(t,"[o2]","�")
				t = replace(t,"[u2]","�")
				t = replace(t,"[A2]","�")
				t = replace(t,"[E2]","�")
				t = replace(t,"[I2]","�")
				t = replace(t,"[O2]","�")
				t = replace(t,"[U2]","�")
				
				t = replace(t,"[n]","�")
				t = replace(t,"[N]","�")
				
				t = replace(t,"[as]","&")
				t = replace(t,"[ig]","=")
			end if
			uncode = t
		end function
		
		' Filtro que adapta los caracteres no v�lidos para XML
		function filtroXML(pCadena)
			dim cadena
			cadena = ""&pCadena
			
			cadena = replace(cadena,"�","E")
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

		' Convierte lo que le pasemos a n�mero, en caso de nos ser un n�mero v�lido devuelve 0
		function numero(n)
			if ""&n <> "" and isNumeric(n) then
				numero = 0+n
			else
				numero = 0
			end if			
		end function

		' Convierte lo que le pasemos a n�mero decimal, en caso de nos ser un n�mero v�lido devuelve 0
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
				t = replace(t,"�","&euro;")
				t = replace(t,"�","&ordf;")
				t = replace(t,"�","&ordm;")
				t = replace(t,"�","&Ccedil;")
				t = replace(t,"�","&ccedil;")
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

		' Obtener la �ltima palabra de una cadena: algo/otracosa/ => "otracosa"
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

				txt = replace(txt,"�","&aacute;")
				txt = replace(txt,"�","&eacute;")
				txt = replace(txt,"�","&iacute;")
				txt = replace(txt,"�","&oacute;")
				txt = replace(txt,"�","&uacute;")

				txt = replace(txt,"�","&Aacute;")
				txt = replace(txt,"�","&Eacute;")
				txt = replace(txt,"�","&Iacute;")
				txt = replace(txt,"�","&Oacute;")
				txt = replace(txt,"�","&Uacute;")

				txt = replace(txt,chr(34),"&quot;")
				txt = replace(txt,"�","&middot;")
				txt = replace(txt,"�","&iquest;")
				txt = replace(txt,"�","&iexcl;")
				txt = replace(txt,"�","&ordm;")
				txt = replace(txt,"�","&ordf;")
				txt = replace(txt,"�","&ccedil;")
				txt = replace(txt,"�","&Ccedil;")
				txt = replace(txt,"�","&ntilde;")
				txt = replace(txt,"�","&Ntilde;")

				txt = replace(txt,vbCrlf,"<br>")

				escribeHtml = txt

			end if
		end function

		' Escribe
		function escribe ( txt )
			txt = replace(txt,"+"," ")
			txt = replace(txt,"%E1","�")
			txt = replace(txt,"%E9","�")
			txt = replace(txt,"%ED","�")
			txt = replace(txt,"%F3","�")
			txt = replace(txt,"%FA","�")
			txt = replace(txt,"%C1","�")
			txt = replace(txt,"%C9","�")
			txt = replace(txt,"%CD","�")
			txt = replace(txt,"%D3","�")
			txt = replace(txt,"%DA","�")
		'	'--
			txt = replace(txt,"%28","(")
			txt = replace(txt,"%29",")")
			txt = replace(txt,"%2C",",")
			txt = replace(txt,"%3B",";")
			txt = replace(txt,"%3A",":")
		'	' --
			txt = replace(txt,"%BA","�")
			txt = replace(txt,"%AA","�")
			txt = replace(txt,"%BF","�")
			txt = replace(txt,"%3F","?")
			txt = replace(txt,"%F1","�")
			txt = replace(txt,"%5B","[")
			txt = replace(txt,"%5D","]")
			txt = replace(txt,"%26","&")
			txt = replace(txt,"%22","&quot;")
			txt = replace(txt,"%27","'")
			txt = replace(txt,"%2F","/")
			txt = replace(txt,"%5C","\")
			txt = replace(txt,"%80","�")
			txt = replace(txt,"%2B","+")
			txt = replace(txt,"%25","%")
			txt = replace(txt,"%3D","=")
		
			txt = replace(txt,"%3C","<")
			txt = replace(txt,"%3E",">")
			txt = replace(txt,"%21","!")
			txt = replace(txt,"%A1","�")
			txt = replace(txt,"%23","#")
			txt = replace(txt,"%E4","�")
			txt = replace(txt,"%EB","�")
			txt = replace(txt,"%EF","�")
			txt = replace(txt,"%F6","�")
			txt = replace(txt,"%FC","�")
			txt = replace(txt,"%A8","�")
			txt = replace(txt,"%C4","�")
			txt = replace(txt,"%CB","�")
			txt = replace(txt,"%CF","�")
			txt = replace(txt,"%D6","�")
			txt = replace(txt,"%DC","�")
			txt = replace(txt,"%E2","�")
			txt = replace(txt,"%EA","�")
			txt = replace(txt,"%EE","�")
			txt = replace(txt,"%F4","�")
			txt = replace(txt,"%FB","�")
			txt = replace(txt,"%C2","�")
			txt = replace(txt,"%CA","�")
			txt = replace(txt,"%CE","�")
			txt = replace(txt,"%D4","�")
			txt = replace(txt,"%DB","�")
			txt = replace(txt,"%E0","�")
			txt = replace(txt,"%E8","�")
			txt = replace(txt,"%EC","�")
			txt = replace(txt,"%F2","�")
			txt = replace(txt,"%F9","�")
			txt = replace(txt,"%C0","�")
			txt = replace(txt,"%C8","�")
			txt = replace(txt,"%CC","�")
			txt = replace(txt,"%D2","�")
			txt = replace(txt,"%D9","�")
			txt = replace(txt,"%7C","|")
			txt = replace(txt,"%AC","�")
			txt = replace(txt,"%0D%0A",vbCrLf)
			escribe = (txt)
		End Function

		'Pasa un texto a UTF
		Function utf (dato)
			valor = dato
			valor = Replace(dato,"�","%C3%A1")
			valor = Replace(valor,"�","%C3%A9")
			valor = Replace(valor,"�","%C3%AD")
			valor = Replace(valor,"�","%C3%B3")
			valor = Replace(valor,"�","%C3%BA")
			valor = Replace(valor,"�","%C3%81")
			valor = Replace(valor,"�","%C3%89")
			valor = Replace(valor,"�","%C3%A9")
			valor = Replace(valor,"�","%C3%8D")
			valor = Replace(valor,"�","%C3%9A")
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
'				txt.WriteLine(texto) ' S�lo una l�nea
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

		' Devuelve true o false seg�n exista o no el archivo que le pasamos
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

		
		'   ***   CALCULOS NUM�RICOS   ***
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
					unerror = true : msgerror = "No hay una conexi�n activa o la cadena de conex�on esta vacia o es incorrecta."
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