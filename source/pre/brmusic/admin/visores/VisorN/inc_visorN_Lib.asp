<%
				' ****************************
				' Visor N - Agrupalia
				' Enero 2005 - Manuel Morante
				' ****************************

			Class VisorN
				Private vUnerror, vMsgerror
				Private vCualid
				Private vIdioma
				Private conn_activa
				Private xml_config
				Private config					' Configuración XML del visor

				Private nav_cuerpo
				Private nav_orden
				Private nav_portada
				Private nav_buscar
				Private nav_subtitulo
				Private nav_fecha
				Private nav_fuente
				Private nav_ampliar
				Private nav_activo_secciones
				Private nav_activo_secciones2
				Private nav_regporpag
				Private nav_regporpag_opt
				Private nav_foto
				Private nav_icono
				Private nav_verenlace
				Private nav_archivo
				Private nav_tienda
				Private nav_nom_titulo
				Private nav_nom_fuente
				Private nav_nom_enlace
				Private nav_nom_fecha
				Private nav_nom_foto
				Private nav_nom_icono
				Private nav_nom_archivo

				Sub Class_Initialize()
					vMsgerror = ""
					vUnerror = false
				End Sub

				' Cualidad
				Public Property Let cualid(valor) vCualid = ""& valor end property
				Public Property Get cualid() cualid = vCualid end property

				' Idioma
				Public Property Let idioma(valor) vIdioma = ""& valor end property
				Public Property Get idioma() idioma = vIdioma end property

				' unerror & msgerror
				Public Property Get unerror() unerror = vUnerror end property
				Public Property Get msgerror() msgerror = vMsgerror end property
				
				' Declarar y abrir una conexion activa
				Public Function activar
					dim conn_

					if ""& vIdioma = "" then
						vUnerror = true : vMsgerror = "[Visor N] Especifique un idioma."
					end if
					if ""& vCualid = "" then
						vUnerror = true : vMsgerror = "[Visor N] Especifique una cualidad."
					end if
					if not vUnerror then
						conn_ = "Driver={Microsoft Access Driver (*.mdb)};DBQ= " & Server.MapPath("\"& c_s &"datos\"& vIdioma &"\"& vCualid &"\"& vCualid &".mdb")
						set conn_activa = server.CreateObject("ADODB.Connection")
						on error resume next
						conn_activa.open conn_
						if err<>0 then
							vUnerror = true : vMsgerror = "[Visor N] No se ha encontrado la base de datos. (Idioma: "& vIdioma &", Cualidad: "& vCualid &", Carpeta Sitio: "& c_s &")."
						end if
						on error goto 0
						
						if not vUnerror then
							' Configuración XML del visor
							' ---------------------------
							set xml_config = CreateObject("MSXML.DOMDocument")
							if not xml_config.Load(Server.MapPath("/"& c_s &"datos/xml_admin_config.xml")) then
								vUnerror = true : vMsgerror = "[Visor N] No se ha encontrado el archivo de configuración. Compruebe 'c_s'."
							else
								set config = xml_config.selectSingleNode("configuracion/"&cualid&"/visor")
								if not typeOK(config) then
									vUnerror = true : vMsgerror = "[Visor N] No se ha encontrado en nodo de configuración para el visor."
								end if
							end if

							nav_cuerpo = ""&config.getAttribute("cuerpo")
							nav_orden = ""&config.getAttribute("orden")
							nav_portada = numero(config.getAttribute("portada"))
							nav_buscar = bool(config.getAttribute("buscar"))
							nav_subtitulo = bool(config.getAttribute("subtitulo"))
							nav_fecha = bool(config.getAttribute("fecha"))
							nav_fuente = bool(config.getAttribute("fuente"))
							nav_ampliar = bool(config.getAttribute("ampliar"))
							nav_activo_secciones = numero(config.getAttribute("secciones"))
							nav_activo_secciones2 = numero(config.getAttribute("secciones2"))
							nav_regporpag = numero(config.getAttribute("regporpag"))
							nav_regporpag_opt = numero(config.getAttribute("regporpag_opt"))
							nav_foto = bool(config.getAttribute("foto"))
							nav_icono = bool(config.getAttribute("icono"))
							nav_verenlace = bool((config.getAttribute("verenlace")))
							nav_archivo = bool(config.getAttribute("archivo"))
							nav_tienda = bool(config.getAttribute("tienda"))
							nav_nom_titulo = ""&config.getAttribute("nombretitulo")
							nav_nom_fuente = ""&config.getAttribute("nombrefuente")
							nav_nom_enlace = ""&config.getAttribute("nombreenlace")
							nav_nom_fecha = ""&config.getAttribute("nombrefecha")
							nav_nom_foto = ""&config.getAttribute("nombrefoto")
							nav_nom_icono = ""&config.getAttribute("nombreicono")
							nav_nom_archivo = ""&config.getAttribute("nombrearchivo")
						end if
					end if
				end function

				Public Function tabla
					dim sql
					dim re
					dim str
					sql = "SELECT * FROM REGISTROS"
					set re = Server.CreateObject("ADODB.Recordset")
					if not vUnerror then
						re.Open sql, conn_activa
							if not re.eof then
								str = "<table>"
								while not re.eof
									str = str & "<tr><td>"& re("R_TITULO") &"</td></tr>"
									re.MoveNext()
								wend
								str = str & "</table>"
							end if
						re.Close()
					end if

					tabla = str
				end Function

			End Class
%>

