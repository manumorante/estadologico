<%
				ruta_xml_config = Server.MapPath("/"& c_s &"datos/xml_admin_config.xml")
				
				public config_iexplorer			' Activo / Inactivo opciones especiales sólo compatibles con IExplorer
				public config_idioma_eng
				public config_idioma_eng_bd
				public config_idioma_esp
				public config_idioma_esp_bd
				public config_sqlextra
				public config_nombrecualid		' Nombre "bonito" (con espacios y acentuación) de la cualidad
				public config_carpetasitio
				public config_nombresitio
				public config_urlsitio
				public config_portada			' Noticias de portada: (0 = no) (1 o mas = límite) (para visionado por parte de usuario)
				public config_creador
				public config_activo
				public nav_visor
				public config_cuerpo			' Indica el campo opcional que debemos tomar como el "cuerpo" para  mostrarlo en el listado de admin, ...
				public config_orden				' Se ordenna por el campo indicado en esta variable ...
				public config_buscar			' Opción de buscar: boolenao
				public config_subtitulo			' Opción de subtitulo: booleano
				public config_fecha				' Mostrar/Editar fechas: 0=no | 1=si | 2=si, siempre
				public config_hora				' Mostrar/Editar fechas: booleano
				public config_fechainifin		' Usar el cuadro de rango de fechas INI FIN: boolenao
				public config_fuente			' Mostrar/Editar fuente de la noticia: booleano (funciona en conjunto con el campo R_ENLACE de la bd)
				public config_ampliar			' Permite o no ampliar los registros: booleano
				public config_maxcarseccion		' Caracteres máximos en el nomrbe de seccion
				public config_regporpag			' Número de registros por página: (0 = predeterminado a 5)
				public config_regporpag_opt		' Número de opciones que aparen en el desplegable para cambiar los reg por pag: (0 = no aparece)
				public config_foto				' Usar fotos: boolenano
				public config_icono				' Usar iconos: boolenano
				public config_icono_seccion		' Usar iconos para secciones: boolenano
				public config_foto_seccion		' Usar fotos para secciones: boolenano
				public config_icono_seccion2	' Usar iconos para secciones2: boolenano
				public config_foto_seccion2		' Usar fotos para secciones2: boolenano
				public config_verenlace			' Usar el boton ver enlace: booleano
				public config_verenlacearchivo	' Mostrar el boton ver enlace para descargar el archivo: booleano
				public config_verenlacefoto		' Mostrar el boton ver enlace para linkar la foto: booleano
				public config_verenlaceicono	' Mostrar el boton ver enlace para linkar el icono: booleano
				public config_archivo			' usar o no un campo de archivo para descargas de archivos
				public config_descripcion		' Descripción de la cualidad
				public config_ordenidioma		' Usar el campo especial relacionado con los idiomas
				public config_nuevos			' Permite o no insertar nuevos registros (usado cuando los registros se insertan desde otra lugar)
				public config_editar			' Permite o no modificar los registros
				public config_eliminar			' Permite o no eliminar los registros
				public config_sqlorden			' Establece el ORDE BY de la SQL
				public config_activo_seccion	' Activo / Inactivo para secciones
				public config_activo_seccion2	' Activo / Inactivo para sub secciones
				public config_activo_seccion3	' Activo / Inactivo para sub sub secciones
				public config_foro				' cualidad foro activa
				public config_posicion_foto		' Activo / Inactivo para la modificar la posición de la foto.
				public config_posicion_icono	' Activo / Inactivo para la modificar la posición del icono.
				public config_pie_foto			' Activo / Inactivo para el pie de foto.
				public config_pie_icono			' Activo / Inactivo para el pie de icono.
				public config_mover				' Activo / Inactivo Pemite mover registros arriba y abajo en el orden general (R_ORDEN).
				public config_mover_seccion		' Activo / Inactivo Pemite mover registros arriba y abajo en el orden R_ORDEN_seccion
				public config_mover_seccion2	' Activo / Inactivo Pemite mover registros arriba y abajo en el orden R_ORDEN_seccion2
				public config_alfinal			' Declara el valor fijo de "al final" para insertar usuarios. Y ya no se muestra.
				public config_estados
				public config_coloresfecha		' Cambia de color los registros según la fecha del mismo.
				public config_infoemail			' si o no, abre popEmail para informar de cambios por email.

				' DEFINICION DE NOMBRES PARA LOS CAMPOS FIJOS
				public config_nom_titulo
				public config_nom_fuente
				public config_nom_enlace
				public config_nom_fecha
				public config_nom_fechaini
				public config_nom_fechafin
				public config_nom_foto
				public config_nom_icono
				public config_nom_archivo
				public config_nom_secciones
				public config_nom_secciones2
				

				' Visores
				public nav_activo
				public nav_portada
				public nav_cuerpo
				public nav_orden				' Se ordenna por el campo indicado en esta variable ...
				public nav_buscar
				public nav_subtitulo
				public nav_fecha
				public nav_fuente
				public nav_ampliar
				public nav_activo_secciones
				public nav_activo_secciones2
				public nav_regporpag
				public nav_regporpag_opt
				public nav_foto
				public nav_icono
				public nav_verenlace
				public nav_archivo
				
				' DEFINICION DE NOMBRES PARA LOS CAMPOS FIJOS
				public nav_nom_titulo
				public nav_nom_fuente
				public nav_nom_enlace
				public nav_nom_fecha
				public nav_nom_foto
				public nav_nom_icono
				public nav_nom_archivo
				
				public xml_config
				public nodoConfig
				public nodoCualid
				public nodoVisor
				public config_str_idiomas		' Contiene una lista de idioma separdos por "|" los cuales indican que la cualidad actual para dicho el cada idioma de la lista esta 'vinculado' a la base de datos actual.
				
				public sub inicia_xml
					' *** CARGA DEL XML DE CONFIGURACIÓN *** (xml de configuración general para todas la cualidades del sistema en la página actual)
					set xml_config = CreateObject("MSXML.DOMDocument")
					if not xml_config.Load(ruta_xml_config) then
						unerror = true : msgerror = "Hay un problema con el archivo de configuración general.<br>Conpruebe su ubicación y que no contenga ningún error.<br>Ruta:"& ruta_xml_config
					else
						set nodoConfig = xml_config.selectSingleNode("configuracion")
						if not typeOK(nodoConfig) then
							unerror = true : msgerror = "Hay un problema con el archivo de configuración general.<br>Conpruebe que no contenga ningún error."
						end if
					end if
				
					if not unerror then
						' VARIABLE DEL SITIO
						config_carpetasitio = ""& nodoConfig.getAttribute("carpetasitio")
				'		Response.Write "<br> >config_carpetasitio: "& config_carpetasitio & "< <br>"
						config_nombresitio = ""& nodoConfig.getAttribute("nombresitio")
						config_urlsitio = ""& nodoConfig.getAttribute("urlsitio")
						if config_carpetasitio = "" then
							'unerror = true : msgerror = "No se ha declarado la carpeta de este sitio web."
						end if
						
						if config_nombresitio = "" then
							unerror = true : msgerror = "No se ha declarado el nombre de este sitio web."
						end if
						
						if config_urlsitio = "" then
							' No es necesaria
				'			unerror = true : msgerror = "No se ha declarado la URL de este sitio web."
						end if
					end if
					
					if not unerror then
						' Buscamos la cualidad solicitada (cualid) en las cualidades de nuestro XML y declaraoms el NODO
						if ""&cualid <> "" then
							set nodoCualid = xml_config.selectSingleNode("configuracion/"&cualid)
							if typeName(nodoCualid) = "Nothing" or typeName(nodoCualid) = "Empty" then
								unerror = true : msgerror = "La zona de aSkipper (cualidad) solicitada no está disponible."
							end if
						else
							unerror = true : msgerror = "No se ha indicado una zona de visionado (Cualidad)."
						end if
					end if
					
					' DEFINICIÓN DE VARIABLES PARA LAS ESPECIFICACIONES EN CADA ZONA
					if not unerror then

						' Busco los nodos que indican estas 'metidos' en la base de datos actual.
						config_str_idiomas = "|"
						set idiomas = nodoCualid.selectNodes("idioma")
						for each a in idiomas
							if ""&a.getAttribute("bd") <> "" then
								if ""&a.getAttribute("bd") = session("idioma") then
									config_str_idiomas = config_str_idiomas & a.getAttribute("nombre") & "|"
								end if
							end if
						next
						set nodo_esp = nodoCualid.selectNodes("idioma[@nombre='esp']").item(0)
						if typeOK(nodo_esp) then
							config_idioma_esp = true
							config_idioma_esp_bd = ""&nodo_esp.getAttribute("bd")
						else
							config_idioma_esp = false
							config_idioma_esp_bd = ""
						end if
						set nodo_esp = Nothing

						set nodo_eng = nodoCualid.selectNodes("idioma[@nombre='eng']").item(0)
						if typeOK(nodo_eng) then
							config_idioma_eng = true
							config_idioma_eng_bd = ""&nodo_eng.getAttribute("bd")
						else
							config_idioma_eng = false
							config_idioma_eng_bd = ""
						end if
						set nodo_eng = Nothing

						config_iexplorer = ""& nodoCualid.getAttribute("iexplorer")
						if config_iexplorer = "" or config_iexplorer = "1" then
							config_iexplorer = true
						else
							config_iexplorer = false
						end if
						config_nombrecualid = ""&nodoCualid.getAttribute("nombre")
						config_cuerpo = ""&nodoCualid.getAttribute("cuerpo")
						config_orden = ""&nodoCualid.getAttribute("orden")
						config_portada = bool(nodoCualid.getAttribute("portada"))
						config_creador = bool(nodoCualid.getAttribute("creador"))
						config_activo = bool(nodoCualid.getAttribute("activo"))
						config_buscar = bool(nodoCualid.getAttribute("buscar"))
						config_subtitulo = bool(nodoCualid.getAttribute("subtitulo"))
						config_fecha = numero(nodoCualid.getAttribute("fecha"))
						config_hora = bool(nodoCualid.getAttribute("hora"))
						config_fechainifin = bool(nodoCualid.getAttribute("fechainifin"))		
						config_fuente = bool(nodoCualid.getAttribute("fuente"))
						config_ampliar = bool(nodoCualid.getAttribute("ampliar"))
						config_sqlextra = ""& nodoCualid.getAttribute("sqlextra")
						config_alfinal = ""&nodoCualid.getAttribute("alfinal")
						config_estados = bool(nodoCualid.getAttribute("estados"))
						config_coloresfecha = bool(nodoCualid.getAttribute("coloresfecha"))
						config_infoemail = bool(nodoCualid.getAttribute("infoemail"))
						
						config_maxcarseccion = numero(nodoCualid.getAttribute("maxcarseccion"))
						if config_maxcarseccion = 0 then
							config_maxcarseccion = 20
						end if
						config_regporpag = numero(nodoCualid.getAttribute("regporpag"))
						config_regporpag_opt = numero(nodoCualid.getAttribute("regporpag_opt"))
						config_foto = bool((nodoCualid.getAttribute("foto")))
						config_icono = bool((nodoCualid.getAttribute("icono")))
						config_verenlace = bool((nodoCualid.getAttribute("verenlace")))
						config_verenlacearchivo = bool((nodoCualid.getAttribute("verenlacearchivo")))
						config_verenlaceicono = bool((nodoCualid.getAttribute("verenlaceicono")))
						config_verenlacefoto = bool((nodoCualid.getAttribute("verenlacefoto")))
						config_archivo = bool((nodoCualid.getAttribute("archivo")))
						config_descripcion = ""&nodoCualid.getAttribute("descripcion")
						config_ordenidioma = bool((nodoCualid.getAttribute("ordenidioma")))
				
						if ""&nodoCualid.getAttribute("nuevos") <> "" then
							config_nuevos = bool((nodoCualid.getAttribute("nuevos")))
						else
							config_nuevos = true
						end if
						if ""&nodoCualid.getAttribute("editar") <> "" then
							config_editar = bool((nodoCualid.getAttribute("editar")))
						else
							config_editar = true
						end if
						if ""&nodoCualid.getAttribute("eliminar") <> "" then
							config_eliminar = bool((nodoCualid.getAttribute("eliminar")))
						else
							config_eliminar = true
						end if
						config_sqlorden = "" & nodoCualid.getAttribute("sqlorden")
						
						config_nom_titulo = ""&nodoCualid.getAttribute("nombretitulo")
						config_nom_fuente = ""&nodoCualid.getAttribute("nombrefuente")
						config_nom_enlace = ""&nodoCualid.getAttribute("nombreenlace")
						config_nom_fecha = ""&nodoCualid.getAttribute("nombrefecha")
						config_nom_fechaini = ""&nodoCualid.getAttribute("nombrefechaini")
						config_nom_fechafin = ""&nodoCualid.getAttribute("nombrefechafin")
						config_nom_foto = ""&nodoCualid.getAttribute("nombrefoto")
						config_nom_icono = ""&nodoCualid.getAttribute("nombreicono")
						config_nom_archivo = ""&nodoCualid.getAttribute("nombrearchivo")
						config_posicion_foto = bool((nodoCualid.getAttribute("posicion_foto")))
						config_posicion_icono = bool((nodoCualid.getAttribute("posicion_icono")))
						config_pie_foto = bool((nodoCualid.getAttribute("pie_foto")))
						config_pie_icono = bool((nodoCualid.getAttribute("pie_icono")))
						if ""&nodoCualid.getAttribute("mover") <> "" then
							config_mover = bool((nodoCualid.getAttribute("mover")))
						else
							config_mover = true
						end if
						
						
						' Atributos de secciones
						set nodoSecciones = nodoCualid.selectSingleNode("secciones")
						if typeOK(nodoSecciones) then
							config_activo_seccion = bool((nodoSecciones.getAttribute("activo")))
							config_foto_seccion = bool((nodoSecciones.getAttribute("foto")))
							config_icono_seccion = bool((nodoSecciones.getAttribute("icono")))
							if ""&nodoSecciones.getAttribute("mover") <> "" then
								config_mover_seccion = bool((nodoSecciones.getAttribute("mover")))
							else
								config_mover_seccion = true
							end if
							config_nom_secciones = ""&nodoSecciones.getAttribute("nombre")
							if config_nom_secciones = "" then
								config_nom_secciones = "Secciones"
							end if
						else
							config_activo_seccion = false
							config_foto_seccion = false
							config_icono_seccion = false
							config_mover_seccion = false
						end if
						
						' Atributos de secciones2
						set nodoSecciones2 = nodoCualid.selectSingleNode("secciones2")
						if typeOK(nodoSecciones2) then
							config_activo_seccion2 = bool((nodoSecciones2.getAttribute("activo")))
							config_foto_seccion2 = bool((nodoSecciones2.getAttribute("foto")))
							config_icono_seccion2 = bool((nodoSecciones2.getAttribute("icono")))
							if ""&nodoSecciones2.getAttribute("mover") <> "" then
								config_mover_seccion2 = bool((nodoSecciones2.getAttribute("mover")))
							else
								config_mover_seccion2 = true
							end if
							config_nom_secciones2 = ""&nodoSecciones2.getAttribute("nombre")
							if config_nom_secciones2 = "" then
								config_nom_secciones2 = "Sub secciones"
							end if
						else
							config_activo_seccion2 = false
							config_foto_seccion2 = false
							config_icono_seccion2 = false
							config_mover_seccion2 = false
						end if
				
						' Foro
						set nodoForo = nodoCualid.selectSingleNode("foro")
						if typeOK(nodoForo) then
							config_foro = bool((nodoForo.getAttribute("activo")))
						else
							config_foro = false
						end if
						
						' VISORES
						' Definición de las variabels NAV_ (de navegación) para los visores
						if nav_visor <> "" then
							set nodoVisor = xml_config.selectSingleNode("configuracion/"&cualid&"/visor"&nav_visor)
						else
							set nodoVisor = xml_config.selectSingleNode("configuracion/"&cualid&"/visor")
						end if
						if typeName(nodoVisor) = "Nothing" or typeName(nodoVisor) = "Empty" then
							unerror = true : msgerror = "No se ha encontrado en nodo de configuración para el visor."
						end if
					end if
					
					if not unerror then
				
						nav_activo = config_activo
						nav_cuerpo = ""&nodoVisor.getAttribute("cuerpo")
						nav_orden = ""&nodoVisor.getAttribute("orden")
						nav_portada = numero(nodoVisor.getAttribute("portada"))
						nav_buscar = bool((nodoVisor.getAttribute("buscar")))
						nav_subtitulo = bool((nodoVisor.getAttribute("subtitulo")))
						nav_fecha = bool((nodoVisor.getAttribute("fecha")))
						nav_fuente = bool((nodoVisor.getAttribute("fuente")))
						nav_ampliar = bool((nodoVisor.getAttribute("ampliar")))
						nav_activo_secciones = numero(nodoVisor.getAttribute("secciones"))
						nav_activo_secciones2 = numero(nodoVisor.getAttribute("secciones2"))
						nav_regporpag = numero(nodoVisor.getAttribute("regporpag"))
						nav_regporpag_opt = numero(nodoVisor.getAttribute("regporpag_opt"))
						nav_foto = bool((nodoVisor.getAttribute("foto")))
						nav_icono = bool((nodoVisor.getAttribute("icono")))
						nav_verenlace = bool((nodoVisor.getAttribute("verenlace")))
						nav_archivo = bool((nodoVisor.getAttribute("archivo")))
						nav_tienda = bool((nodoVisor.getAttribute("tienda")))
						nav_nom_titulo = ""&nodoVisor.getAttribute("nombretitulo")
						nav_nom_fuente = ""&nodoVisor.getAttribute("nombrefuente")
						nav_nom_enlace = ""&nodoVisor.getAttribute("nombreenlace")
						nav_nom_fecha = ""&nodoVisor.getAttribute("nombrefecha")
						nav_nom_foto = ""&nodoVisor.getAttribute("nombrefoto")
						nav_nom_icono = ""&nodoVisor.getAttribute("nombreicono")
						nav_nom_archivo = ""&nodoVisor.getAttribute("nombrearchivo")
					end if
				
				end sub

%>