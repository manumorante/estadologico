<%
				' *******************************
				' XMLShop v0.1 - Agrupalia
				' Noviembre 2004 - Manuel Morante
				' *******************************


			
			Class xmlSessionShop
				Private m_unerror
				Private m_msgerror
				Private xmlObj			' Objeto XML
				Private nodo_carrito	' Nodo carrito

				Sub Class_Initialize()
					if typeOK(session("xmlShop_carrito")) then
						set xmlObj = session("xmlShop_carrito")
						set nodo_carrito = xmlObj.selectSingleNode("carrito")
					else
						set xmlObj = CreateObject("MSXML.DOMDocument")
						set nodo_carrito = xmlObj.createElement("carrito")

						' identificador interno (único, temporal)
						set att_id = xmlObj.createAttribute("maxidx")
						nodo_carrito.setAttributeNode(att_id)
						att_id.nodeValue = 1
	
						xmlObj.appendChild(nodo_carrito)
						guardar
					end if

					m_msgerror = ""
					m_unerror = false
				End Sub

				Public Property Get total()
					total = nodo_carrito.getAttribute("total")
				end property

				Public Property Get nodo()
					set nodo = nodo_carrito
				end property

				Public Property Get unerror()
					unerror = m_unerror
				end property

				Public Property Get msgerror()
					msgerror = m_msgerror
				end property

				Sub Class_Terminate()
				End Sub

				' Añadir una nueva referencia al carrito
				Public Function addItem(idioma, cualidad, id, titulo, precio, cantidad)

					dim nuevo
					dim idx		' Identificador único

					set nuevo = xmlObj.createElement("item")
					nodo_carrito.appendChild(nuevo)

					set attTemp = xmlObj.createAttribute("idioma")		: nuevo.setAttributeNode(attTemp) : attTemp.nodeValue = ""& idioma
					set attTemp = xmlObj.createAttribute("cualidad")	: nuevo.setAttributeNode(attTemp) : attTemp.nodeValue = ""& cualidad
					set attTemp = xmlObj.createAttribute("id")			: nuevo.setAttributeNode(attTemp) : attTemp.nodeValue = ""& id
					set attTemp = xmlObj.createAttribute("cantidad")	: nuevo.setAttributeNode(attTemp) : attTemp.nodeValue = ""&cantidad
					set attTemp = xmlObj.createAttribute("precio")		: nuevo.setAttributeNode(attTemp) : attTemp.nodeValue = ""& precio
					set attTemp = xmlObj.createAttribute("titulo")		: nuevo.setAttributeNode(attTemp) : attTemp.nodeValue = ""& titulo

					' id única
					idx = numero(nodo_carrito.getAttribute("maxidx"))
					set attTemp = xmlObj.createAttribute("idx")
					nuevo.setAttributeNode(attTemp)
					attTemp.nodeValue = idx

					' Incremento contador id única
					set attTemp = xmlObj.createAttribute("maxidx")
					nodo_carrito.setAttributeNode(attTemp)
					attTemp.nodeValue = idx+1

					calcular_total
					guardar
					
					set nuevo = Nothing
					set attTemp = Nothing
					
				end function

				' Eliminar una referencia al carrito
				Public Function delItem(idx)
					dim nodo_borrar
					set nodo_borrar = nodo_carrito.selectNodes("item[@idx='"& idx &"']").item(0)
					if not typeOK(nodo_borrar) then
						m_unerror = true : m_msgerror = "No se ha encontado el nodo con la id especificada."
					else
						nodo_carrito.removeChild(nodo_borrar)
					end if

					calcular_total
					guardar
					
					set nodo_borrar = Nothing
				end function

				' Localiza y devuelve el attributo con el nombre indicado en la id indicada
				Public Function getAtt(idioma, cualidad, id, nomAtt)
					Dim temp
					set temp = nodo_carrito.selectNodes("item[@idioma='"& idioma &"' and @cualidad='"& cualidad &"' and @id='"& id &"']").item(0)
					if typeOK(temp) then
						getAtt = ""&temp.getAttribute(""&nomAtt)
					else
						getAtt = ""
					end if
					set temp = nothing
				end function

				' Añade un atributo a un item
				Public Function addAtt(idx, nomAtt, valorAtt)
					Dim temp, attTemp
					if ""& nomAtt <> "" then
						set temp = nodo_carrito.selectNodes("item[@idx='"& idx &"']").item(0)
						if typeOK(temp) then
							set attTemp = xmlObj.createAttribute(nomAtt)
							temp.setAttributeNode(attTemp)
							attTemp.nodeValue = ""& valorAtt
						else
							addAtt = ""
						end if

						calcular_total
						guardar

						set temp = nothing
						set attTemp = nothing
					end if
				end function
				
				Private sub calcular_total
					Dim item
					Dim total

					total = 0
					for each item in nodo_carrito.childNodes
						total = total + (numero(item.getAttribute("precio")) * numero(item.getAttribute("cantidad")))
					next
					set attTotal = xmlObj.createAttribute("total")
					nodo_carrito.setAttributeNode(attTotal)
					attTotal.nodeValue = total
					set attTotal = Nothing
				end sub

				Private sub guardar
					set session("xmlShop_carrito") = xmlObj
					xmlObj.save Server.MapPath("prueba.xml")
				end sub
				
			End Class
			


%>