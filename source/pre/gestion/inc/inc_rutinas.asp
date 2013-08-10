<%
		' Rutinas
		Function mConsulta(sql,conexion,tipoBloqueo)
	
			dim re
			set re = Server.CreateObject("ADODB.Recordset")
			re.ActiveConnection = conexion
			re.Source = sql
			re.CursorType = 1
			re.CursorLocation = 2
			re.LockType = tipoBloqueo
'			on error resume next
			re.open
			if err<>0 then
				set mConsulta = nothing
			else
				set mConsulta = re
			end if
	
		end function
	
		Sub mCierra(re)
			on error resume next
				re.Close()
				set re = Nothing
			on error goto 0
		end sub

		' Convierte lo que le pasemos a número, en caso de nos ser un número válido devuelve 0
		function mNumero(n)
			if ""&n = "" then
				mNumero = 0
			else
				n = replace(n,".",",")
				if isNumeric(n) then
					mNumero = 0+n
				else
					mNumero = 0
				end if
			end if
		end function

		' Convierte lo que le pasemos a número, en caso de nos ser un número válido devuelve 0
		function mMonedaBD(n)
			mMonedaBD = mNumero(replace(n,",","."))
		end function

		' Cuenta el numero total de veces que encuentra la palabra que le indicamos en una cadena que le indicamos
		function mCPalabras(cadena, palabra)
			dim pos, conta
			pos = 1
			conta = 0
			While inStr(pos, cadena, palabra) > 0
				conta = conta + 1
				pos = inStr(pos, cadena, palabra) + Len(palabra)
			Wend   
			mCPalabras = conta    
		End Function
		
		' Euros
		function mEuros(n)
			mEuros = FormatNumber(mNumero(n),2)
		end function
		
		' Web Objects
		function mWOFecha(pFecha, nombre, clase)
			dim fecha, str, n, mes_arr

			if pFecha <> "" then
				fecha = CDate(pFecha)
			end if
			if not isDate(fecha) then
				fecha = Date()
			end if

			dia = day(fecha)
			mes = month(fecha)
			ano = year(fecha)

			mes_arr = Array("","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")
			if clase = "" then clase = "campo" end if
			str = ""

			' Día ------------------------------------------------------------------------
			str = str &"<select name=dia_"& nombre &" class="& clase &">"& vbCrlf
			for n=1 to 31
				selected = ""
				if dia = n then
					selected = "selected"
				end if
				str = str & "<option value="& n &" "& selected &">"& right("0"& n,2) &"</option>"& vbCrlf
			next
			str = str & "</select>"& vbCrlf

			' Mes ------------------------------------------------------------------------
			str = str &"<select name=mes_"& nombre &" class="& clase &">"& vbCrlf
			for n=1 to 12
				selected = ""
				if mes = n then
					selected = "selected"
				end if
				str = str & "<option value="& n &" "& selected &">"& mes_arr(n) &"</option>"& vbCrlf
			next
			str = str & "</select>"& vbCrlf

			' Año ------------------------------------------------------------------------
			str = str &"<select name=ano_"& nombre &" class="& clase &">"& vbCrlf
			for n=ano-3 to ano+3
				selected = ""
				if ano = n then
					selected = "selected"
				end if
				str = str & "<option value="& n &" "& selected &">"& n &"</option>"& vbCrlf
			next
			str = str & "</select>"& vbCrlf

			mWOFecha = str
		end function

%>