<%


FUNCTION SortArray(varArray)
	For i = UBound(varArray) - 1 To 1 Step - 1
		MaxVal = varArray(i)
		MaxIndex = i
	
		For j = 0 To i
			If varArray(j) > MaxVal Then
				MaxVal = varArray(j)
				MaxIndex = j
			End If
		Next
	
		If MaxIndex < i Then
			varArray(MaxIndex) = varArray(i)
			varArray(i) = MaxVal
		End If
	Next 
END FUNCTION

		function sinNum(cadena)
			cadena = replace(cadena,"1","")
			cadena = replace(cadena,"2","")
			cadena = replace(cadena,"3","")
			cadena = replace(cadena,"4","")
			cadena = replace(cadena,"5","")
			cadena = replace(cadena,"6","")
			cadena = replace(cadena,"7","")
			cadena = replace(cadena,"8","")
			cadena = replace(cadena,"9","")
			cadena = replace(cadena,"0","")
			sinNum = cadena
		end function


Function borrarArchivo (archivo)
	Dim borrararchivo_fso
	set borrararchivo_fso = Server.CreateObject("Scripting.FileSystemObject")
	borrararchivo_fso.DeleteFile archivo, true
	set borrararchivo_fso = nothing
end function

' Cuenta Palabras: -----------------------------------------------------
function cuentaPalabras(cadena, palabra)
    pos = 1
    conta = 0
    While inStr(pos, cadena, palabra) > 0
        conta = conta + 1
        pos = inStr(pos, cadena, palabra) + Len(palabra)
    Wend   
    cuentaPalabras = conta    
End Function
'
' Funcion para validar los formatos de los email
function validarEmail(email) 
	if ""&email = "" then
       validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela." 
       exit function 
	end if

    dim partes, parte, i, c 
    'rompo el email en dos partes, antes y después de la arroba 
    partes = Split(email, "@") 
    if UBound(partes) <> 1 then 
       'si el mayor indice del array es distinto de 1 es que no he obtenido las dos partes 
       validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela." 
       exit function 
    end if 
    'para cada parte, compruebo varias cosas 
    for each parte in partes 
       'Compruebo que tiene algún caracter 
       if Len(parte) <= 0 then 
          validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela."  
          exit function 
       end if 
       'para cada caracter de la parte 
       for i = 1 to Len(parte) 
          'tomo el caracter actual 
          c = Lcase(Mid(parte, i, 1)) 
          'miro a ver si ese caracter es uno de los permitidos 
          if InStr("._-abcdefghijklmnopqrstuvwxyz", c) <= 0 and not IsNumeric(c) then 
             validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela."  
             exit function 
          end if 
       next 
       'si la parte actual acaba o empieza en punto la dirección no es válida 
       if Left(parte, 1) = "." or Right(parte, 1) = "." then 
          validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela."  
          exit function 
       end if 
    next 
    'si en la segunda parte del email no tenemos un punto es que va mal 
    if InStr(partes(1), ".") <= 0 then 
       validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela." 
       exit function 
    end if 
    'calculo cuantos caracteres hay después del último punto de la segunda parte del mail 
    i = Len(partes(1)) - InStrRev(partes(1), ".") 
    'si el número de caracteres es distinto de 2 y 3 
    if not (i = 2 or i = 3) then 
       validarEmail = "La dirección de e-mail no puede ser correcta. Por favor, revísela." 
       exit function 
    end if 
    'si encuentro dos puntos seguidos tampoco va bien 
    if InStr(email, "..") > 0 then 
       validarEmail="Una dirección e-mail no puede contener dos espacios seguidos. Por favor, revíselo." 
       exit function 
    end if 
    validarEmail = true 
end function 

function getEmails(texto)
	Dim largo, c, email, arrReplaces
	texto = lcase(texto)
	texto = replace(texto,vbCrlf," ")
	texto = replace(texto,chr(24),"")
	
	Dim abc, num, car
	abc = "abcdefghijklmnopqrstuvwxyz"
	num = "0123456789"
	sig = "._-@"
	car = abc & num & sig

	arrReplaces = array("?",":",";","estadologico@hotmail.com","info@sellourbano.com","info@estadologico.com","tradepunk@estadologico.com",".servidoresdns.")
	for each repla in arrReplaces
		texto = replace(texto,repla," ")
	next

	presuntos = cuentaPalabras(texto, "@")
	
	texto = replace(texto,".com",".com ")
	texto = replace(texto,".es",".es ")
	texto = replace(texto,".net",".net ")
	texto = replace(texto,".it",".it ")
	texto = replace(texto,".org",".org ")
	texto = replace(texto,".info",".info ")
	texto = replace(texto,".biz",".biz ")
	texto = replace(texto,".name",".name ")
	texto = replace(texto,".eu",".eu")
	texto = replace(texto,".com.ar",".com.ar ")
	texto = replace(texto,".tk",".tk ")

	texto = replace(texto,"  "," ")
	texto = replace(texto,"  "," ")
	texto = replace(texto,"  "," ")
	texto = replace(texto,"  "," ")

	encontrados = 0

	largo = len(texto)
	for n=1 to largo
		c = mid(texto,n,1)
		if inStr(car,c) > 0 then
			email = email & c
		else
			if validarEmail(email) = true then
				if inStr(salida,email) <=0 then
					salida = salida & email & "|"
					encontrados = encontrados + 1
				end if
			end if
			email = ""
		end if
	next
	' 
	if salida <> "" then
		salida = left(salida,len(salida)-1)
	end if
	
	getEmails = salida
end function
%>