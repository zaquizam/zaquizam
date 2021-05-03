<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
'
' g_ConvCalcularCombinacionesxHogaresPaso2.asp - 10abr21 - 12abr21
'
Session.lcid = 1034
Response.CodePage = 65001	
Response.CharSet = "utf-8"
Server.ScriptTimeout = 360
Response.Buffer = True
'Response.Expire = 0

'	

Dim QrySql, arrResultados(10000,3), idMeses
Dim dataArray, rsx1
Dim totalRefrescoHogares, tRefrescoAgua, tRefrescoJugo, tRefrescoTe
Dim totalAguaHogares, tAguaRefresco, tAguaJugo, tAguaTe
Dim totalTeHogares, tTeRefresco, tTeAgua, tTeJugo
Dim totalJugoHogares, tJugoRefresco, tJugoAgua, tJugoTe
'
'idMeses = Request.QueryString("id_Mes")
'
Response.write(Server.ScriptTimeout)
'Response.END
''
idMeses="16,17,18,19,20,21,22,23,24,25,26,27,28" ' Trimestral
'
StartTime = Timer

Calcular_Refrescos
'	
Calcular_Aguas
'
Calcular_Jugo
'
Calcular_Te
'
ElapsedTime = Timer - StartTime
Response.Write "<br><br>Proceso tardo: " & Cstr(ElapsedTime) & " Segundos."
'
SUB Calcular_Refrescos
	'	
	' Calcular Total hogares del Mes Compraron Refresco
	'
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenStatic 
	rsx1.LockType   = 3 'adLockOptimistic 
	'
	' QrySql = vbnullstring
	' " SELECT"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	' QrySql = QrySql & " FROM"
	' QrySql = QrySql & " PH_DataCrudaMensual"
	' QrySql = QrySql & " WHERE"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
	' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"	
	' QrySql = QrySql & " GROUP BY"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	' QrySql = QrySql & " HAVING"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 " '&  IN (1, 3 ,12 , 22)"
	' QrySql = QrySql & " ORDER BY"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	'
	QrySql = vbnullstring
	QrySql = " SELECT" & _
	" PH_DataCrudaMensual.Id_Hogar," & _
	" PH_DataCrudaMensual.Id_Categoria" & _
	" FROM" & _
	" PH_DataCrudaMensual" & _
	" WHERE" & _
	" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
	" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"	 & _
	" GROUP BY" & _
	" PH_DataCrudaMensual.Id_Hogar," & _
	" PH_DataCrudaMensual.Id_Categoria" & _
	" HAVING" & _
	" PH_DataCrudaMensual.Id_Categoria = 1 " & _
	" ORDER BY" & _
	" PH_DataCrudaMensual.Id_Hogar;"
	'
	' Response.write QrySql & "<br>"
	' Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then
		rsx1.close
	else		
		dataArray = rsx1.GetRows
		rsx1.close
		'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
	end if
	'	
	Set rsx1 = nothing
	'
	' Response.write "Recordset Ok<br>"
	' Response.end
	'	
	IF IsArray(dataArray) THEN
		'
		'Response.write "Total: " & ubound(dataArray,2) + 1 
		totalRefrescoHogares = 0
		totalRefrescoHogares = ubound(dataArray,2) + 1 
		'Response.write "<br>"
		'
		' for iReg = 0 to ubound(dataArray,2)
			' Response.write "<br>" &  dataArray(0,iReg)  & "=>" & dataArray(1,iReg)
			' Response.Write "Hogar : " & dataArray(0,i) & " Categoria : " & dataArray(1, i) & "<BR>"
		' next		
		'
		' Calculo Refresco/Agua
		'
		totalAgua=0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)
			' 
			' Response.Write "Hogar : " & hogar & " Categoria : " & Categ & "<BR>"
			' Response.Write "Tipo  : " & TypeName(hogar) & " Categoria : " & TypeName(Categ) & "<BR>"
			'
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenStatic 
			rsx1.LockType = 3 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = " SELECT" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" FROM" & _
			" PH_DataCrudaMensual" & _
			" WHERE" & _
			" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
			" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
			" GROUP BY" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" HAVING" & _
			" PH_DataCrudaMensual.Id_Categoria = 3" & _ 
			" AND PH_DataCrudaMensual.Id_Hogar = "  & Hogar & _ 
			" ORDER BY" & _
			" PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			' Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else				
				rsx1.close
				totalAgua = totalAgua + 1
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'								
		NEXT
		'		
		tRefrescoAgua = 0
		if totalAgua = 0 then
			tRefrescoAgua = 0
		else
			tRefrescoAgua = totalAgua * 100 / totalRefrescoHogares
		end if
		'
		' Response.write "Refresco" & "<br>"
		' response.write "totalagua " & totalAgua & "<br>"
		' response.write "total hogares REF " & totalRefrescoHogares & "<br>"
		' response.write "porcentaje Agua " & FORMATNUMBER(tRefrescoAgua,2) & "<br>"
		' Response.write "<br>"
		'
		' Calculo Refresco/jugo
		'
		totalJugo=0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			
			' 
			' Response.Write "Hogar : " & hogar & " Categoria : " & Categ & "<BR>"
			' Response.Write "Tipo  : " & TypeName(hogar) & " Categoria : " & TypeName(Categ) & "<BR>"
			'
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenStatic 
			rsx1.LockType = 3 'adLockOptimistic 
			' QrySql = vbnullstring
			' QrySql = QrySql & " SELECT"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " FROM"
			' QrySql = QrySql & " PH_DataCrudaMensual"
			' QrySql = QrySql & " WHERE"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			' QrySql = QrySql & " GROUP BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " HAVING"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 22" '&  IN (1, 3 ,12 , 22)"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar '&  IN (1, 3 ,12 , 22)"
			' QrySql = QrySql & " ORDER BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			
			QrySql = vbnullstring
			QrySql = " SELECT" & _
			" PH_DataCrudaMensual.Id_Hogar,"  & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" FROM" & _
			" PH_DataCrudaMensual" & _
			" WHERE" & _
			" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
			" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
			" GROUP BY" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" HAVING" & _
			" PH_DataCrudaMensual.Id_Categoria = 22" & _
			" AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar & _
			" ORDER BY" & _
			" PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else		
				rsx1.close
				totalJugo=totalJugo + 1
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'											
		NEXT	
		tRefrescoJugo = 0
		if totalJugo=0 then
			tRefrescoJugo = 0
		else
			tRefrescoJugo = totalJugo * 100 / totalRefrescoHogares
		end if
		'
		' Response.write "Refresco" & "<br>"
		' response.write "totaljugo " & totaljugo & "<br>"
		' response.write "total hogares REF " & totalRefrescoHogares & "<br>"
		' response.write "porcentaje jugo " & FORMATNUMBER(tRefrescojugo,2) & "<br>"
		' Response.write "<br>"
		'
		' Calculo Refresco/Te
		'
		totalTe=0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			 
			' 
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenStatic 
			rsx1.LockType = 3 'adLockOptimistic 			
			''			
			QrySql = vbnullstring
			QrySql = " SELECT" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" FROM" & _
			" PH_DataCrudaMensual" & _
			" WHERE" & _
			" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
			" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
			" GROUP BY" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" HAVING" & _
			" PH_DataCrudaMensual.Id_Categoria = 12" & _
			" AND PH_DataCrudaMensual.Id_HOGAR="  & Hogar & _
			" ORDER BY" & _
			" PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else						
				rsx1.close
				totalTe = totalTe + 1
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'									
		NEXT	
		'
		tRefrescoTe = 0
		if totalTe=0 then
			tRefrescoTe = 0
		else
			tRefrescoTe = totalTe * 100 / totalRefrescoHogares
		end if
		'
		' Response.write "Refresco" & "<br>"
		' response.write "totalte " & totalte & "<br>"
		' response.write "total hogares REF " & totalRefrescoHogares & "<br>"
		' response.write "porcentaje te " & FORMATNUMBER(tRefrescoTe,2) & "<br>"
		' Response.write "<br>"
		'		
	ELSE		
		'
		' Graficar los resultados en tablas NO HAY DATOS
		'
		Response.Write "<table class='table table-borderless table-hover' style=' margin: auto; width: 50% !important;'>"
			Response.Write "<thead>"
				Response.Write "<tr>"
					Response.Write "<th colspan='5' class='text-center text-danger'><i class='fas fa-check-double'></i>&nbsp;MATRIZ DE CONVIVENCIA CATEGORIAS&nbsp;</th>"	  
				Response.Write "</tr>"				
				Response.Write "<tr>"
					Response.Write "<td class='text-center'></td><td class='text-center'>Refresco</td><td class='text-center'>Agua</td><td class='text-center'>Jugo</td><td class='text-center'>T&eacute;</td>"
				Response.Write "</tr>"	
		   Response.Write "</thead>"
		   Response.Write "<tbody>"
			  Response.Write "<tr>"
				Response.Write "<th colspan='5' class='text-center text-primary'><strong>....NO HAY DATOS PARA EL MES SELECCIONADO....</strong></th>"
			  Response.Write "</tr>"
		   Response.Write "</tbody>"
		Response.Write "</table>"
		'		
	END IF

	Response.Flush
	Response.Clear	

END SUB	
'
SUB Calcular_Aguas
	'	
	' Calcular Total hogares del Mes Compraron Agua
	'
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenStatic 
	rsx1.LockType = 3 'adLockOptimistic 
	'		
	QrySql = vbnullstring
	QrySql = " SELECT" & _
	" PH_DataCrudaMensual.Id_Hogar," & _
	" PH_DataCrudaMensual.Id_Categoria" & _
	" FROM" & _
	" PH_DataCrudaMensual" & _
	" WHERE" & _
	" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
	" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
	" GROUP BY" & _
	" PH_DataCrudaMensual.Id_Hogar," & _
	" PH_DataCrudaMensual.Id_Categoria" & _
	" HAVING" & _
	" PH_DataCrudaMensual.Id_Categoria = 3 " & _
	" ORDER BY" & _
	" PH_DataCrudaMensual.Id_Hogar;"
	'
	' Response.write QrySql & "<br>"
	' Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then
		rsx1.close
	else		
		dataArray = rsx1.GetRows
		rsx1.close
		'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
	end if
	'	
	Set rsx1 = nothing
	'
	' Response.write "Recordset Ok<br>"
	' Response.end
	'	
	IF IsArray(dataArray) THEN
		'
		'Response.write "Total: " & ubound(dataArray,2) + 1 
		totalAguaHogares = 0
		totalAguaHogares = ubound(dataArray,2) + 1 
		'Response.write "<br>"
		'
		' for iReg = 0 to ubound(dataArray,2)
			' Response.write "<br>" &  dataArray(0,iReg)  & "=>" & dataArray(1,iReg)
			' Response.Write "Hogar : " & dataArray(0,i) & " Categoria : " & dataArray(1, i) & "<BR>"
		' next
		'
		' Llenar la matriz Resultante con Ceros
		'	
		' For i = 1 to 10000								
			' arrResultados(i,0) = 0
			' arrResultados(i,1) = 0
			' arrResultados(i,2) = 0
			' arrResultados(i,3) = 0
		' Next
		' '
		' Calculo Agua/Refresco
		'
		totalRefresco=0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)
			' 
			' Response.Write "Hogar : " & hogar & " Categoria : " & Categ & "<BR>"
			' Response.Write "Tipo  : " & TypeName(hogar) & " Categoria : " & TypeName(Categ) & "<BR>"
			'
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenStatic 
			rsx1.LockType = 3 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = " SELECT" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" FROM" & _
			" PH_DataCrudaMensual" & _
			" WHERE" & _
			" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
			" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
			" GROUP BY" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" HAVING" & _
			" PH_DataCrudaMensual.Id_Categoria = 1" & _
			" AND PH_DataCrudaMensual.Id_Hogar = "  & Hogar  & _
			" ORDER BY" & _
			" PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			' Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else				
				rsx1.close
				totalRefresco = totalRefresco + 1
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'								
		NEXT
		'		
		tAguaRefresco = 0
		if totalRefresco = 0 then
			tAguaRefresco = 0
		else
			tAguaRefresco = totalRefresco * 100 / totalAguaHogares
		end if
		'
		' Response.write "Agua" & "<br>"
		' response.write "totalRefresco " & totalRefresco & "<br>"
		' response.write "total hogares Agua " & totalAguaHogares & "<br>"
		' response.write "porcentaje Ref " & FORMATNUMBER(tAguaRefresco,2) & "<br>"
		' Response.write "<br>"
		'
		' Calculo Agua/jugo
		'
		totalJugo=0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			
			' 
			' Response.Write "Hogar : " & hogar & " Categoria : " & Categ & "<BR>"
			' Response.Write "Tipo  : " & TypeName(hogar) & " Categoria : " & TypeName(Categ) & "<BR>"
			'
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenStatic 
			rsx1.LockType = 3 'adLockOptimistic 
			QrySql = vbnullstring
			QrySql = " SELECT" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" FROM" & _
			" PH_DataCrudaMensual" & _
			" WHERE" & _
			" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
			" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
			" GROUP BY" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" HAVING" & _
			" PH_DataCrudaMensual.Id_Categoria = 22" & _
			" AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar  & _
			" ORDER BY" & _
			" PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else		
				rsx1.close
				totalJugo=totalJugo + 1
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'											
		NEXT	
		tAguaJugo = 0
		if totalJugo = 0 then
			tAguaJugo = 0
		else
			tAguaJugo = totalJugo * 100 / totalAguaHogares
		end if
		'
		' Response.write "Agua" & "<br>"
		' response.write "totalJugo " & totalJugo & "<br>"
		' response.write "total hogares Agua " & totalAguaHogares & "<br>"
		' response.write "porcentaje Jugo " & FORMATNUMBER(tAguaJugo,2) & "<br>"
		' Response.write "<br>"
		'
		' Calculo Agua/Te
		'
		totalTe=0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			 
			' 
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenStatic 
			rsx1.LockType = 3 'adLockOptimistic 			
			QrySql = vbnullstring
			QrySql = " SELECT" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" FROM" & _
			" PH_DataCrudaMensual" & _
			" WHERE" & _
			" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
			" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
			" GROUP BY" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" HAVING" & _
			" PH_DataCrudaMensual.Id_Categoria = 12" & _
			" AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar  & _
			" ORDER BY" & _
			" PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else						
				rsx1.close
				totalTe = totalTe + 1
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'									
		NEXT	
		tAguaTe = 0
		if totalTe = 0 then
			tAguaTe = 0
		else
			tAguaTe = totalTe * 100 / totalAguaHogares
		end if
		'
		' Response.write "AGUA " & "<br>"
		' response.write "totalTe " & totalTe & "<br>"
		' response.write "total hogares Agua " & totalAguaHogares & "<br>"
		' response.write "porcentaje Te " & FORMATNUMBER(tAguaTe,2) & "<br>"
		' Response.write "<br>"
		'		
	ELSE		
		'
		' Graficar los resultados en tablas NO HAY DATOS
		'
		Response.Write "<table class='table table-borderless table-hover' style=' margin: auto; width: 50% !important;'>"
			Response.Write "<thead>"
				Response.Write "<tr>"
					Response.Write "<th colspan='5' class='text-center text-danger'><i class='fas fa-check-double'></i>&nbsp;MATRIZ DE CONVIVENCIA CATEGORIAS&nbsp;</th>"	  
				Response.Write "</tr>"				
				Response.Write "<tr>"
					Response.Write "<td class='text-center'></td><td class='text-center'>Refresco</td><td class='text-center'>Agua</td><td class='text-center'>Jugo</td><td class='text-center'>T&eacute;</td>"
				Response.Write "</tr>"	
		   Response.Write "</thead>"
		   Response.Write "<tbody>"
			  Response.Write "<tr>"
				Response.Write "<th colspan='5' class='text-center text-primary'><strong>....NO HAY DATOS PARA EL MES SELECCIONADO....</strong></th>"
			  Response.Write "</tr>"
		   Response.Write "</tbody>"
		Response.Write "</table>"
		'		
	END IF		
	
	Response.Flush
	Response.Clear

END SUB	
'
SUB Calcular_Jugo
	'	
	' Calcular Total hogares del Mes Compraron Jugo
	'
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenStatic 
	rsx1.LockType = 3 'adLockOptimistic 
	'
	QrySql = vbnullstring
	QrySql = " SELECT" & _
	" PH_DataCrudaMensual.Id_Hogar," & _
	" PH_DataCrudaMensual.Id_Categoria" & _
	" FROM" & _
	" PH_DataCrudaMensual" & _
	" WHERE" & _
	" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
	" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
	" GROUP BY" & _
	" PH_DataCrudaMensual.Id_Hogar," & _
	" PH_DataCrudaMensual.Id_Categoria" & _
	" HAVING" & _
	" PH_DataCrudaMensual.Id_Categoria = 22 " & _
	" ORDER BY" & _
	" PH_DataCrudaMensual.Id_Hogar;"
	'
	' Response.write QrySql & "<br>"
	' Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then
		rsx1.close
	else		
		dataArray = rsx1.GetRows
		rsx1.close
		'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
	end if
	'	
	Set rsx1 = nothing
	'
	' Response.write "Recordset Ok<br>"
	' Response.end
	'	
	IF IsArray(dataArray) THEN
		'
		'Response.write "Total: " & ubound(dataArray,2) + 1 
		totalJugoHogares = 0
		totalJugoHogares = ubound(dataArray,2) + 1 
		'Response.write "<br>"
		'
		' for iReg = 0 to ubound(dataArray,2)
			' Response.write "<br>" &  dataArray(0,iReg)  & "=>" & dataArray(1,iReg)
			' Response.Write "Hogar : " & dataArray(0,i) & " Categoria : " & dataArray(1, i) & "<BR>"
		' next
		'
		' Calculo Jugo / Refresco
		'
		totalRefresco=0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)
			' 
			' Response.Write "Hogar : " & hogar & " Categoria : " & Categ & "<BR>"
			' Response.Write "Tipo  : " & TypeName(hogar) & " Categoria : " & TypeName(Categ) & "<BR>"
			'
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenStatic 
			rsx1.LockType = 3 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = " SELECT" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" FROM" & _
			" PH_DataCrudaMensual" & _
			" WHERE" & _
			" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
			" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
			" GROUP BY" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" HAVING" & _
			" PH_DataCrudaMensual.Id_Categoria = 1" & _
			" AND PH_DataCrudaMensual.Id_Hogar = "  & Hogar & _
			" ORDER BY" & _
			" PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			' Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else				
				rsx1.close
				totalRefresco = totalRefresco + 1
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'								
		NEXT
		'		
		tJugoRefresco = 0
		if totalRefresco = 0 then
			tJugoRefresco = 0
		else
			tJugoRefresco = totalRefresco * 100 / totalJugoHogares
		end if
		'
		' Response.write "JUGO " & "<br>"
		' response.write "totalRefresco " & totalRefresco & "<br>"
		' response.write "total hogares Jugo " & totalJugoHogares & "<br>"
		' response.write "porcentaje Ref " & FORMATNUMBER(tJugoRefresco,2) & "<br>"
		' Response.write "<br>"
		'		
		'
		' Calculo Jugo / Agua
		'
		totalAgua = 0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			 
			' 
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenStatic 
			rsx1.LockType = 3 'adLockOptimistic 			
			QrySql = vbnullstring
			QrySql = " SELECT" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" FROM" & _
			" PH_DataCrudaMensual" & _
			" WHERE" & _
			" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
			" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
			" GROUP BY" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" HAVING" & _
			" PH_DataCrudaMensual.Id_Categoria = 3" & _
			" AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar & _
			" ORDER BY" & _
			" PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else						
				rsx1.close
				totalAgua = totalAgua + 1
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'									
		NEXT	
		tJugoAgua = 0
		if totalAgua = 0 then
			tJugoAgua = 0
		else
			tJugoAgua = totalAgua * 100 / totalJugoHogares
		end if
		'
		' Response.write "JUGO " & "<br>"
		' response.write "totalAgua " & totalAgua & "<br>"
		' response.write "total hogares Jugo " & totalJugoHogares & "<br>"
		' response.write "porcentaje AGUA " & FORMATNUMBER(tJugoAgua,2) & "<br>"
		' Response.write "<br>"
		'
		' Calculo Jugo / Te
		'
		totalTe = 0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			 
			' 
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenStatic 
			rsx1.LockType = 3 'adLockOptimistic 			
			QrySql = vbnullstring
			QrySql = " SELECT" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" FROM" & _
			" PH_DataCrudaMensual" & _
			" WHERE" & _
			" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
			" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
			" GROUP BY" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" HAVING" & _
			" PH_DataCrudaMensual.Id_Categoria = 12" & _
			" AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar & _
			" ORDER BY" & _
			" PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else						
				rsx1.close
				totalTe = totalTe + 1
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'									
		NEXT	
		tJugoTe = 0
		if totalTe = 0 then
			tJugoTe = 0
		else
			tJugoTe = totalTe * 100 / totalJugoHogares
		end if
		'
		' Response.write "JUGO " & "<br>"
		' response.write "totalTe " & totalTe & "<br>"
		' response.write "total hogares Jugo " & totalJugoHogares & "<br>"
		' response.write "porcentaje Te " & FORMATNUMBER(tJugoTe,2) & "<br>"
		' Response.write "<br>"
		'
	ELSE		
		'
		' Graficar los resultados en tablas NO HAY DATOS
		'
		Response.Write "<table class='table table-borderless table-hover' style=' margin: auto; width: 50% !important;'>"
			Response.Write "<thead>"
				Response.Write "<tr>"
					Response.Write "<th colspan='5' class='text-center text-danger'><i class='fas fa-check-double'></i>&nbsp;MATRIZ DE CONVIVENCIA CATEGORIAS&nbsp;</th>"	  
				Response.Write "</tr>"				
				Response.Write "<tr>"
					Response.Write "<td class='text-center'></td><td class='text-center'>Refresco</td><td class='text-center'>Agua</td><td class='text-center'>Jugo</td><td class='text-center'>T&eacute;</td>"
				Response.Write "</tr>"	
		   Response.Write "</thead>"
		   Response.Write "<tbody>"
			  Response.Write "<tr>"
				Response.Write "<th colspan='5' class='text-center text-primary'><strong>....NO HAY DATOS PARA EL MES SELECCIONADO....</strong></th>"
			  Response.Write "</tr>"
		   Response.Write "</tbody>"
		Response.Write "</table>"
		'		
	END IF		
	
	Response.Flush
	Response.Clear

END SUB	
'
SUB Calcular_Te
	'	
	' Calcular Total hogares del Mes Compraron Te
	'
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenStatic 
	rsx1.LockType = 3 'adLockOptimistic 
	'
	QrySql = vbnullstring
	QrySql = " SELECT" & _
	" PH_DataCrudaMensual.Id_Hogar," & _
	" PH_DataCrudaMensual.Id_Categoria" & _
	" FROM" & _
	" PH_DataCrudaMensual" & _
	" WHERE" & _
	" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
	" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
	" GROUP BY" & _
	" PH_DataCrudaMensual.Id_Hogar," & _
	" PH_DataCrudaMensual.Id_Categoria" & _
	" HAVING" & _
	" PH_DataCrudaMensual.Id_Categoria = 12 " & _
	" ORDER BY" & _
	" PH_DataCrudaMensual.Id_Hogar;"
	'
	' Response.write QrySql & "<br>"
	' Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then
		rsx1.close
	else		
		dataArray = rsx1.GetRows
		rsx1.close
		'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
	end if
	'	
	Set rsx1 = nothing
	'
	' Response.write "Recordset Ok<br>"
	' Response.end
	'	
	IF IsArray(dataArray) THEN
		'
		'Response.write "Total: " & ubound(dataArray,2) + 1 
		totalTeHogares = 0
		totalTeHogares = ubound(dataArray,2) + 1 
		'Response.write "<br>"
		'
		' for iReg = 0 to ubound(dataArray,2)
			' Response.write "<br>" &  dataArray(0,iReg)  & "=>" & dataArray(1,iReg)
			' Response.Write "Hogar : " & dataArray(0,i) & " Categoria : " & dataArray(1, i) & "<BR>"
		' next		
		'
		' Calculo Te / Refresco
		'
		totalRefresco=0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)
			' 
			' Response.Write "Hogar : " & hogar & " Categoria : " & Categ & "<BR>"
			' Response.Write "Tipo  : " & TypeName(hogar) & " Categoria : " & TypeName(Categ) & "<BR>"
			'
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenStatic 
			rsx1.LockType = 3 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = " SELECT" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" FROM" & _
			" PH_DataCrudaMensual" & _
			" WHERE" & _
			" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
			" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
			" GROUP BY" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" HAVING" & _
			" PH_DataCrudaMensual.Id_Categoria = 1" & _
			" AND PH_DataCrudaMensual.Id_Hogar = "  & Hogar & _
			" ORDER BY" & _
			" PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			' Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else				
				rsx1.close
				totalRefresco = totalRefresco + 1
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'								
		NEXT
		'		
		tTeRefresco = 0
		if totalRefresco = 0 then
			tTeRefresco = 0
		else
			tTeRefresco = totalRefresco * 100 / totalTeHogares
		end if
		'
		' Response.write "TE " & "<br>"
		' response.write "totalRefresco " & totalRefresco & "<br>"
		' response.write "total hogares Te " & totalTeHogares & "<br>"
		' response.write "porcentaje Ref " & FORMATNUMBER(tTeRefresco,2) & "<br>"
		' Response.write "<br>"
		'
		' Calculo Te / Agua
		'
		totalAgua = 0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			 
			' 
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenStatic 
			rsx1.LockType = 3 'adLockOptimistic 			
			QrySql = vbnullstring
			QrySql = " SELECT" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" FROM" & _
			" PH_DataCrudaMensual" & _
			" WHERE" & _
			" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
			" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
			" GROUP BY" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" HAVING" & _
			" PH_DataCrudaMensual.Id_Categoria = 3" & _
			" AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar & _
			" ORDER BY" & _
			" PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else						
				rsx1.close
				totalAgua = totalAgua + 1
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'									
		NEXT	
		tTeAgua = 0
		if totalAgua = 0 then
			tTeAgua = 0
		else
			tTeAgua = totalAgua * 100 / totalTeHogares
		end if
		'
		' Response.write "TE " & "<br>"
		' response.write "totalAgua " & totalAgua & "<br>"
		' response.write "total hogares Te " & totalTeHogares & "<br>"
		' response.write "porcentaje Ref " & FORMATNUMBER(tTeAgua,2) & "<br>"
		' Response.write "<br>"
		'
		'
		' Calculo Te / Jugo
		'
		totalJugo = 0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			 
			' 
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenStatic 
			rsx1.LockType = 3 'adLockOptimistic 			
			QrySql = vbnullstring
			QrySql = " SELECT" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" FROM" & _
			" PH_DataCrudaMensual" & _
			" WHERE" & _
			" PH_DataCrudaMensual.Id_Fabricante <> 0" & _
			" AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")" & _
			" GROUP BY" & _
			" PH_DataCrudaMensual.Id_Hogar," & _
			" PH_DataCrudaMensual.Id_Categoria" & _
			" HAVING" & _
			" PH_DataCrudaMensual.Id_Categoria = 22" & _
			" AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar  & _
			" ORDER BY" & _
			" PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else						
				rsx1.close
				totalJugo = totalJugo + 1
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'									
		NEXT	
		tTeJugo = 0
		if totalJugo = 0 then
			tTeJugo = 0
		else
			tTeJugo = totalJugo * 100 / totalTeHogares
		end if
		'
		' Response.write "TE " & "<br>"
		' response.write "totalJugo " & totalJugo & "<br>"
		' response.write "total hogares te " & totalTeHogares & "<br>"
		' response.write "porcentaje JUGO " & FORMATNUMBER(tTeJugo,2) & "<br>"
		' Response.write "<br>"
		'
	ELSE		
		'
		' Graficar los resultados en tablas NO HAY DATOS
		'
		Response.Write "<table class='table table-borderless table-hover' style=' margin: auto; width: 50% !important;'>"
			Response.Write "<thead>"
				Response.Write "<tr>"
					Response.Write "<th colspan='5' class='text-center text-danger'><i class='fas fa-check-double'></i>&nbsp;MATRIZ DE CONVIVENCIA CATEGORIAS&nbsp;</th>"	  
				Response.Write "</tr>"				
				Response.Write "<tr>"
					Response.Write "<td class='text-center'></td><td class='text-center'>Refresco</td><td class='text-center'>Agua</td><td class='text-center'>Jugo</td><td class='text-center'>T&eacute;</td>"
				Response.Write "</tr>"	
		   Response.Write "</thead>"
		   Response.Write "<tbody>"
			  Response.Write "<tr>"
				Response.Write "<th colspan='5' class='text-center text-primary'><strong>....NO HAY DATOS PARA EL MES SELECCIONADO....</strong></th>"
			  Response.Write "</tr>"
		   Response.Write "</tbody>"
		Response.Write "</table>"
		'		
	END IF		
	
	Response.Flush
	Response.Clear

END SUB	
'
' Graficar los resultados en tablas
'
Response.Write "<strong><table class='table table-borderless table-hover table-condensed' style=' margin: auto; width: 65% !important;'>"
	
	Response.Write "<thead>"
	
		Response.Write "<tr>"
			Response.Write "<th colspan='5' class='text-center text-danger'><i class='fas fa-check-double'></i>&nbsp;MATRIZ DE CONVIVENCIA CATEGORIAS&nbsp;</th>"	  
		Response.Write "</tr>"
		
		Response.Write "<tr>"
			Response.Write "<td class='text-center'></td><td class='text-center'>Refresco</td><td class='text-center'>Agua</td><td class='text-center'>Jugo</td><td class='text-center'>T&eacute;</td>"
		Response.Write "</tr>"				
		
   Response.Write "</thead>"
   
   Response.Write "<tbody>"
   
		'Refresco
		Response.Write "<tr>"
			Response.Write "<td>Refresco</td>"
			Response.Write "<td class='text-center text-danger'> 100%</td><td class='text-center text-primary'>" & FormatNumber(tAguaRefresco,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tJugoRefresco,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tTeRefresco,2) & " %</td>"
		Response.Write "</tr>"
		'Agua
		Response.Write "<tr>"
			Response.Write "<td>Agua</td>"
			Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefrescoAgua,2) & " %</td><td class='text-center text-danger'> 100 %</td><td class='text-center text-primary'>" & FormatNumber(tJugoAgua,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tTeAgua,2) & " %</td>"
		Response.Write "</tr>"
		'Jugo
		Response.Write "<tr>"
			Response.Write "<td>Jugo</td>"
			Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefrescoJugo,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tAguaJugo,2) & " %</td><td class='text-center text-danger'> 100 %</td><td class='text-center text-primary'>" & FormatNumber(tTeJugo,2) & " %</td>"
		Response.Write "</tr>"
		'Te
		Response.Write "<tr>"
			Response.Write "<td>T&eacute;</td>"
			Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefrescoTe,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tAguaTe,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tJugoTe,2) & " %</td><td class='text-center text-danger'> 100 %</td>"
		Response.Write "</tr>"
		
   Response.Write "</tbody>"
   
Response.Write "</table></strong>"
'	
Response.Flush
Response.Clear
'
%>
