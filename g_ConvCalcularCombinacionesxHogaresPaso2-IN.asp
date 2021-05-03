<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
'
' g_ConvCalcularCombinacionesxHogaresPaso2.asp - 10abr21 - 11abr21
'
Session.lcid = 1034
Response.CodePage = 65001	
Response.CharSet = "utf-8"
'	
Dim QrySql, arrResultados(10000,3), idMeses
Dim dataArray, rsx1, rsTotalArray
Dim totalRefrescoHogares, tRefrescoAgua, tRefrescoJugo, tRefrescoTe
Dim totalAguaHogares, tAguaRefresco, tAguaJugo, tAguaTe
Dim totalTeHogares, tTeRefresco, tTeAgua, tTeJugo
Dim totalJugoHogares, tJugoRefresco, tJugoAgua, tJugoTe
'
idMeses = "16,14,18,19" 'Request.QueryString("id_Mes")
'
Calcular_Refrescos
'	
'Calcular_Aguas
'
'Calcular_Jugo
'
'Calcular_Te
'
SUB Calcular_Refrescos
	'	
	' Calcular Total hogares del Mes Compraron Refresco
	'
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType   = 2 'adLockOptimistic 
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_DataCrudaMensual"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
	QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"	
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 " '&  IN (1, 3 ,12 , 22)"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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
		Dim idTotalHogares
		idTotalHogares=vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			idTotalHogares = idTotalHogares + cstr(dataArray(0,iReg)) & ","			
		next
		idTotalHogares = Left(idTotalHogares, Len(idTotalHogares) - 1)
		'
		' Llenar la matriz Resultante con Ceros
		'	
		' For i = 1 to 10000								
			' arrResultados(i,0) = 0
			' arrResultados(i,1) = 0
			' arrResultados(i,2) = 0
			' arrResultados(i,3) = 0
		' Next
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
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 3" '&  IN (1, 3 ,12 , 22)"
			'QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar ="  & Hogar '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN (" & idTotalHogares & ")"					
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			' Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else
				totalAgua = totalAgua + 1
				rsTotalArray = rsx1.GetRows()
				rsx1.close
				totalAgua = UBound(rsTotalArray, 2) + 1 	
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
		' response.write "totalagua " & totalAgua & "<br>"
		' response.write "total hogares " & totalRefrescoHogares & "<br>"
		' response.write "porcentaje Agua " & tRefrescoAgua & "<br>"
		' Response.END
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
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 22" '&  IN (1, 3 ,12 , 22)"
			'QrySql = QrySql & " AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN (" & idTotalHogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else		
				'rsx1.close
				'totalJugo=totalJugo + 1
				rsTotalArray = rsx1.GetRows() 
				rsx1.close
				totalJugo = UBound(rsTotalArray, 2) + 1 
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
		' Calculo Refresco/Te
		'
		totalTe=0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			 
			' 
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 			
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			'QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (16, 17, 18, 19)"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 12" '&  IN (1, 3 ,12 , 22)"
			'QrySql = QrySql & " AND PH_DataCrudaMensual.Id_HOGAR="  & Hogar '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN (" & idTotalHogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			' Response.write QrySql & "<br>"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
			else						
				' rsx1.close
				' totalTe = totalTe + 1
				rsTotalArray = rsx1.GetRows() 
				rsx1.close
				totalTe = UBound(rsTotalArray, 2) + 1 
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'									
		NEXT	
		tRefrescoTe = 0
		if totalTe=0 then
			tRefrescoTe = 0
		else
			tRefrescoTe = totalTe * 100 / totalRefrescoHogares
		end if
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

END SUB	
'
SUB Calcular_Aguas
	'	
	' Calcular Total hogares del Mes Compraron Agua
	'
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType   = 2 'adLockOptimistic 
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_DataCrudaMensual"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
	QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"	
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 3 " '&  IN (1, 3 ,12 , 22)"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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
		For i = 1 to 10000								
			arrResultados(i,0) = 0
			arrResultados(i,1) = 0
			arrResultados(i,2) = 0
			arrResultados(i,3) = 0
		Next
		'
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
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar = "  & Hogar '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 22" '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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
		' Calculo Agua/Te
		'
		totalTe=0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			 
			' 
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 			
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			'QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (16, 17, 18, 19)"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 12" '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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

END SUB	
'
SUB Calcular_Jugo
	'	
	' Calcular Total hogares del Mes Compraron Jugo
	'
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType   = 2 'adLockOptimistic 
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_DataCrudaMensual"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
	QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"	
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 22 " '&  IN (1, 3 ,12 , 22)"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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
		' Llenar la matriz Resultante con Ceros
		'	
		For i = 1 to 10000								
			arrResultados(i,0) = 0
			arrResultados(i,1) = 0
			arrResultados(i,2) = 0
			arrResultados(i,3) = 0
		Next
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
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar = "  & Hogar '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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
		' Calculo Jugo / Agua
		'
		totalAgua = 0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			 
			' 
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 			
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			'QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (16, 17, 18, 19)"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 3" '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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
		' Calculo Jugo / Te
		'
		totalTe = 0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			 
			' 
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 			
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			'QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (16, 17, 18, 19)"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 12" '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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

END SUB	
'
SUB Calcular_Te
	'	
	' Calcular Total hogares del Mes Compraron Te
	'
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType   = 2 'adLockOptimistic 
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_DataCrudaMensual"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
	QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"	
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 12 " '&  IN (1, 3 ,12 , 22)"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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
		' Llenar la matriz Resultante con Ceros
		'	
		For i = 1 to 10000								
			arrResultados(i,0) = 0
			arrResultados(i,1) = 0
			arrResultados(i,2) = 0
			arrResultados(i,3) = 0
		Next
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
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar = "  & Hogar '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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
		' Calculo Te / Agua
		'
		totalAgua = 0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			 
			' 
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 			
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			'QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (16, 17, 18, 19)"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 3" '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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
		' Calculo Te / Jugo
		'
		totalJugo = 0
		FOR  i = 0 to ubound(dataArray,2) 
			'
			hogar  = dataArray(0,i)			 
			' 
			set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 			
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			'QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (16, 17, 18, 19)"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 22" '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_HOGAR = "  & Hogar '&  IN (1, 3 ,12 , 22)"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
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
%>
