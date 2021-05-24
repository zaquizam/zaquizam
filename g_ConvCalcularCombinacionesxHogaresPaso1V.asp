 <%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
'
' g_ConvCalcularCombinacionesxHogaresPaso1.asp - 23abr21 - 
'
Session.lcid = 1034
Response.CodePage = 65001	
Response.CharSet = "utf-8"
Response.Buffer = True
'	
Dim QrySql, idMeses, dataArray, rsx1
'
Dim totalRefrescoHogaresCon_250, tRefresco250_250, tRefresco250_320, tRefresco250_350, tRefresco250_355
Dim tRefresco250_500, tRefresco250_600, tRefresco250_1000, tRefresco250_1250, tRefresco250_1500, tRefresco250_2000, tRefresco250_2500
'
Dim totalRefrescoHogaresCon_320, tRefresco320_250, tRefresco320_320, tRefresco320_350, tRefresco320_355
Dim tRefresco320_500, tRefresco320_600, tRefresco320_1000, tRefresco320_1250, tRefresco320_1500, tRefresco320_2000, tRefresco320_2500
'
Dim totalRefrescoHogaresCon_350, tRefresco350_250, tRefresco350_320, tRefresco350_350, tRefresco350_355
Dim tRefresco350_500, tRefresco350_600, tRefresco350_1000, tRefresco350_1250, tRefresco350_1500, tRefresco350_2000, tRefresco350_2500
'
Dim totalRefrescoHogaresCon_355, tRefresco355_250, tRefresco355_320, tRefresco355_350, tRefresco355_355
Dim tRefresco355_500, tRefresco355_600, tRefresco355_1000, tRefresco355_1250, tRefresco355_1500, tRefresco355_2000, tRefresco355_2500
'
Dim totalRefrescoHogaresCon_500, tRefresco500_250, tRefresco500_320, tRefresco500_350, tRefresco500_355
Dim tRefresco500_500, tRefresco500_600, tRefresco500_1000, tRefresco500_1250, tRefresco500_1500, tRefresco500_2000, tRefresco500_2500
'
Dim totalRefrescoHogaresCon_600, tRefresco600_250, tRefresco600_320, tRefresco600_350, tRefresco600_355
Dim tRefresco600_500, tRefresco600_600, tRefresco600_1000, tRefresco600_1250, tRefresco600_1500, tRefresco600_2000, tRefresco600_2500
'
Dim totalRefrescoHogaresCon_1000, tRefresco1000_250, tRefresco1000_320, tRefresco1000_350, tRefresco1000_355
Dim tRefresco1000_500, tRefresco1000_600, tRefresco1000_1000, tRefresco1000_1250, tRefresco1000_1500, tRefresco1000_2000, tRefresco1000_2500
'
Dim totalRefrescoHogaresCon_1250, tRefresco1250_250, tRefresco1250_320, tRefresco1250_350, tRefresco1250_355
Dim tRefresco1250_500, tRefresco1250_600, tRefresco1250_1000, tRefresco1250_1250, tRefresco1250_1500, tRefresco1250_2000, tRefresco1250_2500
'
Dim totalRefrescoHogaresCon_1500, tRefresco1500_250, tRefresco1500_320, tRefresco1500_350, tRefresco1500_355
Dim tRefresco1500_500, tRefresco1500_600, tRefresco1500_1000, tRefresco1500_1250, tRefresco1500_1500, tRefresco1500_2000, tRefresco1500_2500
'
Dim totalRefrescoHogaresCon_2000, tRefresco2000_250, tRefresco2000_320, tRefresco2000_350, tRefresco2000_355
Dim tRefresco2000_500, tRefresco2000_600, tRefresco2000_1000, tRefresco2000_1250, tRefresco2000_1500, tRefresco2000_2000, tRefresco2000_2500
'
Dim totalRefrescoHogaresCon_2500, tRefresco2500_250, tRefresco2500_320, tRefresco2500_350, tRefresco2500_355
Dim tRefresco2500_500, tRefresco2500_600, tRefresco2500_1000, tRefresco2500_1250, tRefresco2500_1500, tRefresco2500_2000, tRefresco2500_2500
'
'Dim ArrayTamanos : ArrayTamanos = Array(250,320,350,355,1000,1250,11000,2000,2500)
'
'idMeses ="16,17,18,19" ' Mensual
idMeses="16,17,18,19,20,21,22,23,24,25,26,27,28" ' Trimestral
'
'idMeses = Request.QueryString("id_Mes")
'
StartTime = Timer
'
Calcular_Refrescos_250
'	
Calcular_Refrescos_320
'
Calcular_Refrescos_350
'
Calcular_Refrescos_355
'
Calcular_Refrescos_500
'
Calcular_Refrescos_600
'
Calcular_Refrescos_1000
'
Calcular_Refrescos_1250
'
Calcular_Refrescos_1500
'
Calcular_Refrescos_2000
'
Calcular_Refrescos_2500
'
ElapsedTime = Timer - StartTime
Response.Write "<br><br>Proceso tardo: " & Cstr(ElapsedTime) & " Segundos."
'
Graficar_Datos
'
SUB Calcular_Refrescos_250
	Response.write "<BR><br>CALCULAR 250 ML<BR><BR>"
	'	
	' Buscar Todos Los Hogares compraron Tamaño 250 ml
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
	QrySql = QrySql & " AND PH_DataCrudaMensual.TAMANO = 250 "
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 "
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	'	
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then
		rsx1.close
		'
		Response.Write "Tama&ntilde;o 250 No! Convive con Nadie <BR><BR>"
		tRefresco250_250  = 0
		tRefresco250_320  = tRefresco250_350  = tRefresco250_355  = 0
		tRefresco250_500  = tRefresco250_600  = tRefresco250_1000 = 0
		tRefresco250_1500 = tRefresco250_2000 = tRefresco250_2500 = 0	
		'	
	else		
		'Erase dataArray
		dataArray = rsx1.GetRows
		rsx1.close
		'Response.write "<br>Total =: " & ubound(dataArray,2)+1
		'******
		'
		' Calculo total hogares con Refrescos 250
		'		
		totalRefrescoHogaresCon_250 = 0
		totalRefrescoHogaresCon_250 = ubound(dataArray,2) + 1 
		total250=0
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next		
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "<br>Total Compras 250 = " & totalRefrescoHogaresCon_250 & "<br>"
		Response.write "<br>ID Hogares que compraron 250 = " & replace(Hogares,",","-") & "<br>"
		'
		Set rsx1 = CreateObject("ADODB.Recordset")
		rsx1.CursorType = adOpenKeyset 
		rsx1.LockType   = 2 'adLockOptimistic 
		'
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_DataCrudaMensual"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 250"
		QrySql = QrySql & " GROUP BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " HAVING"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
		'
		'Response.write QrySql & "<br>"
		'Response.end
		'
		rsx1.Open QrySql, conexion		
		'
		if rsx1.eof then
			rsx1.close
			total250 = 0
		else
			rsArray = rsx1.GetRows() 
        	total250 = UBound(rsArray, 2) + 1         	
			'total250 = rsx1.recordcount
			rsx1.close			
		end if
		'								
		if total250 > 0 then
			'			
			tRefresco250_250 = total250 * 100 / totalRefrescoHogaresCon_250
			'
			Response.Write "<br>Total compras de 250 = " & total250 & "<br>"
			Response.Write "Total Porcentaje 250/250 = " & tRefresco250_250 & "<br>"
			'Response.End			
			'
			' Calcular Tamaño 320
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 320"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total320 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total320 = UBound(rsArray, 2) + 1         	
				'total320 = rsx1.recordcount
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total320 = 0 then
				'
				tRefresco250_320 = 0
				Response.Write "250 no convive con 320 / "
				'	
			else
				'
				tRefresco250_320 = total320 * 100 / totalRefrescoHogaresCon_250
				'
				Response.Write "<br>Total compras de 320 = " & total320 & "<br>"
				Response.Write "Total Porcentaje 250/320 = " & tRefresco250_320 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 350
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 350"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total350 = 0
			else
				rsArray = rsx1.GetRows() 
        		total350 = UBound(rsArray, 2) + 1         	
				'total350 = rsx1.recordcount
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total350 = 0 then
				'
				tRefresco250_350 = 0				
				Response.Write "250 no convive con 350 / "
				'	
			else
				'
				tRefresco250_350 = total350 * 100 / totalRefrescoHogaresCon_250
				'
				Response.Write "<br>Total compras de 350 = " & total350 & "<br>"
				Response.Write "Total Porcentaje 250/350 = " & tRefresco250_350 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 355
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 355"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total355 = 0
			else
				rsArray = rsx1.GetRows() 
        		total355 = UBound(rsArray, 2) + 1         	
				'total355 = rsx1.recordcount
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total355 = 0 then
				'
				tRefresco250_355 = 0
				Response.Write "250 no convive con 355 / "
				'	
			else
				'
				tRefresco250_355 = total355 * 100 / totalRefrescoHogaresCon_250
				'
				'
				Response.Write "<br>Total compras de 355 = " & total350 & "<br>"
				Response.Write "Total Porcentaje 250/355 = " & tRefresco250_355 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total500 = 0
			else
				rsArray = rsx1.GetRows() 
        		total500 = UBound(rsArray, 2) + 1         	
				'total500 = rsx1.recordcount
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total500 = 0 then
				'
				tRefresco250_500 = 0
				Response.Write "250 no convive con 500 / "
				'	
			else
				'
				tRefresco250_500 = total500 * 100 / totalRefrescoHogaresCon_250
				'
				Response.Write "<br>Total compras de 500 = " & total500 & "<br>"
				Response.Write "Total Porcentaje 250/500 = " & tRefresco250_500 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 600
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 600"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total600 = 0
			else
				rsArray = rsx1.GetRows() 
        		total600 = UBound(rsArray, 2) + 1         				
				'total600 = rsx1.recordcount
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total600 = 0 then
				'
				tRefresco250_600 = 0
				Response.Write "250 no convive con 600 / "
				'	
			else
				'
				tRefresco250_600 = total600 * 100 / totalRefrescoHogaresCon_250
				'
				Response.Write "<br>Total compras de 600 = " & total600 & "<br>"
				Response.Write "Total Porcentaje 250/600 = " & tRefresco250_600 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 1000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1000 = 0
			else
				rsArray = rsx1.GetRows() 
        		total1000 = UBound(rsArray, 2) + 1         	
				'total1000 = rsx1.recordcount
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1000 = 0 then
				'
				tRefresco250_1000 = 0
				Response.Write "250 no convive con 1000 / "
				'	
			else
				'
				tRefresco250_1000 = total1000 * 100 / totalRefrescoHogaresCon_250
				'
				Response.Write "<br>Total compras de 1000 = " & total1000 & "<br>"
				Response.Write "Total Porcentaje 250/1000 = " & tRefresco250_1000 & "<br>"
				'
			end if			
			'
			' Calcular Tamaño 1250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1250 = 0
			else
				rsArray = rsx1.GetRows() 
        		total1250 = UBound(rsArray, 2) + 1         	
				'total1250 = rsx1.recordcount
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1250 = 0 then
				'
				tRefresco250_1250 = 0
				Response.Write "250 no convive con 1250 / "
				'	
			else
				'
				tRefresco250_1250 = total1250 * 100 / totalRefrescoHogaresCon_250
				'
				'
				Response.Write "<br>Total compras de 1250 = " & total1250 & "<br>"
				Response.Write "Total Porcentaje 250/1250 = " & tRefresco250_1250 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 1500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1500 = 0
			else
				rsArray = rsx1.GetRows() 
        		total1500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total1500 = 0 then
				'
				tRefresco250_1500 = 0
				Response.Write "250 no convive con 1500 / "
				'	
			else
				'
				tRefresco250_1500 = total1500 * 100 / totalRefrescoHogaresCon_250				
				'
				Response.Write "<br>Total compras de 1500 = " & total1500 & "<br>"
				Response.Write "Total Porcentaje 250/1500 = " & tRefresco250_1500 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 2000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2000 = 0
			else
				rsArray = rsx1.GetRows() 
        		total2000 = UBound(rsArray, 2) + 1         	
				'total2000 = rsx1.recordcount
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total2000 = 0 then
				'
				tRefresco250_2000 = 0
				Response.Write "250 no convive con 2000 / "
				'	
			else
				'
				tRefresco250_2000 = total2000 * 100 / totalRefrescoHogaresCon_250
				'
				Response.Write "<br>Total compras de 2000 = " & total2000 & "<br>"
				Response.Write "Total Porcentaje 250/2000 = " & tRefresco250_2000 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 2500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2500 = 0
			else
				rsArray = rsx1.GetRows() 
        		total2500 = UBound(rsArray, 2) + 1         					
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total2500 = 0 then
				'
				tRefresco250_2500 = 0
				Response.Write "250 no convive con 2500 / "
				'	
			else
				'
				tRefresco250_2500 = total2500 * 100 / totalRefrescoHogaresCon_250
				'
				Response.Write "<br>Total compras de 2500 = " & total2500 & "<br>"
				Response.Write "Total Porcentaje 250/2500 = " & tRefresco250_2500 & "<br>"
				'
			end if
						
		end if
		'			
	end if
	'	
	Set rsx1 = nothing
	'
END SUB	
'
SUB Calcular_Refrescos_320
	Response.write "<BR><br>CALCULAR 320 ML<BR><BR>"
	'	
	' Buscar Todos Los Hogares compraron Tamaño 320 ml
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
	QrySql = QrySql & " AND PH_DataCrudaMensual.TAMANO = 320 "
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 "
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	'	
	'Response.write QrySql & "<br>"
	'Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then
		rsx1.close
		Response.Write "Tama&ntilde;o 320 No! Convive con nadie <BR><BR>"
		tRefresco320_250  = 0
		tRefresco320_320  = tRefresco320_350  = tRefresco320_355  = 0
		tRefresco320_500  = tRefresco320_600  = tRefresco320_1000 = 0
		tRefresco320_1500 = tRefresco320_2000 = tRefresco320_2500 = 0	
		'
	else
		'Erase dataArray
		dataArray = rsx1.GetRows
		rsx1.close
		'
		' Calculo total hogares con Refrescos 320
		'		
		totalRefrescoHogaresCon_320 = 0
		totalRefrescoHogaresCon_320 = ubound(dataArray,2) + 1 
		total320=0
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "<br>Total Compras 320            = " & totalRefrescoHogaresCon_320 & "<br>"
		Response.write "<br>ID Hogares que compraron 320 = " & replace(Hogares,",","-") & "<br>"
		'		
		Set rsx1 = CreateObject("ADODB.Recordset")
		rsx1.CursorType = adOpenKeyset 
		rsx1.LockType   = 2 'adLockOptimistic 
		'
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_DataCrudaMensual"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 320"
		QrySql = QrySql & " GROUP BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " HAVING"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
		'
		'Response.write QrySql & "<br>"
		'Response.end
		'
		rsx1.Open QrySql, conexion		
		'
		if rsx1.eof then
			rsx1.close
			total320 = 0
		else				
			'total320 = rsx1.recordcount
			rsArray = rsx1.GetRows() 
        	total320 = UBound(rsArray, 2) + 1         	
			rsx1.close			
		end if
		'								
		IF total320 > 0 THEN
			'
			tRefresco320_320 = total320 * 100 / totalRefrescoHogaresCon_320
			'
			Response.Write "<br>Total compras de 320 = " & total320 & "<br>"
			Response.Write "Total Porcentaje 320/320 = " & tRefresco320_320 & "<br>"			
			'			
			'Response.End			
			'
			' Calcular Tamaño 320/250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total250 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total250 = UBound(rsArray, 2) + 1         	
				rsx1.close				
			end if
			'
			if total250 = 0 then
				'
				tRefresco320_250 = 0
				Response.Write "320 no convive con 250 / "
				'	
			else
				'
				tRefresco320_250 = total250 * 100 / totalRefrescoHogaresCon_320
				'
				Response.Write "<br>Total compras de 250  = " & total320 & "<br>"
				Response.Write "Total Porcentaje 320/250  = " & tRefresco320_250 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 320/350
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 350"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total350 = 0
			else				
				'total350 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total350 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total350 = 0 then
				'
				tRefresco320_350 = 0				
				Response.Write "320 no convive con 350 / "
				'	
			else
				'
				tRefresco320_350 = total350 * 100 / totalRefrescoHogaresCon_320
				'
				Response.Write "<br>Total compras de 320  = " & total350 & "<br>"
				Response.Write "Total Porcentaje 320/350  = " & tRefresco320_350 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 320/355
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 355"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & " 355<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total355 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total355 = UBound(rsArray, 2) + 1         	
				rsx1.close				
			end if
			'
			if total355 = 0 then
				'
				tRefresco320_355 = 0
				Response.Write "320 no convive con 355 / "
				'	
			else
				'
				tRefresco320_355 = total355 * 100 / totalRefrescoHogaresCon_320
				'
				Response.Write "<br>Total compras de 355  = " & total355 & "<br>"
				Response.Write "Total Porcentaje 320/355  = " & tRefresco320_355 & "<br>"												
				'
			end if
			'
			' Calcular Tamaño 500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total500 = 0
			else				

				rsArray = rsx1.GetRows() 
        		total500 = UBound(rsArray, 2) + 1         	
				rsx1.close

			end if
			'
			if total500 = 0 then
				'
				tRefresco320_500 = 0
				Response.Write "320 no convive con 500 / "
				'	
			else
				'
				tRefresco320_500 = total500 * 100 / totalRefrescoHogaresCon_320
				Response.Write "<br>Total compras de 500  = " & total500 & "<br>"
				Response.Write "Total Porcentaje 320/500  = " & tRefresco320_500 & "<br>"								
				'
			end if
			'
			' Calcular Tamaño 600
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 600"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total600 = 0
			else				

				rsArray = rsx1.GetRows() 
        		total600 = UBound(rsArray, 2) + 1         	
				rsx1.close

			end if
			'
			if total600 = 0 then
				'
				tRefresco320_600 = 0
				Response.Write " 320 no convive con 600 / "
				'	
			else
				'
				tRefresco320_600 = total600 * 100 / totalRefrescoHogaresCon_320
				Response.Write "<br>Total compras de 600  = " & total600 & "<br>"
				Response.Write "Total Porcentaje 320/600  = " & tRefresco320_600 & "<br>"								
				'
				'
			end if
			'
			' Calcular Tamaño 1000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1000 = 0
			else				
				'total1000 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1000 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1000 = 0 then
				'
				tRefresco320_1000 = 0
				Response.Write "320 no convive con 1000"
				'	
			else
				'
				tRefresco320_1000 = total1000 * 100 / totalRefrescoHogaresCon_320
				Response.Write "<br>Total compras de 1000  = " & total1000 & "<br>"
				Response.Write "Total Porcentaje 320/1000  = " & tRefresco320_1000 & "<br>"								
				'
				'
			end if			
			'
			' Calcular Tamaño 1250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1250 = 0
			else				
				'total1250 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1250 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1250 = 0 then
				'
				tRefresco320_1250 = 0
				Response.Write "320 no convive con 1250"
				'	
			else
				'
				tRefresco320_1250 = total1250 * 100 / totalRefrescoHogaresCon_320
				Response.Write "<br>Total compras de 1250  = " & total1250 & "<br>"
				Response.Write "Total Porcentaje 320/1250  = " & tRefresco320_1250 & "<br>"								
				'
				'
			end if
			'
			' Calcular Tamaño 1500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1500 = 0
			else				
				'total1500 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1500 = 0 then
				'
				tRefresco320_1500 = 0
				Response.Write "320 no convive con 1500"
				'	
			else
				'
				tRefresco320_1500 = total1500 * 100 / totalRefrescoHogaresCon_320
				Response.Write "<br>Total compras de 1500  = " & total1500 & "<br>"
				Response.Write "Total Porcentaje 320/1500  = " & tRefresco320_1500 & "<br>"								
				'
				'
			end if
			'
			' Calcular Tamaño 2000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2000 = 0
			else				
				'total2000 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total2000 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total2000 = 0 then
				'
				tRefresco320_2000 = 0
				Response.Write "320 no convive con 2000"
				'	
			else
				'
				tRefresco320_2000 = total2000 * 100 / totalRefrescoHogaresCon_320
				Response.Write "<br>Total compras de 2000  = " & total2000 & "<br>"
				Response.Write "Total Porcentaje 320/2000  = " & tRefresco320_2000 & "<br>"								
				'
				'
			end if
			'
			' Calcular Tamaño 2500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2500 = 0
			else				
				'total2500 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total2500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total2500 = 0 then
				'
				tRefresco320_2500 = 0
				Response.Write "320 no convive con 2500"
				'	
			else
				'
				tRefresco320_2500 = total2500 * 100 / totalRefrescoHogaresCon_320
				Response.Write "<br>Total compras de 2500  = " & total2500 & "<br>"
				Response.Write "Total Porcentaje 320/2500  = " & tRefresco320_2500 & "<br>"								
				'
				'
			end if
			
		END IF
		'***
	end if
	'	
	Set rsx1 = nothing
	'
END SUB	
'
SUB Calcular_Refrescos_350
	Response.write "<BR><br>CALCULAR 350 ML<BR><BR>"
	'	
	' Buscar Todos Los Hogares compraron Tamaño 350 ml
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
	QrySql = QrySql & " AND PH_DataCrudaMensual.TAMANO = 350 "
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 "
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	'	
	'Response.write QrySql & "<br>"
	'Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then
		'
		rsx1.close
		'
		Response.Write "Tama&ntilde;o 350 No! Convive con nadie <BR><BR>"
		tRefresco350_250  = 0
		tRefresco350_320  = tRefresco350_350  = tRefresco350_355  = 0
		tRefresco350_500  = tRefresco350_600  = tRefresco350_1000 = 0
		tRefresco350_1500 = tRefresco350_2000 = tRefresco350_2500 = 0	
		'	
	else	
		'
		dataArray = rsx1.GetRows
		rsx1.close
		'***
		'
		' Calculo total hogares con Refrescos 350
		'		
		totalRefrescoHogaresCon_350 = 0
		totalRefrescoHogaresCon_350 = ubound(dataArray,2) + 1 
		total350=0
		'
		'Response.write "Total hogares 350 = " & totalRefrescoHogaresCon_350 & "<br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "<br>Total Compras 350 = " & totalRefrescoHogaresCon_350 & "<br>"
		Response.write "<br>ID Hogares que compraron 350 = " & replace(Hogares,",","-") & "<br>"
		'Response.end
		'				
		Set rsx1 = CreateObject("ADODB.Recordset")
		rsx1.CursorType = adOpenKeyset 
		rsx1.LockType   = 2 'adLockOptimistic 
		'
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_DataCrudaMensual"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 350"
		QrySql = QrySql & " GROUP BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " HAVING"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
		'
		'Response.write QrySql & "<br>"
		'Response.end
		'
		rsx1.Open QrySql, conexion		
		'
		if rsx1.eof then
			rsx1.close
			total350 = 0
		else				
			rsArray = rsx1.GetRows() 
        	total350 = UBound(rsArray, 2) + 1         	
			'total350 = rsx1.recordcount			
			rsx1.close			
		end if
		'
		'Response.write " Total RecordCount 350 = " & total350 & "<br>"
		'Response.end
		'
		IF total350 > 0 THEN
			'
			tRefresco350_350 = total350 * 100 / totalRefrescoHogaresCon_350
			'
			Response.Write "Total Compraron      350 =" & total350 & "<br>"
			Response.Write "Total Porcentaje 350/350 =" & tRefresco350_350 & "<br>"
			'Response.End			
			'
			' Calcular Tamaño 250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total250 = 0
			else				
				'total250 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total250 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total250 = 0 then
				'
				tRefresco350_250 = 0
				Response.Write "350 no convive con 250 / "
				'	
			else
				'
				tRefresco350_250 = total250 * 100 / totalRefrescoHogaresCon_350
				Response.Write "<br>Total compras de 250  = " & total250 & "<br>"
				Response.Write "Total Porcentaje 350/250  = " & tRefresco350_250 & "<br>"	
				'
			end if
			'
			' Calcular Tamaño 350
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 320"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total350 = 0
			else				
				'total350 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total320 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total350 = 0 then
				'
				tRefresco350_320 = 0				
				Response.Write "350 no convive con 320 / "
				'	
			else
				'
				tRefresco350_320 = total320 * 100 / totalRefrescoHogaresCon_350
				Response.Write "<br>Total compras de 350  = " & total320 & "<br>"
				Response.Write "Total Porcentaje 350/320  = " & tRefresco350_320 & "<br>"	
				'
				'
			end if
			'
			' Calcular Tamaño 355
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 355"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total355 = 0
			else				
				'total355 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total355 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total355 = 0 then
				'
				tRefresco350_355 = 0
				Response.Write "350 no convive con 355 / "
				'	
			else
				'
				tRefresco350_355 = total355 * 100 / totalRefrescoHogaresCon_350
				Response.Write "<br>Total compras de 355  = " & total355 & "<br>"
				Response.Write "Total Porcentaje 350/355  = " & tRefresco350_355 & "<br>"	
				'
			end if
			'
			' Calcular Tamaño 500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total500 = 0
			else				
				'total500 = rsx1.recordcount
				rsArray  = rsx1.GetRows() 
        		total500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total500 = 0 then
				'
				tRefresco350_500 = 0
				Response.Write "350 no convive con 500 / "
				'	
			else
				'
				tRefresco350_500 = total500 * 100 / totalRefrescoHogaresCon_350
				Response.Write "<br>Total compras de 350  = " & total500 & "<br>"
				Response.Write "Total Porcentaje 350/500  = " & tRefresco350_500 & "<br>"	
				'
			end if
			'
			' Calcular Tamaño 600
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 600"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total600 = 0
			else				
				'total600 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total600 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total600 = 0 then
				'
				tRefresco350_600 = 0
				Response.Write "350 no convive con 600<br>"
				'	
			else
				'
				tRefresco350_600 = total600 * 100 / totalRefrescoHogaresCon_350
				Response.Write "<br>Total compras de 600  = " & total600 & "<br>"
				Response.Write "Total Porcentaje 350/600  = " & tRefresco350_600 & "<br>"	
				'
			end if
			'
			' Calcular Tamaño 1000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1000 = 0
			else				
				'total1000 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1000 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1000 = 0 then
				'
				tRefresco350_1000 = 0
				Response.Write "350 no convive con 1000 / "
				'	
			else
				'
				tRefresco350_1000 = total1000 * 100 / totalRefrescoHogaresCon_350
				Response.Write "<br>Total compras de 1000  = " & total1000 & "<br>"
				Response.Write "Total Porcentaje 350/1000  = " & tRefresco350_1000 & "<br>"	
				'
			end if			
			'
			' Calcular Tamaño 1250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1250 = 0
			else				
				'total1250 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1250 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1250 = 0 then
				'
				tRefresco350_1250 = 0
				Response.Write "350 no convive con 1250 / "
				'	
			else
				'
				tRefresco350_1250 = total1250 * 100 / totalRefrescoHogaresCon_350
				Response.Write "<br>Total compras de 1250  = " & total1250 & "<br>"
				Response.Write "Total Porcentaje 350/1250  = " & tRefresco350_1250 & "<br>"	
				'
				'
			end if
			'
			' Calcular Tamaño 1500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1500 = 0
			else				
				'total1500 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1500 = 0 then
				'
				tRefresco350_1500 = 0
				Response.Write "350 no convive con 1500 / "
				'	
			else
				'
				tRefresco350_1500 = total1500 * 100 / totalRefrescoHogaresCon_350
				Response.Write "<br>Total compras de 1500  = " & total1500 & "<br>"
				Response.Write "Total Porcentaje 350/1500  = " & tRefresco350_1500 & "<br>"	
				'
				'
			end if
			'
			' Calcular Tamaño 2000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2000 = 0
			else				
				'total2000 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total2000 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total2000 = 0 then
				'
				tRefresco350_2000 = 0
				Response.Write "350 no convive con 2000 / "
				'	
			else
				'
				tRefresco350_2000 = total2000 * 100 / totalRefrescoHogaresCon_350
				Response.Write "<br>Total compras de 2000  = " & total2000 & "<br>"
				Response.Write "Total Porcentaje 350/2000  = " & tRefresco350_2000 & "<br>"	
				'
				'
			end if
			'
			' Calcular Tamaño 2500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2500 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total2500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total2500 = 0 then
				'
				tRefresco350_2500 = 0
				Response.Write "350 no convive con 2500 / "
				'	
			else
				'
				tRefresco350_2500 = total2500 * 100 / totalRefrescoHogaresCon_350
				Response.Write "<br>Total compras de 2500  = " & total600 & "<br>"
				Response.Write "Total Porcentaje 350/2500  = " & tRefresco350_2500 & "<br>"	
				'
			end if
			
		END IF
		'	
		
		'***
	end if
	'	
	Set rsx1 = nothing
	'
END SUB	
'
SUB Calcular_Refrescos_355
	Response.write "<BR><br>CALCULAR 355 ML<BR><BR>"
	'	
	' Buscar Todos Los Hogares compraron Tamaño 355 ml
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
	QrySql = QrySql & " AND PH_DataCrudaMensual.TAMANO = 355 "
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 "
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	'	
	'Response.write QrySql & "<br>"
	'Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then
		rsx1.close
	else
		'Erase dataArray
		dataArray = rsx1.GetRows
		rsx1.close
	end if
	'	
	Set rsx1 = nothing
	'
	IF IsArray(dataArray) THEN	
		'
		' Calculo total hogares con Refrescos 355
		'		
		totalRefrescoHogaresCon_355 = 0
		totalRefrescoHogaresCon_355 = ubound(dataArray,2) + 1 
		total355=0
		'
		Response.write "Total hogares 355 = " & totalRefrescoHogaresCon_355 & "<br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		'
		Response.write "<br>Total Compras 			 355 = " & totalRefrescoHogaresCon_355 & "<br>"
		Response.write "<br>ID Hogares que compraron 355 = " & replace(Hogares,",","-") & "<br>"
		'	
		'Response.end
		'				
		Set rsx1 = CreateObject("ADODB.Recordset")
		rsx1.CursorType = adOpenKeyset 
		rsx1.LockType   = 2 'adLockOptimistic 
		'
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_DataCrudaMensual"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 355"
		QrySql = QrySql & " GROUP BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " HAVING"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
		'
		'Response.write QrySql & "<br>"
		'Response.end
		'
		rsx1.Open QrySql, conexion		
		'
		if rsx1.eof then
			rsx1.close
			total355 = 0
		else				
			rsArray = rsx1.GetRows() 
        	total355 = UBound(rsArray, 2) + 1         	
			'total355 = rsx1.recordcount			
			rsx1.close			
		end if
		'
		'Response.write " Total RecordCount 355 = " & total355 & "<br>"
		'Response.end
		'
		IF total355 > 0 THEN
			'
			tRefresco355_355 = total355 * 100 / totalRefrescoHogaresCon_355
			'
			Response.Write "Total Compra de 355   = " & total355 & "<br>"
			Response.Write "Total porcentaje 355/355 = " & tRefresco355_355 & "<br>"
			'Response.End			
			'
			' Calcular Tamaño 250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total250 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total250 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total250 = 0 then
				'
				tRefresco355_250 = 0
				Response.Write "355 no convive con 250 / "
				'	
			else
				'
				tRefresco355_250 = total250 * 100 / totalRefrescoHogaresCon_355
				Response.Write "<br>Total compras de 250  = " & total250 & "<br>"
				Response.Write "Total Porcentaje 355/250  = " & tRefresco355_250 & "<br>"	
				'
			end if
			'
			' Calcular Tamaño 355/320
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 320"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & " 320 en hogares con 355<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total320 = 0
			else				
				'total320= rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total320 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total320 = 0 then
				'
				tRefresco355_320 = 0				
				Response.Write "355 no convive con 320 / "
				'	
			else
				'
				tRefresco355_320 = total320 * 100 / totalRefrescoHogaresCon_355
				Response.Write "<br>Total compras de 320  = " & total320 & "<br>"
				Response.Write "Total Porcentaje 355/320  = " & tRefresco355_320 & "<br>"	
				'
			end if
			' '
			' ' Calcular Tamaño 355
			' '
			' Set rsx1 = CreateObject("ADODB.Recordset")
			' rsx1.CursorType = adOpenKeyset 
			' rsx1.LockType   = 2 'adLockOptimistic 
			' '
			' QrySql = vbnullstring
			' QrySql = QrySql & " SELECT"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " FROM"
			' QrySql = QrySql & " PH_DataCrudaMensual"
			' QrySql = QrySql & " WHERE"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 355"
			' QrySql = QrySql & " GROUP BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " HAVING"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			' QrySql = QrySql & " ORDER BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			' '
			' 'Response.write QrySql & "<br>"
			' 'Response.end
			' '
			' rsx1.Open QrySql, conexion		
			' '
			' if rsx1.eof then
				' rsx1.close
				' total355 = 0
			' else				
				' 'total355 = rsx1.recordcount
				' rsArray = rsx1.GetRows() 
        		' total355 = UBound(rsArray, 2) + 1         	
				' rsx1.close
				' 'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			' end if
			' '
			' if total355 = 0 then
				' '
				' tRefresco355_355 = 0
				' Response.Write "355 no convive con 355 / "
				' '	
			' else
				' '
				' tRefresco355_355 = total355 * 100 / totalRefrescoHogaresCon_355
				' '
			' end if
			'
			' Calcular Tamaño 500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total500 = 0
			else				
				'total500 = rsx1.recordcount
				rsArray  = rsx1.GetRows() 
        		total500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total500 = 0 then
				'
				tRefresco355_500 = 0
				Response.Write "355 no convive con 500 / "
				'	
			else
				'
				tRefresco355_500 = total500 * 100 / totalRefrescoHogaresCon_355
				Response.Write "<br>Total compras de 500  = " & total500 & "<br>"
				Response.Write "Total Porcentaje 355/500  = " & tRefresco355_500 & "<br>"	
				'
			end if
			'
			' Calcular Tamaño 600
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 600"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total600 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total600 = UBound(rsArray, 2) + 1         	
				rsx1.close
				
			end if
			'
			if total600 = 0 then
				'
				tRefresco355_600 = 0
				Response.Write "355 no convive con 600 / "
				'	
			else
				'
				tRefresco355_600 = total600 * 100 / totalRefrescoHogaresCon_355
				Response.Write "<br>Total compras de 600  = " & total600 & "<br>"
				Response.Write "Total Porcentaje 355/600  = " & tRefresco355_600 & "<br>"	
				'
			end if
			'
			' Calcular Tamaño 1000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1000 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total1000 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1000 = 0 then
				'
				tRefresco355_1000 = 0
				Response.Write "355 no convive con 1000 / "
				'	
			else
				'
				tRefresco355_1000 = total1000 * 100 / totalRefrescoHogaresCon_355
				Response.Write "<br>Total compras de 1000  = " & total1000 & "<br>"
				Response.Write "Total Porcentaje 355/1000  = " & tRefresco355_1000 & "<br>"	
				'
			end if			
			'
			' Calcular Tamaño 1250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1250 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total1250 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total1250 = 0 then
				'
				tRefresco355_1250 = 0
				Response.Write "355 no convive con 1250 / "
				'	
			else
				'
				tRefresco355_1250 = total1250 * 100 / totalRefrescoHogaresCon_355
				Response.Write "<br>Total compras de 1250  = " & total1250 & "<br>"
				Response.Write "Total Porcentaje 355/1250  = " & tRefresco355_1250 & "<br>"	
				'
			end if
			'
			' Calcular Tamaño 1500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1500 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total1500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1500 = 0 then
				'
				tRefresco355_1500 = 0
				Response.Write "355 no convive con 1500 / "
				'	
			else
				'
				tRefresco355_1500 = total1500 * 100 / totalRefrescoHogaresCon_355
				Response.Write "<br>Total compras de 1500  = " & total1500 & "<br>"
				Response.Write "Total Porcentaje 355/1500  = " & tRefresco355_1500 & "<br>"	
				'
			end if
			'
			' Calcular Tamaño 2000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2000 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total2000 = UBound(rsArray, 2) + 1         	
				rsx1.close				
			end if
			'
			if total2000 = 0 then
				'
				tRefresco355_2000 = 0
				Response.Write "355 no convive con 2000 / "
				'	
			else
				'
				tRefresco355_2000 = total2000 * 100 / totalRefrescoHogaresCon_355
				Response.Write "<br>Total compras de 2000  = " & total2000 & "<br>"
				Response.Write "Total Porcentaje 355/2000  = " & tRefresco355_2000 & "<br>"	
				'
			end if
			'
			' Calcular Tamaño 2500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2500 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total2500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total2500 = 0 then
				'
				tRefresco355_2500 = 0
				Response.Write "355 no convive con 2500 / "
				'	
			else
				'
				tRefresco355_2500 = total2500 * 100 / totalRefrescoHogaresCon_355
				Response.Write "<br>Total compras de 2500  = " & total2500 & "<br>"
				Response.Write "Total Porcentaje 355/2500  = " & tRefresco355_2500 & "<br>"	
				'
			end if
			
		END IF
		'	
	ELSE
		'
		Response.Write "Tama&ntilde;o 355 No! Convive con nadie <BR><BR>"
		tRefresco355_250  = 0
		tRefresco355_320  = tRefresco355_350  = tRefresco355_355  = 0
		tRefresco355_500  = tRefresco355_600  = tRefresco355_1000 = 0
		tRefresco355_1500 = tRefresco355_2000 = tRefresco355_2500 = 0	
		'	
	END IF		

END SUB	
'
SUB Calcular_Refrescos_500
	Response.write "<BR><br>CALCULAR 500 ML<BR><BR>"
	'	
	' Buscar Todos Los Hogares compraron Tamaño 500 ml
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
	QrySql = QrySql & " AND PH_DataCrudaMensual.TAMANO = 500 "
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 "
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	'	
	'Response.write QrySql & "<br>"
	'Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then		
		rsx1.close
		'
		Response.Write "<br>Tama&ntilde;o 500 No! Convive con nadie <BR><BR>"
		tRefresco500_250  = 0
		tRefresco500_320  = tRefresco500_350  = tRefresco500_355  = 0
		tRefresco500_500  = tRefresco500_600  = tRefresco500_1000 = 0
		tRefresco500_1500 = tRefresco500_2000 = tRefresco500_2500 = 0	
		'			
	else
		'Erase dataArray
		dataArray = rsx1.GetRows
		rsx1.close
		'
		' Calculo total hogares con Refrescos 500
		'		
		totalRefrescoHogaresCon_500 = 0
		totalRefrescoHogaresCon_500 = ubound(dataArray,2) + 1 
		total500=0
		'
		'Response.write "<br>Total hogares 500 = " & totalRefrescoHogaresCon_500 & "<br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "<br>Total Compras 500 = " & totalRefrescoHogaresCon_500 & "<br>"
		Response.write "<br>ID Hogares que compraron 500 = " & replace(Hogares,",","-") & "<br>"
		'	
		'Response.end
		'				
		Set rsx1 = CreateObject("ADODB.Recordset")
		rsx1.CursorType = adOpenKeyset 
		rsx1.LockType   = 2 'adLockOptimistic 
		'
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_DataCrudaMensual"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 500"
		QrySql = QrySql & " GROUP BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " HAVING"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
		'
		'Response.write QrySql & "<br>"
		'Response.end
		'
		rsx1.Open QrySql, conexion		
		'
		if rsx1.eof then
			rsx1.close
			total500 = 0
		else				
			rsArray = rsx1.GetRows() 
        	total500 = UBound(rsArray, 2) + 1         	
			rsx1.close			
		end if
		'		
		'Response.end
		'
		IF total500 > 0 THEN
			'
			tRefresco500_500 = total500 * 100 / totalRefrescoHogaresCon_500
			'
			Response.Write "Total Compra de  500 = " & total500 & "<br>"
			Response.Write "Total Porcentaje 500/500 = " & tRefresco500_500 & "<br>"
			'
			' Calcular Tamaño 250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total250 = 0
			else				
				'total250 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total250 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total250 = 0 then
				'
				tRefresco500_250 = 0
				Response.Write "500 no convive con 250 / "
				'	
			else
				'
				tRefresco500_250 = total250 * 100 / totalRefrescoHogaresCon_500
				Response.Write "Total Compra de 250 = " & total250 & "<br>"
				Response.Write "Total Porcentaje 500/250 = " & tRefresco500_250 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 320
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 320"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total320 = 0
			else				
				'total320= rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total320 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total320 = 0 then
				'
				tRefresco500_320 = 0				
				Response.Write "500 no convive con 320 <br>"
				'	
			else
				'
				tRefresco500_320 = total320 * 100 / totalRefrescoHogaresCon_500
				Response.Write "Total Compra de 320 = " & total320 & "<br>"
				Response.Write "Total Porcentaje 500/320 = " & tRefresco500_320 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 355
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 355"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total355 = 0
			else				
				'total355 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total355 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total355 = 0 then
				'
				tRefresco500_355 = 0
				Response.Write "500 no convive con 355 / "
				'	
			else
				'
				tRefresco500_355 = total355 * 100 / totalRefrescoHogaresCon_500
				Response.Write "Total Compra de 355 = " & total355 & "<br>"
				Response.Write "Total Porcentaje 500/355 = " & tRefresco500_355 & "<br>"				
				'
			end if
			'
			' Calcular Tamaño 500
			'
			' Set rsx1 = CreateObject("ADODB.Recordset")
			' rsx1.CursorType = adOpenKeyset 
			' rsx1.LockType   = 2 'adLockOptimistic 
			' '
			' QrySql = vbnullstring
			' QrySql = QrySql & " SELECT"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " FROM"
			' QrySql = QrySql & " PH_DataCrudaMensual"
			' QrySql = QrySql & " WHERE"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 500"
			' QrySql = QrySql & " GROUP BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " HAVING"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			' QrySql = QrySql & " ORDER BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			' '
			' 'Response.write QrySql & "<br>"
			' 'Response.end
			' '
			' rsx1.Open QrySql, conexion		
			' '
			' if rsx1.eof then
				' rsx1.close
				' total500 = 0
			' else				
				' 'total500 = rsx1.recordcount
				' rsArray  = rsx1.GetRows() 
        		' total500 = UBound(rsArray, 2) + 1         	
				' rsx1.close
				' 'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			' end if
			' '
			' if total500 = 0 then
				' '
				' tRefresco500_500 = 0
				' Response.Write "500 no convive con 500<br>"
				' '	
			' else
				' '
				' tRefresco500_500 = total500 * 100 / totalRefrescoHogaresCon_500
				' '
			' end if
			'
			' Calcular Tamaño 600
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 600"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total600 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total600 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total600 = 0 then
				'
				tRefresco500_600 = 0
				Response.Write "500 no convive con 600<br>"
				'	
			else
				'
				tRefresco500_600 = total600 * 100 / totalRefrescoHogaresCon_500
				Response.Write "Total Compra de 600 = " & total600 & "<br>"
				Response.Write "Total Porcentaje 500/600 = " & tRefresco500_600 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 1000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1000 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total1000 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total1000 = 0 then
				'
				tRefresco500_1000 = 0
				Response.Write "500 no convive con 1000 / "
				'	
			else
				'
				tRefresco500_1000 = total1000 * 100 / totalRefrescoHogaresCon_500
				Response.Write "Total Compra de 1000 = " & total600 & "<br>"
				Response.Write "Total Porcentaje 500/1000 = " & tRefresco500_1000 & "<br>"
				'
			end if			
			'
			' Calcular Tamaño 1250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1250 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total1250 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total1250 = 0 then
				'
				tRefresco500_1250 = 0
				Response.Write "500 no convive con 1250 / "
				'	
			else
				'
				tRefresco500_1250 = total1250 * 100 / totalRefrescoHogaresCon_500
				Response.Write "Total Compra de 1250 = " & total1250 & "<br>"
				Response.Write "Total Porcentaje 500/1250 = " & tRefresco500_1250 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 1500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1500 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total1500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total1500 = 0 then
				'
				tRefresco500_1500 = 0
				Response.Write "500 no convive con 1500<br>"
				'	
			else
				'
				tRefresco500_1500 = total1500 * 100 / totalRefrescoHogaresCon_500
				Response.Write "Total Compra de 1500 = " & total1500 & "<br>"
				Response.Write "Total Porcentaje 500/1500 = " & tRefresco500_1500 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 2000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2000 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total2000 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total2000 = 0 then
				'
				tRefresco500_2000 = 0
				Response.Write "500 no convive con 2000 / "
				'	
			else
				'
				tRefresco500_2000 = total2000 * 100 / totalRefrescoHogaresCon_500
				Response.Write "Total Compra de 2000 = " & total2000 & "<br>"
				Response.Write "Total Porcentaje 500/2000 = " & tRefresco500_2000 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 2500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2500 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total2500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total2500 = 0 then
				'
				tRefresco500_2500 = 0
				Response.Write "500 no convive con 2500 / "
				'	 
			else
				'
				tRefresco500_2500 = total2500 * 100 / totalRefrescoHogaresCon_500
				Response.Write "Total Compra de 2500 = " & total2500 & "<br>"
				Response.Write "Total Porcentaje 500/2500 = " & tRefresco500_2500 & "<br>"
				'
			end if
			
		END IF
		'
		
	end if
	'	
	Set rsx1 = nothing
	'
END SUB
'
SUB Calcular_Refrescos_600
	Response.write "<BR><br>CALCULAR 600 ML<BR><BR>"
	'	
	' Buscar Todos Los Hogares compraron Tamaño 600 ml
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
	QrySql = QrySql & " AND PH_DataCrudaMensual.TAMANO = 600 "
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 "
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	'	
	'Response.write QrySql & "<br>"
	'Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then		
		rsx1.close
		'
		Response.Write "<br>Tama&ntilde;o 600 No! Convive con nadie <BR><BR>"
		tRefresco600_250  = 0
		tRefresco600_320  = tRefresco600_350  = tRefresco600_355  = 0
		tRefresco600_500  = tRefresco600_600  = tRefresco600_1000 = 0
		tRefresco600_1500 = tRefresco600_2000 = tRefresco600_2500 = 0	
		'			
	else
		'Erase dataArray
		dataArray = rsx1.GetRows
		rsx1.close
		'
		' Calculo total hogares con Refrescos 600
		'		
		totalRefrescoHogaresCon_600 = 0
		totalRefrescoHogaresCon_600 = ubound(dataArray,2) + 1 
		total600=0
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "<br>Total Compras 600 = " & totalRefrescoHogaresCon_600 & "<br>"
		Response.write "<br>ID Hogares que compraron 600 = " & replace(Hogares,",","-") & "<br>"
		'	
		Response.write "hogares con 600 = " & replace(Hogares,",","-") & "<br>"
		'Response.end
		'				
		Set rsx1 = CreateObject("ADODB.Recordset")
		rsx1.CursorType = adOpenKeyset 
		rsx1.LockType   = 2 'adLockOptimistic 
		'
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_DataCrudaMensual"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 600"
		QrySql = QrySql & " GROUP BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " HAVING"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
		'
		'Response.write QrySql & "<br>"
		'Response.end
		'
		rsx1.Open QrySql, conexion		
		'
		if rsx1.eof then
			rsx1.close
			total600 = 0
		else				
			rsArray = rsx1.GetRows() 
        	total600 = UBound(rsArray, 2) + 1         	
			'total600 = rsx1.recordcount			
			rsx1.close			
		end if			
		'
		IF total600 > 0 THEN
			'
			tRefresco600_600 = total600 * 100 / totalRefrescoHogaresCon_600
			'
			Response.Write "Total Compra de 600 = " & total600 & "<br>"
			Response.Write "Total Porcentaje 600/600 = " & tRefresco600_600 & "<br>"
			'				
			'Response.End			
			'
			' Calcular Tamaño 250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total250 = 0
			else				
				'total250 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total250 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total250 = 0 then
				'
				tRefresco600_250 = 0
				Response.Write "600 no convive con 250 / "
				'	
			else
				'
				tRefresco600_250 = total250 * 100 / totalRefrescoHogaresCon_600
				Response.Write "Total Compra de 600 = " & total250 & "<br>"
				Response.Write "Total Porcentaje 600/250 = " & tRefresco600_250 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 320
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 320"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total320 = 0
			else				
				'total320= rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total320 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total320 = 0 then
				'
				tRefresco600_320 = 0				
				Response.Write "600 no convive con 320 / "
				'	
			else
				'
				tRefresco600_320 = total320 * 100 / totalRefrescoHogaresCon_600
				Response.Write "Total Compra de 600 = " & total320 & "<br>"
				Response.Write "Total Porcentaje 600/320 = " & tRefresco600_320 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 355
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 355"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total355 = 0
			else				
				'total355 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total355 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total355 = 0 then
				'
				tRefresco600_355 = 0
				Response.Write "600 no convive con 355 / "
				'	
			else
				'
				tRefresco600_355 = total355 * 100 / totalRefrescoHogaresCon_600
				Response.Write "Total Compra de 600 = " & total355 & "<br>"
				Response.Write "Total Porcentaje 600/355 = " & tRefresco600_355 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total500 = 0
			else				
				rsArray  = rsx1.GetRows() 
        		total500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total500 = 0 then
				'
				tRefresco600_500 = 0
				Response.Write "600 no convive con 500 / "
				'	
			else
				'
				tRefresco600_500 = total500 * 100 / totalRefrescoHogaresCon_600
				Response.Write "Total Compra de 500 = " & total500 & "<br>"
				Response.Write "Total Porcentaje 600/500 = " & tRefresco600_500 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 600
			'
			' Set rsx1 = CreateObject("ADODB.Recordset")
			' rsx1.CursorType = adOpenKeyset 
			' rsx1.LockType   = 2 'adLockOptimistic 
			' '
			' QrySql = vbnullstring
			' QrySql = QrySql & " SELECT"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " FROM"
			' QrySql = QrySql & " PH_DataCrudaMensual"
			' QrySql = QrySql & " WHERE"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 600"
			' QrySql = QrySql & " GROUP BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " HAVING"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			' QrySql = QrySql & " ORDER BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			' '
			' 'Response.write QrySql & "<br>"
			' 'Response.end
			' '
			' rsx1.Open QrySql, conexion		
			' '
			' if rsx1.eof then
				' rsx1.close
				' total600 = 0
			' else				
				' 'total600 = rsx1.recordcount
				' rsArray = rsx1.GetRows() 
        		' total600 = UBound(rsArray, 2) + 1         	
				' rsx1.close
				' 'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			' end if
			' '
			' if total600 = 0 then
				' '
				' tRefresco600_600 = 0
				' Response.Write "600 no convive con 600<br>"
				' '	
			' else
				' '
				' tRefresco600_600 = total600 * 100 / totalRefrescoHogaresCon_600
				' '
			' end if
			'
			' Calcular Tamaño 1000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1000 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total1000 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total1000 = 0 then
				'
				tRefresco600_1000 = 0
				Response.Write "600 no convive con 1000 / "
				'	
			else
				'
				tRefresco600_1000 = total1000 * 100 / totalRefrescoHogaresCon_600
				Response.Write "<br>Total compras de 600 =" & total1000 & "<br>"
				Response.Write "Total Porcentaje 600/1000 =" & tRefresco600_1000 & "<br>"
				'				
			end if			
			'
			' Calcular Tamaño 1250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1250 = 0
			else				
				'total1250 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1250 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1250 = 0 then
				'
				tRefresco600_1250 = 0
				Response.Write "600 no convive con 1250 / "
				'	
			else
				'
				tRefresco600_1250 = total1250 * 100 / totalRefrescoHogaresCon_600
				Response.Write "Total Compra de 1250 = " & total1250 & "<br>"
				Response.Write "Total Porcentaje 600/1250 = " & tRefresco600_1250 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 1500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1500 = 0
			else	
				rsArray = rsx1.GetRows() 
        		total1500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total1500 = 0 then
				'
				tRefresco600_1500 = 0
				Response.Write "600 no convive con 1500 / "
				'	
			else
				'
				tRefresco600_1500 = total1500 * 100 / totalRefrescoHogaresCon_600
				Response.Write "Total Compra de 1500 = " & total1500 & "<br>"
				Response.Write "Total Porcentaje 600/1500 = " & tRefresco600_1500 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 2000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2000 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total2000 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total2000 = 0 then
				'
				tRefresco600_2000 = 0
				Response.Write "600 no convive con 2000 / "
				'	
			else
				'
				tRefresco600_2000 = total2000 * 100 / totalRefrescoHogaresCon_600
				Response.Write "Total Compra de 2000 = " & total2000 & "<br>"
				Response.Write "Total Porcentaje 600/2000 = " & tRefresco600_2000 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 2500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2500 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total2500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total2500 = 0 then
				'
				tRefresco600_2500 = 0
				Response.Write "600 no convive con 2500 / "
				'	
			else
				'
				tRefresco600_2500 = total2500 * 100 / totalRefrescoHogaresCon_600
				Response.Write "Total Compra de 2500 = " & total2500 & "<br>"
				Response.Write "Total Porcentaje 600/2500 = " & tRefresco600_2500 & "<br>"
				'
			end if
			
		END IF
		'
		
	end if
	'	
	Set rsx1 = nothing
	
END SUB
'
SUB Calcular_Refrescos_1000
	Response.write "<BR><br>CALCULAR 1000 ML<BR><BR>"
	'	
	' Buscar Todos Los Hogares compraron Tamaño 1000 ml
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
	QrySql = QrySql & " AND PH_DataCrudaMensual.TAMANO = 1000 "
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 "
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	'	
	'Response.write QrySql & "<br>"
	'Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then		
		rsx1.close
		'
		Response.Write "<br>Tama&ntilde;o 1000 No! Convive con nadie <BR><BR>"
		tRefresco1000_250  = 0
		tRefresco1000_320  = tRefresco1000_350  = tRefresco1000_355  = 0
		tRefresco1000_500  = tRefresco1000_600  = tRefresco1000_1000 = 0
		tRefresco1000_1500 = tRefresco1000_2000 = tRefresco1000_2500 = 0	
		'			
	else
		'Erase dataArray
		dataArray = rsx1.GetRows
		rsx1.close
		'
		' Calculo total hogares con Refrescos 1000
		'		
		totalRefrescoHogaresCon_1000 = 0
		totalRefrescoHogaresCon_1000 = ubound(dataArray,2) + 1 
		total1000=0
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "<br>Total Compras 1000 = " & totalRefrescoHogaresCon_1000 & "<br>"
		Response.write "<br>ID Hogares que compraron 1000 = " & replace(Hogares,",","-") & "<br>"
		'	
		Set rsx1 = CreateObject("ADODB.Recordset")
		rsx1.CursorType = adOpenKeyset 
		rsx1.LockType   = 2 'adLockOptimistic 
		'
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_DataCrudaMensual"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1000"
		QrySql = QrySql & " GROUP BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " HAVING"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
		'
		'Response.write QrySql & "<br>"
		'Response.end
		'
		rsx1.Open QrySql, conexion		
		'
		if rsx1.eof then
			rsx1.close
			total1000 = 0
		else				
			rsArray = rsx1.GetRows() 
        	total1000 = UBound(rsArray, 2) + 1         	
			'total1000 = rsx1.recordcount			
			rsx1.close			
		end if
		'
		'Response.write " Total RecordCount 1000 = " & total1000 & "<br>"
		'Response.end
		'
		IF total1000 > 0 THEN
			'
			tRefresco1000_1000 = total1000 * 100 / totalRefrescoHogaresCon_1000
			'			
			Response.Write "Total Compra de 1000 = " & total1000 & "<br>"
			Response.Write "Total Porcentaje 1000/1000 = " & tRefresco1000_1000 & "<br>"			
			'
			' Calcular Tamaño 250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total250 = 0
			else				
				'total250 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total250 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total250 = 0 then
				'
				tRefresco1000_250 = 0
				Response.Write "1000 no convive con 250 / "
				'	
			else
				'
				tRefresco1000_250 = total250 * 100 / totalRefrescoHogaresCon_1000
				Response.Write "Total Compra de 250 = " & total250 & "<br>"
				Response.Write "Total Porcentaje 1000/250 = " & tRefresco1000_250 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 320
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 320"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total320 = 0
			else				
				'total320= rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total320 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total320 = 0 then
				'
				tRefresco1000_320 = 0				
				Response.Write "1000 no convive con 320 / "
				'	
			else
				'
				tRefresco1000_320 = total320 * 100 / totalRefrescoHogaresCon_1000
				Response.Write "Total Compra de 320 = " & total320 & "<br>"
				Response.Write "Total Porcentaje 1000/320 = " & tRefresco1000_320 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 355
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 355"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total355 = 0
			else				
				'total355 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total355 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total355 = 0 then
				'
				tRefresco1000_355 = 0
				Response.Write "1000 no convive con 355 / "
				'	
			else
				'
				tRefresco1000_355 = total355 * 100 / totalRefrescoHogaresCon_1000
				Response.Write "Total Compra de 355 = " & total355 & "<br>"
				Response.Write "Total Porcentaje 1000/355= " & tRefresco1000_355 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total500 = 0
			else				
				rsArray  = rsx1.GetRows() 
        		total500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total500 = 0 then
				'
				tRefresco1000_500 = 0
				Response.Write "1000 no convive con 500 / "
				'	
			else
				'
				tRefresco1000_500 = total500 * 100 / totalRefrescoHogaresCon_1000
				Response.Write "Total Compra de 500 = " & total500 & "<br>"
				Response.Write "Total Porcentaje 1000/500 = " & tRefresco1000_500 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 600
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 600"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total600 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total600 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total600 = 0 then
				'
				tRefresco1000_600 = 0
				Response.Write "1000 no convive con 600 / "
				'	
			else
				'
				tRefresco1000_600 = total600 * 100 / totalRefrescoHogaresCon_1000
				'
				Response.Write "<br>Total compras de 600 =" & total600 & "<br>"
				Response.Write "Total Porcentaje 1000/600 =" & tRefresco1000_600 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 1000
			'
			' Set rsx1 = CreateObject("ADODB.Recordset")
			' rsx1.CursorType = adOpenKeyset 
			' rsx1.LockType   = 2 'adLockOptimistic 
			' '
			' QrySql = vbnullstring
			' QrySql = QrySql & " SELECT"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " FROM"
			' QrySql = QrySql & " PH_DataCrudaMensual"
			' QrySql = QrySql & " WHERE"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1000"
			' QrySql = QrySql & " GROUP BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " HAVING"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			' QrySql = QrySql & " ORDER BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			' '
			' 'Response.write QrySql & "<br>"
			' 'Response.end
			' '
			' rsx1.Open QrySql, conexion		
			' '
			' if rsx1.eof then
				' rsx1.close
				' total1000 = 0
			' else				
				' 'total1000 = rsx1.recordcount
				' rsArray = rsx1.GetRows() 
        		' total1000 = UBound(rsArray, 2) + 1         	
				' rsx1.close
				' 'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			' end if
			' '
			' if total1000 = 0 then
				' '
				' tRefresco1000_600 = 0
				' Response.Write "1000 no convive con 1000<br>"
				' '	
			' else
				' '
				' tRefresco1000_600 = total1000 * 100 / totalRefrescoHogaresCon_1000
				' '
			' end if			
			'
			' Calcular Tamaño 1250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1250 = 0
			else								
				rsArray = rsx1.GetRows() 
        		total1250 = UBound(rsArray, 2) + 1         	
				rsx1.close				
			end if
			'
			if total1250 = 0 then
				'
				tRefresco1000_1250 = 0
				Response.Write "1000 no convive con 1250 / "
				'	
			else
				'
				tRefresco1000_1250 = total1250 * 100 / totalRefrescoHogaresCon_1000
				Response.Write "Total Compra de 1250 = " & total1250 & "<br>"
				Response.Write "Total Porcentaje 1000/1250 = " & tRefresco1000_1250 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 1500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1500 = 0
			else				
				'total1500 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1500 = 0 then
				'
				tRefresco1000_1500 = 0
				Response.Write "1000 no convive con 1500 / "
				'	
			else
				'
				tRefresco1000_1500 = total1500 * 100 / totalRefrescoHogaresCon_1000
				Response.Write "Total Compra de 1500 = " & total1500 & "<br>"
				Response.Write "Total Porcentaje 1000/1500 = " & tRefresco1000_1500 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 2000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2000 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total2000 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total2000 = 0 then
				'
				tRefresco1000_2000 = 0
				Response.Write "1000 no convive con 2000 / "
				'	
			else
				'
				tRefresco1000_2000 = total2000 * 100 / totalRefrescoHogaresCon_1000
				Response.Write "Total Compra de 2000 = " & total2000 & "<br>"
				Response.Write "Total Porcentaje 1000/2000 = " & tRefresco1000_2000 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 2500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2500 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total2500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total2500 = 0 then
				'
				tRefresco1000_2500 = 0
				Response.Write "1000 no convive con 2500 / "
				'	
			else
				'
				tRefresco1000_2500 = total2500 * 100 / totalRefrescoHogaresCon_1000
				Response.Write "Total Compra de 2500 = " & total2500 & "<br>"
				Response.Write "Total Porcentaje 1000/2500 = " & tRefresco1000_2500 & "<br>"
				'
			end if
			
		END IF
		'
		
	end if
	'	
	Set rsx1 = nothing
	
END SUB
'
SUB Calcular_Refrescos_1250
	Response.write "<BR><br>CALCULAR 1250 ML<BR><BR>"
	'	
	' Buscar Todos Los Hogares compraron Tamaño 1250 ml
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
	QrySql = QrySql & " AND PH_DataCrudaMensual.TAMANO = 1250 "
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 "
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	'	
	'Response.write QrySql & "<br>"
	'Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then		
		rsx1.close
		'
		Response.Write "<br>Tama&ntilde;o 1250 No! Convive con nadie <BR><BR>"
		tRefresco1250_250  = tRefresco1250_1250 = 0
		tRefresco1250_320  = tRefresco1250_350  = tRefresco1250_355  = 0
		tRefresco1250_500  = tRefresco1250_600  = tRefresco1250_1000 = 0
		tRefresco1250_1500 = tRefresco1250_2000 = tRefresco1250_2500 = 0	
		'			
	else
		'Erase dataArray
		dataArray = rsx1.GetRows
		rsx1.close
		'
		' Calculo total hogares con Refrescos 1250
		'		
		totalRefrescoHogaresCon_1250 = 0
		totalRefrescoHogaresCon_1250 = ubound(dataArray,2) + 1 
		total1250=0
		'
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "<br>Total Compras 1250= " & totalRefrescoHogaresCon_1250 & "<br>"
		Response.write "<br>ID Hogares que compraron 1250 = " & replace(Hogares,",","-") & "<br>"
		'
		'Response.end
		'				
		Set rsx1 = CreateObject("ADODB.Recordset")
		rsx1.CursorType = adOpenKeyset 
		rsx1.LockType   = 2 'adLockOptimistic 
		'
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_DataCrudaMensual"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1250"
		QrySql = QrySql & " GROUP BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " HAVING"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
		'
		'Response.write QrySql & "<br>"
		'Response.end
		'
		rsx1.Open QrySql, conexion		
		'
		if rsx1.eof then
			rsx1.close
			total1250 = 0
		else				
			rsArray = rsx1.GetRows() 
        	total1250 = UBound(rsArray, 2) + 1         	
			'total1250 = rsx1.recordcount			
			rsx1.close			
		end if
		'		
		'Response.end
		'
		IF total1250 > 0 THEN
			'
			tRefresco1250_1250 = total1250 * 100 / totalRefrescoHogaresCon_1250
			'
			Response.Write "<br>"
			Response.Write "Total Compra de  1250      = " & total1250 & "<br>"
			
			Response.Write "total Refresco Hogares Con_1250  = " & totalRefrescoHogaresCon_1250 & "<br>"
			
			Response.Write "Total Porcentaje 1250/1250 = " & tRefresco1250_1250 & "<br>"
			Response.Write "<br>"
			'
			' Response.End			
			'
			' Calcular Tamaño 250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total250 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total250 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total250 = 0 then
				'
				tRefresco1250_250 = 0
				Response.Write "1250 no convive con 250 / "
				'	
			else
				'
				tRefresco1250_250 = total250 * 100 / totalRefrescoHogaresCon_1250
				Response.Write "Total Compra de  250      = " & total250 & "<br>"
				Response.Write "Total Porcentaje 1250/250 = " & tRefresco1250_250 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 320
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 320"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total320 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total320 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total320 = 0 then
				'
				tRefresco1250_320 = 0				
				Response.Write "1250 no convive con 320 / "
				'	
			else
				'
				tRefresco1250_320 = total320 * 100 / totalRefrescoHogaresCon_1250
				Response.Write "Total Compra de  320       = " & total320 & "<br>"
				Response.Write "Total Porcentaje 1250/320 = " & tRefresco1250_320 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 355
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 355"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total355 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total355 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total355 = 0 then
				'
				tRefresco1250_355 = 0
				Response.Write "1250 no convive con 355 / "
				'	
			else
				'
				tRefresco1250_355 = total355 * 100 / totalRefrescoHogaresCon_1250
				Response.Write "Total Compra de  355      = " & total355 & "<br>"
				Response.Write "Total Porcentaje 1250/355 = " & tRefresco1250_355 & "<br>"								
				'
			end if
			'
			' Calcular Tamaño 500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total500 = 0
			else				
				rsArray  = rsx1.GetRows() 
        		total500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total500 = 0 then
				'
				tRefresco1250_500 = 0
				Response.Write "1250 no convive con 500 / "
				'	
			else
				'
				tRefresco1250_500 = total500 * 100 / totalRefrescoHogaresCon_1250
				Response.Write "Total Compra de  500      = " & total500 & "<br>"
				Response.Write "Total Porcentaje 1250/500 = " & tRefresco1250_500 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 600
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 600"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total600 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total600 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total600 = 0 then
				'
				tRefresco1250_600 = 0
				Response.Write "1250 no convive con 600 / "
				'	
			else
				'
				tRefresco1250_600 = total600 * 100 / totalRefrescoHogaresCon_1250
				Response.Write "Total Compra de  600      = " & total600 & "<br>"
				Response.Write "Total Porcentaje 1250/600 = " & tRefresco1250_600 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 1000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1000 = 0
			else				
				'total1000 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1000 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1000 = 0 then
				'
				tRefresco1250_1000 = 0
				Response.Write "1250 no convive con 1000 / "
				'	
			else
				'
				tRefresco1250_1000 = total1000 * 100 / totalRefrescoHogaresCon_1250
				Response.Write "Total Compra de  1000      = " & total1000 & "<br>"
				Response.Write "Total Porcentaje 1250/1000 = " & tRefresco1250_1000 & "<br>"
				'
			end if			
			'
			' Calcular Tamaño 1250
			'
			' Set rsx1 = CreateObject("ADODB.Recordset")
			' rsx1.CursorType = adOpenKeyset 
			' rsx1.LockType   = 2 'adLockOptimistic 
			' '
			' QrySql = vbnullstring
			' QrySql = QrySql & " SELECT"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " FROM"
			' QrySql = QrySql & " PH_DataCrudaMensual"
			' QrySql = QrySql & " WHERE"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1250"
			' QrySql = QrySql & " GROUP BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " HAVING"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			' QrySql = QrySql & " ORDER BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			' '
			' Response.write "<br><br>"
			' Response.write QrySql & "<br>"
			' 'Response.end
			' '
			' rsx1.Open QrySql, conexion		
			' '
			' if rsx1.eof then
				' rsx1.close
				' total1250 = 0
			' else				
				' rsArray = rsx1.GetRows() 
        		' total1250 = UBound(rsArray, 2) + 1         	
				' rsx1.close
			' end if
			' '
			' if total1250 = 0 then
				' '
				' 'tRefresco1250_1250 = 0
				' Response.Write "1250 no convive con 1250 / "
				' '	
			' else
				' '
				' Response.write "<br> REVISION <br>"
				' 'tRefresco1250_1250 = total1250 * 100 / totalRefrescoHogaresCon_1250
				' Response.Write "Total Compra de  1250      = " & total1250 & "<br>"
				' Response.Write "Total Porcentaje 1250/1250 = " & tRefresco1250_1250 & "<br>"
				' Response.write "<br><br>"
				' '
			' end if
			'
			' Calcular Tamaño 1500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1500 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total1500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total1500 = 0 then
				'
				tRefresco1250_1500 = 0
				Response.Write "1250 no convive con 1500 / "
				'	
			else
				'
				tRefresco1250_1500 = total1500 * 100 / totalRefrescoHogaresCon_1250
				Response.Write "Total Compra de  1500      = " & total1500 & "<br>"
				Response.Write "Total Porcentaje 1250/1500 = " & tRefresco1250_1500 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 2000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2000 = 0
			else				
				'total2000 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total2000 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total2000 = 0 then
				'
				tRefresco1250_2000 = 0
				Response.Write "1250 no convive con 2000 / "
				'	
			else
				'
				tRefresco1250_2000 = total2000 * 100 / totalRefrescoHogaresCon_1250
				Response.Write "Total Compra de  2000      = " & total2000 & "<br>"
				Response.Write "Total Porcentaje 1250/2000 = " & tRefresco1250_2000 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 2500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2500 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total2500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total2500 = 0 then
				'
				tRefresco1250_2500 = 0
				Response.Write "1250 no convive con 2500 / "
				'	
			else
				'
				tRefresco1250_2500 = total2500 * 100 / totalRefrescoHogaresCon_1250
				Response.Write "Total Compra de  2500      = " & total2500 & "<br>"
				Response.Write "Total Porcentaje 1250/2500 = " & tRefresco1250_2500 & "<br>"
				'
			end if
			
		END IF
		'
		
	end if
	'	
	Set rsx1 = nothing
	
END SUB
'
SUB Calcular_Refrescos_1500
	Response.write "<BR><br>CALCULAR 1500 ML<BR><BR>"
	'	
	' Buscar Todos Los Hogares compraron Tamaño 1500 ml
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
	QrySql = QrySql & " AND PH_DataCrudaMensual.TAMANO = 1500 "
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 "
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	'	
	'Response.write QrySql & "<br>"
	'Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then		
		rsx1.close
		'
		Response.Write "<br>Tama&ntilde;o 1500 No! Convive con nadie <BR><BR>"
		tRefresco1500_250  = 0
		tRefresco1500_320  = tRefresco1500_350  = tRefresco1500_355  = 0
		tRefresco1500_500  = tRefresco1500_600  = tRefresco1500_1000 = 0
		tRefresco1500_1250 = tRefresco1500_2000 = tRefresco1500_2500 = 0	
		'			
	else
		'Erase dataArray
		dataArray = rsx1.GetRows
		rsx1.close
		'
		' Calculo total hogares con Refrescos 1500
		'		
		totalRefrescoHogaresCon_1500 = 0
		totalRefrescoHogaresCon_1500 = ubound(dataArray,2) + 1 
		total1500=0
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "<br>Total Compras 1500 = " & totalRefrescoHogaresCon_1500 & "<br>"
		Response.write "<br>ID Hogares que compraron 1500 = " & replace(Hogares,",","-") & "<br>"
		'		
		'Response.end
		'				
		Set rsx1 = CreateObject("ADODB.Recordset")
		rsx1.CursorType = adOpenKeyset 
		rsx1.LockType   = 2 'adLockOptimistic 
		'
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_DataCrudaMensual"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1500"
		QrySql = QrySql & " GROUP BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " HAVING"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
		'
		'Response.write QrySql & "<br>"
		'Response.end
		'
		rsx1.Open QrySql, conexion		
		'
		if rsx1.eof then
			rsx1.close
			total1500 = 0
		else				
			rsArray = rsx1.GetRows() 
        	total1500 = UBound(rsArray, 2) + 1         	
			rsx1.close			
		end if
		'		
		'Response.end
		'
		IF total1500 > 0 THEN
			'
			tRefresco1500_1500 = total1500 * 100 / totalRefrescoHogaresCon_1500
			'
			Response.Write "Total Compra de       1500 = " & total1500 & "<br>"
			Response.Write "Total Porcentaje 1500/1500 = " & tRefresco1500_1500 & "<br>"
			'Response.End			
			'
			' Calcular Tamaño 250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total250 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total250 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total250 = 0 then
				'
				tRefresco1500_250 = 0
				Response.Write "1500 no convive con 250 / "
				'	
			else
				'
				tRefresco1500_250 = total250 * 100 / totalRefrescoHogaresCon_1500
				Response.Write "Total Compra de  250      = " & total250 & "<br>"
				Response.Write "Total Porcentaje 1500/250 = " & tRefresco1500_250 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 320
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 320"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total320 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total320 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total320 = 0 then
				'
				tRefresco1500_320 = 0				
				Response.Write "1500 no convive con 320 / "
				'	
			else
				'
				tRefresco1500_320 = total320 * 100 / totalRefrescoHogaresCon_1500
				Response.Write "Total Compra de  320      = " & total320 & "<br>"
				Response.Write "Total Porcentaje 1500/320 = " & tRefresco1500_320 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 355
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 355"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total355 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total355 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total355 = 0 then
				'
				tRefresco1500_355 = 0
				Response.Write "1500 no convive con 355 / "
				'	
			else
				'
				tRefresco1500_355 = total355 * 100 / totalRefrescoHogaresCon_1500
				Response.Write "Total Compra de  355      = " & total355 & "<br>"
				Response.Write "Total Porcentaje 1500/355 = " & tRefresco1500_355 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total500 = 0
			else				
				rsArray  = rsx1.GetRows() 
        		total500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total500 = 0 then
				'
				tRefresco1500_500 = 0
				Response.Write "1500 no convive con 500 / "
				'	
			else
				'
				tRefresco1500_500 = total500 * 100 / totalRefrescoHogaresCon_1500
				Response.Write "Total Compra de  500      = " & total500 & "<br>"
				Response.Write "Total Porcentaje 1500/500 = " & tRefresco1500_500 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 600
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 600"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total600 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total600 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total600 = 0 then
				'
				tRefresco1500_600 = 0
				Response.Write "1500 no convive con 600 / "
				'	
			else
				'
				tRefresco1500_600 = total600 * 100 / totalRefrescoHogaresCon_1500
				Response.Write "Total Compra de  600      = " & total600 & "<br>"
				Response.Write "Total Porcentaje 1500/600 = " & tRefresco1500_600 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 1000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1000 = 0
			else								
				rsArray = rsx1.GetRows() 
        		total1000 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total1000 = 0 then
				'
				tRefresco1500_1000 = 0
				Response.Write "1250 no convive con 1000 / "
				'	
			else
				'
				tRefresco1500_1000 = total1000 * 100 / totalRefrescoHogaresCon_1500
				Response.Write "Total Compra de  1000      = " & total1000 & "<br>"
				Response.Write "Total Porcentaje 1500/1000 = " & tRefresco1500_1000 & "<br>"
				'
			end if			
			'
			' Calcular Tamaño 1250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1250 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total1250 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total1250 = 0 then
				'
				tRefresco1500_1250 = 0
				Response.Write "1500 no convive con 1250 / "
				'	
			else
				'
				tRefresco1500_1250 = total1250 * 100 / totalRefrescoHogaresCon_1500
				Response.Write "Total Compra de  1250      = " & total1250 & "<br>"
				Response.Write "Total Porcentaje 1500/1250 = " & tRefresco1500_1250 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 1500
			'
			' Set rsx1 = CreateObject("ADODB.Recordset")
			' rsx1.CursorType = adOpenKeyset 
			' rsx1.LockType   = 2 'adLockOptimistic 
			' '
			' QrySql = vbnullstring
			' QrySql = QrySql & " SELECT"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " FROM"
			' QrySql = QrySql & " PH_DataCrudaMensual"
			' QrySql = QrySql & " WHERE"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1500"
			' QrySql = QrySql & " GROUP BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			' QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			' QrySql = QrySql & " HAVING"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			' QrySql = QrySql & " ORDER BY"
			' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			' '
			' 'Response.write QrySql & "<br>"
			' 'Response.end
			' '
			' rsx1.Open QrySql, conexion		
			' '
			' if rsx1.eof then
				' rsx1.close
				' total1500 = 0
			' else				
				' rsArray = rsx1.GetRows() 
        		' total1500 = UBound(rsArray, 2) + 1         	
				' rsx1.close
			' end if
			' '
			' if total1500 = 0 then
				' '
				' tRefresco1500_1500 = 0
				' Response.Write "1500 no convive con 1500 / "
				' '	
			' else
				' '
				' tRefresco1500_1500 = total1500 * 100 / totalRefrescoHogaresCon_1500
				' Response.Write "Total Compra de  1500      = " & total1250 & "<br>"
				' Response.Write "Total Porcentaje 1500/1500 = " & tRefresco1500_1500 & "<br>"
				' '
			' end if
			'
			' Calcular Tamaño 2000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2000 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total2000 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total2000 = 0 then
				'
				tRefresco1500_2000 = 0
				Response.Write "1500 no convive con 2000 / "
				'	
			else
				'
				tRefresco1500_2000 = total2000 * 100 / totalRefrescoHogaresCon_1500
				Response.Write "Total Compra de 2000      = " & total2000 & "<br>"
				Response.Write "Total Porcentaje 1500/2000 = " & tRefresco1500_2000 & "<br>"
				'
			end if
			'
			' Calcular Tamaño 2500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2500 = 0
			else				
				rsArray = rsx1.GetRows() 
        		total2500 = UBound(rsArray, 2) + 1         	
				rsx1.close
			end if
			'
			if total2500 = 0 then
				'
				tRefresco1500_2500 = 0
				Response.Write "1500 no convive con 2500 / "
				'	
			else
				'
				tRefresco1500_2500 = total2500 * 100 / totalRefrescoHogaresCon_1500
				Response.Write "Total Compra de  2500      = " & total2500 & "<br>"
				Response.Write "Total Porcentaje 1500/2500 = " & tRefresco1500_2500 & "<br>"
				'
			end if
			
		END IF
		'
		
	end if
	'	
	Set rsx1 = nothing
	
END SUB
'
SUB Calcular_Refrescos_2000
	'	
	' Buscar Todos Los Hogares compraron Tamaño 2000 ml
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
	QrySql = QrySql & " AND PH_DataCrudaMensual.TAMANO = 2000 "
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 "
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	'	
	'Response.write QrySql & "<br>"
	'Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then		
		rsx1.close
		'
		Response.Write "<br>Tama&ntilde;o 2000 No! Convive con nadie <BR><BR>"
		tRefresco2000_250  = 0
		tRefresco2000_320  = tRefresco2000_350  = tRefresco2000_355  = 0
		tRefresco2000_500  = tRefresco2000_600  = tRefresco2000_1000 = 0
		tRefresco2000_1250 = tRefresco2000_2000 = tRefresco2000_2500 = 0	
		'			
	else
		'Erase dataArray
		dataArray = rsx1.GetRows
		rsx1.close
		'
		' Calculo total hogares con Refrescos 2000
		'		
		totalRefrescoHogaresCon_2000 = 0
		totalRefrescoHogaresCon_2000 = ubound(dataArray,2) + 1 
		total2000=0
		'
		Response.write "<br>Total hogares 2000 = " & totalRefrescoHogaresCon_2000 & "<br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "ID hogares con 2000 = " & replace(Hogares,",","-") & "<br>"
		'Response.end
		'				
		Set rsx1 = CreateObject("ADODB.Recordset")
		rsx1.CursorType = adOpenKeyset 
		rsx1.LockType   = 2 'adLockOptimistic 
		'
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_DataCrudaMensual"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2000"
		QrySql = QrySql & " GROUP BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " HAVING"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
		'
		'Response.write QrySql & "<br>"
		'Response.end
		'
		rsx1.Open QrySql, conexion		
		'
		if rsx1.eof then
			rsx1.close
			total2000 = 0
		else				
			rsArray = rsx1.GetRows() 
        	total2000 = UBound(rsArray, 2) + 1         	
			'total2000 = rsx1.recordcount			
			rsx1.close			
		end if
		'
		'Response.write " Total RecordCount 2000 = " & total2000 & "<br>"
		'Response.end
		'
		IF total2000 > 0 THEN
			'
			tRefresco2000_2000 = total2000 * 100 / totalRefrescoHogaresCon_2000
			'
			Response.Write "Total Compra de 2000 = " & total2000 & "<br>"
			Response.Write "Total Porcentaje 2000 = " & tRefresco2000_2000 & "<br>"
			'Response.End			
			'
			' Calcular Tamaño 250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total250 = 0
			else				
				'total250 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total250 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total250 = 0 then
				'
				tRefresco2000_250 = 0
				Response.Write "2000 no convive con 250 <br>"
				'	
			else
				'
				tRefresco2000_250 = total250 * 100 / totalRefrescoHogaresCon_2000				
				'
			end if
			'
			' Calcular Tamaño 320
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 320"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total320 = 0
			else				
				'total320= rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total320 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total320 = 0 then
				'
				tRefresco2000_320 = 0				
				Response.Write "2000 no convive con 320 <br>"
				'	
			else
				'
				tRefresco2000_320 = total320 * 100 / totalRefrescoHogaresCon_2000
				'
			end if
			'
			' Calcular Tamaño 355
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 355"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total355 = 0
			else				
				'total355 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total355 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total355 = 0 then
				'
				tRefresco2000_355 = 0
				Response.Write "2000 no convive con 355<br>"
				'	
			else
				'
				tRefresco2000_355 = total355 * 100 / totalRefrescoHogaresCon_2000
				'
			end if
			'
			' Calcular Tamaño 500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total500 = 0
			else				
				'total500 = rsx1.recordcount
				rsArray  = rsx1.GetRows() 
        		total500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total500 = 0 then
				'
				tRefresco2000_500 = 0
				Response.Write "2000 no convive con 500<br>"
				'	
			else
				'
				tRefresco2000_500 = total500 * 100 / totalRefrescoHogaresCon_2000
				'
			end if
			'
			' Calcular Tamaño 600
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 600"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total600 = 0
			else				
				'total600 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total600 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total600 = 0 then
				'
				tRefresco2000_600 = 0
				Response.Write "2000 no convive con 600<br>"
				'	
			else
				'
				tRefresco2000_600 = total600 * 100 / totalRefrescoHogaresCon_2000
				'
			end if
			'
			' Calcular Tamaño 1000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1000 = 0
			else				
				'total1000 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1000 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1000 = 0 then
				'
				tRefresco2000_1000 = 0
				Response.Write "2000 no convive con 1000<br>"
				'	
			else
				'
				tRefresco2000_1000 = total1000 * 100 / totalRefrescoHogaresCon_2000
				'
			end if			
			'
			' Calcular Tamaño 1250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1250 = 0
			else				
				'total1250 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1250 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1250 = 0 then
				'
				tRefresco2000_1250 = 0
				Response.Write "2000 no convive con 1250 / "
				'	
			else
				'
				tRefresco2000_1250 = total1250 * 100 / totalRefrescoHogaresCon_2000
				'
			end if
			'
			' Calcular Tamaño 1500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1500 = 0
			else				
				'total1500 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1500 = 0 then
				'
				tRefresco2000_1500 = 0
				Response.Write "2000 no convive con 1500<br>"
				'	
			else
				'
				tRefresco2000_1500 = total1500 * 100 / totalRefrescoHogaresCon_2000
				'
			end if
			'
			' Calcular Tamaño 2000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2000 = 0
			else				
				'total2000 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total2000 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total2000 = 0 then
				'
				tRefresco2000_2000 = 0
				Response.Write "2000 no convive con 2000 / "
				'	
			else
				'
				tRefresco2000_2000 = total2000 * 100 / totalRefrescoHogaresCon_2000
				'
			end if
			'
			' Calcular Tamaño 2500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2500 = 0
			else				
				'total2500 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total2500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total2500 = 0 then
				'
				tRefresco2000_2500 = 0
				Response.Write "2000 no convive con 2500 / "
				'	
			else
				'
				tRefresco2000_2500 = total2500 * 100 / totalRefrescoHogaresCon_2000
				'
			end if
			
		END IF
		'
		
	end if
	'	
	Set rsx1 = nothing
	
END SUB
'
SUB Calcular_Refrescos_2500
	'	
	' Buscar Todos Los Hogares compraron Tamaño 2500 ml
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
	QrySql = QrySql & " AND PH_DataCrudaMensual.TAMANO = 2500 "
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	QrySql = QrySql & " HAVING"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1 "
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	'	
	'Response.write QrySql & "<br>"
	'Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then		
		rsx1.close
		'
		Response.Write "<br>Tama&ntilde;o 2500 No! Convive con nadie <BR><BR>"
		tRefresco2500_250  = 0
		tRefresco2500_320  = tRefresco2500_350  = tRefresco2500_355  = 0
		tRefresco2500_500  = tRefresco2500_600  = tRefresco2500_1000 = 0
		tRefresco2500_1250 = tRefresco2500_2000 = tRefresco2500_2500 = 0	
		'			
	else
		'Erase dataArray
		dataArray = rsx1.GetRows
		rsx1.close
		'
		' Calculo total hogares con Refrescos 2500
		'		
		totalRefrescoHogaresCon_2500 = 0
		totalRefrescoHogaresCon_2500 = ubound(dataArray,2) + 1 
		total2500=0
		'
		Response.write "<br>Total hogares 2500 = " & totalRefrescoHogaresCon_2500 & "<br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "ID hogares con 2500 = " & replace(Hogares,",","-") & "<br>"
		'Response.end
		'				
		Set rsx1 = CreateObject("ADODB.Recordset")
		rsx1.CursorType = adOpenKeyset 
		rsx1.LockType   = 2 'adLockOptimistic 
		'
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_DataCrudaMensual"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
		QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2500"
		QrySql = QrySql & " GROUP BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
		QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
		QrySql = QrySql & " HAVING"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
		QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
		'
		'Response.write QrySql & "<br>"
		'Response.end
		'
		rsx1.Open QrySql, conexion		
		'
		if rsx1.eof then
			rsx1.close
			total2500 = 0
		else				
			rsArray = rsx1.GetRows() 
        	total2500 = UBound(rsArray, 2) + 1         	
			'total2500 = rsx1.recordcount			
			rsx1.close			
		end if
		'
		'Response.write " Total RecordCount 2500 = " & total2500 & "<br>"
		'Response.end
		'
		IF total2500 > 0 THEN
			'
			tRefresco2500_2500 = total2500 * 100 / totalRefrescoHogaresCon_2500
			'
			Response.Write "Total Compra de  2500 = " & total2500 & "<br>"
			Response.Write "Total Porcentaje 2500 = " & tRefresco2500_2500 & "<br>"
			'Response.End			
			'
			' Calcular Tamaño 250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total250 = 0
			else				
				'total250 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total250 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total250 = 0 then
				'
				tRefresco2500_250 = 0
				Response.Write "2500 no convive con 250 <br>"
				'	
			else
				'
				tRefresco2500_250 = total250 * 100 / totalRefrescoHogaresCon_2500
				'
			end if
			'
			' Calcular Tamaño 320
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 320"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total320 = 0
			else				
				'total320= rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total320 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total320 = 0 then
				'
				tRefresco2500_320 = 0				
				Response.Write "2500 no convive con 320 <br>"
				'	
			else
				'
				tRefresco2500_320 = total320 * 100 / totalRefrescoHogaresCon_2500
				'
			end if
			'
			' Calcular Tamaño 355
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 355"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total355 = 0
			else				
				'total355 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total355 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total355 = 0 then
				'
				tRefresco2500_355 = 0
				Response.Write "2500 no convive con 355<br>"
				'	
			else
				'
				tRefresco2500_355 = total355 * 100 / totalRefrescoHogaresCon_2500
				'
			end if
			'
			' Calcular Tamaño 500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total500 = 0
			else				
				'total500 = rsx1.recordcount
				rsArray  = rsx1.GetRows() 
        		total500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total500 = 0 then
				'
				tRefresco2500_500 = 0
				Response.Write "2500 no convive con 500<br>"
				'	
			else
				'
				tRefresco2500_500 = total500 * 100 / totalRefrescoHogaresCon_2500
				'
			end if
			'
			' Calcular Tamaño 600
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 600"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total600 = 0
			else				
				'total600 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total600 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total600 = 0 then
				'
				tRefresco2500_600 = 0
				Response.Write "2500 no convive con 600<br>"
				'	
			else
				'
				tRefresco2500_600 = total600 * 100 / totalRefrescoHogaresCon_2500
				'
			end if
			'
			' Calcular Tamaño 1000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1000 = 0
			else				
				'total1000 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1000 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1000 = 0 then
				'
				tRefresco2500_1000 = 0
				Response.Write "2500 no convive con 1000<br>"
				'	
			else
				'
				tRefresco2500_1000 = total1000 * 100 / totalRefrescoHogaresCon_2500
				'
			end if			
			'
			' Calcular Tamaño 1250
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1250"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1250 = 0
			else				
				'total1250 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1250 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1250 = 0 then
				'
				tRefresco2500_1250 = 0
				Response.Write "2500 no convive con 1250 / "
				'	
			else
				'
				tRefresco2500_1250 = total1250 * 100 / totalRefrescoHogaresCon_2500
				'
			end if
			'
			' Calcular Tamaño 1500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 1500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total1500 = 0
			else				
				'total1500 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total1500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1500 = 0 then
				'
				tRefresco2500_1500 = 0
				Response.Write "2500 no convive con 1500 / "
				'	
			else
				'
				tRefresco2500_1500 = total1500 * 100 / totalRefrescoHogaresCon_2500
				'
			end if
			'
			' Calcular Tamaño 2000
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2000"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( "  & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2000 = 0
			else				
				'total2000 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total2000 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total2000 = 0 then
				'
				tRefresco2500_2000 = 0
				Response.Write "2500 no convive con 2000 / "
				'	
			else
				'
				tRefresco2500_2000 = total2000 * 100 / totalRefrescoHogaresCon_2500
				'
			end if
			'
			' Calcular Tamaño 2500
			'
			Set rsx1 = CreateObject("ADODB.Recordset")
			rsx1.CursorType = adOpenKeyset 
			rsx1.LockType   = 2 'adLockOptimistic 
			'
			QrySql = vbnullstring
			QrySql = QrySql & " SELECT"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.Tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " FROM"
			QrySql = QrySql & " PH_DataCrudaMensual"
			QrySql = QrySql & " WHERE"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante <> 0"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (" & idMeses & ")"
			QrySql = QrySql & " AND PH_DataCrudaMensual.Tamano = 2500"
			QrySql = QrySql & " GROUP BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
			QrySql = QrySql & " PH_DataCrudaMensual.tamano,"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
			QrySql = QrySql & " HAVING"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria = 1" 
			QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Hogar IN ( " & Hogares & ")"
			QrySql = QrySql & " ORDER BY"
			QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
			'
			'Response.write QrySql & "<br>"
			'Response.end
			'
			rsx1.Open QrySql, conexion		
			'
			if rsx1.eof then
				rsx1.close
				total2500 = 0
			else				
				'total2500 = rsx1.recordcount
				rsArray = rsx1.GetRows() 
        		total2500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total2500 = 0 then
				'
				tRefresco2500_2500 = 0
				Response.Write "2500 no convive con 2500 / "
				'	
			else
				'
				tRefresco2500_2500 = total2500 * 100 / totalRefrescoHogaresCon_2500
				'
			end if
			
		END IF
		'
		
	end if
	'	
	Set rsx1 = nothing
	
END SUB
'
'**************
'*  GRAFICAR  *
'**************
'
SUB Graficar_Datos
	'
	' Graficar los resultados en tablas
	'
	Response.Write "<strong><table class='table table-borderless table-hover table-condensed' style=' margin: auto; width: 65% !important;'>"
		
		Response.Write "<thead>"
		
			Response.Write "<tr>"
				Response.Write "<th colspan='11' class='text-center text-primary'><i class='fas fa-check-double'></i>&nbsp;MATRIZ DE CONVIVENCIA TAMA&Ntilde;OS&nbsp;</th>"	  
			Response.Write "</tr>"
			
			Response.Write "<tr>"
				Response.Write "<td class='text-center'></td><td class='text-center'>250 ml</td><td class='text-center'>320 ml</td><td class='text-center'>350 ml</td><td class='text-center'>355 ml</td><td class='text-center'>500 ml</td>"
				Response.Write "<td class='text-center'>600 ml</td><td class='text-center'>1000 ml</td><td class='text-center'>1250 ml</td><td class='text-center'>1500 ml</td><td class='text-center'>2000 ml</td><td class='text-center'>2500 ml</td>"
			Response.Write "</tr>"				
			
	   Response.Write "</thead>"
	   
	   Response.Write "<tbody>"
	   
			'250
			Response.Write "<tr>"
				Response.Write "<td>250 ml</td>"
				Response.Write "<td class='text-center text-danger'>"  & FormatNumber(tRefresco250_250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco320_250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco350_250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco355_250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco500_250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco600_250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1000_250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1250_250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1500_250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2000_250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2500_250,2) & " %</td>"
			Response.Write "</tr>"
			'320
			Response.Write "<tr>"
				Response.Write "<td>320 ml</td>"
				Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefresco250_320,2) & " %</td><td class='text-center text-danger'>" & FormatNumber(tRefresco320_320,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco350_320,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco355_320,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco500_320,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco600_320,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1000_320,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1250_320,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1500_320,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2000_320,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2500_320,2) & " %</td>"
			Response.Write "</tr>"
			'350
			Response.Write "<tr>"
				Response.Write "<td>350 ml</td>"
				Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefresco250_350,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco320_350,2) & " %</td><td class='text-center text-danger'>" & FormatNumber(tRefresco350_350,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco355_350,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco500_350,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco600_350,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1000_350,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1250_350,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1500_350,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2000_350,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2500_350,2) & " %</td>"
			Response.Write "</tr>"
			'355
			Response.Write "<tr>"
				Response.Write "<td>355 ml</td>"
				Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefresco250_355,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco320_355,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco350_355,2) & " %</td><td class='text-center text-danger'>" & FormatNumber(tRefresco355_355,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco500_355,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco600_355,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1000_355,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1250_355,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1500_355,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2000_355,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2500_355,2) & " %</td>"
			Response.Write "</tr>"
			'500
			Response.Write "<tr>"
				Response.Write "<td>500 ml</td>"
				Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefresco250_500,2) & " %</td><td class='text-center text-primary'> " & FormatNumber(tRefresco320_500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco350_500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco355_500,2) & " %</td><td class='text-center text-danger'>" & FormatNumber(tRefresco500_500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco600_500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1000_500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1250_500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1500_500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2000_500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2500_500,2) & " %</td>"
			Response.Write "</tr>"
			'600
			Response.Write "<tr>"
				Response.Write "<td>600 ml</td>"
				Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefresco250_600,2) & " %</td><td class='text-center text-primary'> " & FormatNumber(tRefresco320_600,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco350_600,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco355_600,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco500_600,2) & " %</td><td class='text-center text-danger'>" & FormatNumber(tRefresco600_600,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1000_600,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1250_600,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1500_600,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2000_600,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2500_600,2) & " %</td>"
			Response.Write "</tr>"
			'1000
			Response.Write "<tr>"
				Response.Write "<td>1000 ml</td>"
				Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefresco250_1000,2) & " %</td><td class='text-center text-primary'> " & FormatNumber(tRefresco320_1000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco350_1000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco355_1000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco500_1000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco600_1000,2) & " %</td><td class='text-center text-danger'>" & FormatNumber(tRefresco1000_1000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1250_1000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1500_1000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2000_1000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2500_1000,2) & " %</td>"
			Response.Write "</tr>"
			'1250
			Response.Write "<tr>"
				Response.Write "<td>1250 ml</td>"
				Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefresco250_1250,2) & " %</td><td class='text-center text-primary'> " & FormatNumber(tRefresco320_1250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco350_1250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco355_1250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco500_1250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco600_1250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1000_1250,2) & " %</td><td class='text-center text-danger'>" & FormatNumber(tRefresco1250_1250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1500_1250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2000_1250,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2500_1250,2) & " %</td>"
			Response.Write "</tr>"
			'1500
			Response.Write "<tr>"
				Response.Write "<td>1500 ml</td>"
				Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefresco250_1500,2) & " %</td><td class='text-center text-primary'> " & FormatNumber(tRefresco320_1500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco350_1500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco355_1500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco500_1500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco600_1500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1000_1500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1250_1500,2) & " %</td><td class='text-center text-danger'>" & FormatNumber(tRefresco1500_1500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2000_1500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2500_1500,2) & " %</td>"
			Response.Write "</tr>"
			'2000
			Response.Write "<tr>"
				Response.Write "<td>2000 ml</td>"
				Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefresco250_2000,2) & " %</td><td class='text-center text-primary'> " & FormatNumber(tRefresco320_2000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco350_2000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco355_2000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco500_2000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco600_2000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1000_2000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1250_2000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1500_2000,2) & " %</td><td class='text-center text-danger'>" & FormatNumber(tRefresco2000_2000,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2500_2000,2) & " %</td>"
			Response.Write "</tr>"
			'2500
			Response.Write "<tr>"
				Response.Write "<td>2500 ml</td>"
				Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefresco250_2500,2) & " %</td><td class='text-center text-primary'> " & FormatNumber(tRefresco320_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco350_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco355_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco500_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco600_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1000_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1250_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1500_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2000_2500,2) & " %</td><td class='text-center text-danger'>" & FormatNumber(tRefresco2500_2500,2) & " %</td>"
			Response.Write "</tr>"					
			
	   Response.Write "</tbody>"
	   
	Response.Write "</table></strong>"
	'
END SUB
'	
%>



