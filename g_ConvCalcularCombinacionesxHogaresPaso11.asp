<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
'
' g_ConvCalcularCombinacionesxHogaresPaso1.asp - 12abr21 - 13abr21
'
Session.lcid = 1034
Response.CodePage = 65001	
Response.CharSet = "utf-8"
'	
Dim QrySql, idMeses, dataArray, rsx1
'
Dim totalRefrescoHogares250, tRefresco250_250, tRefresco250_320, tRefresco250_350, tRefresco250_355
Dim tRefresco250_500, tRefresco250_600, tRefresco250_1000, tRefresco250_1500, tRefresco250_2000, tRefresco250_2500
'
Dim totalRefrescoHogares320, tRefresco320_250, tRefresco320_320, tRefresco320_350, tRefresco320_355
Dim tRefresco320_500, tRefresco320_600, tRefresco320_1000, tRefresco320_1500, tRefresco320_2000, tRefresco320_2500
'
Dim totalRefrescoHogares350, tRefresco350_250, tRefresco350_320, tRefresco350_350, tRefresco350_355
Dim tRefresco350_500, tRefresco350_600, tRefresco350_1000, tRefresco350_1500, tRefresco350_2000, tRefresco350_2500
'
Dim totalRefrescoHogares355, tRefresco355_250, tRefresco355_320, tRefresco355_350, tRefresco355_355
Dim tRefresco355_500, tRefresco355_600, tRefresco355_1000, tRefresco355_1500, tRefresco355_2000, tRefresco355_2500
'
Dim totalRefrescoHogares500, tRefresco500_250, tRefresco500_320, tRefresco500_350, tRefresco500_355
Dim tRefresco500_500, tRefresco500_600, tRefresco500_1000, tRefresco500_1500, tRefresco500_2000, tRefresco500_2500
'
Dim totalRefrescoHogares600, tRefresco600_250, tRefresco600_320, tRefresco600_350, tRefresco600_355
Dim tRefresco600_500, tRefresco600_600, tRefresco600_1000, tRefresco600_1500, tRefresco600_2000, tRefresco600_2500
'
Dim totalRefrescoHogares1000, tRefresco1000_250, tRefresco1000_320, tRefresco1000_350, tRefresco1000_355
Dim tRefresco1000_500, tRefresco1000_600, tRefresco1000_1000, tRefresco1000_1500, tRefresco1000_2000, tRefresco1000_2500
'
Dim totalRefrescoHogares1250, tRefresco1250_250, tRefresco1250_320, tRefresco1250_350, tRefresco1250_355
Dim tRefresco1250_500, tRefresco1250_600, tRefresco1250_1000, tRefresco1250_1500, tRefresco1250_2000, tRefresco1250_2500
'
Dim totalRefrescoHogares1500, tRefresco1500_250, tRefresco1500_320, tRefresco1500_350, tRefresco1500_355
Dim tRefresco1500_500, tRefresco1500_600, tRefresco1500_1000, tRefresco1500_1500, tRefresco1500_2000, tRefresco1500_2500
'
Dim totalRefrescoHogares2000, tRefresco2000_250, tRefresco2000_320, tRefresco2000_350, tRefresco2000_355
Dim tRefresco2000_500, tRefresco2000_600, tRefresco2000_1000, tRefresco2000_1500, tRefresco2000_2000, tRefresco2000_2500
'
Dim totalRefrescoHogares2500, tRefresco2500_250, tRefresco2500_320, tRefresco2500_350, tRefresco2500_355
Dim tRefresco2500_500, tRefresco2500_600, tRefresco2500_1000, tRefresco2500_1500, tRefresco2500_2000, tRefresco2500_2500
'
idMeses ="16,17,18,19" ' Request.QueryString("id_Mes")
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
Response.Write "<br>Proceso tardo: " & Cstr(ElapsedTime) & " Segundos."
'
Graficar_Datos
'
SUB Calcular_Refrescos_250
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
		Response.Write "Tama&ntilde;o 250 No! Convive con nadie <BR><BR>"
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
		totalRefrescoHogares250 = 0
		totalRefrescoHogares250 = ubound(dataArray,2) + 1 
		total250=0
		'
		'FOR  i = 0 to ubound(dataArray,2) 
		'	Hogar  = dataArray(0,i)
		'	HOGARES = HOGARES + cstr(Hogar) ","
		'NEXT
		'
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
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
			tRefresco250_250 = total250 * 100 / totalRefrescoHogares250
			'
			Response.Write "Total Reg 250 " & total250 & "<br><br>"
			Response.Write "Total %%% 250 " & tRefresco250_250 & "<br><br>"
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
				Response.Write "250 no convive con 320"
				'	
			else
				'
				tRefresco250_320 = total320 * 100 / totalRefrescoHogares250
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
				Response.Write "250 no convive con 350"
				'	
			else
				'
				tRefresco250_350 = total350 * 100 / totalRefrescoHogares250
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
				Response.Write "250 no convive con 355"
				'	
			else
				'
				tRefresco250_355 = total355 * 100 / totalRefrescoHogares250
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
				Response.Write "250 no convive con 500"
				'	
			else
				'
				tRefresco250_500 = total500 * 100 / totalRefrescoHogares250
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
				Response.Write "250 no convive con 600"
				'	
			else
				'
				tRefresco250_600 = total600 * 100 / totalRefrescoHogares250
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
				Response.Write "250 no convive con 1000"
				'	
			else
				'
				tRefresco250_1000 = total1000 * 100 / totalRefrescoHogares250
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
				Response.Write "250 no convive con 1250"
				'	
			else
				'
				tRefresco250_1250 = total1250 * 100 / totalRefrescoHogares250
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
				'total1500 = rsx1.recordcount
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total1500 = 0 then
				'
				tRefresco250_1500 = 0
				Response.Write "250 no convive con 1500"
				'	
			else
				'
				tRefresco250_1500 = total1500 * 100 / totalRefrescoHogares250
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
				Response.Write "250 no convive con 2000"
				'	
			else
				'
				tRefresco250_2000 = total2000 * 100 / totalRefrescoHogares250
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
				total2500 = 0
			else
				rsArray = rsx1.GetRows() 
        		total2500 = UBound(rsArray, 2) + 1         	
				'total2500 = rsx1.recordcount
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total2500 = 0 then
				'
				tRefresco250_2500 = 0
				Response.Write "250 no convive con 2500"
				'	
			else
				'
				tRefresco250_2500 = total2500 * 100 / totalRefrescoHogares250
				'
			end if
						
		end if
		'	
		
		
		
		'*****
	end if
	'	
	Set rsx1 = nothing
	'
END SUB	
'
SUB Calcular_Refrescos_320
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
		'***
		'
		' Calculo total hogares con Refrescos 320
		'		
		totalRefrescoHogares320 = 0
		totalRefrescoHogares320 = ubound(dataArray,2) + 1 
		total320=0
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
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
			tRefresco320_320 = total320 * 100 / totalRefrescoHogares320
			Response.Write "Total Reg 250 " & total320 & "<br><br>"
			Response.Write "Total %%% 250 " & tRefresco320_320 & "<br><br>"
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
				tRefresco320_250 = 0
				Response.Write "320 no convive con 250"
				'	
			else
				'
				tRefresco320_250 = total250 * 100 / totalRefrescoHogares320
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
				Response.Write "250 no convive con 350"
				'	
			else
				'
				tRefresco320_350 = total350 * 100 / totalRefrescoHogares320
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
				tRefresco320_355 = 0
				Response.Write "320 no convive con 355"
				'	
			else
				'
				tRefresco320_355 = total355 * 100 / totalRefrescoHogares320
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
				rsArray = rsx1.GetRows() 
        		total500 = UBound(rsArray, 2) + 1         	
				rsx1.close
				'Response.write "<br>Total =: " & ubound(dataArray,2)+1		
			end if
			'
			if total500 = 0 then
				'
				tRefresco320_500 = 0
				Response.Write "320 no convive con 500"
				'	
			else
				'
				tRefresco320_500 = total500 * 100 / totalRefrescoHogares320
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
				tRefresco320_600 = 0
				Response.Write "250 no convive con 600"
				'	
			else
				'
				tRefresco320_600 = total600 * 100 / totalRefrescoHogares320
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
				tRefresco320_1000 = total1000 * 100 / totalRefrescoHogares320
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
				tRefresco320_1250 = total1250 * 100 / totalRefrescoHogares320
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
				tRefresco320_1500 = total1500 * 100 / totalRefrescoHogares320
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
				tRefresco320_2000 = total2000 * 100 / totalRefrescoHogares320
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
				tRefresco320_2500 = total2500 * 100 / totalRefrescoHogares320
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
	'Response.write QrySql & "<br><br>"
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
		totalRefrescoHogares350 = 0
		totalRefrescoHogares350 = ubound(dataArray,2) + 1 
		total350=0
		'
		Response.write "Total hogares 350 = " & totalRefrescoHogares350 & "<br><br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "hogares con 350 = " & Hogares & "<br><br>"
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
		Response.write " Total RecordCount 350 = " & total350 & "<br><br>"
		'Response.end
		'
		IF total350 > 0 THEN
			'
			tRefresco350_350 = total350 * 100 / totalRefrescoHogares350
			'
			Response.Write "Total Reg 350 " & total350 & "<br><br>"
			Response.Write "Total %%% 350 " & tRefresco350_350 & "<br><br>"
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
				Response.Write "350 no convive con 250 <br>"
				'	
			else
				'
				tRefresco350_250 = total250 * 100 / totalRefrescoHogares350
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
				tRefresco350_350 = 0				
				Response.Write "350 no convive con 350 <br>"
				'	
			else
				'
				tRefresco350_350 = total350 * 100 / totalRefrescoHogares350
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
				Response.Write "350 no convive con 355<br>"
				'	
			else
				'
				tRefresco350_355 = total355 * 100 / totalRefrescoHogares350
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
				Response.Write "350 no convive con 500<br>"
				'	
			else
				'
				tRefresco350_500 = total500 * 100 / totalRefrescoHogares350
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
				tRefresco350_600 = total600 * 100 / totalRefrescoHogares350
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
				Response.Write "350 no convive con 1000<br>"
				'	
			else
				'
				tRefresco350_1000 = total1000 * 100 / totalRefrescoHogares350
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
				Response.Write "350 no convive con 1250<br>"
				'	
			else
				'
				tRefresco350_1250 = total1250 * 100 / totalRefrescoHogares350
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
				Response.Write "350 no convive con 1500<br>"
				'	
			else
				'
				tRefresco350_1500 = total1500 * 100 / totalRefrescoHogares350
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
				Response.Write "350 no convive con 2000<br>"
				'	
			else
				'
				tRefresco350_2000 = total2000 * 100 / totalRefrescoHogares350
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
				tRefresco350_2500 = 0
				Response.Write "350 no convive con 2500<br>"
				'	
			else
				'
				tRefresco350_2500 = total2500 * 100 / totalRefrescoHogares350
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
	'Response.write QrySql & "<br><br>"
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
		totalRefrescoHogares355 = 0
		totalRefrescoHogares355 = ubound(dataArray,2) + 1 
		total355=0
		'
		Response.write "Total hogares 355 = " & totalRefrescoHogares355 & "<br><br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "hogares con 355 = " & Hogares & "<br><br>"
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
		Response.write " Total RecordCount 355 = " & total355 & "<br><br>"
		'Response.end
		'
		IF total355 > 0 THEN
			'
			tRefresco355_355 = total355 * 100 / totalRefrescoHogares355
			'
			Response.Write "Total Reg 355 = " & total355 & "<br><br>"
			Response.Write "Total %%% 355 = " & tRefresco355_355 & "<br><br>"
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
				tRefresco355_250 = 0
				Response.Write "355 no convive con 250 <br>"
				'	
			else
				'
				tRefresco355_250 = total250 * 100 / totalRefrescoHogares355
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
				tRefresco355_320 = 0				
				Response.Write "355 no convive con 320 <br>"
				'	
			else
				'
				tRefresco355_320 = total320 * 100 / totalRefrescoHogares355
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
				tRefresco355_355 = 0
				Response.Write "355 no convive con 355<br>"
				'	
			else
				'
				tRefresco355_355 = total355 * 100 / totalRefrescoHogares355
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
				tRefresco355_500 = 0
				Response.Write "355 no convive con 500<br>"
				'	
			else
				'
				tRefresco355_500 = total500 * 100 / totalRefrescoHogares355
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
				tRefresco355_600 = 0
				Response.Write "355 no convive con 600<br>"
				'	
			else
				'
				tRefresco355_600 = total600 * 100 / totalRefrescoHogares355
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
				tRefresco355_1000 = 0
				Response.Write "355 no convive con 1000<br>"
				'	
			else
				'
				tRefresco355_1000 = total1000 * 100 / totalRefrescoHogares355
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
				tRefresco355_1250 = 0
				Response.Write "355 no convive con 1250<br>"
				'	
			else
				'
				tRefresco355_1250 = total1250 * 100 / totalRefrescoHogares355
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
				tRefresco355_1500 = 0
				Response.Write "355 no convive con 1500<br>"
				'	
			else
				'
				tRefresco355_1500 = total1500 * 100 / totalRefrescoHogares355
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
				tRefresco355_2000 = 0
				Response.Write "355 no convive con 2000<br>"
				'	
			else
				'
				tRefresco355_2000 = total2000 * 100 / totalRefrescoHogares355
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
				tRefresco355_2500 = 0
				Response.Write "355 no convive con 2500<br>"
				'	
			else
				'
				tRefresco355_2500 = total2500 * 100 / totalRefrescoHogares355
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
	'Response.write QrySql & "<br><br>"
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
		totalRefrescoHogares500 = 0
		totalRefrescoHogares500 = ubound(dataArray,2) + 1 
		total500=0
		'
		Response.write "<br>Total hogares 500 = " & totalRefrescoHogares500 & "<br><br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "hogares con 500 = " & Hogares & "<br><br>"
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
		Response.write QrySql & "<br>"
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
		end if
		'
		Response.write " Total RecordCount 500 = " & total500 & "<br><br>"
		'Response.end
		'
		IF total500 > 0 THEN
			'
			tRefresco500_500 = total500 * 100 / totalRefrescoHogares500
			'
			Response.Write "Total Reg 500 = " & total500 & "<br><br>"
			Response.Write "Total %%% 500 = " & tRefresco500_500 & "<br><br>"
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
				tRefresco500_250 = 0
				Response.Write "500 no convive con 250 <br>"
				'	
			else
				'
				tRefresco500_250 = total250 * 100 / totalRefrescoHogares500
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
				tRefresco500_320 = total320 * 100 / totalRefrescoHogares500
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
				Response.Write "355 no convive con 355<br>"
				'	
			else
				'
				tRefresco500_355 = total355 * 100 / totalRefrescoHogares500
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
				tRefresco500_500 = 0
				Response.Write "355 no convive con 500<br>"
				'	
			else
				'
				tRefresco500_500 = total500 * 100 / totalRefrescoHogares500
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
				tRefresco500_600 = 0
				Response.Write "500 no convive con 600<br>"
				'	
			else
				'
				tRefresco500_600 = total600 * 100 / totalRefrescoHogares500
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
				tRefresco500_1000 = 0
				Response.Write "500 no convive con 1000<br>"
				'	
			else
				'
				tRefresco500_1000 = total1000 * 100 / totalRefrescoHogares500
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
				tRefresco500_1250 = 0
				Response.Write "500 no convive con 1250<br>"
				'	
			else
				'
				tRefresco500_1250 = total1250 * 100 / totalRefrescoHogares500
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
				tRefresco500_1500 = 0
				Response.Write "500 no convive con 1500<br>"
				'	
			else
				'
				tRefresco500_1500 = total1500 * 100 / totalRefrescoHogares500
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
				tRefresco500_2000 = 0
				Response.Write "500 no convive con 2000<br>"
				'	
			else
				'
				tRefresco500_2000 = total2000 * 100 / totalRefrescoHogares500
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
				tRefresco500_2500 = 0
				Response.Write "500 no convive con 2500<br>"
				'	
			else
				'
				tRefresco500_2500 = total2500 * 100 / totalRefrescoHogares500
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
	'Response.write QrySql & "<br><br>"
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
		totalRefrescoHogares600 = 0
		totalRefrescoHogares600 = ubound(dataArray,2) + 1 
		total600=0
		'
		Response.write "<br>Total hogares 600 = " & totalRefrescoHogares600 & "<br><br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "hogares con 600 = " & Hogares & "<br><br>"
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
		Response.write QrySql & "<br>"
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
		Response.write " Total RecordCount 600 = " & total600 & "<br><br>"
		'Response.end
		'
		IF total600 > 0 THEN
			'
			tRefresco600_600 = total600 * 100 / totalRefrescoHogares600
			'
			Response.Write "Total Reg 600 = " & total600 & "<br><br>"
			Response.Write "Total %%% 600 = " & tRefresco600_600 & "<br><br>"
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
				Response.Write "600 no convive con 250 <br>"
				'	
			else
				'
				tRefresco600_250 = total250 * 100 / totalRefrescoHogares600
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
				Response.Write "600 no convive con 320 <br>"
				'	
			else
				'
				tRefresco600_320 = total320 * 100 / totalRefrescoHogares600
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
				Response.Write "600 no convive con 355<br>"
				'	
			else
				'
				tRefresco600_355 = total355 * 100 / totalRefrescoHogares600
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
				tRefresco600_500 = 0
				Response.Write "600 no convive con 500<br>"
				'	
			else
				'
				tRefresco600_500 = total500 * 100 / totalRefrescoHogares600
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
				tRefresco600_600 = 0
				Response.Write "600 no convive con 600<br>"
				'	
			else
				'
				tRefresco600_600 = total600 * 100 / totalRefrescoHogares600
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
				tRefresco600_1000 = 0
				Response.Write "600 no convive con 1000<br>"
				'	
			else
				'
				tRefresco600_1000 = total1000 * 100 / totalRefrescoHogares600
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
				Response.Write "600 no convive con 1250<br>"
				'	
			else
				'
				tRefresco600_1250 = total1250 * 100 / totalRefrescoHogares600
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
				tRefresco600_1500 = 0
				Response.Write "600 no convive con 1500<br>"
				'	
			else
				'
				tRefresco600_1500 = total1500 * 100 / totalRefrescoHogares600
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
				tRefresco600_2000 = 0
				Response.Write "600 no convive con 2000<br>"
				'	
			else
				'
				tRefresco600_2000 = total2000 * 100 / totalRefrescoHogares600
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
				tRefresco600_2500 = 0
				Response.Write "600 no convive con 2500<br>"
				'	
			else
				'
				tRefresco600_2500 = total2500 * 100 / totalRefrescoHogares600
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
	'Response.write QrySql & "<br><br>"
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
		totalRefrescoHogares1000 = 0
		totalRefrescoHogares1000 = ubound(dataArray,2) + 1 
		total1000=0
		'
		Response.write "<br>Total hogares 1000 = " & totalRefrescoHogares1000 & "<br><br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "hogares con 1000 = " & Hogares & "<br><br>"
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
		Response.write " Total RecordCount 1000 = " & total1000 & "<br><br>"
		'Response.end
		'
		IF total1000 > 0 THEN
			'
			tRefresco1000_1000 = total1000 * 100 / totalRefrescoHogares1000
			'
			Response.Write "Total Reg 1000 = " & total1000 & "<br><br>"
			Response.Write "Total %%% 1000 = " & tRefresco1000_1000 & "<br><br>"
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
				tRefresco1000_250 = 0
				Response.Write "1000 no convive con 250 <br>"
				'	
			else
				'
				tRefresco1000_250 = total250 * 100 / totalRefrescoHogares1000
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
				Response.Write "1000 no convive con 320 <br>"
				'	
			else
				'
				tRefresco1000_320 = total320 * 100 / totalRefrescoHogares1000
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
				Response.Write "1000 no convive con 355<br>"
				'	
			else
				'
				tRefresco1000_355 = total355 * 100 / totalRefrescoHogares1000
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
				tRefresco1000_500 = 0
				Response.Write "1000 no convive con 500<br>"
				'	
			else
				'
				tRefresco1000_500 = total500 * 100 / totalRefrescoHogares1000
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
				tRefresco1000_600 = 0
				Response.Write "1000 no convive con 600<br>"
				'	
			else
				'
				tRefresco1000_600 = total600 * 100 / totalRefrescoHogares1000
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
				tRefresco600_1000 = 0
				Response.Write "1000 no convive con 1000<br>"
				'	
			else
				'
				tRefresco600_1000 = total1000 * 100 / totalRefrescoHogares1000
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
				tRefresco1000_1250 = 0
				Response.Write "1000 no convive con 1250<br>"
				'	
			else
				'
				tRefresco1000_1250 = total1250 * 100 / totalRefrescoHogares1000
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
				Response.Write "1000 no convive con 1500<br>"
				'	
			else
				'
				tRefresco1000_1500 = total1500 * 100 / totalRefrescoHogares1000
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
				tRefresco1000_2000 = 0
				Response.Write "1000 no convive con 2000<br>"
				'	
			else
				'
				tRefresco1000_2000 = total2000 * 100 / totalRefrescoHogares1000
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
				tRefresco1000_2500 = 0
				Response.Write "1000 no convive con 2500<br>"
				'	
			else
				'
				tRefresco1000_2500 = total2500 * 100 / totalRefrescoHogares1000
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
	'Response.write QrySql & "<br><br>"
	'Response.end
	'
    rsx1.Open QrySql, conexion		
	'
	if rsx1.eof then		
		rsx1.close
		'
		Response.Write "<br>Tama&ntilde;o 1250 No! Convive con nadie <BR><BR>"
		tRefresco1250_250  = 0
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
		totalRefrescoHogares1250 = 0
		totalRefrescoHogares1250 = ubound(dataArray,2) + 1 
		total1250=0
		'
		Response.write "<br>Total hogares 1250 = " & totalRefrescoHogares1250 & "<br><br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "hogares con 1250 = " & Hogares & "<br><br>"
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
		Response.write " Total RecordCount 1250 = " & total1250 & "<br><br>"
		'Response.end
		'
		IF total1250 > 0 THEN
			'
			tRefresco1250_1250 = total1250 * 100 / totalRefrescoHogares1250
			'
			Response.Write "Total Reg 1250 = " & total1250 & "<br><br>"
			Response.Write "Total %%% 1250 = " & tRefresco1250_1250 & "<br><br>"
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
				tRefresco1250_250 = 0
				Response.Write "1250 no convive con 250 <br>"
				'	
			else
				'
				tRefresco1250_250 = total250 * 100 / totalRefrescoHogares1250
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
				tRefresco1250_320 = 0				
				Response.Write "1250 no convive con 320 <br>"
				'	
			else
				'
				tRefresco1250_320 = total320 * 100 / totalRefrescoHogares1250
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
				tRefresco1250_355 = 0
				Response.Write "1250 no convive con 355<br>"
				'	
			else
				'
				tRefresco1250_355 = total355 * 100 / totalRefrescoHogares1250
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
				tRefresco1250_500 = 0
				Response.Write "1250 no convive con 500<br>"
				'	
			else
				'
				tRefresco1250_500 = total500 * 100 / totalRefrescoHogares1250
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
				tRefresco1250_600 = 0
				Response.Write "1250 no convive con 600<br>"
				'	
			else
				'
				tRefresco1250_600 = total600 * 100 / totalRefrescoHogares1250
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
				Response.Write "1250 no convive con 1000<br>"
				'	
			else
				'
				tRefresco1250_1000 = total1000 * 100 / totalRefrescoHogares1250
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
				tRefresco1250_1250 = 0
				Response.Write "1250 no convive con 1250<br>"
				'	
			else
				'
				tRefresco1250_1250 = total1250 * 100 / totalRefrescoHogares1250
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
				tRefresco1250_1500 = 0
				Response.Write "1250 no convive con 1500<br>"
				'	
			else
				'
				tRefresco1250_1500 = total1500 * 100 / totalRefrescoHogares1250
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
				Response.Write "1250 no convive con 2000<br>"
				'	
			else
				'
				tRefresco1250_2000 = total2000 * 100 / totalRefrescoHogares1250
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
				tRefresco1250_2500 = 0
				Response.Write "1250 no convive con 2500<br>"
				'	
			else
				'
				tRefresco1250_2500 = total2500 * 100 / totalRefrescoHogares1250
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
	'Response.write QrySql & "<br><br>"
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
		totalRefrescoHogares1500 = 0
		totalRefrescoHogares1500 = ubound(dataArray,2) + 1 
		total1500=0
		'
		Response.write "<br>Total hogares 1500 = " & totalRefrescoHogares1500 & "<br><br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "hogares con 1500 = " & Hogares & "<br><br>"
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
			'total1500 = rsx1.recordcount			
			rsx1.close			
		end if
		'
		Response.write " Total RecordCount 1500 = " & total1500 & "<br><br>"
		'Response.end
		'
		IF total1500 > 0 THEN
			'
			tRefresco1500_1500 = total1500 * 100 / totalRefrescoHogares1500
			'
			Response.Write "Total Reg 1500 = " & total1500 & "<br><br>"
			Response.Write "Total %%% 1500 = " & tRefresco1500_1500 & "<br><br>"
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
				tRefresco1500_250 = 0
				Response.Write "1500 no convive con 250 <br>"
				'	
			else
				'
				tRefresco1500_250 = total250 * 100 / totalRefrescoHogares1500
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
				tRefresco1500_320 = 0				
				Response.Write "1500 no convive con 320 <br>"
				'	
			else
				'
				tRefresco1500_320 = total320 * 100 / totalRefrescoHogares1500
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
				tRefresco1500_355 = 0
				Response.Write "1500 no convive con 355<br>"
				'	
			else
				'
				tRefresco1500_355 = total355 * 100 / totalRefrescoHogares1500
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
				tRefresco1500_500 = 0
				Response.Write "1500 no convive con 500<br>"
				'	
			else
				'
				tRefresco1500_500 = total500 * 100 / totalRefrescoHogares1500
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
				tRefresco1500_600 = 0
				Response.Write "1500 no convive con 600<br>"
				'	
			else
				'
				tRefresco1500_600 = total600 * 100 / totalRefrescoHogares1500
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
				tRefresco1500_1000 = 0
				Response.Write "1250 no convive con 1000<br>"
				'	
			else
				'
				tRefresco1500_1000 = total1000 * 100 / totalRefrescoHogares1500
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
				tRefresco1500_1250 = 0
				Response.Write "1500 no convive con 1250<br>"
				'	
			else
				'
				tRefresco1500_1250 = total1250 * 100 / totalRefrescoHogares1500
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
				tRefresco1500_1500 = 0
				Response.Write "1500 no convive con 1500<br>"
				'	
			else
				'
				tRefresco1500_1500 = total1500 * 100 / totalRefrescoHogares1500
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
				tRefresco1500_2000 = 0
				Response.Write "1500 no convive con 2000<br>"
				'	
			else
				'
				tRefresco1500_2000 = total2000 * 100 / totalRefrescoHogares1500
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
				tRefresco1500_2500 = 0
				Response.Write "1500 no convive con 2500<br>"
				'	
			else
				'
				tRefresco1500_2500 = total2500 * 100 / totalRefrescoHogares1500
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
	'Response.write QrySql & "<br><br>"
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
		totalRefrescoHogares2000 = 0
		totalRefrescoHogares2000 = ubound(dataArray,2) + 1 
		total2000=0
		'
		Response.write "<br>Total hogares 2000 = " & totalRefrescoHogares2000 & "<br><br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "hogares con 2000 = " & Hogares & "<br><br>"
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
		Response.write " Total RecordCount 2000 = " & total2000 & "<br><br>"
		'Response.end
		'
		IF total2000 > 0 THEN
			'
			tRefresco2000_2000 = total2000 * 100 / totalRefrescoHogares2000
			'
			Response.Write "Total Reg 2000 = " & total2000 & "<br><br>"
			Response.Write "Total %%% 2000 = " & tRefresco2000_2000 & "<br><br>"
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
				tRefresco2000_250 = total250 * 100 / totalRefrescoHogares2000
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
				tRefresco2000_320 = total320 * 100 / totalRefrescoHogares2000
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
				tRefresco2000_355 = total355 * 100 / totalRefrescoHogares2000
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
				tRefresco2000_500 = total500 * 100 / totalRefrescoHogares2000
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
				tRefresco2000_600 = total600 * 100 / totalRefrescoHogares2000
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
				tRefresco2000_1000 = total1000 * 100 / totalRefrescoHogares2000
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
				Response.Write "2000 no convive con 1250<br>"
				'	
			else
				'
				tRefresco2000_1250 = total1250 * 100 / totalRefrescoHogares2000
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
				tRefresco2000_1500 = total1500 * 100 / totalRefrescoHogares2000
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
				Response.Write "2000 no convive con 2000<br>"
				'	
			else
				'
				tRefresco2000_2000 = total2000 * 100 / totalRefrescoHogares2000
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
				Response.Write "2000 no convive con 2500<br>"
				'	
			else
				'
				tRefresco2000_2500 = total2500 * 100 / totalRefrescoHogares2000
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
	'Response.write QrySql & "<br><br>"
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
		totalRefrescoHogares2500 = 0
		totalRefrescoHogares2500 = ubound(dataArray,2) + 1 
		total2500=0
		'
		Response.write "<br>Total hogares 2500 = " & totalRefrescoHogares2500 & "<br><br>"
		'Response.end
		'		
		Dim Hogares
		Hogares = vbnullstring
		for iReg = 0 to ubound(dataArray,2)
			Hogares = Hogares + cstr(dataArray(0,iReg)) & ","			
		next
		Hogares = Left(Hogares, Len(Hogares) - 1)
		'
		Response.write "hogares con 2500 = " & Hogares & "<br><br>"
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
		Response.write " Total RecordCount 2500 = " & total2500 & "<br><br>"
		'Response.end
		'
		IF total2500 > 0 THEN
			'
			tRefresco2500_2500 = total2500 * 100 / totalRefrescoHogares2500
			'
			Response.Write "Total Reg 2500 = " & total2500 & "<br><br>"
			Response.Write "Total %%% 2500 = " & tRefresco2500_2500 & "<br><br>"
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
				tRefresco2500_250 = total250 * 100 / totalRefrescoHogares2500
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
				tRefresco2500_320 = total320 * 100 / totalRefrescoHogares2500
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
				tRefresco2500_355 = total355 * 100 / totalRefrescoHogares2500
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
				tRefresco2500_500 = total500 * 100 / totalRefrescoHogares2500
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
				tRefresco2500_600 = total600 * 100 / totalRefrescoHogares2500
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
				tRefresco2500_1000 = total1000 * 100 / totalRefrescoHogares2500
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
				Response.Write "2500 no convive con 1250<br>"
				'	
			else
				'
				tRefresco2500_1250 = total1250 * 100 / totalRefrescoHogares2500
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
				Response.Write "2500 no convive con 1500<br>"
				'	
			else
				'
				tRefresco2500_1500 = total1500 * 100 / totalRefrescoHogares2500
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
				Response.Write "2500 no convive con 2000<br>"
				'	
			else
				'
				tRefresco2500_2000 = total2000 * 100 / totalRefrescoHogares2500
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
				Response.Write "2500 no convive con 2500<br>"
				'	
			else
				'
				tRefresco2500_2500 = total2500 * 100 / totalRefrescoHogares2500
				'
			end if
			
		END IF
		'
		
	end if
	'	
	Set rsx1 = nothing
	
END SUB
'
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
				Response.Write "<td class='text-center'>600 ml</td><td class='text-center'>1000 ml</td><td class='text-center'>1250 ml</td><td class='text-center'>1500 ml</td><td class='text-center'>2000 ml</td><td class='text-center'>3000 ml</td>"
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
			'3000
			Response.Write "<tr>"
				Response.Write "<td>3000 ml</td>"
				Response.Write "<td class='text-center text-primary'>" & FormatNumber(tRefresco250_2500,2) & " %</td><td class='text-center text-primary'> " & FormatNumber(tRefresco320_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco350_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco355_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco500_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco600_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1000_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1250_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco1500_2500,2) & " %</td><td class='text-center text-primary'>" & FormatNumber(tRefresco2000_2500,2) & " %</td><td class='text-center text-danger'>" & FormatNumber(tRefresco2500_2500,2) & " %</td>"
			Response.Write "</tr>"					
			
	   Response.Write "</tbody>"
	   
	Response.Write "</table></strong>"
	'
END SUB
'	
%>



